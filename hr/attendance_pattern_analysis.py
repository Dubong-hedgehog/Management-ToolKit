"""
attendance_pattern_analysis.py
개인별/부서별 근태 "이상 패턴" 탐지 도구.

attendance_dashboard.py가 "이번 기간 vs 직전기/전년동기" 숫자 비교라면,
이 스크립트는 "누가, 어느 부서가 비정상적인 패턴을 보이는가"를 찾아낸다.

1) 지각 패턴
   - 개인 요일 집중형: 지각이 특정 요일에 비정상적으로 몰리는 사람
       (예: "이 사람은 지각의 대부분이 월요일에 몰려 있다")
   - 개인 추세 악화형: 최근 3개월 지각률이 그 이전 기간 대비 뚜렷하게
       나빠진 사람
   - 부서 이상치: 특정 월에 그 부서 지각률이 전사 평균 대비 비정상적으로
       높은 경우

2) 야근 위험 단계
   고용노동부 산재보험법 시행령 별표3(근로복지공단 과로사·뇌심혈관질환
   업무상 질병 인정기준)을 그대로 위험 구간으로 사용한다.
       - 발병 전 4주 평균 주 64시간 초과  -> 업무 관련성 "강함" (매우위험)
       - 발병 전 12주 평균 주 60시간 초과 -> 업무 관련성 "강함" (위험)
       - 발병 전 12주 평균 주 52시간 초과 -> 법정 연장근로 한도 초과 (주의)
       - 그 외                            -> 정상
   개인별 최근 4주/12주 평균 총근로시간(정규 8h/일 + 초과근무시간)을 계산해
   단계를 매기고, 부서 단위로는 위험군 인원 비율과, 부서 평균 초과근무시간이
   전사 평균 대비 특정 월에 비정상적으로 튀는지를 함께 본다.

   ※ 이 기준은 "발병한 질병과 업무의 관련성"을 평가하는 산재 인정기준이며,
   그 자체로 "위험이 몇 배"라는 통계치를 제공하는 것은 아니다. 참고로
   WHO·ILO 공동연구(2021)는 주 55시간 이상 근무 시 35~40시간 근무 대비
   허혈성 심장질환 사망위험 17%, 뇌졸중 사망위험 35% 증가로 추정한다.
   (출처: 근로복지공단 업무상질병 인정기준, WHO/ILO "Long working hours
   increasing deaths from heart disease and stroke", 2021)

Usage:
    python hr/attendance_pattern_analysis.py
    python hr/attendance_pattern_analysis.py --as-of 2026-06-30
    python hr/attendance_pattern_analysis.py --input hr/sample_data/sample_attendance.csv

Input:
    hr/sample_data/sample_attendance.csv (또는 --input / .env로 설정한 구글시트)
    컬럼: 사번, 이름, 부서, 거래일자, 출근시각, 퇴근시각, 지각여부(Y/N), 초과근무시간

Output:
    hr/output/attendance_pattern_report.xlsx
        - 개인_지각이상패턴 / 부서_지각이상치 / 개인_야근위험단계 / 부서_야근이상치
    hr/output/late_weekday_heatmap.png   (직원 x 요일 지각률 히트맵)
    hr/output/overtime_risk_trend.png    (위험군 직원 최근 12주 근로시간 추이)
    hr/output/dept_monthly_trend.png     (부서별 월별 지각률/초과근무 추이)
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.chart_style import PALETTE, save_chart, setup_style  # noqa: E402
from common.excel_io import load_csv, save_excel_report  # noqa: E402
from common.format_utils import print_section  # noqa: E402
from common.sheet_io import is_configured as sheet_is_configured  # noqa: E402
from common.sheet_io import load_transactions_from_sheet  # noqa: E402

HERE = Path(__file__).resolve().parent

# ---- 지각 패턴 탐지 임계값 ----
WEEKDAY_MIN_LATE_COUNT = 8       # 이 이상 지각해야 요일집중형 후보로 본다
WEEKDAY_CONCENTRATION = 0.40     # 특정 요일이 전체 지각의 40% 이상을 차지하면 집중형
WEEKDAY_MIN_ON_DAY = 5           # 해당 요일 지각 횟수 최소치(우연 배제)

TREND_MIN_PRIOR_DAYS = 40        # 이전기간 최소 근무일수(표본 부족 방지)
TREND_MIN_RECENT_DAYS = 15       # 최근 3개월 최소 근무일수
TREND_RATIO = 2.0                # 최근 지각률이 이전 대비 몇 배 이상이면 악화
TREND_ABS_DIFF = 0.10            # 그리고 최소 10%p 이상 벌어져야 함

DEPT_LATE_RATIO = 1.5
DEPT_LATE_ABS_DIFF = 0.08        # 8%p 이상

DEPT_OT_RATIO = 1.8
DEPT_OT_ABS_DIFF = 1.0           # 1.0h 이상

# ---- 야근 위험 단계 (근로복지공단 과로사 인정기준) ----
WEEKLY_STANDARD_HOURS = 40.0
DAILY_REGULAR_HOURS = 8.0
TIER_VERY_HIGH_4W = 64.0   # 4주 평균 64h 초과 -> 매우위험
TIER_HIGH_12W = 60.0       # 12주 평균 60h 초과 -> 위험
TIER_CAUTION_12W = 52.0    # 12주 평균 52h 초과 -> 주의 (법정 연장한도 초과)

# ---- 지각 시간(분) 기준 출근시각 ----
STANDARD_START_MIN = 9 * 60  # 09:00 기준, 이후 도착분(분)을 지각시간으로 본다


# ------------------------------------------------------------------
# 데이터 로드
# ------------------------------------------------------------------
def load_attendance(input_path: str | Path) -> pd.DataFrame:
    if sheet_is_configured():
        df = load_transactions_from_sheet()
    else:
        df = load_csv(input_path)
    df["거래일자"] = pd.to_datetime(df["거래일자"])
    df["초과근무시간"] = pd.to_numeric(df["초과근무시간"], errors="coerce").fillna(0)
    df["지각"] = (df["지각여부"] == "Y").astype(int)
    df["요일"] = df["거래일자"].dt.day_name()

    def _to_minutes(t: str) -> int:
        h, m = str(t).split(":")
        return int(h) * 60 + int(m)

    start_min = df["출근시각"].map(_to_minutes)
    # 지각이 아닌 날은 정의상 지각시간 0으로 둔다(정시 도착이 09:00을 살짝
    # 넘는 랜덤 노이즈로 잡히는 걸 방지 — 지각여부 컬럼을 기준으로 판단).
    df["지각분"] = np.where(df["지각"] == 1, (start_min - STANDARD_START_MIN).clip(lower=0), 0)
    return df


# ------------------------------------------------------------------
# 1) 개인 지각 - 요일 집중형
# ------------------------------------------------------------------
WEEKDAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
WEEKDAY_KO = {"Monday": "월", "Tuesday": "화", "Wednesday": "수", "Thursday": "목", "Friday": "금"}


def detect_weekday_concentration(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for (emp_id, name, dept), g in df.groupby(["사번", "이름", "부서"]):
        late = g[g["지각"] == 1]
        total_late = len(late)
        if total_late < WEEKDAY_MIN_LATE_COUNT:
            continue
        by_dow = late["요일"].value_counts()
        top_dow = by_dow.idxmax()
        top_count = int(by_dow.max())
        concentration = top_count / total_late
        if concentration >= WEEKDAY_CONCENTRATION and top_count >= WEEKDAY_MIN_ON_DAY:
            rows.append({
                "사번": emp_id, "이름": name, "부서": dept,
                "패턴": "요일 집중형",
                "집중 요일": WEEKDAY_KO[top_dow],
                "총 지각 횟수": total_late,
                "해당 요일 지각 횟수": top_count,
                "집중도": round(concentration, 3),
                "설명": f"전체 지각 {total_late}건 중 {top_count}건({concentration:.0%})이 {WEEKDAY_KO[top_dow]}요일에 발생",
            })
    return pd.DataFrame(rows).sort_values("집중도", ascending=False) if rows else pd.DataFrame(
        columns=["사번", "이름", "부서", "패턴", "집중 요일", "총 지각 횟수", "해당 요일 지각 횟수", "집중도", "설명"]
    )


# ------------------------------------------------------------------
# 2) 개인 지각 - 최근 추세 악화형
# ------------------------------------------------------------------
def detect_trend_worsening(df: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    recent_start = (as_of - pd.DateOffset(months=3)) + pd.Timedelta(days=1)
    rows = []
    for (emp_id, name, dept), g in df.groupby(["사번", "이름", "부서"]):
        recent = g[(g["거래일자"] >= recent_start) & (g["거래일자"] <= as_of)]
        prior = g[g["거래일자"] < recent_start]
        if len(recent) < TREND_MIN_RECENT_DAYS or len(prior) < TREND_MIN_PRIOR_DAYS:
            continue
        recent_rate = recent["지각"].mean()
        prior_rate = prior["지각"].mean()
        diff = recent_rate - prior_rate
        ratio = recent_rate / prior_rate if prior_rate > 0 else np.inf
        if diff >= TREND_ABS_DIFF and ratio >= TREND_RATIO:
            rows.append({
                "사번": emp_id, "이름": name, "부서": dept,
                "패턴": "최근 추세 악화형",
                "이전 지각률": round(prior_rate, 3),
                "최근 3개월 지각률": round(recent_rate, 3),
                "증감(%p)": round(diff * 100, 1),
                "배수": round(ratio, 2) if np.isfinite(ratio) else None,
                "설명": f"이전 {prior_rate:.1%} -> 최근 3개월 {recent_rate:.1%}로 지각률 급등",
            })
    return pd.DataFrame(rows).sort_values("증감(%p)", ascending=False) if rows else pd.DataFrame(
        columns=["사번", "이름", "부서", "패턴", "이전 지각률", "최근 3개월 지각률", "증감(%p)", "배수", "설명"]
    )


# ------------------------------------------------------------------
# 3) 부서 지각 이상치 (월별, 전사 평균 대비)
# ------------------------------------------------------------------
def detect_dept_late_anomaly(df: pd.DataFrame) -> pd.DataFrame:
    m = df.copy()
    m["월"] = m["거래일자"].dt.to_period("M").astype(str)
    company_by_month = m.groupby("월")["지각"].mean()
    dept_by_month = m.groupby(["부서", "월"])["지각"].mean()

    rows = []
    for (dept, month), dept_rate in dept_by_month.items():
        company_rate = company_by_month[month]
        if company_rate <= 0:
            continue
        diff = dept_rate - company_rate
        ratio = dept_rate / company_rate
        if ratio >= DEPT_LATE_RATIO and diff >= DEPT_LATE_ABS_DIFF:
            rows.append({
                "부서": dept, "월": month,
                "부서 지각률": round(dept_rate, 3),
                "전사 평균 지각률": round(company_rate, 3),
                "차이(%p)": round(diff * 100, 1),
                "배수": round(ratio, 2),
                "설명": f"{month} {dept} 지각률 {dept_rate:.1%} (전사 평균 {company_rate:.1%} 대비 {ratio:.1f}배)",
            })
    return pd.DataFrame(rows).sort_values("차이(%p)", ascending=False) if rows else pd.DataFrame(
        columns=["부서", "월", "부서 지각률", "전사 평균 지각률", "차이(%p)", "배수", "설명"]
    )


# ------------------------------------------------------------------
# 4) 개인 야근 위험 단계 (근로복지공단 기준)
# ------------------------------------------------------------------
def _weekly_hours_series(g: pd.DataFrame) -> pd.Series:
    """완결된(근무일 5일) 캘린더 주 단위로 총근로시간(정규+초과) 시계열을 만든다.

    부분 주(연휴 등으로 근무일이 5일이 안 되는 주)는 평균을 왜곡할 수 있어
    제외한다. 총근로시간 = 근무일수 * 8h(정규) + 그 주 초과근무시간 합.
    """
    w = g.copy()
    w["주"] = w["거래일자"].dt.to_period("W-SUN")  # 월~일 기준 ISO 유사 주
    weekly = w.groupby("주").agg(근무일수=("거래일자", "count"), 초과합=("초과근무시간", "sum"))
    weekly = weekly[weekly["근무일수"] >= 5]  # 완결 주만
    weekly["총근로시간"] = weekly["근무일수"] * DAILY_REGULAR_HOURS + weekly["초과합"]
    return weekly["총근로시간"].sort_index()


def _tier_from_averages(avg_4w: float | None, avg_12w: float | None) -> str:
    if avg_4w is not None and avg_4w > TIER_VERY_HIGH_4W:
        return "매우위험"
    if avg_12w is not None and avg_12w > TIER_HIGH_12W:
        return "위험"
    if avg_12w is not None and avg_12w > TIER_CAUTION_12W:
        return "주의"
    if avg_4w is None and avg_12w is None:
        return "데이터부족"
    return "정상"


def detect_overtime_risk(df: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    rows = []
    for (emp_id, name, dept), g in df.groupby(["사번", "이름", "부서"]):
        g = g[g["거래일자"] <= as_of]
        weekly = _weekly_hours_series(g)
        last_4 = weekly.tail(4)
        last_12 = weekly.tail(12)
        avg_4w = round(last_4.mean(), 1) if len(last_4) >= 4 else None
        avg_12w = round(last_12.mean(), 1) if len(last_12) >= 8 else None  # 8주 이상 확보시만 신뢰
        tier = _tier_from_averages(avg_4w, avg_12w)
        rows.append({
            "사번": emp_id, "이름": name, "부서": dept,
            "최근4주 평균 주당근로시간": avg_4w,
            "최근12주 평균 주당근로시간": avg_12w,
            "위험단계": tier,
        })
    out = pd.DataFrame(rows)
    tier_order = {"매우위험": 0, "위험": 1, "주의": 2, "정상": 3, "데이터부족": 4}
    out["_순서"] = out["위험단계"].map(tier_order)
    return out.sort_values(["_순서", "최근4주 평균 주당근로시간"], ascending=[True, False]).drop(columns="_순서")


# ------------------------------------------------------------------
# 5) 부서 야근 이상치 (월별, 전사 평균 대비 + 위험군 비율)
# ------------------------------------------------------------------
def detect_dept_overtime_anomaly(df: pd.DataFrame) -> pd.DataFrame:
    m = df.copy()
    m["월"] = m["거래일자"].dt.to_period("M").astype(str)
    company_by_month = m.groupby("월")["초과근무시간"].mean()
    dept_by_month = m.groupby(["부서", "월"])["초과근무시간"].mean()

    rows = []
    for (dept, month), dept_val in dept_by_month.items():
        company_val = company_by_month[month]
        if company_val <= 0:
            continue
        diff = dept_val - company_val
        ratio = dept_val / company_val
        if ratio >= DEPT_OT_RATIO and diff >= DEPT_OT_ABS_DIFF:
            rows.append({
                "부서": dept, "월": month,
                "부서 평균 초과근무(h/일)": round(dept_val, 2),
                "전사 평균 초과근무(h/일)": round(company_val, 2),
                "차이(h)": round(diff, 2),
                "배수": round(ratio, 2),
                "설명": f"{month} {dept} 평균 초과근무 {dept_val:.1f}h (전사 평균 {company_val:.1f}h 대비 {ratio:.1f}배)",
            })
    return pd.DataFrame(rows).sort_values("차이(h)", ascending=False) if rows else pd.DataFrame(
        columns=["부서", "월", "부서 평균 초과근무(h/일)", "전사 평균 초과근무(h/일)", "차이(h)", "배수", "설명"]
    )


# ------------------------------------------------------------------
# 6) 개인 종합 현황표 (전체 / 위험군 / 이상패턴 필터용)
# ------------------------------------------------------------------
def build_person_summary(
    df: pd.DataFrame,
    weekday_df: pd.DataFrame,
    trend_df: pd.DataFrame,
    risk_df: pd.DataFrame,
) -> pd.DataFrame:
    """사람별로 지각/야근/위험단계를 한 줄에 모은 종합표.

    엑셀에서 '위험군' / '이상패턴' 컬럼으로 자동필터를 걸면 전체·위험군만·
    이상패턴만을 그 자리에서 골라볼 수 있다(save_excel_report의
    autofilter=True로 헤더에 필터 버튼이 붙는다).
    """
    pattern_by_emp: dict[int, list[str]] = {}
    for _, r in weekday_df.iterrows():
        pattern_by_emp.setdefault(r["사번"], []).append(f"요일집중형({r['집중 요일']})")
    for _, r in trend_df.iterrows():
        pattern_by_emp.setdefault(r["사번"], []).append("추세악화형")

    late_agg = df.groupby(["사번", "이름", "부서"]).agg(
        지각횟수=("지각", "sum"),
        총지각시간_분=("지각분", "sum"),
        야근총시간_h=("초과근무시간", "sum"),
    ).reset_index()
    late_agg["평균지각시간_분"] = np.where(
        late_agg["지각횟수"] > 0,
        (late_agg["총지각시간_분"] / late_agg["지각횟수"]).round(1),
        0.0,
    )

    out = late_agg.merge(
        risk_df[["사번", "최근4주 평균 주당근로시간", "최근12주 평균 주당근로시간", "위험단계"]],
        on="사번", how="left",
    )
    out["이상패턴"] = out["사번"].map(lambda e: ", ".join(pattern_by_emp.get(e, [])) or "-")
    out["위험군"] = out["위험단계"].isin(["위험", "매우위험"]).map({True: "Y", False: "N"})
    out["이상패턴여부"] = out["사번"].map(lambda e: "Y" if e in pattern_by_emp else "N")

    out = out[[
        "사번", "이름", "부서",
        "지각횟수", "총지각시간_분", "평균지각시간_분",
        "야근총시간_h", "최근4주 평균 주당근로시간", "최근12주 평균 주당근로시간",
        "위험단계", "위험군", "이상패턴", "이상패턴여부",
    ]]
    return out.sort_values(["위험군", "이상패턴여부", "총지각시간_분"], ascending=[False, False, False]).reset_index(drop=True)


TIER_COLOR = {"매우위험": "#C0102A", "위험": "#E8630A", "주의": "#F2B705", "정상": "#3A7D44", "데이터부족": "#999999"}


def plot_person_overview(person_df: pd.DataFrame, output_path, scope: str = "위험군+이상패턴", max_rows: int = 16) -> Path:
    """사람별 지각/야근/위험단계 종합표를 이미지(표)로 렌더링.

    scope="전체": 전 인원, scope="위험군+이상패턴": 위험군이거나 이상패턴이
    있는 사람만 골라 보여준다(README 미리보기용 기본값 — 이 도구가 실제로
    무엇을 잡아내는지 한눈에 보여주는 게 목적이라 전체 20명보다 임팩트 있음).
    엑셀 리포트의 '개인_종합현황' 시트는 자동필터가 걸려 있어 전체/위험군/
    이상패턴을 직접 걸러볼 수 있다.
    """
    setup_style()
    if scope == "위험군+이상패턴":
        shown = person_df[(person_df["위험군"] == "Y") | (person_df["이상패턴여부"] == "Y")]
        title_scope = "위험군 + 이상패턴 대상자"
    else:
        shown = person_df
        title_scope = "전체 인원"
    shown = shown.head(max_rows)

    col_labels = ["이름", "부서", "지각횟수", "총지각시간(분)", "야근총시간(h)", "위험단계", "이상패턴"]
    cell_data = []
    for _, r in shown.iterrows():
        cell_data.append([
            r["이름"], r["부서"], f"{int(r['지각횟수'])}",
            f"{r['총지각시간_분']:.0f}", f"{r['야근총시간_h']:.1f}",
            r["위험단계"], r["이상패턴"],
        ])

    fig_h = max(2.0, 0.6 * (len(cell_data) + 1))
    fig, ax = plt.subplots(figsize=(14, fig_h))
    ax.axis("off")
    tbl = ax.table(cellText=cell_data, colLabels=col_labels, loc="center", cellLoc="center")
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(13)
    tbl.scale(1, 2.4)
    tbl.auto_set_column_width(col=list(range(len(col_labels))))

    for j in range(len(col_labels)):
        cell = tbl[0, j]
        cell.set_facecolor("#2E5EAA")
        cell.set_text_props(color="white", fontweight="bold")

    tier_col_idx = col_labels.index("위험단계")
    for i, row in enumerate(cell_data, start=1):
        tier = row[tier_col_idx]
        color = TIER_COLOR.get(tier, "#FFFFFF")
        cell = tbl[i, tier_col_idx]
        cell.set_facecolor(color)
        cell.set_text_props(color="white", fontweight="bold")

    ax.set_title(f"직원별 지각·야근 현황 및 위험단계 ({title_scope})", fontsize=16, pad=16)
    fig.tight_layout()
    return save_chart(fig, output_path)


# ------------------------------------------------------------------
# 차트
# ------------------------------------------------------------------
def plot_weekday_heatmap(df: pd.DataFrame, output_path) -> Path:
    setup_style()
    pivot = df.groupby(["이름", "요일"])["지각"].mean().unstack().reindex(columns=WEEKDAY_ORDER)
    pivot.columns = [WEEKDAY_KO[c] for c in pivot.columns]
    pivot = pivot.sort_index()

    fig, ax = plt.subplots(figsize=(6.5, 8))
    im = ax.imshow(pivot.values * 100, cmap="OrRd", aspect="auto", vmin=0, vmax=max(60, pivot.values.max() * 100))
    ax.set_xticks(range(len(pivot.columns)))
    ax.set_xticklabels(pivot.columns)
    ax.set_yticks(range(len(pivot.index)))
    ax.set_yticklabels(pivot.index, fontsize=8)
    for i in range(pivot.shape[0]):
        for j in range(pivot.shape[1]):
            val = pivot.values[i, j] * 100
            ax.text(j, i, f"{val:.0f}", ha="center", va="center",
                    color="white" if val > 35 else "black", fontsize=7)
    ax.set_title("직원별 요일별 지각률(%)")
    fig.colorbar(im, ax=ax, label="지각률 (%)", fraction=0.046, pad=0.04)
    return save_chart(fig, output_path)


def plot_overtime_risk_trend(df: pd.DataFrame, risk_table: pd.DataFrame, as_of: pd.Timestamp, output_path) -> Path:
    setup_style()
    flagged = risk_table[risk_table["위험단계"].isin(["위험", "매우위험"])]
    fig, ax = plt.subplots(figsize=(9, 5))
    if flagged.empty:
        ax.text(0.5, 0.5, "위험/매우위험 단계 인원 없음", ha="center", va="center", transform=ax.transAxes)
    else:
        for i, (_, row) in enumerate(flagged.iterrows()):
            g = df[(df["사번"] == row["사번"]) & (df["거래일자"] <= as_of)]
            weekly = _weekly_hours_series(g).tail(12)
            x = list(range(len(weekly)))
            ax.plot(x, weekly.values, marker="o", label=f"{row['이름']}({row['부서']})",
                    color=PALETTE[i % len(PALETTE)])
    ax.axhline(52, color="#F2B705", linestyle="--", linewidth=1, label="52h (법정 상한)")
    ax.axhline(60, color="#E8630A", linestyle="--", linewidth=1, label="60h (12주평균 - 위험)")
    ax.axhline(64, color="#C0102A", linestyle="--", linewidth=1, label="64h (4주평균 - 매우위험)")
    ax.set_title(f"위험군 직원 최근 12주 주당 총근로시간 추이 ({as_of.date()} 기준)")
    ax.set_xlabel("최근 12주 (좌: 과거 -> 우: 최신)")
    ax.set_ylabel("주당 총근로시간 (h)")
    ax.legend(fontsize=8, loc="upper left")
    return save_chart(fig, output_path)


def plot_dept_monthly_trend(df: pd.DataFrame, output_path) -> Path:
    setup_style()
    m = df.copy()
    m["월"] = m["거래일자"].dt.to_period("M").astype(str)
    late = m.groupby(["부서", "월"])["지각"].mean().unstack(0)
    ot = m.groupby(["부서", "월"])["초과근무시간"].mean().unstack(0)

    fig, axes = plt.subplots(2, 1, figsize=(10, 8), sharex=True)
    for i, dept in enumerate(late.columns):
        axes[0].plot(late.index, late[dept] * 100, marker="o", label=dept, color=PALETTE[i % len(PALETTE)])
        axes[1].plot(ot.index, ot[dept], marker="o", label=dept, color=PALETTE[i % len(PALETTE)])
    axes[0].set_title("부서별 월별 지각률 추이")
    axes[0].set_ylabel("지각률 (%)")
    axes[0].legend(fontsize=8)
    axes[1].set_title("부서별 월별 평균 초과근무시간 추이")
    axes[1].set_ylabel("평균 초과근무시간 (h/일)")
    axes[1].tick_params(axis="x", rotation=60)
    axes[1].legend(fontsize=8)
    fig.tight_layout()
    return save_chart(fig, output_path)


# ------------------------------------------------------------------
# 콘솔 요약
# ------------------------------------------------------------------
def print_summary(weekday_df, trend_df, dept_late_df, risk_df, dept_ot_df, as_of) -> None:
    print_section(f"근태 이상 패턴 분석 ({as_of.date()} 기준)")

    print("\n[1] 개인 지각 - 요일 집중형")
    if weekday_df.empty:
        print("  해당 없음")
    else:
        for _, r in weekday_df.iterrows():
            print(f"  - {r['이름']}({r['부서']}): {r['설명']}")

    print("\n[2] 개인 지각 - 최근 추세 악화형")
    if trend_df.empty:
        print("  해당 없음")
    else:
        for _, r in trend_df.iterrows():
            print(f"  - {r['이름']}({r['부서']}): {r['설명']}")

    print("\n[3] 부서 지각 이상치")
    if dept_late_df.empty:
        print("  해당 없음")
    else:
        for _, r in dept_late_df.iterrows():
            print(f"  - {r['설명']}")

    print("\n[4] 개인 야근 위험 단계 (근로복지공단 과로사 인정기준: 52h/60h/64h)")
    risky = risk_df[risk_df["위험단계"].isin(["위험", "매우위험"])]
    if risky.empty:
        print("  위험/매우위험 단계 없음")
    else:
        for _, r in risky.iterrows():
            print(f"  - {r['이름']}({r['부서']}) [{r['위험단계']}] "
                  f"최근4주 {r['최근4주 평균 주당근로시간']}h / 최근12주 {r['최근12주 평균 주당근로시간']}h")

    print("\n[5] 부서 야근 이상치")
    if dept_ot_df.empty:
        print("  해당 없음")
    else:
        for _, r in dept_ot_df.iterrows():
            print(f"  - {r['설명']}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=HERE / "sample_data" / "sample_attendance.csv")
    parser.add_argument("--as-of", default=None, help="YYYY-MM-DD (기본값: 데이터의 최신 날짜)")
    args = parser.parse_args()

    df = load_attendance(args.input)
    as_of = pd.Timestamp(args.as_of) if args.as_of else df["거래일자"].max()

    weekday_df = detect_weekday_concentration(df)
    trend_df = detect_trend_worsening(df, as_of)
    dept_late_df = detect_dept_late_anomaly(df)
    risk_df = detect_overtime_risk(df, as_of)
    dept_ot_df = detect_dept_overtime_anomaly(df)
    person_df = build_person_summary(df, weekday_df, trend_df, risk_df)

    excel_path = save_excel_report({
        "개인_종합현황": person_df,
        "개인_지각_요일집중형": weekday_df,
        "개인_지각_추세악화형": trend_df,
        "부서_지각이상치": dept_late_df,
        "개인_야근위험단계": risk_df,
        "부서_야근이상치": dept_ot_df,
    }, HERE / "output" / "attendance_pattern_report.xlsx")

    heatmap_path = plot_weekday_heatmap(df, HERE / "output" / "late_weekday_heatmap.png")
    risk_trend_path = plot_overtime_risk_trend(df, risk_df, as_of, HERE / "output" / "overtime_risk_trend.png")
    dept_trend_path = plot_dept_monthly_trend(df, HERE / "output" / "dept_monthly_trend.png")
    person_overview_path = plot_person_overview(
        person_df, HERE / "output" / "person_overview.png", scope="위험군+이상패턴"
    )
    person_overview_all_path = plot_person_overview(
        person_df, HERE / "output" / "person_overview_all.png", scope="전체"
    )

    print_summary(weekday_df, trend_df, dept_late_df, risk_df, dept_ot_df, as_of)
    print(f"\n엑셀 리포트 저장: {excel_path} (개인_종합현황 시트에서 '위험군'/'이상패턴여부' 컬럼 필터로 전체/위험군/이상패턴을 골라볼 수 있음)")
    print(f"차트 저장: {heatmap_path}, {risk_trend_path}, {dept_trend_path}")
    print(f"인원 종합표 이미지: {person_overview_path} (README 미리보기용), {person_overview_all_path} (전체 인원)")


if __name__ == "__main__":
    main()
