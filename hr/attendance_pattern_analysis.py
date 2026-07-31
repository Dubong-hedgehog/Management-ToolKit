"""
attendance_pattern_analysis.py
개인별/부서별 근태 "이상 패턴" 탐지 도구.

attendance_dashboard.py가 "이번 기간 vs 직전기/전년동기" 숫자 비교라면,
이 스크립트는 "누가, 어느 부서가 비정상적인 패턴을 보이는가"를 찾아낸다.

※ 2026-07, 실제 회사 근태 대시보드를 만들면서 확정된 규칙을 그대로 반영한다.

1) 지각/초과근무 기준 (유연근무 적용)
   - 지각 = 출근시각이 오전 10:00보다 늦은 만큼(분)
   - 초과근무 = 출근시각이 09:30~10:00 유연근무 구간일 때만 인정.
     정규퇴근시각 = 출근시각 + 9시간, 초과근무 = 실제퇴근시각 - 정규퇴근시각
   이 값들은 sample_data 생성 단계(generate_attendance_sample.py)에서 이미
   계산되어 CSV에 들어있고, 이 스크립트는 그 값을 그대로 신뢰한다 — 실제
   회사 시트에 이미 검증된 수식(L열=지각, M열=초과근무)이 있으면 Apps Script
   에서 재계산하지 않고 그대로 참조한 것과 같은 원칙이다.

2) 개인별 지각 - 월별 경고/개선 추적
   누적 합계(총 지각횟수 등)로 보면 근속연수가 길수록 불리해지는 착시가
   생긴다. 그래서 "이번달 vs 지난달" 비교로 설계했다.
       - 월 지각 5회 이상 또는 누적 지각 60분 이상 -> 그 달은 "경고"
       - 경고 -> 다음 달도 경고: 경고 지속
       - 경고 -> 다음 달 정상: 개선됨
       - 정상 -> 경고: 신규 경고
       - 그 외: 정상

3) 부서 이상치
   특정 월에 그 부서 지각률이 전사 평균 대비 비정상적으로 높은 경우.

4) 야근 위험 단계 (지각 상태와는 별개의 독립적인 기준)
   고용노동부 산재보험법 시행령 별표3(근로복지공단 과로사·뇌심혈관질환
   업무상 질병 인정기준)을 위험 구간으로 사용한다.
       - 발병 전 4주 평균 주 64시간 초과  -> 업무 관련성 "강함" (매우위험)
       - 발병 전 12주 평균 주 60시간 초과 -> 업무 관련성 "강함" (위험)
       - 발병 전 12주 평균 주 52시간 초과 -> 법정 연장근로 한도 초과 (주의)
       - 그 외                            -> 정상
   지각 상태(경고/개선됨)와 야근 위험단계(주의/위험/매우위험)는 서로 다른
   기준이므로 하나로 합치지 않고 항상 별도 컬럼으로 나란히 보여준다.

   ※ 이 기준은 "발병한 질병과 업무의 관련성"을 평가하는 산재 인정기준이며,
   그 자체로 "위험이 몇 배"라는 통계치를 제공하는 것은 아니다. 참고로
   WHO·ILO 공동연구(2021)는 주 55시간 이상 근무 시 35~40시간 근무 대비
   허혈성 심장질환 사망위험 17%, 뇌졸중 사망위험 35% 증가로 추정한다.
   (출처: 근로복지공단 업무상질병 인정기준, WHO/ILO "Long working hours
   increasing deaths from heart disease and stroke", 2021)

Usage:
    python hr/attendance_pattern_analysis.py
    python hr/attendance_pattern_analysis.py --as-of 2026-07-31
    python hr/attendance_pattern_analysis.py --input hr/sample_data/sample_attendance.csv

Input:
    hr/sample_data/sample_attendance.csv (또는 --input / .env로 설정한 구글시트)
    컬럼: 사번, 이름, 부서, 거래일자, 출근시각, 퇴근시각, 지각여부(Y/N), 초과근무시간
    (지각여부/초과근무시간은 이미 검증된 값으로 취급하고 재계산하지 않는다.
    지각"분"만 출근시각으로부터 같은 10:00 기준으로 추가 계산한다.)

Output:
    hr/output/attendance_pattern_report.xlsx
        - 개인_월간현황 / 부서_지각이상치 / 개인_야근위험단계 / 부서_야근이상치
    hr/output/person_monthly_examples.png  (경고/개선/야근위험 예시 인물 월별 추이)
    hr/output/overtime_risk_trend.png      (위험군 직원 최근 12주 근로시간 추이)
    hr/output/dept_monthly_trend.png       (부서별 월별 지각률/초과근무 추이)
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.patches import FancyBboxPatch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.chart_style import PALETTE, save_chart, setup_style  # noqa: E402
from common.excel_io import load_csv, save_excel_report  # noqa: E402
from common.format_utils import print_section  # noqa: E402
from common.sheet_io import is_configured as sheet_is_configured  # noqa: E402
from common.sheet_io import load_transactions_from_sheet  # noqa: E402

HERE = Path(__file__).resolve().parent

# ---- 지각 기준 (유연근무: 오전 10:00) ----
LATE_THRESHOLD_MIN = 10 * 60

# ---- 월별 경고 기준 ----
WARN_LATE_COUNT = 5       # 월 지각 5회 이상이면 경고
WARN_LATE_SUM_MIN = 60    # 또는 월 누적 지각 60분 이상이면 경고

# ---- 부서 이상치 ----
DEPT_LATE_RATIO = 1.5
DEPT_LATE_ABS_DIFF = 0.08        # 8%p 이상

DEPT_OT_RATIO = 1.8
DEPT_OT_ABS_DIFF = 1.0           # 1.0h 이상

# ---- 야근 위험 단계 (근로복지공단 과로사 인정기준, 지각 상태와는 별개 지표) ----
DAILY_REGULAR_HOURS = 8.0
TIER_VERY_HIGH_4W = 64.0   # 4주 평균 64h 초과 -> 매우위험
TIER_HIGH_12W = 60.0       # 12주 평균 60h 초과 -> 위험
TIER_CAUTION_12W = 52.0    # 12주 평균 52h 초과 -> 주의 (법정 연장한도 초과)

# 실제 배포된 근태 대시보드(Index.html)의 배지 색상과 동일하게 맞춘다.
TIER_COLOR = {"매우위험": "#C0102A", "위험": "#E8630A", "주의": "#EDA100", "정상": "#1BAF7A", "데이터부족": "#999999"}
STATUS_COLOR = {"경고 지속": "#C0102A", "신규 경고": "#E8630A", "개선됨": "#1BAF7A", "정상": "#999999"}

# 실제 대시보드 디자인 시스템(크림 배경 + 카드) 톤
PAGE_BG = "#FAF9F5"
CARD_BG = "#EEF0E8"
ROW_LINE = "#EDEBE3"
TEXT_DARK = "#33322C"
TEXT_MUTED = "#6B6A63"


# ------------------------------------------------------------------
# 데이터 로드
# ------------------------------------------------------------------
def load_attendance(input_path: str | Path) -> pd.DataFrame:
    if sheet_is_configured():
        df = load_transactions_from_sheet()
    else:
        df = load_csv(input_path)
    df["거래일자"] = pd.to_datetime(df["거래일자"])
    # 지각여부/초과근무시간은 이미 검증된 값으로 취급 - 재계산하지 않는다.
    df["초과근무시간"] = pd.to_numeric(df["초과근무시간"], errors="coerce").fillna(0)
    df["지각"] = (df["지각여부"] == "Y").astype(int)
    df["월"] = df["거래일자"].dt.to_period("M").astype(str)

    def _to_minutes(t: str) -> int:
        h, m = str(t).split(":")
        return int(h) * 60 + int(m)

    start_min = df["출근시각"].map(_to_minutes)
    # 지각"분"만 출근시각 기준으로 추가 계산(지각여부 Y/N과 같은 10:00 기준).
    df["지각분"] = np.where(df["지각"] == 1, (start_min - LATE_THRESHOLD_MIN).clip(lower=0), 0)
    return df


# ------------------------------------------------------------------
# 1) 부서 지각 이상치 (월별, 전사 평균 대비)
# ------------------------------------------------------------------
def detect_dept_late_anomaly(df: pd.DataFrame) -> pd.DataFrame:
    company_by_month = df.groupby("월")["지각"].mean()
    dept_by_month = df.groupby(["부서", "월"])["지각"].mean()

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
# 2) 개인 야근 위험 단계 (근로복지공단 기준) - 지각 상태와는 별개 지표
# ------------------------------------------------------------------
def _weekly_hours_series(g: pd.DataFrame) -> pd.Series:
    """완결된(근무일 5일) 캘린더 주 단위로 총근로시간(정규+초과) 시계열을 만든다.

    부분 주(연휴 등으로 근무일이 5일이 안 되는 주)는 평균을 왜곡할 수 있어
    제외한다. 총근로시간 = 근무일수 * 8h(정규) + 그 주 초과근무시간 합.
    """
    w = g.copy()
    w["주"] = w["거래일자"].dt.to_period("W-SUN")
    weekly = w.groupby("주").agg(근무일수=("거래일자", "count"), 초과합=("초과근무시간", "sum"))
    weekly = weekly[weekly["근무일수"] >= 5]
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
        avg_12w = round(last_12.mean(), 1) if len(last_12) >= 8 else None
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
# 3) 부서 야근 이상치 (월별, 전사 평균 대비)
# ------------------------------------------------------------------
def detect_dept_overtime_anomaly(df: pd.DataFrame) -> pd.DataFrame:
    company_by_month = df.groupby("월")["초과근무시간"].mean()
    dept_by_month = df.groupby(["부서", "월"])["초과근무시간"].mean()

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
# 4) 개인 월별 지각/초과근무 집계 + 경고/개선 추적
# ------------------------------------------------------------------
def build_monthly_stats(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby(["사번", "이름", "부서", "월"]).agg(
        지각횟수=("지각", "sum"),
        지각분합=("지각분", "sum"),
        초과근무합=("초과근무시간", "sum"),
    ).reset_index()


def _is_warn_month(row: pd.Series) -> bool:
    return row["지각횟수"] >= WARN_LATE_COUNT or row["지각분합"] >= WARN_LATE_SUM_MIN


def build_person_monthly_trend(monthly: pd.DataFrame, risk_df: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    """개인별 "이번달 vs 지난달" 지각 상태 + 초과근무 위험단계(별도 컬럼).

    누적 합계 대신 월 단위 경고/개선 추적으로 설계한 이유: 총 지각횟수 같은
    누적치는 근속연수가 길수록 불리해지는 착시를 만든다. 신규 입사자와
    5년차 직원을 같은 "총 지각횟수"로 비교하는 건 공정하지 않다.
    """
    cur_ym = as_of.strftime("%Y-%m")
    risk_lookup = risk_df.set_index("사번")["위험단계"].to_dict()

    rows = []
    for (emp_id, name, dept), g in monthly.groupby(["사번", "이름", "부서"]):
        g = g.sort_values("월")
        yms = g["월"].tolist()
        cur = g[g["월"] == cur_ym]
        cur_row = cur.iloc[0] if len(cur) else pd.Series({"지각횟수": 0, "지각분합": 0, "초과근무합": 0.0})

        prev_row = None
        if cur_ym in yms:
            idx = yms.index(cur_ym)
            if idx > 0:
                prev_row = g[g["월"] == yms[idx - 1]].iloc[0]

        cur_warn = _is_warn_month(cur_row)
        prev_warn = _is_warn_month(prev_row) if prev_row is not None else False

        if cur_warn and prev_warn:
            status = "경고 지속"
        elif cur_warn and not prev_warn:
            status = "신규 경고"
        elif not cur_warn and prev_warn:
            status = "개선됨"
        else:
            status = "정상"

        rows.append({
            "사번": emp_id, "이름": name, "부서": dept,
            "이번달_지각횟수": int(cur_row["지각횟수"]),
            "이번달_지각분": int(cur_row["지각분합"]),
            "전월_지각분": int(prev_row["지각분합"]) if prev_row is not None else None,
            "전월대비_지각분": (int(cur_row["지각분합"]) - int(prev_row["지각분합"])) if prev_row is not None else None,
            "이번달_초과근무_h": round(float(cur_row["초과근무합"]), 1),
            "지각상태": status,
            "초과근무위험단계": risk_lookup.get(emp_id, "데이터부족"),
        })

    out = pd.DataFrame(rows)
    status_order = {"경고 지속": 0, "신규 경고": 1, "개선됨": 2, "정상": 3}
    tier_order = {"매우위험": 0, "위험": 1, "주의": 2, "정상": 3, "데이터부족": 4}
    out["_s"] = out["지각상태"].map(status_order)
    out["_t"] = out["초과근무위험단계"].map(tier_order)
    out["_심각도"] = out[["_s", "_t"]].min(axis=1)
    out = out.sort_values(["_심각도", "이번달_지각분", "이번달_초과근무_h"], ascending=[True, False, False])
    return out.drop(columns=["_s", "_t", "_심각도"]).reset_index(drop=True)


def get_recent_series(monthly: pd.DataFrame, emp_id: int, as_of: pd.Timestamp, n_months: int = 3) -> pd.DataFrame:
    """이번달 포함 최근 n개월(데이터 있는 만큼만) 시리즈."""
    cur_ym = as_of.strftime("%Y-%m")
    g = monthly[(monthly["사번"] == emp_id) & (monthly["월"] <= cur_ym)].sort_values("월")
    return g.tail(n_months)


def pick_example_people(person_trend_df: pd.DataFrame, max_examples: int = 4) -> list[int]:
    """지각상태(경고 지속/신규 경고/개선됨)와 야근 위험단계를 골고루 보여줄 예시 인원 선정."""
    picks: list[int] = []
    seen_status: set[str] = set()
    for _, r in person_trend_df.iterrows():
        status = r["지각상태"]
        if status != "정상" and status not in seen_status:
            seen_status.add(status)
            picks.append(int(r["사번"]))
        if len(picks) >= max_examples:
            break
    if len(picks) < max_examples:
        risky = person_trend_df[person_trend_df["초과근무위험단계"].isin(["매우위험", "위험", "주의"])]
        for _, r in risky.iterrows():
            emp = int(r["사번"])
            if emp not in picks:
                picks.append(emp)
                break
    return picks[:max_examples]


# ------------------------------------------------------------------
# 차트
# ------------------------------------------------------------------
def plot_person_monthly_examples(monthly: pd.DataFrame, person_trend_df: pd.DataFrame, as_of: pd.Timestamp, output_path) -> Path:
    """예시 인물 몇 명의 최근 3개월 지각(분)/초과근무(h) 추이를 나란히 보여준다.

    실제 배포된 근태 대시보드에서 "이름 클릭 -> 개인별 그래프" 형태로 보여준
    것과 같은 그림을, 여기서는 대표 사례 몇 명을 뽑아 한 이미지에 모았다.
    """
    setup_style()
    emp_ids = pick_example_people(person_trend_df, max_examples=4)
    if not emp_ids:
        emp_ids = person_trend_df["사번"].head(4).astype(int).tolist()

    n = len(emp_ids)
    fig, axes = plt.subplots(1, n, figsize=(4.6 * n, 4.2))
    if n == 1:
        axes = [axes]

    info_by_emp = person_trend_df.set_index("사번").to_dict("index")

    for ax, emp_id in zip(axes, emp_ids):
        series = get_recent_series(monthly, emp_id, as_of, n_months=3)
        info = info_by_emp.get(emp_id, {})
        name = series["이름"].iloc[0] if len(series) else str(emp_id)
        dept = series["부서"].iloc[0] if len(series) else ""

        x = list(range(len(series)))
        late_vals = series["지각분합"].tolist()
        ot_vals = series["초과근무합"].tolist()

        bar_color = "#E4572E"
        line_color = "#2F6FED"

        ax.bar(x, late_vals, color=bar_color, alpha=0.85, width=0.5, label="지각(분)", zorder=2)
        late_max = max(late_vals) if late_vals else 0
        ax.set_ylim(0, late_max * 1.35 + 5)
        for i, v in enumerate(late_vals):
            if v > 0:
                ax.annotate(f"{v:.0f}분", (i, v), textcoords="offset points", xytext=(0, 6),
                            ha="center", va="bottom", fontsize=9, color=bar_color, fontweight="bold")

        ax2 = ax.twinx()
        ax2.plot(x, ot_vals, marker="o", color=line_color, linewidth=2, zorder=3, label="초과근무(h)")
        ot_max = max(ot_vals) if ot_vals else 0
        ot_min_val = min(ot_vals) if ot_vals else 0
        pad = max(ot_max * 0.3, 1)
        ax2.set_ylim(min(0, ot_min_val) - pad * 0.2, ot_max + pad)
        for i, v in enumerate(ot_vals):
            ax2.annotate(f"{v:.1f}h", (i, v), textcoords="offset points", xytext=(0, 14),
                         ha="center", va="bottom", fontsize=9, color=line_color, fontweight="bold")

        ax.set_xticks(x)
        ax.set_xticklabels(series["월"].tolist(), fontsize=9)
        status = info.get("지각상태", "-")
        tier = info.get("초과근무위험단계", "-")
        ax.set_title(f"{name}({dept})\n지각상태: {status} / 야근위험: {tier}", fontsize=10)
        ax.set_ylabel("지각(분)", fontsize=8, color=bar_color)
        ax2.set_ylabel("초과근무(h)", fontsize=8, color=line_color)
        ax.tick_params(axis="y", labelcolor=bar_color)
        ax2.tick_params(axis="y", labelcolor=line_color)

    fig.suptitle(f"개인별 지각·초과근무 월간 추이 예시 ({as_of.date()} 기준)", fontsize=13)
    fig.tight_layout(rect=[0, 0, 1, 0.92])
    return save_chart(fig, output_path)


def _draw_pill(ax, cx: float, cy: float, text: str, color: str, w: float = 1.15, h: float = 0.34) -> None:
    box = FancyBboxPatch((cx - w / 2, cy - h / 2), w, h,
                          boxstyle=f"round,pad=0,rounding_size={h / 2}",
                          linewidth=0, facecolor=color, zorder=3)
    ax.add_patch(box)
    ax.text(cx, cy, text, fontsize=9.5, color="white", fontweight="bold",
            ha="center", va="center", zorder=4)


def plot_person_overview(person_trend_df: pd.DataFrame, as_of: pd.Timestamp, output_path,
                          scope: str = "관심대상", max_rows: int = 16) -> Path:
    """실제 배포된 근태 대시보드(Index.html)와 같은 카드+표 형태로 렌더링한다.

    상단에 신규경고/경고지속/개선됨 카드 3개(전 인원 기준 집계), 그 아래 표에
    지각 상태 / 야근 위험단계를 배지(pill) 색상으로 보여준다 — matplotlib
    ax.table의 파란 헤더 표 대신, 실사용 대시보드와 동일한 크림 배경 +
    카드 + 배지 스타일을 그대로 재현했다.

    scope="관심대상": 지각상태가 정상이 아니거나 야근위험단계가 주의 이상인
    사람만 표에 보여준다(카드 집계는 scope와 무관하게 항상 전 인원 기준).
    scope="전체": 표에 전 인원을 보여준다.
    """
    setup_style()
    if scope == "관심대상":
        shown = person_trend_df[
            (person_trend_df["지각상태"] != "정상") |
            (person_trend_df["초과근무위험단계"].isin(["주의", "위험", "매우위험"]))
        ]
    else:
        shown = person_trend_df
    shown = shown.head(max_rows).reset_index(drop=True)

    new_warn = int((person_trend_df["지각상태"] == "신규 경고").sum())
    ongoing = int((person_trend_df["지각상태"] == "경고 지속").sum())
    improved = int((person_trend_df["지각상태"] == "개선됨").sum())

    n_rows = len(shown)
    row_h = 0.6
    fig_w = 13.6
    # 헤더 영역 높이(제목 0.5~0.9 + 카드 y0 1.05 + 카드높이 1.25 + 캡션여백 0.4
    # + 헤더행 0.42 + 헤더밑줄 0.3)를 실제 그리는 좌표와 동일하게 맞춰서 계산 -
    # 안 맞으면 마지막 행이 하단 각주와 겹친다.
    header_area_h = 1.05 + 1.25 + 0.4 + 0.42 + 0.3
    footer_margin = 0.9
    fig_h = header_area_h + row_h * max(n_rows, 1) + footer_margin

    fig = plt.figure(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor(PAGE_BG)
    ax = fig.add_axes((0, 0, 1, 1))
    ax.set_xlim(0, fig_w)
    ax.set_ylim(0, fig_h)
    ax.invert_yaxis()
    ax.axis("off")
    ax.set_facecolor(PAGE_BG)

    # 제목
    ax.text(0.35, 0.5, "이번 달 근태 경고 현황 (전월 대비)", fontsize=15, fontweight="bold", color=TEXT_DARK, va="center")

    # 카드 3개
    card_y0, card_h, gap = 1.05, 1.25, 0.3
    card_w = (fig_w - 0.7 - 2 * gap) / 3
    for i, (label, val) in enumerate([("신규 경고", new_warn), ("경고 지속", ongoing), ("개선됨", improved)]):
        x0 = 0.35 + i * (card_w + gap)
        ax.add_patch(FancyBboxPatch((x0, card_y0), card_w, card_h,
                                     boxstyle="round,pad=0,rounding_size=0.12",
                                     linewidth=0, facecolor=CARD_BG))
        ax.text(x0 + 0.28, card_y0 + 0.36, label, fontsize=11, color=TEXT_MUTED, va="center")
        ax.text(x0 + 0.28, card_y0 + 0.92, f"{val}명", fontsize=23, fontweight="bold", color=TEXT_DARK, va="center")

    # 기준월 캡션
    caption_y = card_y0 + card_h + 0.4
    ax.text(0.35, caption_y, f"기준월: {as_of.strftime('%Y-%m')}", fontsize=10, color=TEXT_MUTED, va="center")

    # 표
    col_defs = [
        ("이름", 0.35, "left"), ("부서", 1.75, "left"),
        ("이번달 지각횟수", 4.35, "right"), ("이번달 지각(분)", 5.95, "right"),
        ("전월대비", 7.15, "right"), ("이번달 초과근무(h)", 8.95, "right"),
        ("지각 상태", 10.65, "center"), ("야근 위험단계", 12.35, "center"),
    ]
    header_y = caption_y + 0.42
    for label, x, align in col_defs:
        ax.text(x, header_y, label, fontsize=10, fontweight="bold", color=TEXT_MUTED,
                 ha=align if align != "center" else "center", va="center")
    header_rule_y = header_y + 0.3
    ax.plot([0.35, fig_w - 0.35], [header_rule_y, header_rule_y], color="#D9D6CC", linewidth=1.1, zorder=2)

    for i, r in shown.iterrows():
        y = header_rule_y + row_h * (i + 0.62)
        diff = r["전월대비_지각분"]
        diff_str = "-" if pd.isna(diff) else (f"+{int(diff)}분" if diff > 0 else (f"{int(diff)}분" if diff < 0 else "±0분"))

        ax.text(col_defs[0][1], y, str(r["이름"]), fontsize=10, color=TEXT_DARK, ha="left", va="center")
        ax.text(col_defs[1][1], y, str(r["부서"]), fontsize=10, color=TEXT_DARK, ha="left", va="center")
        ax.text(col_defs[2][1], y, str(int(r["이번달_지각횟수"])), fontsize=10, color=TEXT_DARK, ha="right", va="center")
        ax.text(col_defs[3][1], y, str(int(r["이번달_지각분"])), fontsize=10, color=TEXT_DARK, ha="right", va="center")
        ax.text(col_defs[4][1], y, diff_str, fontsize=10, color=TEXT_DARK, ha="right", va="center")
        ax.text(col_defs[5][1], y, f"{r['이번달_초과근무_h']:.1f}", fontsize=10, color=TEXT_DARK, ha="right", va="center")

        _draw_pill(ax, col_defs[6][1], y, r["지각상태"], STATUS_COLOR.get(r["지각상태"], "#999999"))
        _draw_pill(ax, col_defs[7][1], y, r["초과근무위험단계"], TIER_COLOR.get(r["초과근무위험단계"], "#999999"))

        row_rule_y = header_rule_y + row_h * (i + 1)
        ax.plot([0.35, fig_w - 0.35], [row_rule_y, row_rule_y], color=ROW_LINE, linewidth=0.8, zorder=1)

    footer_y = header_rule_y + row_h * (n_rows + 0.6)
    ax.text(0.35, footer_y,
            "경고 기준: 월 지각 5회 이상 또는 누적 60분 이상. 지각/초과근무 값은 시트 지각·초과근무 열 값을 그대로 사용합니다.",
            fontsize=8.5, color=TEXT_MUTED, va="center")

    return save_chart(fig, output_path, tight=False)


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
    late = df.groupby(["부서", "월"])["지각"].mean().unstack(0)
    ot = df.groupby(["부서", "월"])["초과근무시간"].mean().unstack(0)

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
def print_summary(person_trend_df, dept_late_df, risk_df, dept_ot_df, as_of) -> None:
    print_section(f"근태 이상 패턴 분석 ({as_of.date()} 기준)")

    print("\n[1] 개인 지각 상태 - 신규 경고 / 경고 지속")
    flagged = person_trend_df[person_trend_df["지각상태"].isin(["신규 경고", "경고 지속"])]
    if flagged.empty:
        print("  해당 없음")
    else:
        for _, r in flagged.iterrows():
            print(f"  - {r['이름']}({r['부서']}) [{r['지각상태']}] 이번달 지각 {r['이번달_지각횟수']}회 / {r['이번달_지각분']}분")

    print("\n[2] 개인 지각 상태 - 개선됨")
    improved = person_trend_df[person_trend_df["지각상태"] == "개선됨"]
    if improved.empty:
        print("  해당 없음")
    else:
        for _, r in improved.iterrows():
            print(f"  - {r['이름']}({r['부서']}): 지난달 경고 -> 이번달 {r['이번달_지각분']}분으로 개선")

    print("\n[3] 부서 지각 이상치")
    if dept_late_df.empty:
        print("  해당 없음")
    else:
        for _, r in dept_late_df.iterrows():
            print(f"  - {r['설명']}")

    print("\n[4] 개인 야근 위험 단계 (근로복지공단 과로사 인정기준: 52h/60h/64h, 지각 상태와 별개 지표)")
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

    dept_late_df = detect_dept_late_anomaly(df)
    risk_df = detect_overtime_risk(df, as_of)
    dept_ot_df = detect_dept_overtime_anomaly(df)

    monthly = build_monthly_stats(df)
    person_trend_df = build_person_monthly_trend(monthly, risk_df, as_of)

    excel_path = save_excel_report({
        "개인_월간현황": person_trend_df,
        "부서_지각이상치": dept_late_df,
        "개인_야근위험단계": risk_df,
        "부서_야근이상치": dept_ot_df,
    }, HERE / "output" / "attendance_pattern_report.xlsx")

    examples_path = plot_person_monthly_examples(monthly, person_trend_df, as_of, HERE / "output" / "person_monthly_examples.png")
    risk_trend_path = plot_overtime_risk_trend(df, risk_df, as_of, HERE / "output" / "overtime_risk_trend.png")
    dept_trend_path = plot_dept_monthly_trend(df, HERE / "output" / "dept_monthly_trend.png")
    person_overview_path = plot_person_overview(person_trend_df, as_of, HERE / "output" / "person_overview.png", scope="관심대상")
    person_overview_all_path = plot_person_overview(person_trend_df, as_of, HERE / "output" / "person_overview_all.png", scope="전체")

    print_summary(person_trend_df, dept_late_df, risk_df, dept_ot_df, as_of)
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {examples_path}, {risk_trend_path}, {dept_trend_path}")
    print(f"인원 현황표 이미지: {person_overview_path} (README 미리보기용), {person_overview_all_path} (전체 인원)")


if __name__ == "__main__":
    main()
