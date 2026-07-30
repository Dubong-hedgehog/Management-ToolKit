"""
contract_wbs_tracker.py
계약 전~종료까지 전 단계를 WBS(작업분류체계)로 관리하는 도구.

계약 하나마다 4단계(계약전 / 계약체결 / 이행 / 종료)로 표준 태스크를
두고, 각 태스크의 계획일/실제완료일을 비교해 지연 여부를 판단한다.
계약체결 단계에는 "계약이행보증/선급금보증 가입확인"을, 종료 단계 마지막에는
"실적증명서 수취"(공공기관 발주 계약 대상)를 표준 체크 항목으로 포함한다 —
둘 다 실무에서 자주 놓치는 항목이라 별도 체크리스트 시트로 뽑아준다.

계약 만료가 임박(D-30/D-60)한 진행중 계약은 자동으로 감지해서 알림을
보낸다. 알림 채널은 .env의 NOTIFY_CHANNEL로 정한다(기본값 email, slack/teams
로 변경 가능 — common/notify_utils.py 참고). 채널 자격증명이 없으면 콘솔
출력 + 로그 파일 기록으로 대체되므로 별도 설정 없이도 항상 실행된다.

Usage:
    python procurement/contract_wbs_tracker.py
    python procurement/contract_wbs_tracker.py --as-of 2026-07-29
    python procurement/contract_wbs_tracker.py --delay-threshold-days 3

Input:
    procurement/sample_data/contracts.csv
        계약번호, 계약명, 공급업체, 계약유형(물품/용역/공사), 발주기관유형(민간/공공),
        계약금액, 계약시작일, 계약종료일, 계약상태(진행중/종료/해지)
    procurement/sample_data/wbs_tasks.csv
        계약번호, 단계, 순서, 작업명, 담당자, 계획일, 실제완료일, 상태(참고용, 스크립트가
        as-of 기준으로 재계산함), 결과메모

Output:
    procurement/output/contract_wbs_report.xlsx
        - 계약_진행현황 / 지연_태스크 / 계약만료_임박 / 보증_실적증명서_체크리스트
    procurement/output/contract_gantt.png       (진행중 계약 간트차트)
    procurement/output/notify_log.csv           (알림 채널 미설정 시 발송 이력 대체 기록)
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.chart_style import save_chart, setup_style  # noqa: E402
from common.excel_io import load_csv, save_excel_report  # noqa: E402
from common.format_utils import print_section  # noqa: E402
from common.notify_utils import send_alert  # noqa: E402

HERE = Path(__file__).resolve().parent

DELAY_THRESHOLD_DAYS_DEFAULT = 3   # 계획일보다 이 일수 이상 지나도 미완료면 "지연"
EXPIRY_WARN_DAYS = 30              # D-30 경고
EXPIRY_ALERT_DAYS = 60             # D-60 예고(참고용)

GUARANTEE_TASKS = {"계약이행보증 가입확인", "선급금보증 가입확인"}
CERT_TASK = "실적증명서 수취"
PROBLEM_MEMOS = {"면제(소액계약)": False, "해당없음(선급금 미지급)": False, "해당없음(민간발주)": False}


# ------------------------------------------------------------------
# 로드
# ------------------------------------------------------------------
def load_contracts(path) -> pd.DataFrame:
    df = load_csv(path)
    df["계약시작일"] = pd.to_datetime(df["계약시작일"])
    df["계약종료일"] = pd.to_datetime(df["계약종료일"])
    return df


def load_wbs(path) -> pd.DataFrame:
    df = load_csv(path)
    df["계획일"] = pd.to_datetime(df["계획일"])
    df["실제완료일"] = pd.to_datetime(df["실제완료일"], errors="coerce")
    df["결과메모"] = df["결과메모"].fillna("")
    return df


# ------------------------------------------------------------------
# 상태 재계산 (as-of 기준 — CSV의 상태 컬럼은 참고용일 뿐 신뢰하지 않는다)
# ------------------------------------------------------------------
def compute_task_status(wbs: pd.DataFrame, as_of: pd.Timestamp, delay_threshold_days: int) -> pd.DataFrame:
    w = wbs.copy()

    def _status(row):
        if pd.notna(row["실제완료일"]):
            return "완료"
        if row["계획일"] + pd.Timedelta(days=delay_threshold_days) < as_of:
            return "지연"
        if row["계획일"] <= as_of:
            return "진행중"
        return "예정"

    w["상태_재계산"] = w.apply(_status, axis=1)
    return w


# ------------------------------------------------------------------
# 계약별 진행률
# ------------------------------------------------------------------
def build_contract_progress(contracts: pd.DataFrame, wbs: pd.DataFrame) -> pd.DataFrame:
    agg = wbs.groupby("계약번호").agg(
        전체태스크=("작업명", "count"),
        완료태스크=("상태_재계산", lambda s: (s == "완료").sum()),
        지연태스크=("상태_재계산", lambda s: (s == "지연").sum()),
    ).reset_index()
    agg["진행률"] = (agg["완료태스크"] / agg["전체태스크"]).round(3)

    out = contracts.merge(agg, on="계약번호", how="left")
    return out[[
        "계약번호", "계약명", "공급업체", "계약유형", "발주기관유형", "계약상태",
        "계약금액", "계약시작일", "계약종료일", "전체태스크", "완료태스크", "지연태스크", "진행률",
    ]].sort_values(["계약상태", "진행률"])


# ------------------------------------------------------------------
# 지연 태스크 리스트
# ------------------------------------------------------------------
def build_delay_list(wbs: pd.DataFrame, contracts: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    delayed = wbs[wbs["상태_재계산"] == "지연"].copy()
    if delayed.empty:
        return pd.DataFrame(columns=["계약번호", "계약명", "공급업체", "단계", "작업명", "담당자", "계획일", "지연일수"])
    delayed = delayed.merge(contracts[["계약번호", "계약명", "공급업체"]], on="계약번호", how="left")
    delayed["지연일수"] = (as_of - delayed["계획일"]).dt.days
    return delayed[["계약번호", "계약명", "공급업체", "단계", "작업명", "담당자", "계획일", "지연일수"]].sort_values(
        "지연일수", ascending=False
    )


# ------------------------------------------------------------------
# 계약만료 임박
# ------------------------------------------------------------------
def build_expiry_alerts(contracts: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    active = contracts[contracts["계약상태"] == "진행중"].copy()
    active["잔여일수"] = (active["계약종료일"] - as_of).dt.days
    active = active[active["잔여일수"] <= EXPIRY_ALERT_DAYS]

    def _bucket(days):
        if days < 0:
            return "만료됨(즉시 확인 필요)"
        if days <= EXPIRY_WARN_DAYS:
            return f"D-{EXPIRY_WARN_DAYS} 이내(긴급)"
        return f"D-{EXPIRY_ALERT_DAYS} 이내(예고)"

    active["긴급도"] = active["잔여일수"].map(_bucket)
    return active[[
        "계약번호", "계약명", "공급업체", "계약유형", "계약종료일", "잔여일수", "긴급도",
    ]].sort_values("잔여일수")


def notify_expiry(expiry_df: pd.DataFrame, as_of: pd.Timestamp) -> str:
    if expiry_df.empty:
        return "만료 임박 계약 없음 - 알림 미발송"
    lines = [f"{as_of.date()} 기준 계약만료 임박 알림 ({len(expiry_df)}건)"]
    for _, r in expiry_df.iterrows():
        dday = f"D-{r['잔여일수']}" if r["잔여일수"] >= 0 else f"D+{-r['잔여일수']} 경과"
        lines.append(f"- [{r['긴급도']}] {r['계약명']}({r['공급업체']}) 종료일 {r['계약종료일'].date()} ({dday})")
    body = "\n".join(lines)
    result = send_alert(
        subject=f"[계약만료 임박] {len(expiry_df)}건 확인 필요",
        body=body,
        fallback_log_path=HERE / "output" / "notify_log.csv",
    )
    return result


# ------------------------------------------------------------------
# 보증보험/실적증명서 체크리스트
# ------------------------------------------------------------------
def build_guarantee_cert_checklist(wbs: pd.DataFrame, contracts: pd.DataFrame) -> pd.DataFrame:
    target = wbs[wbs["작업명"].isin(GUARANTEE_TASKS | {CERT_TASK})].copy()
    target = target.merge(contracts[["계약번호", "계약명", "공급업체", "발주기관유형"]], on="계약번호", how="left")

    def _flag_problem(memo: str) -> str:
        if memo in ("미가입(확인필요)", "미수취(누락)"):
            return "확인 필요"
        if memo == "":
            return "미완료(진행중)"
        return "정상"

    target["점검결과"] = target["결과메모"].map(_flag_problem)
    return target[[
        "계약번호", "계약명", "공급업체", "발주기관유형", "작업명", "담당자", "계획일", "결과메모", "점검결과",
    ]].sort_values(["점검결과", "계약번호"])


# ------------------------------------------------------------------
# 간트차트 (진행중 계약, 계약 단위)
# ------------------------------------------------------------------
def plot_contract_gantt(contracts: pd.DataFrame, wbs: pd.DataFrame, as_of: pd.Timestamp, output_path) -> Path:
    setup_style()
    active = contracts[contracts["계약상태"] == "진행중"].sort_values("계약종료일").copy()
    has_delay = wbs[wbs["상태_재계산"] == "지연"]["계약번호"].unique()
    active["지연있음"] = active["계약번호"].isin(has_delay)

    fig, ax = plt.subplots(figsize=(10, max(3, 0.5 * len(active) + 1)))
    for i, (_, r) in enumerate(active.iterrows()):
        start_num = mdates.date2num(r["계약시작일"])
        end_num = mdates.date2num(r["계약종료일"])
        color = "#C0102A" if r["지연있음"] else "#2E5EAA"
        ax.broken_barh([(start_num, end_num - start_num)], (i - 0.4, 0.8), facecolors=color, alpha=0.85)
        ax.text(end_num + 2, i, f"{r['계약명']} ({r['공급업체']})", va="center", fontsize=8)

    as_of_num = mdates.date2num(as_of)
    ax.axvline(as_of_num, color="black", linestyle="-", linewidth=1, label="기준일(오늘)")
    ax.axvline(mdates.date2num(as_of + pd.Timedelta(days=EXPIRY_WARN_DAYS)), color="#F2B705", linestyle="--", linewidth=1, label=f"D+{EXPIRY_WARN_DAYS}")
    ax.axvline(mdates.date2num(as_of + pd.Timedelta(days=EXPIRY_ALERT_DAYS)), color="#E8630A", linestyle="--", linewidth=1, label=f"D+{EXPIRY_ALERT_DAYS}")

    ax.xaxis_date()
    ax.set_yticks(range(len(active)))
    ax.set_yticklabels([r["계약번호"] for _, r in active.iterrows()])
    ax.set_title(f"진행중 계약 현황 ({as_of.date()} 기준, 빨강=지연 태스크 있음)")
    ax.legend(fontsize=8, loc="upper left")
    fig.tight_layout()
    return save_chart(fig, output_path)


# ------------------------------------------------------------------
def print_summary(progress_df, delay_df, expiry_df, checklist_df, notify_result, as_of) -> None:
    print_section(f"구매/계약 WBS 현황 ({as_of.date()} 기준)")

    print("\n[1] 진행중 계약 진행률")
    active = progress_df[progress_df["계약상태"] == "진행중"]
    for _, r in active.iterrows():
        print(f"  - {r['계약명']}({r['공급업체']}): {r['진행률']:.0%} "
              f"({r['완료태스크']}/{r['전체태스크']}), 지연 {r['지연태스크']}건")

    print("\n[2] 지연 태스크")
    if delay_df.empty:
        print("  해당 없음")
    else:
        for _, r in delay_df.iterrows():
            print(f"  - {r['계약명']} [{r['단계']}] {r['작업명']} - 계획일 {r['계획일'].date()} 대비 {r['지연일수']}일 지연 (담당 {r['담당자']})")

    print("\n[3] 계약만료 임박")
    if expiry_df.empty:
        print("  해당 없음")
    else:
        for _, r in expiry_df.iterrows():
            dday = f"D-{r['잔여일수']}" if r["잔여일수"] >= 0 else f"D+{-r['잔여일수']} 경과"
            print(f"  - [{r['긴급도']}] {r['계약명']}({r['공급업체']}) 종료 {r['계약종료일'].date()} ({dday})")
    print(f"  알림 발송 결과: {notify_result}")

    print("\n[4] 보증보험/실적증명서 확인 필요 항목")
    problems = checklist_df[checklist_df["점검결과"] == "확인 필요"]
    if problems.empty:
        print("  해당 없음")
    else:
        for _, r in problems.iterrows():
            print(f"  - {r['계약명']}({r['공급업체']}) - {r['작업명']}: {r['결과메모']}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--contracts-input", default=HERE / "sample_data" / "contracts.csv")
    parser.add_argument("--wbs-input", default=HERE / "sample_data" / "wbs_tasks.csv")
    parser.add_argument("--as-of", default=None, help="YYYY-MM-DD (기본값: 오늘)")
    parser.add_argument("--delay-threshold-days", type=int, default=DELAY_THRESHOLD_DAYS_DEFAULT)
    args = parser.parse_args()

    as_of = pd.Timestamp(args.as_of) if args.as_of else pd.Timestamp.today().normalize()

    contracts = load_contracts(args.contracts_input)
    wbs = load_wbs(args.wbs_input)
    wbs = compute_task_status(wbs, as_of, args.delay_threshold_days)

    progress_df = build_contract_progress(contracts, wbs)
    delay_df = build_delay_list(wbs, contracts, as_of)
    expiry_df = build_expiry_alerts(contracts, as_of)
    checklist_df = build_guarantee_cert_checklist(wbs, contracts)
    notify_result = notify_expiry(expiry_df, as_of)

    excel_path = save_excel_report({
        "계약_진행현황": progress_df,
        "지연_태스크": delay_df,
        "계약만료_임박": expiry_df,
        "보증_실적증명서_체크리스트": checklist_df,
    }, HERE / "output" / "contract_wbs_report.xlsx")

    gantt_path = plot_contract_gantt(contracts, wbs, as_of, HERE / "output" / "contract_gantt.png")

    print_summary(progress_df, delay_df, expiry_df, checklist_df, notify_result, as_of)
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"간트차트 저장: {gantt_path}")


if __name__ == "__main__":
    main()
