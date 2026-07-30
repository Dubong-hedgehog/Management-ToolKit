"""
general_affairs_ledger.py
총무 업무 중 "정교한 이상탐지"보다는 정리·조회가 더 중요한 6개 영역을
하나의 관리대장으로 묶어서 관리하는 도구.

    - 업무환경(청소·미화 점검) + 우편·물류(수발신 로그)  -> 일자별 처리 로그
    - 임대차 + 청소/경비/정수기 등 각종 용역계약           -> 계약 현황 + 만료임박 알림
    - 문서·인장 관리                                        -> 보존연한 관리 + 폐기대상 알림
    - 행사·의전관리 + 복리후생                              -> 예산/집행 대장

asset_lifecycle_tracker.py, safety_training_tracker.py처럼 이상탐지 로직을
따로 만들 만큼 데이터가 무겁지 않은 영역들이라, 각 영역을 시트로 나눠 한
파일로 관리하고, 계약 만료/문서 폐기처럼 "날짜가 지나면 챙겨야 하는" 항목만
공통으로 알림을 보낸다(계약/자산 모듈과 같은 notify_utils 패턴).

Usage:
    python general_affairs/general_affairs_ledger.py
    python general_affairs/general_affairs_ledger.py --as-of 2026-07-29

Input (모두 general_affairs/sample_data/):
    facility_ops_log.csv         구분, 일자, 항목명, 담당자, 상태, 비고
    facility_contracts.csv       계약명, 계약유형, 업체명, 계약시작일, 계약종료일, 월계약금액, 상태
    document_registry.csv        문서명, 문서유형, 관리번호, 등록일, 보존연한, 보존만료일, 처리상태, 인장사용여부
    event_and_welfare_ledger.csv 구분, 항목명, 일자, 부서/대상, 예산, 집행액, 비고

Output:
    general_affairs/output/general_affairs_report.xlsx
        - 업무환경_우편물류_로그 / 시설계약_현황 / 계약만료_임박 /
          문서_보존현황 / 문서_폐기대상 / 행사_복리후생_집행현황
    general_affairs/output/facility_contract_status.png
    general_affairs/output/event_welfare_budget.png
    general_affairs/output/notify_log.csv (알림 채널 미설정 시 대체 기록)
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.chart_style import PALETTE, save_chart, setup_style  # noqa: E402
from common.excel_io import load_csv, save_excel_report  # noqa: E402
from common.format_utils import print_section  # noqa: E402
from common.notify_utils import send_alert  # noqa: E402

HERE = Path(__file__).resolve().parent
CONTRACT_WARN_DAYS = 30
CONTRACT_ALERT_DAYS = 60


def _dday_bucket(days: int, warn=CONTRACT_WARN_DAYS, alert=CONTRACT_ALERT_DAYS) -> str:
    if days < 0:
        return "만료됨(즉시 확인 필요)"
    if days <= warn:
        return f"D-{warn} 이내(긴급)"
    return f"D-{alert} 이내(예고)"


# ------------------------------------------------------------------
# 시설계약 (임대차 + 각종 용역)
# ------------------------------------------------------------------
def load_contracts(path) -> pd.DataFrame:
    df = load_csv(path)
    df["계약시작일"] = pd.to_datetime(df["계약시작일"])
    df["계약종료일"] = pd.to_datetime(df["계약종료일"])
    return df


def build_contract_expiry(contracts: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    d = contracts.copy()
    d["잔여일수"] = (d["계약종료일"] - as_of).dt.days
    d = d[d["잔여일수"] <= CONTRACT_ALERT_DAYS]
    d["긴급도"] = d["잔여일수"].map(_dday_bucket)
    return d[["계약명", "계약유형", "업체명", "계약종료일", "월계약금액", "잔여일수", "긴급도"]].sort_values("잔여일수")


def notify_contract_expiry(expiry: pd.DataFrame, as_of: pd.Timestamp) -> str:
    if expiry.empty:
        return "만료 임박 시설계약 없음 - 알림 미발송"
    lines = [f"{as_of.date()} 기준 시설계약 만료 임박 알림 ({len(expiry)}건)"]
    for _, r in expiry.iterrows():
        lines.append(f"- [{r['긴급도']}] {r['계약명']}({r['업체명']}) 종료일 {r['계약종료일'].date()}")
    return send_alert(
        subject=f"[시설계약 만료 임박] {len(expiry)}건 확인 필요",
        body="\n".join(lines),
        fallback_log_path=HERE / "output" / "notify_log.csv",
    )


# ------------------------------------------------------------------
# 문서·인장 관리
# ------------------------------------------------------------------
def load_documents(path) -> pd.DataFrame:
    df = load_csv(path)
    df["등록일"] = pd.to_datetime(df["등록일"])
    df["보존만료일"] = pd.to_datetime(df["보존만료일"])
    return df


def build_disposal_targets(documents: pd.DataFrame) -> pd.DataFrame:
    targets = documents[documents["처리상태"] == "폐기대상"].copy()
    return targets[["문서명", "문서유형", "관리번호", "등록일", "보존연한", "보존만료일", "인장사용여부"]]


# ------------------------------------------------------------------
# 행사·의전 + 복리후생
# ------------------------------------------------------------------
def load_ledger(path) -> pd.DataFrame:
    df = load_csv(path)
    df["일자"] = pd.to_datetime(df["일자"])
    df["집행률"] = (df["집행액"] / df["예산"]).round(3).where(df["예산"] > 0)
    return df


# ------------------------------------------------------------------
# 차트
# ------------------------------------------------------------------
def plot_contract_status(contracts: pd.DataFrame, as_of: pd.Timestamp, output_path) -> Path:
    setup_style()
    d = contracts.sort_values("계약종료일")
    fig, ax = plt.subplots(figsize=(11, max(3, 0.5 * len(d) + 1)))
    for i, (_, r) in enumerate(d.iterrows()):
        start_num = mdates.date2num(r["계약시작일"])
        end_num = mdates.date2num(r["계약종료일"])
        soon = (r["계약종료일"] - as_of).days <= CONTRACT_ALERT_DAYS
        color = "#C0102A" if soon else "#2E5EAA"
        ax.broken_barh([(start_num, end_num - start_num)], (i - 0.4, 0.8), facecolors=color, alpha=0.85)
        ax.text(end_num + 3, i, f"{r['계약명']} ({r['계약유형']})", va="center", fontsize=8)
    ax.axvline(mdates.date2num(as_of), color="black", linewidth=1, label="기준일(오늘)")
    ax.xaxis_date()
    ax.set_yticks(range(len(d)))
    ax.set_yticklabels(d["업체명"])
    ax.set_title(f"시설/용역 계약 현황 ({as_of.date()} 기준, 빨강=D-60 이내)")
    ax.legend(fontsize=8, loc="upper left")
    fig.tight_layout()
    return save_chart(fig, output_path)


def plot_event_welfare_budget(ledger: pd.DataFrame, output_path) -> Path:
    setup_style()
    fig, axes = plt.subplots(1, 2, figsize=(13, 5.5))

    events = ledger[ledger["구분"] == "사내행사"]
    x = range(len(events))
    axes[0].bar([i - 0.2 for i in x], events["예산"], width=0.4, label="예산", color=PALETTE[0])
    axes[0].bar([i + 0.2 for i in x], events["집행액"], width=0.4, label="집행액", color=PALETTE[1])
    axes[0].set_xticks(list(x))
    axes[0].set_xticklabels(events["항목명"], rotation=20, ha="right", fontsize=8)
    axes[0].set_title("사내행사 예산 대비 집행")
    axes[0].set_ylabel("금액 (원)")
    axes[0].legend(fontsize=8)

    by_type = ledger[ledger["구분"] != "사내행사"].groupby("구분")["집행액"].sum()
    axes[1].pie(by_type.values, labels=by_type.index, autopct="%1.0f%%", colors=PALETTE[2:4],
                textprops={"fontsize": 10})
    axes[1].set_title("경조사지원/복리후생 집행 비중")

    fig.tight_layout()
    return save_chart(fig, output_path)


def print_summary(expiry, disposal, ledger, notify_result, as_of) -> None:
    print_section(f"총무 관리대장 현황 ({as_of.date()} 기준)")

    print("\n[1] 시설/용역 계약 만료 임박")
    if expiry.empty:
        print("  해당 없음")
    else:
        for _, r in expiry.iterrows():
            print(f"  - [{r['긴급도']}] {r['계약명']}({r['업체명']}) 종료 {r['계약종료일'].date()}")
    print(f"  알림 발송 결과: {notify_result}")

    print("\n[2] 문서 폐기 대상 (보존연한 경과)")
    if disposal.empty:
        print("  해당 없음")
    else:
        for _, r in disposal.iterrows():
            print(f"  - {r['문서명']}({r['관리번호']}) 보존연한 {r['보존연한']}년 만료 {r['보존만료일'].date()}")

    print("\n[3] 행사/경조사/복리후생 집행 현황")
    for gubun, g in ledger.groupby("구분"):
        total_budget = g["예산"].sum()
        total_actual = g["집행액"].sum()
        print(f"  - {gubun}: {len(g)}건, 예산 {total_budget:,.0f}원 / 집행 {total_actual:,.0f}원")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ops-log-input", default=HERE / "sample_data" / "facility_ops_log.csv")
    parser.add_argument("--contracts-input", default=HERE / "sample_data" / "facility_contracts.csv")
    parser.add_argument("--documents-input", default=HERE / "sample_data" / "document_registry.csv")
    parser.add_argument("--ledger-input", default=HERE / "sample_data" / "event_and_welfare_ledger.csv")
    parser.add_argument("--as-of", default=None, help="YYYY-MM-DD (기본값: 오늘)")
    args = parser.parse_args()

    as_of = pd.Timestamp(args.as_of) if args.as_of else pd.Timestamp.today().normalize()

    ops_log = load_csv(args.ops_log_input)
    contracts = load_contracts(args.contracts_input)
    documents = load_documents(args.documents_input)
    ledger = load_ledger(args.ledger_input)

    expiry = build_contract_expiry(contracts, as_of)
    disposal = build_disposal_targets(documents)
    notify_result = notify_contract_expiry(expiry, as_of)

    excel_path = save_excel_report({
        "업무환경_우편물류_로그": ops_log,
        "시설계약_현황": contracts,
        "계약만료_임박": expiry,
        "문서_보존현황": documents,
        "문서_폐기대상": disposal,
        "행사_복리후생_집행현황": ledger,
    }, HERE / "output" / "general_affairs_report.xlsx")

    contract_chart = plot_contract_status(contracts, as_of, HERE / "output" / "facility_contract_status.png")
    budget_chart = plot_event_welfare_budget(ledger, HERE / "output" / "event_welfare_budget.png")

    print_summary(expiry, disposal, ledger, notify_result, as_of)
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {contract_chart}, {budget_chart}")


if __name__ == "__main__":
    main()
