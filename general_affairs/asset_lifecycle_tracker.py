"""
asset_lifecycle_tracker.py
비품/소모품 + 고정자산 + 법인차량을 하나로 묶어 "취득 -> 사용(승계 포함) ->
불용/폐기"까지 전 생애주기를 추적하는 도구. 셋 다 물리적 자산 하나가
태어나서 없어지기까지를 관리한다는 점에서 같은 패턴이라 통합했다
(CONVENTIONS.md 5번 원칙: 성격이 같으면 억지로 쪼개지 않는다).

정액법으로 감가상각 장부가액을 추정하고, 내용연수를 넘겼는데도 아직
'사용중'/'불용대기'인 자산을 폐기 검토 대상으로 잡아준다. 화재보험/배상책임
보험/자동차보험처럼 자산에 걸린 보험의 갱신일이 임박하면
common/notify_utils.py로 알림을 보낸다(계약만료 임박 알림과 동일한 패턴).

Usage:
    python general_affairs/asset_lifecycle_tracker.py
    python general_affairs/asset_lifecycle_tracker.py --as-of 2026-07-29

Input:
    general_affairs/sample_data/assets.csv
        자산번호, 자산명, 카테고리(비품/소모품/고정자산/법인차량), 부서, 사용자,
        직전사용자(승계 이력, 없으면 빈칸), 취득일, 취득가액, 내용연수(년),
        상태(사용중/불용대기/폐기완료), 폐기일, 보험종류, 보험갱신일, 담당자

Output:
    general_affairs/output/asset_report.xlsx
        - 자산_현황 / 폐기검토_대상 / 승계_이력 / 보험갱신_임박
    general_affairs/output/asset_by_category.png   (카테고리별 자산 현황)
    general_affairs/output/notify_log.csv           (알림 채널 미설정 시 대체 기록)
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.chart_style import PALETTE, save_chart, setup_style  # noqa: E402
from common.excel_io import load_csv, save_excel_report  # noqa: E402
from common.format_utils import print_section  # noqa: E402
from common.notify_utils import send_alert  # noqa: E402

HERE = Path(__file__).resolve().parent
INSURANCE_WARN_DAYS = 30
INSURANCE_ALERT_DAYS = 60


def load_assets(path) -> pd.DataFrame:
    df = load_csv(path)
    df["취득일"] = pd.to_datetime(df["취득일"])
    df["폐기일"] = pd.to_datetime(df["폐기일"], errors="coerce")
    df["보험갱신일"] = pd.to_datetime(df["보험갱신일"], errors="coerce")
    return df


def compute_asset_status(df: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    """자산별 경과연수/장부가액/폐기검토 필요 여부를 계산한다.

    감가상각은 정액법(잔존가치 0)으로 근사한다 - 실제 세법상 상각률/상각방법과
    다를 수 있어 참고용 추정치임을 명시한다.
    """
    d = df.copy()
    ref_date = d["폐기일"].fillna(as_of)
    d["경과연수"] = ((ref_date - d["취득일"]).dt.days / 365).round(2)
    d["상각률"] = (d["경과연수"] / d["내용연수"]).clip(upper=1.0)
    d["감가상각누계액"] = (d["취득가액"] * d["상각률"]).round(0)
    d["장부가액"] = (d["취득가액"] - d["감가상각누계액"]).round(0)
    d["내용연수초과"] = d["경과연수"] > d["내용연수"]
    d["폐기검토필요"] = d["내용연수초과"] & (d["상태"] != "폐기완료")
    d["승계이력있음"] = d["직전사용자"].fillna("") != ""
    return d


def build_disposal_review_list(d: pd.DataFrame) -> pd.DataFrame:
    review = d[d["폐기검토필요"]].copy()
    review["초과연수"] = (review["경과연수"] - review["내용연수"]).round(1)
    return review[[
        "자산번호", "자산명", "카테고리", "부서", "사용자", "취득일", "내용연수", "경과연수", "초과연수", "장부가액", "상태",
    ]].sort_values("초과연수", ascending=False)


def build_succession_list(d: pd.DataFrame) -> pd.DataFrame:
    succ = d[d["승계이력있음"]].copy()
    return succ[["자산번호", "자산명", "카테고리", "부서", "직전사용자", "사용자", "담당자"]]


def build_insurance_alerts(d: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    ins = d[d["보험종류"].fillna("") != ""].copy()
    ins = ins[ins["보험갱신일"].notna()]
    ins["잔여일수"] = (ins["보험갱신일"] - as_of).dt.days
    ins = ins[ins["잔여일수"] <= INSURANCE_ALERT_DAYS]

    def _bucket(days):
        if days < 0:
            return "만료됨(즉시 확인 필요)"
        return f"D-{INSURANCE_WARN_DAYS} 이내(긴급)" if days <= INSURANCE_WARN_DAYS else f"D-{INSURANCE_ALERT_DAYS} 이내(예고)"

    ins["긴급도"] = ins["잔여일수"].map(_bucket)
    return ins[["자산번호", "자산명", "카테고리", "부서", "보험종류", "보험갱신일", "잔여일수", "긴급도"]].sort_values("잔여일수")


def notify_insurance(alerts: pd.DataFrame, as_of: pd.Timestamp) -> str:
    if alerts.empty:
        return "보험 갱신 임박 건 없음 - 알림 미발송"
    lines = [f"{as_of.date()} 기준 보험 갱신 임박 알림 ({len(alerts)}건)"]
    for _, r in alerts.iterrows():
        dday = f"D-{r['잔여일수']}" if r["잔여일수"] >= 0 else f"D+{-r['잔여일수']} 경과"
        lines.append(f"- [{r['긴급도']}] {r['자산명']}({r['자산번호']}, {r['부서']}) {r['보험종류']} 갱신일 {r['보험갱신일'].date()} ({dday})")
    return send_alert(
        subject=f"[보험 갱신 임박] {len(alerts)}건 확인 필요",
        body="\n".join(lines),
        fallback_log_path=HERE / "output" / "notify_log.csv",
    )


def plot_asset_by_category(d: pd.DataFrame, output_path) -> Path:
    setup_style()
    fig, axes = plt.subplots(1, 2, figsize=(13, 5.5))

    by_cat_value = d.groupby("카테고리")["장부가액"].sum().sort_values(ascending=False)
    axes[0].bar(by_cat_value.index, by_cat_value.values, color=PALETTE[: len(by_cat_value)])
    axes[0].set_title("카테고리별 장부가액 합계")
    axes[0].set_ylabel("장부가액 (원)")
    axes[0].tick_params(axis="x", rotation=20)

    status_ct = d.groupby(["카테고리", "상태"]).size().unstack(fill_value=0)
    status_ct = status_ct.reindex(columns=["사용중", "불용대기", "폐기완료"], fill_value=0)
    bottom = None
    colors = {"사용중": "#2E5EAA", "불용대기": "#F2B705", "폐기완료": "#999999"}
    for status in status_ct.columns:
        axes[1].bar(status_ct.index, status_ct[status], bottom=bottom, label=status, color=colors[status])
        bottom = status_ct[status] if bottom is None else bottom + status_ct[status]
    axes[1].set_title("카테고리별 상태 분포")
    axes[1].set_ylabel("자산 수")
    axes[1].tick_params(axis="x", rotation=20)
    axes[1].legend(fontsize=8)

    fig.tight_layout()
    return save_chart(fig, output_path)


def print_summary(d, review, succ, alerts, notify_result, as_of) -> None:
    print_section(f"자산 생애주기 현황 ({as_of.date()} 기준)")
    print(f"\n총 자산: {len(d)}건 (사용중 {(d['상태']=='사용중').sum()} / 불용대기 {(d['상태']=='불용대기').sum()} / 폐기완료 {(d['상태']=='폐기완료').sum()})")

    print("\n[1] 폐기 검토 대상 (내용연수 초과)")
    if review.empty:
        print("  해당 없음")
    else:
        for _, r in review.iterrows():
            print(f"  - {r['자산명']}({r['자산번호']}, {r['부서']}): 내용연수 {r['내용연수']}년, 경과 {r['경과연수']}년 (초과 {r['초과연수']}년)")

    print("\n[2] 승계 이력")
    if succ.empty:
        print("  해당 없음")
    else:
        for _, r in succ.iterrows():
            print(f"  - {r['자산명']}({r['자산번호']}): {r['직전사용자']} -> {r['사용자']}")

    print("\n[3] 보험 갱신 임박")
    if alerts.empty:
        print("  해당 없음")
    else:
        for _, r in alerts.iterrows():
            dday = f"D-{r['잔여일수']}" if r["잔여일수"] >= 0 else f"D+{-r['잔여일수']} 경과"
            print(f"  - [{r['긴급도']}] {r['자산명']}({r['자산번호']}) {r['보험종류']} 갱신일 {r['보험갱신일'].date()} ({dday})")
    print(f"  알림 발송 결과: {notify_result}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=HERE / "sample_data" / "assets.csv")
    parser.add_argument("--as-of", default=None, help="YYYY-MM-DD (기본값: 오늘)")
    args = parser.parse_args()

    as_of = pd.Timestamp(args.as_of) if args.as_of else pd.Timestamp.today().normalize()

    raw = load_assets(args.input)
    d = compute_asset_status(raw, as_of)
    review = build_disposal_review_list(d)
    succ = build_succession_list(d)
    alerts = build_insurance_alerts(d, as_of)
    notify_result = notify_insurance(alerts, as_of)

    excel_path = save_excel_report({
        "자산_현황": d.drop(columns=["내용연수초과", "폐기검토필요", "승계이력있음"]),
        "폐기검토_대상": review,
        "승계_이력": succ,
        "보험갱신_임박": alerts,
    }, HERE / "output" / "asset_report.xlsx")

    chart_path = plot_asset_by_category(d, HERE / "output" / "asset_by_category.png")

    print_summary(d, review, succ, alerts, notify_result, as_of)
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {chart_path}")


if __name__ == "__main__":
    main()
