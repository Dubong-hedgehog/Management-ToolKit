"""
safety_training_tracker.py
산업안전보건법상 법정의무교육(정기 안전보건교육, 관리감독자 안전보건교육
등) 이수현황을 관리하는 도구. 개인별/부서별 이수율을 보고, 이수기한이
지났는데도 아직 안 들은 사람을 놓치지 않게 잡아준다.

법정의무교육은 사업주의 법적 의무 사항이라(산업안전보건법 제29조 등) 미이수
적발 시 과태료 대상이 될 수 있다 — "그냥 권장사항"이 아니라는 점에서
HR의 지각/야근 이상탐지보다 더 엄격하게(기한 경과=즉시 위반 상태) 다룬다.

Usage:
    python general_affairs/safety_training_tracker.py
    python general_affairs/safety_training_tracker.py --as-of 2026-07-29

Input:
    general_affairs/sample_data/training_records.csv
        사번, 이름, 부서, 교육명, 대상기간(분기 또는 연도), 이수기한,
        이수일(미이수면 빈칸), 이수여부(Y/N), 이수시간

Output:
    general_affairs/output/safety_training_report.xlsx
        - 부서별_이수율 / 개인별_이수현황 / 미이수_명단
    general_affairs/output/training_completion_by_dept.png
    general_affairs/output/notify_log.csv (알림 채널 미설정 시 대체 기록)
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
DEADLINE_SOON_DAYS = 14  # 이수기한이 이 안으로 남으면 "마감임박"


def load_training(path) -> pd.DataFrame:
    df = load_csv(path)
    df["이수기한"] = pd.to_datetime(df["이수기한"])
    df["이수일"] = pd.to_datetime(df["이수일"], errors="coerce")
    return df


def build_dept_completion(df: pd.DataFrame) -> pd.DataFrame:
    agg = df.groupby("부서").agg(
        대상건수=("이수여부", "count"),
        이수건수=("이수여부", lambda s: (s == "Y").sum()),
    ).reset_index()
    agg["이수율"] = (agg["이수건수"] / agg["대상건수"]).round(3)
    return agg.sort_values("이수율")


def build_person_completion(df: pd.DataFrame) -> pd.DataFrame:
    agg = df.groupby(["사번", "이름", "부서"]).agg(
        대상건수=("이수여부", "count"),
        이수건수=("이수여부", lambda s: (s == "Y").sum()),
        미이수건수=("이수여부", lambda s: (s == "N").sum()),
    ).reset_index()
    agg["이수율"] = (agg["이수건수"] / agg["대상건수"]).round(3)
    return agg.sort_values("이수율")


def build_incomplete_list(df: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    missing = df[df["이수여부"] == "N"].copy()
    missing["기한경과일수"] = (as_of - missing["이수기한"]).dt.days

    def _status(days):
        if days > 0:
            return "기한경과(위반 상태)"
        if days >= -DEADLINE_SOON_DAYS:
            return "마감임박"
        return "예정"

    missing["상태"] = missing["기한경과일수"].map(_status)
    return missing[[
        "사번", "이름", "부서", "교육명", "대상기간", "이수기한", "기한경과일수", "상태",
    ]].sort_values("기한경과일수", ascending=False)


def notify_overdue(missing: pd.DataFrame, as_of: pd.Timestamp) -> str:
    overdue = missing[missing["상태"] == "기한경과(위반 상태)"]
    if overdue.empty:
        return "기한 경과 미이수 없음 - 알림 미발송"
    lines = [f"{as_of.date()} 기준 법정의무교육 기한경과 미이수자 알림 ({len(overdue)}건)"]
    for _, r in overdue.iterrows():
        lines.append(f"- {r['이름']}({r['부서']}) {r['교육명']} [{r['대상기간']}] 기한 {r['이수기한'].date()} 대비 {r['기한경과일수']}일 경과")
    return send_alert(
        subject=f"[법정의무교육 미이수] {len(overdue)}건 확인 필요",
        body="\n".join(lines),
        fallback_log_path=HERE / "output" / "notify_log.csv",
    )


def plot_completion(dept_df: pd.DataFrame, df: pd.DataFrame, output_path) -> Path:
    setup_style()
    fig, axes = plt.subplots(1, 2, figsize=(13, 5.5))

    colors = ["#C0102A" if v < 0.8 else "#2E5EAA" for v in dept_df["이수율"]]
    axes[0].barh(dept_df["부서"], dept_df["이수율"] * 100, color=colors)
    axes[0].axvline(80, color="#F2B705", linestyle="--", linewidth=1, label="80% 기준선")
    axes[0].set_title("부서별 법정의무교육 이수율")
    axes[0].set_xlabel("이수율 (%)")
    axes[0].legend(fontsize=8)
    for i, v in enumerate(dept_df["이수율"] * 100):
        axes[0].text(v + 1, i, f"{v:.0f}%", va="center", fontsize=9)

    q = df[df["교육명"] == "정기 안전보건교육"].copy()
    trend = q.groupby("대상기간")["이수여부"].apply(lambda s: (s == "Y").mean() * 100)
    axes[1].plot(trend.index, trend.values, marker="o", color=PALETTE[0])
    axes[1].axhline(80, color="#F2B705", linestyle="--", linewidth=1)
    axes[1].set_title("분기별 정기 안전보건교육 이수율 추이")
    axes[1].set_ylabel("이수율 (%)")
    axes[1].tick_params(axis="x", rotation=30)

    fig.tight_layout()
    return save_chart(fig, output_path)


def print_summary(dept_df, missing, notify_result, as_of) -> None:
    print_section(f"법정의무교육 이수현황 ({as_of.date()} 기준)")

    print("\n[1] 부서별 이수율")
    for _, r in dept_df.iterrows():
        flag = " (80% 미달)" if r["이수율"] < 0.8 else ""
        print(f"  - {r['부서']}: {r['이수율']:.0%} ({r['이수건수']}/{r['대상건수']}){flag}")

    print("\n[2] 미이수 현황")
    if missing.empty:
        print("  전원 이수 완료")
    else:
        for _, r in missing.iterrows():
            print(f"  - [{r['상태']}] {r['이름']}({r['부서']}) {r['교육명']}[{r['대상기간']}] 기한 {r['이수기한'].date()}")
    print(f"  알림 발송 결과: {notify_result}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=HERE / "sample_data" / "training_records.csv")
    parser.add_argument("--as-of", default=None, help="YYYY-MM-DD (기본값: 오늘)")
    args = parser.parse_args()

    as_of = pd.Timestamp(args.as_of) if args.as_of else pd.Timestamp.today().normalize()

    df = load_training(args.input)
    dept_df = build_dept_completion(df)
    person_df = build_person_completion(df)
    missing = build_incomplete_list(df, as_of)
    notify_result = notify_overdue(missing, as_of)

    excel_path = save_excel_report({
        "부서별_이수율": dept_df,
        "개인별_이수현황": person_df,
        "미이수_명단": missing,
    }, HERE / "output" / "safety_training_report.xlsx")

    chart_path = plot_completion(dept_df, df, HERE / "output" / "training_completion_by_dept.png")

    print_summary(dept_df, missing, notify_result, as_of)
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {chart_path}")


if __name__ == "__main__":
    main()
