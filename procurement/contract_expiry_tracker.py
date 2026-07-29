"""
contract_expiry_tracker.py
경영지원팀 구매/총무 업무 예제: 계약 현황 데이터에서 만료 임박 계약을 찾아
색상으로 강조한 엑셀 리포트와 D-day 차트를 자동 생성한다.

사용법:
    python procurement/contract_expiry_tracker.py
    python procurement/contract_expiry_tracker.py --as-of 2026-08-15 --warn-days 60

입력: procurement/sample_data/sample_contracts.csv
      (컬럼: 계약번호, 거래처, 계약구분, 시작일, 종료일, 계약금액, 담당자)
출력: procurement/output/contract_expiry_report.xlsx (만료 임박순 정렬 + 색상 강조)
      procurement/output/expiry_dday_chart.png       (계약별 D-day 차트)
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.chart_style import save_chart, setup_style  # noqa: E402
from common.excel_io import load_csv  # noqa: E402
from common.format_utils import krw, print_section  # noqa: E402

HERE = Path(__file__).resolve().parent


def _status(days_left: int, warn_days: int) -> str:
    if days_left < 0:
        return "만료됨"
    if days_left <= warn_days:
        return "임박"
    if days_left <= warn_days * 2:
        return "주의"
    return "여유"


def build_report(df: pd.DataFrame, as_of: date, warn_days: int) -> pd.DataFrame:
    df = df.copy()
    df["종료일"] = pd.to_datetime(df["종료일"]).dt.date
    df["D-day"] = df["종료일"].apply(lambda d: (d - as_of).days)
    df["상태"] = df["D-day"].apply(lambda d: _status(d, warn_days))
    return df.sort_values("D-day")


def plot_dday(report: pd.DataFrame, output_path: Path) -> Path:
    setup_style()
    color_map = {"만료됨": "#B0B0B0", "임박": "#E4572E", "주의": "#F4A300", "여유": "#76B041"}
    colors = report["상태"].map(color_map)

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.barh(report["거래처"], report["D-day"], color=colors)
    ax.axvline(0, color="black", linewidth=0.8)
    ax.set_title("계약별 만료까지 남은 일수 (D-day)")
    ax.set_xlabel("D-day")
    ax.invert_yaxis()
    return save_chart(fig, output_path)


def save_highlighted_report(report: pd.DataFrame, output_path: Path) -> Path:
    """상태(임박/만료됨)에 따라 행 색상을 다르게 칠한 엑셀 리포트 저장."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        cols = ["계약번호", "거래처", "계약구분", "시작일", "종료일", "계약금액", "담당자", "D-day", "상태"]
        report[cols].to_excel(writer, sheet_name="계약만료현황", index=False)

        workbook = writer.book
        worksheet = writer.sheets["계약만료현황"]

        fmt_expired = workbook.add_format({"bg_color": "#EAEAEA"})
        fmt_urgent = workbook.add_format({"bg_color": "#FADBD1", "bold": True})
        fmt_watch = workbook.add_format({"bg_color": "#FDEBCB"})

        status_col = cols.index("상태")
        for row_idx, status in enumerate(report["상태"], start=1):  # +1: 헤더행
            if status == "만료됨":
                worksheet.set_row(row_idx, cell_format=fmt_expired)
            elif status == "임박":
                worksheet.set_row(row_idx, cell_format=fmt_urgent)
            elif status == "주의":
                worksheet.set_row(row_idx, cell_format=fmt_watch)

        for col_idx, col in enumerate(cols):
            max_len = max([len(str(col))] + [len(str(v)) for v in report[col].astype(str)])
            worksheet.set_column(col_idx, col_idx, min(max_len + 2, 40))

    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="계약 데이터에서 만료 임박 건을 찾아 리포트를 생성한다.")
    parser.add_argument(
        "--input", default=HERE / "sample_data" / "sample_contracts.csv",
        help="계약 현황 CSV 경로 (컬럼: 계약번호, 거래처, 계약구분, 시작일, 종료일, 계약금액, 담당자)",
    )
    parser.add_argument("--as-of", default=None, help="기준일 (YYYY-MM-DD), 기본값: 오늘")
    parser.add_argument("--warn-days", type=int, default=90, help="며칠 이내를 '임박'으로 볼지 (기본 90일)")
    args = parser.parse_args()

    as_of = datetime.strptime(args.as_of, "%Y-%m-%d").date() if args.as_of else date.today()

    df = load_csv(args.input)
    report = build_report(df, as_of, args.warn_days)

    excel_path = save_highlighted_report(report, HERE / "output" / "contract_expiry_report.xlsx")
    chart_path = plot_dday(report, HERE / "output" / "expiry_dday_chart.png")

    print_section(f"계약 만료 현황 (기준일: {as_of})")
    urgent = report[report["상태"].isin(["임박", "만료됨"])]
    if urgent.empty:
        print(f"{args.warn_days}일 이내 만료 예정 계약이 없습니다.")
    else:
        for _, row in urgent.iterrows():
            print(f"[{row['상태']}] {row['거래처']:12s} {row['계약구분']:6s} "
                  f"종료일 {row['종료일']} (D{row['D-day']:+d})  계약금액 {krw(row['계약금액'])}")

    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {chart_path}")


if __name__ == "__main__":
    main()
