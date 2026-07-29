"""
income_statement_generator.py
경영지원팀 회계/재무 업무 예제: 월별 거래 내역 엑셀을 읽어 간이 손익계산서와
매출/영업이익 추이 차트를 자동 생성한다.

사용법:
    python finance/income_statement_generator.py
    python finance/income_statement_generator.py --input 다른파일.xlsx

입력: finance/sample_data/sample_transactions.xlsx
      (컬럼: 거래월, 계정과목, 구분, 금액)
출력: finance/output/income_statement.xlsx  (월별 손익계산서)
      finance/output/revenue_trend.png      (매출/영업이익 추이 차트)
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

# 레포 루트를 path에 추가해서 common 패키지를 어디서 실행하든 import 가능하게 함
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.chart_style import PALETTE, save_chart, setup_style  # noqa: E402
from common.excel_io import load_excel, save_excel_report  # noqa: E402
from common.format_utils import krw, print_section  # noqa: E402

HERE = Path(__file__).resolve().parent

# 계정과목을 손익계산서 항목으로 묶는 매핑.
# 회사마다 계정과목 체계가 다르므로, 이 부분만 바꾸면 다른 회사 데이터에도 재사용 가능하다.
COGS_ACCOUNTS = {"매출원가"}
OPEX_ACCOUNTS = {"급여", "임차료", "광고선전비", "소모품비", "지급수수료"}
REVENUE_ACCOUNTS = {"매출"}


def build_income_statement(df: pd.DataFrame) -> pd.DataFrame:
    """거래 내역(long format)을 월별 손익계산서(wide format)로 변환."""
    pivot = df.pivot_table(index="계정과목", columns="거래월", values="금액", aggfunc="sum").fillna(0)

    months = sorted(pivot.columns)
    statement = pd.DataFrame(index=["매출", "매출원가", "매출총이익", "판매비와관리비", "영업이익"], columns=months)

    for m in months:
        revenue = pivot.loc[pivot.index.isin(REVENUE_ACCOUNTS), m].sum()
        cogs = pivot.loc[pivot.index.isin(COGS_ACCOUNTS), m].sum()
        opex = pivot.loc[pivot.index.isin(OPEX_ACCOUNTS), m].sum()
        gross_profit = revenue - cogs
        operating_income = gross_profit - opex

        statement.loc["매출", m] = revenue
        statement.loc["매출원가", m] = cogs
        statement.loc["매출총이익", m] = gross_profit
        statement.loc["판매비와관리비", m] = opex
        statement.loc["영업이익", m] = operating_income

    statement.index.name = "항목"
    return statement.reset_index()


def plot_trend(statement: pd.DataFrame, output_path: Path) -> Path:
    """매출 / 영업이익 추이를 꺾은선 그래프로 저장."""
    setup_style()
    months = statement.columns[1:]
    revenue = statement.loc[statement["항목"] == "매출", months].iloc[0]
    operating_income = statement.loc[statement["항목"] == "영업이익", months].iloc[0]

    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.plot(months, revenue, marker="o", color=PALETTE[0], label="매출")
    ax.plot(months, operating_income, marker="o", color=PALETTE[1], label="영업이익")
    ax.set_title("월별 매출 / 영업이익 추이")
    ax.set_ylabel("금액 (원)")
    ax.yaxis.set_major_formatter(lambda x, _: f"{x/1e8:.1f}억")
    ax.legend()

    return save_chart(fig, output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="월별 거래 내역으로 간이 손익계산서를 생성한다.")
    parser.add_argument(
        "--input", default=HERE / "sample_data" / "sample_transactions.xlsx",
        help="거래 내역 엑셀 파일 경로 (컬럼: 거래월, 계정과목, 구분, 금액)",
    )
    args = parser.parse_args()

    df = load_excel(args.input)
    statement = build_income_statement(df)

    excel_path = save_excel_report({"손익계산서": statement}, HERE / "output" / "income_statement.xlsx")
    chart_path = plot_trend(statement, HERE / "output" / "revenue_trend.png")

    print_section("월별 손익계산서 (요약)")
    print(statement.to_string(index=False, formatters={c: krw for c in statement.columns[1:]}))

    latest_month = statement.columns[-1]
    latest_operating_income = statement.loc[statement["항목"] == "영업이익", latest_month].iloc[0]
    print(f"\n최근월({latest_month}) 영업이익: {krw(latest_operating_income)}")
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {chart_path}")


if __name__ == "__main__":
    main()
