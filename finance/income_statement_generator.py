"""
income_statement_generator.py
K-IFRS 다단계 손익계산서 + 기간(년/반기/분기/월/주) 자유 비교 + 구글시트/로컬 겸용.
사용법:
    python finance/income_statement_generator.py
    python finance/income_statement_generator.py --period-type quarter --period 2026-Q2
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
from common.format_utils import krw, pct, print_section, signed_krw  # noqa: E402
from common.period_utils import (  # noqa: E402
    add_period_label, available_periods, compute_change,
    latest_period_label, previous_period_label, year_over_year_label,
)
from common.sheet_io import is_configured as sheet_is_configured  # noqa: E402
from common.sheet_io import load_transactions_from_sheet  # noqa: E402
from common.tax_utils import estimate_corporate_tax, estimate_vat_reference  # noqa: E402

HERE = Path(__file__).resolve().parent

REVENUE_ACCOUNTS = {"매출"}
COGS_ACCOUNTS = {"매출원가"}
OPEX_ACCOUNTS = {"급여", "임차료", "광고선전비", "소모품비", "지급수수료"}
NON_OP_INCOME_ACCOUNTS = {"이자수익"}
NON_OP_EXPENSE_ACCOUNTS = {"이자비용", "잡손실"}
VAT_DEDUCTIBLE_EXPENSE_ACCOUNTS = (COGS_ACCOUNTS | OPEX_ACCOUNTS) - {"급여"}
PERIODS_PER_YEAR = {"year": 1, "half": 2, "quarter": 4, "month": 12, "week": 52}
STATEMENT_LINES = [
    "매출액", "매출원가", "매출총이익", "판매비와관리비", "영업이익",
    "영업외수익", "영업외비용", "법인세비용차감전순이익", "법인세비용(추정)", "당기순이익",
]


def load_transactions(input_path: str | Path) -> pd.DataFrame:
    if sheet_is_configured():
        df = load_transactions_from_sheet()
    else:
        df = load_csv(input_path)
    df["금액"] = pd.to_numeric(df["금액"], errors="coerce")
    df["거래일자"] = pd.to_datetime(df["거래일자"])
    return df


def _period_fiscal_year(period_label: str, period_type: str) -> int:
    return int(period_label[:4])


def build_statement(df, period_col, period_label, period_type, include_local_surtax=True) -> dict[str, float]:
    period_df = df[df[period_col] == period_label]
    revenue = period_df.loc[period_df["계정과목"].isin(REVENUE_ACCOUNTS), "금액"].sum()
    cogs = period_df.loc[period_df["계정과목"].isin(COGS_ACCOUNTS), "금액"].sum()
    opex = period_df.loc[period_df["계정과목"].isin(OPEX_ACCOUNTS), "금액"].sum()
    non_op_income = period_df.loc[period_df["계정과목"].isin(NON_OP_INCOME_ACCOUNTS), "금액"].sum()
    non_op_expense = period_df.loc[period_df["계정과목"].isin(NON_OP_EXPENSE_ACCOUNTS), "금액"].sum()

    gross_profit = revenue - cogs
    operating_income = gross_profit - opex
    pretax_income = operating_income + non_op_income - non_op_expense

    factor = PERIODS_PER_YEAR[period_type]
    fiscal_year = _period_fiscal_year(period_label, period_type)
    annualized_pretax = pretax_income * factor
    annual_tax = estimate_corporate_tax(annualized_pretax, fiscal_year, include_local_surtax)
    effective_rate = (annual_tax / annualized_pretax) if annualized_pretax > 0 else 0.0
    tax_expense = max(pretax_income, 0) * effective_rate
    net_income = pretax_income - tax_expense

    return {
        "매출액": revenue, "매출원가": cogs, "매출총이익": gross_profit, "판매비와관리비": opex,
        "영업이익": operating_income, "영업외수익": non_op_income, "영업외비용": non_op_expense,
        "법인세비용차감전순이익": pretax_income, "법인세비용(추정)": tax_expense, "당기순이익": net_income,
    }


def build_comparison_table(df, period_col, period_label, period_type) -> pd.DataFrame:
    existing_periods = set(available_periods(df, period_col))
    prev_label = previous_period_label(period_label, period_type)
    yoy_label = year_over_year_label(period_label, period_type)
    prev_label = prev_label if prev_label in existing_periods else None
    yoy_label = yoy_label if (yoy_label and yoy_label in existing_periods) else None

    current = build_statement(df, period_col, period_label, period_type)
    previous = build_statement(df, period_col, prev_label, period_type) if prev_label else None
    yoy = build_statement(df, period_col, yoy_label, period_type) if yoy_label else None

    rows = []
    for line in STATEMENT_LINES:
        cur_val = current[line]
        prev_val = previous[line] if previous else None
        yoy_val = yoy[line] if yoy else None
        prev_diff, prev_pct = compute_change(cur_val, prev_val)
        yoy_diff, yoy_pct = compute_change(cur_val, yoy_val)
        rows.append({
            "항목": line, f"당기({period_label})": cur_val, f"직전기({prev_label or '-'})": prev_val,
            "직전기 증감액": prev_diff if previous else None, "직전기 증감률": prev_pct,
            f"전년동기({yoy_label or '-'})": yoy_val, "전년동기 증감액": yoy_diff if yoy else None,
            "전년동기 증감률": yoy_pct,
        })
    return pd.DataFrame(rows)


def plot_trend(df, period_col, period_type, output_path, max_periods=12) -> Path:
    setup_style()
    periods = available_periods(df, period_col)[-max_periods:]
    series = {"매출액": [], "영업이익": [], "당기순이익": []}
    for p in periods:
        stmt = build_statement(df, period_col, p, period_type)
        series["매출액"].append(stmt["매출액"])
        series["영업이익"].append(stmt["영업이익"])
        series["당기순이익"].append(stmt["당기순이익"])

    fig, ax = plt.subplots(figsize=(9, 4.8))
    for i, (name, values) in enumerate(series.items()):
        ax.plot(periods, values, marker="o", label=name, color=PALETTE[i % len(PALETTE)])
    ax.set_title(f"기간별({period_type}) 매출·영업이익·당기순이익 추이")
    ax.set_ylabel("금액 (원)")
    ax.yaxis.set_major_formatter(lambda x, _: f"{x/1e8:.1f}억")
    ax.tick_params(axis="x", rotation=45)
    ax.legend()
    return save_chart(fig, output_path)


def print_statement(table, period_label) -> None:
    print_section(f"K-IFRS 손익계산서 ({period_label} 기준)")
    cur_col = [c for c in table.columns if c.startswith("당기")][0]
    prev_col = [c for c in table.columns if c.startswith("직전기(")][0]
    yoy_col = [c for c in table.columns if c.startswith("전년동기(")][0]
    for _, row in table.iterrows():
        line = row["항목"]
        cur = krw(row[cur_col])
        prev = f"{krw(row[prev_col])} ({signed_krw(row['직전기 증감액'])}, {pct(row['직전기 증감률'])})" if pd.notna(row[prev_col]) else "데이터 없음"
        yoy = f"{krw(row[yoy_col])} ({signed_krw(row['전년동기 증감액'])}, {pct(row['전년동기 증감률'])})" if pd.notna(row[yoy_col]) else "데이터 없음"
        print(f"{line:16s} {cur:>16s}  |  직전기 {prev}  |  전년동기 {yoy}")
    print("\n※ 법인세비용은 추정치이며 실제 신고세액과 다를 수 있습니다.")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=HERE / "sample_data" / "sample_transactions.csv")
    parser.add_argument("--period-type", choices=["year", "half", "quarter", "month", "week"], default="month")
    parser.add_argument("--period", default=None)
    parser.add_argument("--no-local-surtax", action="store_true")
    args = parser.parse_args()

    df = load_transactions(args.input)
    df = add_period_label(df, "거래일자", args.period_type)
    period_label = args.period or latest_period_label(df)
    if period_label not in set(available_periods(df)):
        raise SystemExit(f"'{period_label}' 기간 데이터가 없습니다. 사용 가능: {available_periods(df)}")

    table = build_comparison_table(df, "기간", period_label, args.period_type)
    period_df = df[df["기간"] == period_label]
    taxable_revenue = period_df.loc[period_df["계정과목"].isin(REVENUE_ACCOUNTS), "금액"].sum()
    deductible_purchases = period_df.loc[period_df["계정과목"].isin(VAT_DEDUCTIBLE_EXPENSE_ACCOUNTS), "금액"].sum()
    vat_ref = estimate_vat_reference(taxable_revenue, deductible_purchases)
    vat_df = pd.DataFrame([
        {"구분": "매출세액", "금액": vat_ref["매출세액"]},
        {"구분": "매입세액", "금액": vat_ref["매입세액"]},
        {"구분": "납부예상세액", "금액": vat_ref["납부예상세액"]},
    ])

    excel_path = save_excel_report({"손익계산서": table, "부가세_참고": vat_df}, HERE / "output" / "income_statement.xlsx")
    chart_path = plot_trend(df, "기간", args.period_type, HERE / "output" / "trend_chart.png")
    print_statement(table, period_label)
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {chart_path}")


if __name__ == "__main__":
    main()
