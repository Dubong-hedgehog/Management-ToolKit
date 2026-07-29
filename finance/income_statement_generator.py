"""
income_statement_generator.py
경영지원팀 회계/재무 업무 예제 (v2): 거래 건별(전표 단위) 데이터를 읽어
K-IFRS 다단계 손익계산서를 홈택스/DART에서 보는 것과 동일한 형식으로 만들고,
원하는 기간 단위(년/반기/분기/월/주)로 직전기·전년동기와 비교한다.

데이터 소스는 두 가지 중 하나를 자동으로 고른다:
  1) .env에 구글 시트 정보가 설정돼 있으면 -> 구글 스프레드시트에서 실시간으로 읽음
  2) 아니면 -> finance/sample_data/sample_transactions.csv 같은 로컬 파일을 읽음
  (회사/계정이 바뀌면 .env만 새로 채우면 되고, 이 코드는 안 건드려도 됨)

사용법:
    python finance/income_statement_generator.py
    python finance/income_statement_generator.py --period-type month --period 2026-05
    python finance/income_statement_generator.py --period-type quarter --period 2026-Q2
    python finance/income_statement_generator.py --period-type week --period 2026-W18

입력 컬럼: 거래일자, 계정과목, 금액 (구분 컬럼은 있어도 되지만 사용하지 않음 ―
          계정과목 이름으로 자동 분류)
출력: finance/output/income_statement.xlsx  (손익계산서 + 부가세 참고 시트)
      finance/output/trend_chart.png        (선택한 기간 단위 기준 최근 추이)
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional

import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.chart_style import PALETTE, save_chart, setup_style  # noqa: E402
from common.excel_io import load_csv, save_excel_report  # noqa: E402
from common.format_utils import krw, pct, print_section, signed_krw  # noqa: E402
from common.period_utils import (  # noqa: E402
    add_period_label,
    available_periods,
    compute_change,
    latest_period_label,
    previous_period_label,
    year_over_year_label,
)
from common.sheet_io import is_configured as sheet_is_configured  # noqa: E402
from common.sheet_io import load_transactions_from_sheet  # noqa: E402
from common.tax_utils import estimate_corporate_tax, estimate_vat_reference  # noqa: E402

HERE = Path(__file__).resolve().parent

# 계정과목 -> 손익계산서 구분 매핑. 회사마다 계정과목 체계가 다르므로
# 이 다섯 줄만 실제 계정과목명으로 바꾸면 다른 회사 데이터에도 재사용 가능하다.
REVENUE_ACCOUNTS = {"매출"}
COGS_ACCOUNTS = {"매출원가"}
OPEX_ACCOUNTS = {"급여", "임차료", "광고선전비", "소모품비", "지급수수료"}
NON_OP_INCOME_ACCOUNTS = {"이자수익"}
NON_OP_EXPENSE_ACCOUNTS = {"이자비용", "잡손실"}

# 부가세 참고 추정에서 '매입세액 공제 가능'으로 볼 계정 (인건비는 부가세 대상이 아니라 제외)
VAT_DEDUCTIBLE_EXPENSE_ACCOUNTS = (COGS_ACCOUNTS | OPEX_ACCOUNTS) - {"급여"}

# 기간 단위별 1년 안에 몇 개가 들어가는지 (법인세 연 환산에 사용)
PERIODS_PER_YEAR = {"year": 1, "half": 2, "quarter": 4, "month": 12, "week": 52}

STATEMENT_LINES = [
    "매출액", "매출원가", "매출총이익", "판매비와관리비", "영업이익",
    "영업외수익", "영업외비용", "법인세비용차감전순이익", "법인세비용(추정)", "당기순이익",
]


def load_transactions(input_path: str | Path) -> pd.DataFrame:
    """구글 시트가 설정돼 있으면 시트에서, 아니면 로컬 파일에서 거래 데이터를 읽는다."""
    if sheet_is_configured():
        print("[데이터소스] .env에 구글 시트 설정이 있어 구글 스프레드시트에서 읽습니다.")
        df = load_transactions_from_sheet()
    else:
        print(f"[데이터소스] 구글 시트 설정이 없어 로컬 파일을 사용합니다: {input_path}")
        df = load_csv(input_path)

    df["금액"] = pd.to_numeric(df["금액"], errors="coerce")
    df["거래일자"] = pd.to_datetime(df["거래일자"])

    unknown = set(df["계정과목"].unique()) - (
        REVENUE_ACCOUNTS | COGS_ACCOUNTS | OPEX_ACCOUNTS | NON_OP_INCOME_ACCOUNTS | NON_OP_EXPENSE_ACCOUNTS
    )
    if unknown:
        print(f"[경고] 손익계산서 분류에 없는 계정과목 {unknown} 은(는) 집계에서 제외됩니다. "
              "스크립트 상단의 *_ACCOUNTS 집합에 추가해주세요.")
    return df


def _period_fiscal_year(period_label: str, period_type: str) -> int:
    """세율 구간 조회에 쓸 귀속연도. 라벨 맨 앞 4자리 연도를 그대로 사용."""
    return int(period_label[:4])


def build_statement(
    df: pd.DataFrame,
    period_col: str,
    period_label: str,
    period_type: str,
    include_local_surtax: bool = True,
) -> dict[str, float]:
    """한 기간의 거래 데이터로 K-IFRS 다단계 손익계산서 한 줄짜리 표를 만든다."""
    period_df = df[df[period_col] == period_label]

    revenue = period_df.loc[period_df["계정과목"].isin(REVENUE_ACCOUNTS), "금액"].sum()
    cogs = period_df.loc[period_df["계정과목"].isin(COGS_ACCOUNTS), "금액"].sum()
    opex = period_df.loc[period_df["계정과목"].isin(OPEX_ACCOUNTS), "금액"].sum()
    non_op_income = period_df.loc[period_df["계정과목"].isin(NON_OP_INCOME_ACCOUNTS), "금액"].sum()
    non_op_expense = period_df.loc[period_df["계정과목"].isin(NON_OP_EXPENSE_ACCOUNTS), "금액"].sum()

    gross_profit = revenue - cogs
    operating_income = gross_profit - opex
    pretax_income = operating_income + non_op_income - non_op_expense

    # 법인세 추정: 이 기간 손익을 그대로 연 환산(annualize)해서 어느 구간에 해당하는지
    # 확인한 뒤, 그 유효세율을 이 기간 실제 손익에 다시 적용한다.
    # (기간 손익을 바로 연간 누진세율표에 넣으면 짧은 기간일수록 항상 최저구간으로
    #  잡혀 세액이 크게 과소추정되기 때문)
    factor = PERIODS_PER_YEAR[period_type]
    fiscal_year = _period_fiscal_year(period_label, period_type)
    annualized_pretax = pretax_income * factor
    annual_tax = estimate_corporate_tax(annualized_pretax, fiscal_year, include_local_surtax)
    effective_rate = (annual_tax / annualized_pretax) if annualized_pretax > 0 else 0.0
    tax_expense = max(pretax_income, 0) * effective_rate

    net_income = pretax_income - tax_expense

    return {
        "매출액": revenue,
        "매출원가": cogs,
        "매출총이익": gross_profit,
        "판매비와관리비": opex,
        "영업이익": operating_income,
        "영업외수익": non_op_income,
        "영업외비용": non_op_expense,
        "법인세비용차감전순이익": pretax_income,
        "법인세비용(추정)": tax_expense,
        "당기순이익": net_income,
    }


def build_comparison_table(
    df: pd.DataFrame,
    period_col: str,
    period_label: str,
    period_type: str,
) -> pd.DataFrame:
    """당기 / 직전기 / 전년동기를 나란히 놓고 증감액·증감률까지 계산한 표."""
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
            "항목": line,
            f"당기({period_label})": cur_val,
            f"직전기({prev_label or '-'})": prev_val,
            "직전기 증감액": prev_diff if previous else None,
            "직전기 증감률": prev_pct,
            f"전년동기({yoy_label or '-'})": yoy_val,
            "전년동기 증감액": yoy_diff if yoy else None,
            "전년동기 증감률": yoy_pct,
        })

    return pd.DataFrame(rows)


def plot_trend(df: pd.DataFrame, period_col: str, period_type: str, output_path: Path, max_periods: int = 12) -> Path:
    """선택한 기간 단위 기준으로, 데이터에 존재하는 최근 N개 기간의 매출/영업이익/당기순이익 추이."""
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


def print_statement(table: pd.DataFrame, period_label: str) -> None:
    print_section(f"K-IFRS 손익계산서 ({period_label} 기준)")
    for _, row in table.iterrows():
        line = row["항목"]
        cur_col = [c for c in table.columns if c.startswith("당기")][0]
        prev_col = [c for c in table.columns if c.startswith("직전기(")][0]
        yoy_col = [c for c in table.columns if c.startswith("전년동기(")][0]

        cur = krw(row[cur_col])
        prev = f"{krw(row[prev_col])} ({signed_krw(row['직전기 증감액'])}, {pct(row['직전기 증감률'])})" \
            if pd.notna(row[prev_col]) else "데이터 없음"
        yoy = f"{krw(row[yoy_col])} ({signed_krw(row['전년동기 증감액'])}, {pct(row['전년동기 증감률'])})" \
            if pd.notna(row[yoy_col]) else "데이터 없음"

        print(f"{line:16s} {cur:>16s}  |  직전기 {prev}  |  전년동기 {yoy}")

    print("\n※ 법인세비용은 이 기간 손익을 연 환산해 국세청 세율 구간(지방소득세 포함)을 적용한 "
          "추정치이며, 세무조정·이월결손금 등은 반영되지 않았습니다. 실제 신고세액과 다를 수 있습니다.")


def main() -> None:
    parser = argparse.ArgumentParser(description="거래 데이터로 K-IFRS 손익계산서(기간 비교 포함)를 생성한다.")
    parser.add_argument(
        "--input", default=HERE / "sample_data" / "sample_transactions.csv",
        help="구글 시트 미설정 시 사용할 로컬 거래 데이터 경로 (컬럼: 거래일자, 계정과목, 금액)",
    )
    parser.add_argument(
        "--period-type", choices=["year", "half", "quarter", "month", "week"], default="month",
        help="비교 기준 기간 단위 (기본값: month)",
    )
    parser.add_argument(
        "--period", default=None,
        help="조회할 기간 라벨 (예: 2026-05, 2026-Q2, 2026-H1, 2026, 2026-W18). "
             "생략하면 데이터에 있는 가장 최근 기간 사용",
    )
    parser.add_argument(
        "--no-local-surtax", action="store_true",
        help="법인세비용 추정에서 지방소득세(10%)를 제외하고 국세만 계산",
    )
    args = parser.parse_args()

    df = load_transactions(args.input)
    df = add_period_label(df, "거래일자", args.period_type)

    period_label = args.period or latest_period_label(df)
    if period_label not in set(available_periods(df)):
        raise SystemExit(f"'{period_label}' 기간 데이터가 없습니다. "
                          f"사용 가능한 기간: {available_periods(df)}")

    table = build_comparison_table(df, "기간", period_label, args.period_type)

    # 부가세는 손익계산서 항목이 아니므로 별도 시트로 분리
    period_df = df[df["기간"] == period_label]
    taxable_revenue = period_df.loc[period_df["계정과목"].isin(REVENUE_ACCOUNTS), "금액"].sum()
    deductible_purchases = period_df.loc[period_df["계정과목"].isin(VAT_DEDUCTIBLE_EXPENSE_ACCOUNTS), "금액"].sum()
    vat_ref = estimate_vat_reference(taxable_revenue, deductible_purchases)
    vat_df = pd.DataFrame([
        {"구분": "매출세액 (매출액의 10%)", "금액": vat_ref["매출세액"]},
        {"구분": "매입세액 (공제가능 매입의 10%, 참고용)", "금액": vat_ref["매입세액"]},
        {"구분": "납부예상세액 (참고용, 실제 신고와 다를 수 있음)", "금액": vat_ref["납부예상세액"]},
    ])

    excel_path = save_excel_report(
        {"손익계산서": table, "부가세_참고": vat_df},
        HERE / "output" / "income_statement.xlsx",
    )
    chart_path = plot_trend(df, "기간", args.period_type, HERE / "output" / "trend_chart.png")

    print_statement(table, period_label)
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {chart_path}")


if __name__ == "__main__":
    main()
