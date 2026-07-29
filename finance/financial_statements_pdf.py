"""
financial_statements_pdf.py
경영지원팀 회계/재무 업무 예제: 홈택스/DART에서 보는 것과 같은 정식 서식으로
손익계산서·재무상태표를 PDF로 출력한다. 대표이사가 서류를 받았을 때 바로
알아볼 수 있는 형태를 목표로 한다 (그래프는 여기서는 안 씀 — 이 문서 자체가
공식 보고 양식이라 표가 그래프보다 우선이다).

income_statement_generator.py와의 차이:
  - income_statement_generator.py: 년/반기/분기/월/주 등 원하는 단위로 자유롭게
    직전기·전년동기와 비교하는 "운영/관리용" 도구
  - financial_statements_pdf.py: 정식 재무제표 관행대로 "당기(연초~기준일 누적)
    vs 전기(직전 사업연도 전체/기말)"로 비교하는 "대외 보고용" 도구

사용법:
    python finance/financial_statements_pdf.py
    python finance/financial_statements_pdf.py --as-of 2026-03-31

입력: finance/sample_data/sample_transactions.csv   (손익계산서용, 거래 건별)
      finance/sample_data/sample_balance_sheet.csv  (재무상태표용, 월말 잔액 스냅샷)
출력: finance/output/손익계산서.pdf
      finance/output/재무상태표.pdf
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.excel_io import load_csv  # noqa: E402
from common.pdf_statement import StatementRow, render_statement_pdf  # noqa: E402
from common.tax_utils import estimate_corporate_tax  # noqa: E402

HERE = Path(__file__).resolve().parent

# ---- 회사 정보 (실제 사용 시 이 두 줄만 바꾸면 됨) ----
COMPANY_NAME = "주식회사 다올예제"
FIRST_FISCAL_YEAR = 2019  # 제1기가 시작된 연도 (제N기 번호 계산용)

REVENUE_ACCOUNTS = {"매출"}
COGS_ACCOUNTS = {"매출원가"}
OPEX_ACCOUNTS = {"급여", "임차료", "광고선전비", "소모품비", "지급수수료"}
NON_OP_INCOME_ACCOUNTS = {"이자수익"}
NON_OP_EXPENSE_ACCOUNTS = {"이자비용", "잡손실"}
OPEX_LABELS = {
    "급여": "직 원 급 여", "임차료": "지 급 임 차 료", "광고선전비": "광 고 선 전 비",
    "소모품비": "소 모 품 비", "지급수수료": "지 급 수 수 료",
}

ASSET_GROUPS = {
    "당좌자산": ["보통예금", "외상매출금", "미수금", "선급금", "선급비용", "선납세금"],
}
NONCURRENT_ASSET_GROUPS = {
    "유형자산": ["비품"],
    "무형자산": ["무형자산(개발비등)"],
    "기타비유동자산": ["임차보증금등"],
}
CURRENT_LIAB_ACCOUNTS = ["외상매입금", "미지급금", "예수금", "부가세예수금", "단기차입금", "미지급비용"]
NONCURRENT_LIAB_ACCOUNTS = ["장기차입금"]


def fiscal_year_no(year: int) -> int:
    return year - FIRST_FISCAL_YEAR + 1


def sum_pnl(df: pd.DataFrame, accounts: set[str]) -> float:
    return df.loc[df["계정과목"].isin(accounts), "금액"].sum()


def compute_pnl(df: pd.DataFrame, start, end, fiscal_year: int) -> dict[str, float]:
    period_df = df[(df["거래일자"] >= start) & (df["거래일자"] <= end)]
    revenue = sum_pnl(period_df, REVENUE_ACCOUNTS)
    cogs = sum_pnl(period_df, COGS_ACCOUNTS)
    gross_profit = revenue - cogs
    opex_detail = {acc: sum_pnl(period_df, {acc}) for acc in OPEX_ACCOUNTS}
    opex_total = sum(opex_detail.values())
    operating_income = gross_profit - opex_total
    noi_detail = {acc: sum_pnl(period_df, {acc}) for acc in NON_OP_INCOME_ACCOUNTS}
    noi_total = sum(noi_detail.values())
    noe_detail = {acc: sum_pnl(period_df, {acc}) for acc in NON_OP_EXPENSE_ACCOUNTS}
    noe_total = sum(noe_detail.values())
    pretax = operating_income + noi_total - noe_total
    tax = estimate_corporate_tax(pretax, fiscal_year) if pretax > 0 else 0.0
    net_income = pretax - tax
    return {
        "revenue": revenue, "cogs": cogs, "gross_profit": gross_profit,
        "opex_detail": opex_detail, "opex_total": opex_total, "operating_income": operating_income,
        "noi_detail": noi_detail, "noi_total": noi_total, "noe_detail": noe_detail, "noe_total": noe_total,
        "pretax": pretax, "tax": tax, "net_income": net_income,
    }


def build_income_statement_rows(current: dict, prior: dict) -> list[StatementRow]:
    def signed_label(base_positive: str, base_negative: str, value: float) -> str:
        return base_positive if value >= 0 else base_negative

    rows = []
    rows.append(StatementRow("Ⅰ. 매    출    액", level=0, total_current=current["revenue"], total_prior=prior["revenue"], bold=True))
    rows.append(StatementRow("서 비 스 매 출", level=2, detail_current=current["revenue"], detail_prior=prior["revenue"]))
    rows.append(StatementRow("Ⅱ. 매 출 원 가", level=0, total_current=current["cogs"], total_prior=prior["cogs"], bold=True))
    rows.append(StatementRow("Ⅲ. 매 출 총 이 익", level=0, total_current=current["gross_profit"], total_prior=prior["gross_profit"], bold=True))
    rows.append(StatementRow("Ⅳ. 판 매 비 와 관 리 비", level=0, total_current=current["opex_total"], total_prior=prior["opex_total"], bold=True))
    for acc, label in OPEX_LABELS.items():
        rows.append(StatementRow(label, level=2, detail_current=current["opex_detail"].get(acc, 0), detail_prior=prior["opex_detail"].get(acc, 0)))

    op_label = signed_label("Ⅴ. 영    업    이    익", "Ⅴ. 영    업    손    실", current["operating_income"])
    rows.append(StatementRow(op_label, level=0, total_current=abs(current["operating_income"]), total_prior=abs(prior["operating_income"]), bold=True))

    rows.append(StatementRow("Ⅵ. 영 업 외 수 익", level=0, total_current=current["noi_total"], total_prior=prior["noi_total"], bold=True))
    for acc in NON_OP_INCOME_ACCOUNTS:
        rows.append(StatementRow(acc.replace("", " ").strip(), level=2, detail_current=current["noi_detail"].get(acc, 0), detail_prior=prior["noi_detail"].get(acc, 0)))

    rows.append(StatementRow("Ⅶ. 영 업 외 비 용", level=0, total_current=current["noe_total"], total_prior=prior["noe_total"], bold=True))
    for acc in NON_OP_EXPENSE_ACCOUNTS:
        rows.append(StatementRow(acc, level=2, detail_current=current["noe_detail"].get(acc, 0), detail_prior=prior["noe_detail"].get(acc, 0)))

    pretax_label = signed_label("Ⅷ. 법인세차감전순이익", "Ⅷ. 법인세차감전순손실", current["pretax"])
    rows.append(StatementRow(pretax_label, level=0, total_current=abs(current["pretax"]), total_prior=abs(prior["pretax"]), bold=True))
    rows.append(StatementRow("Ⅸ. 법  인  세  등", level=0, total_current=current["tax"], total_prior=prior["tax"], bold=True))
    net_label = signed_label("Ⅹ. 당   기   순   이   익", "Ⅹ. 당   기   순   손   실", current["net_income"])
    rows.append(StatementRow(net_label, level=0, total_current=abs(current["net_income"]), total_prior=abs(prior["net_income"]), bold=True))
    return rows


def build_balance_sheet_rows(current: pd.Series, prior: pd.Series) -> list[StatementRow]:
    def g(series, acc):
        return float(series.get(acc, 0) or 0)

    rows = []
    rows.append(StatementRow("자        산", level=0, bold=True))

    current_assets_cur = sum(g(current, a) for a in ASSET_GROUPS["당좌자산"])
    current_assets_pri = sum(g(prior, a) for a in ASSET_GROUPS["당좌자산"])
    rows.append(StatementRow("Ⅰ. 유 동 자 산", level=0, total_current=current_assets_cur, total_prior=current_assets_pri, bold=True))
    rows.append(StatementRow("(1) 당 좌 자 산", level=1, total_current=current_assets_cur, total_prior=current_assets_pri, bold=True))
    for acc in ASSET_GROUPS["당좌자산"]:
        rows.append(StatementRow(acc, level=2, detail_current=g(current, acc), detail_prior=g(prior, acc)))

    noncurrent_total_cur = sum(g(current, a) for grp in NONCURRENT_ASSET_GROUPS.values() for a in grp)
    noncurrent_total_pri = sum(g(prior, a) for grp in NONCURRENT_ASSET_GROUPS.values() for a in grp)
    rows.append(StatementRow("Ⅱ. 비 유 동 자 산", level=0, total_current=noncurrent_total_cur, total_prior=noncurrent_total_pri, bold=True))
    group_labels = {"유형자산": "(2) 유 형 자 산", "무형자산": "(3) 무 형 자 산", "기타비유동자산": "(4) 기타비유동자산"}
    for grp, accs in NONCURRENT_ASSET_GROUPS.items():
        grp_cur = sum(g(current, a) for a in accs)
        grp_pri = sum(g(prior, a) for a in accs)
        rows.append(StatementRow(group_labels[grp], level=1, total_current=grp_cur, total_prior=grp_pri, bold=True))
        for acc in accs:
            rows.append(StatementRow(acc, level=2, detail_current=g(current, acc), detail_prior=g(prior, acc)))

    total_assets_cur = current_assets_cur + noncurrent_total_cur
    total_assets_pri = current_assets_pri + noncurrent_total_pri
    rows.append(StatementRow("자    산    총    계", level=0, total_current=total_assets_cur, total_prior=total_assets_pri, bold=True))

    rows.append(StatementRow("부        채", level=0, bold=True))
    cur_liab_cur = sum(g(current, a) for a in CURRENT_LIAB_ACCOUNTS)
    cur_liab_pri = sum(g(prior, a) for a in CURRENT_LIAB_ACCOUNTS)
    rows.append(StatementRow("Ⅰ. 유 동 부 채", level=0, total_current=cur_liab_cur, total_prior=cur_liab_pri, bold=True))
    for acc in CURRENT_LIAB_ACCOUNTS:
        rows.append(StatementRow(acc, level=2, detail_current=g(current, acc), detail_prior=g(prior, acc)))

    noncur_liab_cur = sum(g(current, a) for a in NONCURRENT_LIAB_ACCOUNTS)
    noncur_liab_pri = sum(g(prior, a) for a in NONCURRENT_LIAB_ACCOUNTS)
    rows.append(StatementRow("Ⅱ. 비 유 동 부 채", level=0, total_current=noncur_liab_cur, total_prior=noncur_liab_pri, bold=True))
    for acc in NONCURRENT_LIAB_ACCOUNTS:
        rows.append(StatementRow(acc, level=2, detail_current=g(current, acc), detail_prior=g(prior, acc)))

    total_liab_cur = cur_liab_cur + noncur_liab_cur
    total_liab_pri = cur_liab_pri + noncur_liab_pri
    rows.append(StatementRow("부    채    총    계", level=0, total_current=total_liab_cur, total_prior=total_liab_pri, bold=True))

    rows.append(StatementRow("자        본", level=0, bold=True))
    rows.append(StatementRow("Ⅰ. 자   본   금", level=0, total_current=g(current, "자본금"), total_prior=g(prior, "자본금"), bold=True))
    rows.append(StatementRow("Ⅱ. 자 본 잉 여 금", level=0, total_current=g(current, "자본잉여금"), total_prior=g(prior, "자본잉여금"), bold=True))
    re_cur, re_pri = g(current, "이익잉여금"), g(prior, "이익잉여금")
    re_label = "Ⅲ. 이 익 잉 여 금" if re_cur >= 0 else "Ⅲ. 결        손        금"
    rows.append(StatementRow(re_label, level=0, total_current=abs(re_cur), total_prior=abs(re_pri), bold=True))

    total_equity_cur = g(current, "자본금") + g(current, "자본잉여금") + re_cur
    total_equity_pri = g(prior, "자본금") + g(prior, "자본잉여금") + re_pri
    rows.append(StatementRow("자    본    총    계", level=0, total_current=total_equity_cur, total_prior=total_equity_pri, bold=True))
    rows.append(StatementRow("부채 및 자본총계", level=0, total_current=total_liab_cur + total_equity_cur, total_prior=total_liab_pri + total_equity_pri, bold=True))
    return rows


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--tx-input", default=HERE / "sample_data" / "sample_transactions.csv")
    parser.add_argument("--bs-input", default=HERE / "sample_data" / "sample_balance_sheet.csv")
    parser.add_argument("--as-of", default=None, help="기준일 YYYY-MM-DD (생략하면 데이터의 가장 최근 일자)")
    args = parser.parse_args()

    tx = load_csv(args.tx_input)
    tx["거래일자"] = pd.to_datetime(tx["거래일자"])
    bs = load_csv(args.bs_input)
    bs["기준일"] = pd.to_datetime(bs["기준일"])

    as_of = pd.Timestamp(args.as_of) if args.as_of else tx["거래일자"].max()
    fy = as_of.year
    fy_no = fiscal_year_no(fy)
    prior_fy_no = fy_no - 1

    # ---- 손익계산서: 당기(연초~기준일 누적) vs 전기(직전 사업연도 전체) ----
    current_pnl = compute_pnl(tx, pd.Timestamp(fy, 1, 1), as_of, fy)
    prior_year_end = pd.Timestamp(fy - 1, 12, 31)
    prior_pnl = compute_pnl(tx, pd.Timestamp(fy - 1, 1, 1), prior_year_end, fy - 1)

    pnl_rows = build_income_statement_rows(current_pnl, prior_pnl)
    pnl_period_lines = [
        f"제 {fy_no}(당)기  {fy}년   1월   1일부터   {as_of.year}년 {as_of.month}월 {as_of.day}일까지",
        f"제 {prior_fy_no}(전)기  {fy-1}년   1월   1일부터   {fy-1}년 12월 31일까지",
    ]
    pnl_path = render_statement_pdf(
        pnl_rows, "손 익 계 산 서", pnl_period_lines, COMPANY_NAME,
        HERE / "output" / "손익계산서.pdf",
        current_header=f"제{fy_no}(당)기", prior_header=f"제{prior_fy_no}(전)기",
    )

    # ---- 재무상태표: 당기(기준일 현재) vs 전기(직전 사업연도 말 현재) ----
    bs_current = bs[bs["기준일"] == as_of].set_index("계정과목")["금액"]
    bs_prior_date = bs[bs["기준일"] <= prior_year_end]["기준일"].max()
    if pd.isna(bs_prior_date):
        raise SystemExit(f"전기말({prior_year_end.date()} 이전) 재무상태표 데이터가 없습니다.")
    bs_prior = bs[bs["기준일"] == bs_prior_date].set_index("계정과목")["금액"]

    if bs_current.empty:
        raise SystemExit(f"'{as_of.date()}' 기준 재무상태표 데이터가 없습니다. "
                          f"사용 가능한 기준일: {sorted(bs['기준일'].dt.strftime('%Y-%m-%d').unique())}")

    bs_rows = build_balance_sheet_rows(bs_current, bs_prior)
    bs_period_lines = [
        f"제 {fy_no}기   {as_of.year}년 {as_of.month:02d}월 {as_of.day:02d}일   현재",
        f"제 {prior_fy_no}기   {prior_year_end.year}년 12월 31일   현재",
    ]
    bs_path = render_statement_pdf(
        bs_rows, "재 무 상 태 표", bs_period_lines, COMPANY_NAME,
        HERE / "output" / "재무상태표.pdf",
        current_header=f"제{fy_no}(당)기", prior_header=f"제{prior_fy_no}(전)기",
    )

    print(f"손익계산서 PDF 저장: {pnl_path}")
    print(f"재무상태표 PDF 저장: {bs_path}")
    print(f"\n당기순이익(YTD, {fy}.1.1~{as_of.date()}): {current_pnl['net_income']:,.0f}원")
    print(f"※ 법인세는 추정치이며 실제 신고세액과 다를 수 있습니다.")


if __name__ == "__main__":
    main()
