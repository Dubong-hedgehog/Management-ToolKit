"""
tax_utils.py
법인세비용 / 부가가치세를 원본 거래 데이터만으로 '추정'하는 모듈. (내부 관리용 추정치)
"""
from __future__ import annotations

from typing import Optional

CORPORATE_TAX_BRACKETS: dict[int, list[tuple[Optional[int], float]]] = {
    2026: [
        (200_000_000, 0.10),
        (20_000_000_000, 0.20),
        (300_000_000_000, 0.22),
        (None, 0.25),
    ],
    2000: [
        (200_000_000, 0.09),
        (20_000_000_000, 0.19),
        (300_000_000_000, 0.21),
        (None, 0.24),
    ],
}

LOCAL_SURTAX_RATE = 0.10
VAT_RATE = 0.10


def _brackets_for_year(fiscal_year: int) -> list[tuple[Optional[int], float]]:
    applicable_years = sorted(y for y in CORPORATE_TAX_BRACKETS if y <= fiscal_year)
    key = applicable_years[-1] if applicable_years else min(CORPORATE_TAX_BRACKETS)
    return CORPORATE_TAX_BRACKETS[key]


def national_corporate_tax(pretax_income: float, fiscal_year: int) -> float:
    if pretax_income <= 0:
        return 0.0
    brackets = _brackets_for_year(fiscal_year)
    tax = 0.0
    prev_threshold = 0
    for threshold, rate in brackets:
        if threshold is None or pretax_income <= threshold:
            tax += (pretax_income - prev_threshold) * rate
            break
        tax += (threshold - prev_threshold) * rate
        prev_threshold = threshold
    return tax


def estimate_corporate_tax(pretax_income: float, fiscal_year: int, include_local_surtax: bool = True) -> float:
    national = national_corporate_tax(pretax_income, fiscal_year)
    if include_local_surtax:
        return national * (1 + LOCAL_SURTAX_RATE)
    return national


def estimate_vat_reference(taxable_revenue: float, deductible_purchases: float, vat_rate: float = VAT_RATE) -> dict[str, float]:
    output_tax = taxable_revenue * vat_rate
    input_tax = deductible_purchases * vat_rate
    return {"매출세액": output_tax, "매입세액": input_tax, "납부예상세액": output_tax - input_tax}
