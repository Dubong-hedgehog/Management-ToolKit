"""
tax_utils.py
법인세비용 / 부가가치세를 원본 거래 데이터만으로 '추정'하는 모듈.

주의 (반드시 읽을 것):
  이 모듈이 계산하는 세액은 어디까지나 내부 관리 리포트용 추정치입니다.
  실제 법인세는 세무조정(익금/손금 산입·불산입), 이월결손금, 세액공제·감면 등을
  반영해서 계산되며, 부가세는 세금계산서 발행 기준/공급시기/불공제 항목 등을
  별도로 따져야 합니다. 여기서는 "거래 데이터에 찍힌 손익만으로 대략 얼마
  나오겠다"를 빠르게 가늠하는 용도로만 쓰고, 실제 신고 세액과는 다를 수 있습니다.

세율 출처: 국세청 법인세 세율표, 2026년 귀속분부터 전 구간 1%p 인상
  (2025년 귀속: 9/19/21/24%, 2026년 귀속: 10/20/22/25%, 국세 기준)
법인세율은 세법 개정으로 계속 바뀌므로, 아래 CORPORATE_TAX_BRACKETS만 고치면
새 세율에 맞출 수 있도록 연도별 표로 분리해뒀습니다.
"""
from __future__ import annotations

from typing import Optional

# {귀속연도 시작: [(과세표준 상한, 국세 세율), ...]}
# 상한을 None으로 두면 '그 이상 전부'를 의미.
CORPORATE_TAX_BRACKETS: dict[int, list[tuple[Optional[int], float]]] = {
    2026: [  # 2026년 1월 1일 이후 개시 사업연도부터 (2027년 3월 신고분)
        (200_000_000, 0.10),
        (20_000_000_000, 0.20),
        (300_000_000_000, 0.22),
        (None, 0.25),
    ],
    2000: [  # 2025년 귀속 이하 (2026년 3월 신고분까지) - 기본값
        (200_000_000, 0.09),
        (20_000_000_000, 0.19),
        (300_000_000_000, 0.21),
        (None, 0.24),
    ],
}

LOCAL_SURTAX_RATE = 0.10  # 지방소득세 = 법인세(국세)의 10%
VAT_RATE = 0.10  # 부가가치세 표준세율


def _brackets_for_year(fiscal_year: int) -> list[tuple[Optional[int], float]]:
    """해당 귀속연도에 적용되는 세율 구간표를 찾는다 (없으면 가장 가까운 과거 표 사용)."""
    applicable_years = sorted(y for y in CORPORATE_TAX_BRACKETS if y <= fiscal_year)
    key = applicable_years[-1] if applicable_years else min(CORPORATE_TAX_BRACKETS)
    return CORPORATE_TAX_BRACKETS[key]


def national_corporate_tax(pretax_income: float, fiscal_year: int) -> float:
    """과세표준(=법인세비용차감전순이익으로 근사)에 누진세율을 적용한 국세 법인세."""
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


def estimate_corporate_tax(
    pretax_income: float,
    fiscal_year: int,
    include_local_surtax: bool = True,
) -> float:
    """법인세비용(지방소득세 포함 여부 선택) 추정치를 반환.

    Args:
        pretax_income: 법인세비용차감전순이익 (연 환산 금액을 넣는 것을 권장)
        fiscal_year: 귀속 사업연도 (예: 2026)
        include_local_surtax: True면 지방소득세(법인세의 10%)까지 포함해서 반환
    """
    national = national_corporate_tax(pretax_income, fiscal_year)
    if include_local_surtax:
        return national * (1 + LOCAL_SURTAX_RATE)
    return national


def estimate_vat_reference(
    taxable_revenue: float,
    deductible_purchases: float,
    vat_rate: float = VAT_RATE,
) -> dict[str, float]:
    """부가가치세 참고 추정치 (매출세액 - 매입세액).

    손익계산서 항목이 아니라 별도 참고용으로만 제공한다 (부가세는 예수금 성격이라
    당기순이익 계산에 포함하지 않음). 접대비 등 매입세액 불공제 항목은
    deductible_purchases에서 미리 제외하고 넣어야 정확도가 올라간다.
    """
    output_tax = taxable_revenue * vat_rate
    input_tax = deductible_purchases * vat_rate
    return {
        "매출세액": output_tax,
        "매입세액": input_tax,
        "납부예상세액": output_tax - input_tax,
    }
