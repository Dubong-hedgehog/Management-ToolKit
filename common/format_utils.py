"""
format_utils.py
콘솔/리포트 출력에서 반복적으로 쓰는 포맷팅 함수 모음 (통화, 퍼센트, 구분선 등).
"""
from __future__ import annotations

import math
from typing import Optional


def _is_missing(value) -> bool:
    """None은 물론, pandas DataFrame을 거치며 None이 NaN으로 바뀐 경우까지 함께 걸러낸다."""
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return False


def krw(amount: float) -> str:
    """숫자를 '1,234,000원' 형태의 원화 문자열로 변환. 값이 없으면 'N/A'."""
    if _is_missing(amount):
        return "N/A"
    return f"{amount:,.0f}원"


def pct(value: Optional[float], digits: int = 1) -> str:
    """비율(0.1 = 10%)을 퍼센트 문자열로 변환. 값이 없으면 'N/A'."""
    if _is_missing(value):
        return "N/A"
    return f"{value * 100:.{digits}f}%"


def signed_krw(amount: float) -> str:
    """증감액처럼 부호가 중요한 금액을 '+1,234,000원' / '-500,000원' 형태로 표시. 값이 없으면 'N/A'."""
    if _is_missing(amount):
        return "N/A"
    sign = "+" if amount >= 0 else ""
    return f"{sign}{amount:,.0f}원"


def print_section(title: str) -> None:
    """콘솔 출력에서 구간을 구분하는 헤더를 찍는다."""
    line = "=" * max(40, len(title) + 4)
    print(f"\n{line}\n  {title}\n{line}")
