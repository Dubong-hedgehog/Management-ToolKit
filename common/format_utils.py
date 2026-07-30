"""
format_utils.py
콘솔/리포트 출력에서 반복적으로 쓰는 포맷팅 함수 모음 (통화, 퍼센트, 구분선 등).
"""
from __future__ import annotations

import math
from typing import Optional


def _is_missing(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return False


def krw(amount: float) -> str:
    if _is_missing(amount):
        return "N/A"
    return f"{amount:,.0f}원"


def pct(value: Optional[float], digits: int = 1) -> str:
    if _is_missing(value):
        return "N/A"
    return f"{value * 100:.{digits}f}%"


def signed_krw(amount: float) -> str:
    if _is_missing(amount):
        return "N/A"
    sign = "+" if amount >= 0 else ""
    return f"{sign}{amount:,.0f}원"


def print_section(title: str) -> None:
    line = "=" * max(40, len(title) + 4)
    print(f"\n{line}\n  {title}\n{line}")
