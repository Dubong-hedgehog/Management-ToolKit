"""
format_utils.py
콘솔/리포트 출력에서 반복적으로 쓰는 포맷팅 함수 모음 (통화, 퍼센트, 구분선 등).
"""
from __future__ import annotations


def krw(amount: float) -> str:
    """숫자를 '1,234,000원' 형태의 원화 문자열로 변환."""
    return f"{amount:,.0f}원"


def pct(value: float, digits: int = 1) -> str:
    """비율(0.1 = 10%)을 퍼센트 문자열로 변환."""
    return f"{value * 100:.{digits}f}%"


def print_section(title: str) -> None:
    """콘솔 출력에서 구간을 구분하는 헤더를 찍는다."""
    line = "=" * max(40, len(title) + 4)
    print(f"\n{line}\n  {title}\n{line}")
