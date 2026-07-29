"""
period_utils.py
년/반기/분기/월/주 단위로 자유롭게 기간을 나누고, "직전기 대비" / "전년 동기 대비"를
계산하기 위한 공통 유틸.

핵심 규칙: 같은 단위끼리만 비교한다 (월 vs 월, 분기 vs 분기 ...). 주 단위와
월 단위처럼 길이가 다른 기간은 비교 자체가 성립하지 않으므로 이 모듈은
아예 그런 조합을 만들 수 없게 되어 있다 (period_type 하나를 고르면 그 안에서만
직전기/전년동기를 계산).

주(week) 기준 '전년 동기'는 ISO 주차 번호가 아니라 정확히 364일(52주) 전 날짜를
기준으로 계산한다. ISO 주차는 해마다 52주/53주로 길이가 달라 같은 주차 번호가
반드시 같은 계절을 가리키지 않기 때문에, 요일까지 맞아떨어지는 '52주 전' 방식이
실무적으로 더 정확한 비교가 된다.
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Optional

import pandas as pd

PERIOD_TYPES = ("year", "half", "quarter", "month", "week")


def add_period_label(df: pd.DataFrame, date_col: str, period_type: str, out_col: str = "기간") -> pd.DataFrame:
    """거래일자 컬럼을 보고 년/반기/분기/월/주 라벨을 붙인 새 컬럼을 추가한다."""
    if period_type not in PERIOD_TYPES:
        raise ValueError(f"period_type은 {PERIOD_TYPES} 중 하나여야 합니다: {period_type!r}")

    df = df.copy()
    dates = pd.to_datetime(df[date_col])

    if period_type == "year":
        df[out_col] = dates.dt.year.astype(str)
    elif period_type == "half":
        half = ((dates.dt.month - 1) // 6) + 1
        df[out_col] = dates.dt.year.astype(str) + "-H" + half.astype(str)
    elif period_type == "quarter":
        df[out_col] = dates.dt.year.astype(str) + "-Q" + dates.dt.quarter.astype(str)
    elif period_type == "month":
        df[out_col] = dates.dt.strftime("%Y-%m")
    elif period_type == "week":
        iso = dates.dt.isocalendar()
        df[out_col] = iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2)

    return df


def _parse_label(label: str, period_type: str) -> tuple[int, Optional[int]]:
    if period_type == "year":
        return int(label), None
    if period_type == "half":
        y, h = label.split("-H")
        return int(y), int(h)
    if period_type == "quarter":
        y, q = label.split("-Q")
        return int(y), int(q)
    if period_type == "month":
        y, m = label.split("-")
        return int(y), int(m)
    if period_type == "week":
        y, w = label.split("-W")
        return int(y), int(w)
    raise ValueError(period_type)


def previous_period_label(label: str, period_type: str) -> str:
    """같은 단위 기준 '직전' 기간 라벨. 예: 2026-05 -> 2026-04, 2026-Q1 -> 2025-Q4."""
    if period_type == "year":
        y, _ = _parse_label(label, period_type)
        return str(y - 1)
    if period_type == "half":
        y, h = _parse_label(label, period_type)
        return f"{y-1}-H2" if h == 1 else f"{y}-H1"
    if period_type == "quarter":
        y, q = _parse_label(label, period_type)
        return f"{y-1}-Q4" if q == 1 else f"{y}-Q{q-1}"
    if period_type == "month":
        y, m = _parse_label(label, period_type)
        return f"{y-1}-12" if m == 1 else f"{y}-{m-1:02d}"
    if period_type == "week":
        y, w = _parse_label(label, period_type)
        monday = date.fromisocalendar(y, w, 1)
        prev_monday = monday - timedelta(days=7)
        iso = prev_monday.isocalendar()
        return f"{iso.year}-W{iso.week:02d}"
    raise ValueError(period_type)


def year_over_year_label(label: str, period_type: str) -> Optional[str]:
    """'전년 동기' 라벨. year 타입은 전년동기 개념이 없어 None을 반환한다."""
    if period_type == "year":
        return None
    if period_type == "half":
        y, h = _parse_label(label, period_type)
        return f"{y-1}-H{h}"
    if period_type == "quarter":
        y, q = _parse_label(label, period_type)
        return f"{y-1}-Q{q}"
    if period_type == "month":
        y, m = _parse_label(label, period_type)
        return f"{y-1}-{m:02d}"
    if period_type == "week":
        y, w = _parse_label(label, period_type)
        monday = date.fromisocalendar(y, w, 1)
        year_ago_monday = monday - timedelta(weeks=52)
        iso = year_ago_monday.isocalendar()
        return f"{iso.year}-W{iso.week:02d}"
    raise ValueError(period_type)


def available_periods(df: pd.DataFrame, period_col: str = "기간") -> list[str]:
    """데이터에 실제로 존재하는 기간 라벨을 시간순으로 정렬해서 반환.

    라벨 형식이 전부 'YYYY(-단위+제로패딩)' 구조라 문자열 정렬이 곧 시간순 정렬이다.
    """
    return sorted(df[period_col].dropna().unique().tolist())


def latest_period_label(df: pd.DataFrame, period_col: str = "기간") -> str:
    """데이터에 존재하는 가장 최근 기간 라벨."""
    periods = available_periods(df, period_col)
    if not periods:
        raise ValueError("데이터에서 기간 라벨을 찾을 수 없습니다.")
    return periods[-1]


def compute_change(current: float, previous: Optional[float]) -> tuple[float, Optional[float]]:
    """증감액과 증감률을 계산. 비교 대상 기간 데이터가 없으면 (증감액, None)을 반환."""
    if previous is None:
        return current, None
    diff = current - previous
    pct = (diff / previous) if previous != 0 else None
    return diff, pct
