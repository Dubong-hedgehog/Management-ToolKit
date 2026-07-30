"""
period_utils.py
년/반기/분기/월/주 단위로 자유롭게 기간을 나누고, "직전기 대비" / "전년 동기 대비"를
계산하기 위한 공통 유틸.
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Optional

import pandas as pd

PERIOD_TYPES = ("year", "half", "quarter", "month", "week")


def add_period_label(df: pd.DataFrame, date_col: str, period_type: str, out_col: str = "기간") -> pd.DataFrame:
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
    return sorted(df[period_col].dropna().unique().tolist())


def latest_period_label(df: pd.DataFrame, period_col: str = "기간") -> str:
    periods = available_periods(df, period_col)
    if not periods:
        raise ValueError("데이터에서 기간 라벨을 찾을 수 없습니다.")
    return periods[-1]


def compute_change(current: float, previous: Optional[float]) -> tuple[float, Optional[float]]:
    if previous is None:
        return current, None
    diff = current - previous
    pct = (diff / previous) if previous != 0 else None
    return diff, pct
