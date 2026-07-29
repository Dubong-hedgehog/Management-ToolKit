"""
attendance_dashboard.py
경영지원팀 인사/총무 업무 예제 (v2): 일별 출퇴근 기록을 읽어 부서별 지각률·
초과근무 현황을 집계하고, 원하는 기간 단위(년/반기/분기/월/주)로 직전기·
전년동기와 비교한다. finance 모듈과 동일한 데이터소스/기간비교 구조를 쓴다.

데이터 소스는 두 가지 중 하나를 자동으로 고른다:
  1) .env에 구글 시트 정보가 설정돼 있으면 -> 구글 스프레드시트에서 실시간으로 읽음
  2) 아니면 -> hr/sample_data/sample_attendance.csv 같은 로컬 파일을 읽음

사용법:
    python hr/attendance_dashboard.py
    python hr/attendance_dashboard.py --period-type quarter --period 2026-Q2
    python hr/attendance_dashboard.py --period-type week --period 2026-W18

입력 컬럼: 사번, 이름, 부서, 거래일자, 지각여부(Y/N), 초과근무시간
출력: hr/output/attendance_summary.xlsx (부서별 기간비교 표)
      hr/output/late_by_dept.png        (선택 기간 부서별 지각률)
      hr/output/overtime_trend.png      (선택 기간 단위 기준 부서별 초과근무 추이)
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
from common.format_utils import pct, print_section  # noqa: E402
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

HERE = Path(__file__).resolve().parent

METRICS = ["지각률", "평균초과근무시간"]


def load_attendance(input_path: str | Path) -> pd.DataFrame:
    """구글 시트가 설정돼 있으면 시트에서, 아니면 로컬 파일에서 출퇴근 데이터를 읽는다."""
    if sheet_is_configured():
        print("[데이터소스] .env에 구글 시트 설정이 있어 구글 스프레드시트에서 읽습니다.")
        df = load_transactions_from_sheet()
    else:
        print(f"[데이터소스] 구글 시트 설정이 없어 로컬 파일을 사용합니다: {input_path}")
        df = load_csv(input_path)

    df["거래일자"] = pd.to_datetime(df["거래일자"])
    df["초과근무시간"] = pd.to_numeric(df["초과근무시간"], errors="coerce").fillna(0)
    return df


def dept_metrics_for_period(df: pd.DataFrame, period_col: str, period_label: str | None) -> pd.DataFrame:
    """한 기간의 부서별 근무일수/지각률/평균초과근무시간을 계산 (부서를 인덱스로)."""
    if period_label is None:
        return pd.DataFrame(columns=["근무일수", "지각건수", "지각률", "평균초과근무시간"])

    period_df = df[df[period_col] == period_label]
    grouped = period_df.groupby("부서").agg(
        근무일수=("거래일자", "count"),
        지각건수=("지각여부", lambda s: (s == "Y").sum()),
        총초과근무시간=("초과근무시간", "sum"),
    )
    grouped["지각률"] = grouped["지각건수"] / grouped["근무일수"]
    grouped["평균초과근무시간"] = (grouped["총초과근무시간"] / grouped["근무일수"]).round(2)
    return grouped


def build_comparison_table(df: pd.DataFrame, period_col: str, period_label: str, period_type: str) -> pd.DataFrame:
    """부서 x 지표(지각률/평균초과근무시간)별로 당기·직전기·전년동기를 비교한 표."""
    existing_periods = set(available_periods(df, period_col))

    prev_label = previous_period_label(period_label, period_type)
    yoy_label = year_over_year_label(period_label, period_type)
    prev_label = prev_label if prev_label in existing_periods else None
    yoy_label = yoy_label if (yoy_label and yoy_label in existing_periods) else None

    current = dept_metrics_for_period(df, period_col, period_label)
    previous = dept_metrics_for_period(df, period_col, prev_label)
    yoy = dept_metrics_for_period(df, period_col, yoy_label)

    rows = []
    for dept in sorted(df["부서"].unique()):
        for metric in METRICS:
            cur_val = current.loc[dept, metric] if dept in current.index else 0.0
            prev_val = previous.loc[dept, metric] if dept in previous.index else None
            yoy_val = yoy.loc[dept, metric] if dept in yoy.index else None

            prev_diff, prev_pct = compute_change(cur_val, prev_val)
            yoy_diff, yoy_pct = compute_change(cur_val, yoy_val)

            rows.append({
                "부서": dept,
                "지표": metric,
                f"당기({period_label})": cur_val,
                f"직전기({prev_label or '-'})": prev_val,
                "직전기 증감": prev_diff if prev_label else None,
                "직전기 증감률": prev_pct,
                f"전년동기({yoy_label or '-'})": yoy_val,
                "전년동기 증감": yoy_diff if yoy_label else None,
                "전년동기 증감률": yoy_pct,
            })

    return pd.DataFrame(rows)


def plot_late_by_dept(df: pd.DataFrame, period_col: str, period_label: str, output_path: Path) -> Path:
    setup_style()
    current = dept_metrics_for_period(df, period_col, period_label).sort_values("지각률", ascending=False)

    fig, ax = plt.subplots(figsize=(7, 4.5))
    ax.bar(current.index, current["지각률"] * 100, color=PALETTE[: len(current)])
    ax.set_title(f"부서별 지각률 ({period_label})")
    ax.set_ylabel("지각률 (%)")
    for i, v in enumerate(current["지각률"] * 100):
        ax.text(i, v + 0.3, f"{v:.1f}%", ha="center")
    return save_chart(fig, output_path)


def plot_overtime_trend(df: pd.DataFrame, period_col: str, period_type: str, output_path: Path, max_periods: int = 12) -> Path:
    setup_style()
    periods = available_periods(df, period_col)[-max_periods:]

    fig, ax = plt.subplots(figsize=(9, 4.8))
    for i, dept in enumerate(sorted(df["부서"].unique())):
        values = []
        for p in periods:
            m = dept_metrics_for_period(df, period_col, p)
            values.append(m.loc[dept, "평균초과근무시간"] if dept in m.index else 0)
        ax.plot(periods, values, marker="o", label=dept, color=PALETTE[i % len(PALETTE)])

    ax.set_title(f"기간별({period_type}) 부서별 평균 초과근무시간 추이")
    ax.set_ylabel("평균 초과근무시간 (h)")
    ax.tick_params(axis="x", rotation=45)
    ax.legend()
    return save_chart(fig, output_path)


def print_comparison(table: pd.DataFrame, period_label: str) -> None:
    print_section(f"부서별 근태 현황 ({period_label} 기준)")
    cur_col = [c for c in table.columns if c.startswith("당기")][0]
    prev_col = [c for c in table.columns if c.startswith("직전기(")][0]
    yoy_col = [c for c in table.columns if c.startswith("전년동기(")][0]

    for dept, group in table.groupby("부서", sort=False):
        print(f"\n[{dept}]")
        for _, row in group.iterrows():
            metric = row["지표"]
            is_rate = metric == "지각률"
            fmt = pct if is_rate else (lambda v: "N/A" if pd.isna(v) else f"{v:.2f}h")

            cur = fmt(row[cur_col])
            prev = f"{fmt(row[prev_col])} ({pct(row['직전기 증감률'])})" if pd.notna(row[prev_col]) else "데이터 없음"
            yoy = f"{fmt(row[yoy_col])} ({pct(row['전년동기 증감률'])})" if pd.notna(row[yoy_col]) else "데이터 없음"

            print(f"  {metric:12s} 당기 {cur:>8s}  |  직전기 {prev}  |  전년동기 {yoy}")


def main() -> None:
    parser = argparse.ArgumentParser(description="출퇴근 데이터로 부서별 근태 현황(기간 비교 포함)을 생성한다.")
    parser.add_argument(
        "--input", default=HERE / "sample_data" / "sample_attendance.csv",
        help="구글 시트 미설정 시 사용할 로컬 출퇴근 데이터 경로 (컬럼: 부서, 거래일자, 지각여부, 초과근무시간)",
    )
    parser.add_argument(
        "--period-type", choices=["year", "half", "quarter", "month", "week"], default="month",
        help="비교 기준 기간 단위 (기본값: month)",
    )
    parser.add_argument(
        "--period", default=None,
        help="조회할 기간 라벨 (예: 2026-05, 2026-Q2). 생략하면 데이터에 있는 가장 최근 기간 사용",
    )
    args = parser.parse_args()

    df = load_attendance(args.input)
    df = add_period_label(df, "거래일자", args.period_type)

    period_label = args.period or latest_period_label(df)
    if period_label not in set(available_periods(df)):
        raise SystemExit(f"'{period_label}' 기간 데이터가 없습니다. "
                          f"사용 가능한 기간: {available_periods(df)}")

    table = build_comparison_table(df, "기간", period_label, args.period_type)

    excel_path = save_excel_report({"근태_기간비교": table}, HERE / "output" / "attendance_summary.xlsx")
    late_chart = plot_late_by_dept(df, "기간", period_label, HERE / "output" / "late_by_dept.png")
    overtime_chart = plot_overtime_trend(df, "기간", args.period_type, HERE / "output" / "overtime_trend.png")

    print_comparison(table, period_label)
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {late_chart}, {overtime_chart}")


if __name__ == "__main__":
    main()
