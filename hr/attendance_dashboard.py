"""
attendance_dashboard.py
경영지원팀 인사/총무 업무 예제: 출퇴근 원본 기록을 읽어 부서별 지각/초과근무
현황을 집계하고, 그래프로 시각화한다.

사용법:
    python hr/attendance_dashboard.py
    python hr/attendance_dashboard.py --input 다른파일.csv

입력: hr/sample_data/sample_attendance.csv
      (컬럼: 사번, 이름, 부서, 날짜, 출근시각, 퇴근시각, 지각여부, 초과근무시간)
출력: hr/output/attendance_summary.xlsx  (부서별/월별 집계표)
      hr/output/late_by_dept.png         (부서별 지각 건수 막대그래프)
      hr/output/overtime_trend.png       (부서별 월별 초과근무 추이)
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

HERE = Path(__file__).resolve().parent


def summarize_by_dept(df: pd.DataFrame) -> pd.DataFrame:
    """부서별 근태 요약: 근무일수, 지각건수, 지각률, 총/평균 초과근무시간."""
    grouped = df.groupby("부서").agg(
        근무일수=("날짜", "count"),
        지각건수=("지각여부", lambda s: (s == "Y").sum()),
        총초과근무시간=("초과근무시간", "sum"),
    )
    grouped["지각률"] = grouped["지각건수"] / grouped["근무일수"]
    grouped["평균초과근무시간"] = (grouped["총초과근무시간"] / grouped["근무일수"]).round(2)
    return grouped.reset_index().sort_values("지각률", ascending=False)


def summarize_by_month(df: pd.DataFrame) -> pd.DataFrame:
    """부서 x 월 단위 초과근무시간 합계 (추이 파악용)."""
    df = df.copy()
    df["월"] = pd.to_datetime(df["날짜"]).dt.strftime("%Y-%m")
    return df.groupby(["부서", "월"])["초과근무시간"].sum().reset_index()


def plot_late_by_dept(dept_summary: pd.DataFrame, output_path: Path) -> Path:
    setup_style()
    fig, ax = plt.subplots(figsize=(7, 4.5))
    ax.bar(dept_summary["부서"], dept_summary["지각률"] * 100, color=PALETTE[: len(dept_summary)])
    ax.set_title("부서별 지각률")
    ax.set_ylabel("지각률 (%)")
    for i, v in enumerate(dept_summary["지각률"] * 100):
        ax.text(i, v + 0.3, f"{v:.1f}%", ha="center")
    return save_chart(fig, output_path)


def plot_overtime_trend(monthly: pd.DataFrame, output_path: Path) -> Path:
    setup_style()
    fig, ax = plt.subplots(figsize=(8, 4.5))
    for i, (dept, sub) in enumerate(monthly.groupby("부서")):
        sub = sub.sort_values("월")
        ax.plot(sub["월"], sub["초과근무시간"], marker="o", label=dept, color=PALETTE[i % len(PALETTE)])
    ax.set_title("부서별 월간 초과근무시간 추이")
    ax.set_ylabel("초과근무시간 합계 (h)")
    ax.legend()
    return save_chart(fig, output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="출퇴근 기록으로 부서별 근태 대시보드를 생성한다.")
    parser.add_argument(
        "--input", default=HERE / "sample_data" / "sample_attendance.csv",
        help="출퇴근 기록 CSV 경로 (컬럼: 부서, 날짜, 지각여부, 초과근무시간 등)",
    )
    args = parser.parse_args()

    df = load_csv(args.input)
    dept_summary = summarize_by_dept(df)
    monthly = summarize_by_month(df)

    excel_path = save_excel_report(
        {"부서별_요약": dept_summary, "부서_월별_초과근무": monthly},
        HERE / "output" / "attendance_summary.xlsx",
    )
    late_chart = plot_late_by_dept(dept_summary, HERE / "output" / "late_by_dept.png")
    overtime_chart = plot_overtime_trend(monthly, HERE / "output" / "overtime_trend.png")

    print_section("부서별 근태 요약")
    for _, row in dept_summary.iterrows():
        print(f"{row['부서']:8s}  지각률 {pct(row['지각률'])}  "
              f"(지각 {int(row['지각건수'])}건 / {int(row['근무일수'])}일)  "
              f"평균 초과근무 {row['평균초과근무시간']}h")

    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {late_chart}, {overtime_chart}")


if __name__ == "__main__":
    main()
