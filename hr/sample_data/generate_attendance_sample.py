"""
generate_attendance_sample.py
sample_attendance.csv를 생성하는 스크립트.

실제 배포된 근태 대시보드에서 검증한 규칙을 그대로 사용해서 샘플 데이터를
만든다 (2026-07 실사 과정에서 확정된 규칙):

    지각 = 출근시각이 오전 10:00보다 늦은 만큼(분)
    초과근무 = 출근시각이 09:30~10:00 유연근무 구간일 때만 인정.
               정규퇴근시각 = 출근시각 + 9시간, 초과근무 = 실제퇴근 - 정규퇴근시각

즉 08:50에 출근해서 22:00에 퇴근해도(유연근무 구간 밖이라) 초과근무는 0으로
잡힌다 — 실제 회사 규정이 그렇다. 이 스크립트는 지각여부/초과근무시간 컬럼을
이 규칙으로 직접 계산해서 CSV에 저장한다. attendance_pattern_analysis.py는
(실제 시트의 L/M열처럼) 이미 계산된 이 값을 그대로 신뢰하고 재계산하지 않는다.

아래 몇 명은 일부러 특정 패턴을 심어서, 이상탐지 로직(신규경고/경고지속/
개선됨, 부서 이상치, 야근 위험단계)이 실제로 뭔가를 잡아내는 걸 보여준다.
나머지는 정상 베이스라인이다.

Usage:
    python hr/sample_data/generate_attendance_sample.py
"""
from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUTPUT = HERE / "sample_attendance.csv"

START_DATE = date(2025, 1, 1)
END_DATE = date(2026, 7, 31)

LATE_THRESHOLD_MIN = 10 * 60       # 오전 10:00
FLEX_START_MIN = 9 * 60 + 30       # 09:30
FLEX_END_MIN = 10 * 60             # 10:00
STANDARD_SPAN_MIN = 9 * 60         # 출근 + 9시간 = 정규퇴근

EMPLOYEES = [
    (1001, "영업직원1", "영업팀"), (1002, "영업직원2", "영업팀"), (1003, "영업직원3", "영업팀"),
    (1004, "영업직원4", "영업팀"), (1005, "영업직원5", "영업팀"),
    (1006, "개발직원1", "개발팀"), (1007, "개발직원2", "개발팀"), (1008, "개발직원3", "개발팀"),
    (1009, "개발직원4", "개발팀"), (1010, "개발직원5", "개발팀"),
    (1011, "경영직원1", "경영지원팀"), (1012, "경영직원2", "경영지원팀"), (1013, "경영직원3", "경영지원팀"),
    (1014, "경영직원4", "경영지원팀"), (1015, "경영직원5", "경영지원팀"),
    (1016, "마케직원1", "마케팅팀"), (1017, "마케직원2", "마케팅팀"), (1018, "마케직원3", "마케팅팀"),
    (1019, "마케직원4", "마케팅팀"), (1020, "마케직원5", "마케팅팀"),
]

SCENARIO_ONGOING_WARNING = 1003
SCENARIO_NEW_WARNING = 1007
SCENARIO_IMPROVED = 1016
DEPT_ANOMALY_DEPT = "경영지원팀"
DEPT_ANOMALY_MONTH = (2026, 3)

OT_VERY_HIGH = 1009
OT_HIGH = 1010
OT_CAUTION = 1013

random.seed(42)


def business_days(start: date, end: date):
    d = start
    while d <= end:
        if d.weekday() < 5:
            yield d
        d += timedelta(days=1)


def clamp_min(m: int) -> int:
    return max(0, min(m, 27 * 60))


def fmt_time(total_min: int) -> str:
    total_min = clamp_min(total_min)
    h, m = divmod(total_min, 60)
    return f"{h}:{m:02d}"


def gen_checkin_baseline() -> int:
    r = random.random()
    if r < 0.05:
        return random.randint(FLEX_START_MIN, FLEX_END_MIN)
    if r < 0.07:
        return random.randint(LATE_THRESHOLD_MIN + 1, LATE_THRESHOLD_MIN + 30)
    return int(random.gauss(8 * 60 + 55, 12))


def gen_checkout_baseline(checkin: int) -> int:
    return int(random.gauss(18 * 60 + 30, 40))


def gen_checkin_late_burst() -> int:
    return random.randint(LATE_THRESHOLD_MIN + 1, LATE_THRESHOLD_MIN + 40)


def gen_checkin_flex_heavy(mean: int, sd: int) -> int:
    return int(max(FLEX_START_MIN, min(FLEX_END_MIN, random.gauss(mean, sd))))


def build_rows():
    rows = []
    days = list(business_days(START_DATE, END_DATE))

    for emp_id, name, dept in EMPLOYEES:
        for d in days:
            ym = (d.year, d.month)
            checkin = gen_checkin_baseline()
            checkout = gen_checkout_baseline(checkin)

            if emp_id == SCENARIO_ONGOING_WARNING and ym >= (2026, 5):
                if random.random() < 0.55:
                    checkin = gen_checkin_late_burst()
                    checkout = int(random.gauss(18 * 60 + 40, 30))

            if emp_id == SCENARIO_NEW_WARNING and ym == (2026, 7):
                if random.random() < 0.60:
                    checkin = gen_checkin_late_burst()
                    checkout = int(random.gauss(18 * 60 + 40, 30))

            if emp_id == SCENARIO_IMPROVED and ym == (2026, 6):
                if random.random() < 0.55:
                    checkin = gen_checkin_late_burst()
                    checkout = int(random.gauss(18 * 60 + 40, 30))

            if dept == DEPT_ANOMALY_DEPT and ym == DEPT_ANOMALY_MONTH:
                if random.random() < 0.50:
                    checkin = gen_checkin_late_burst()

            heavy_ot_start = date(2026, 5, 11)
            if d >= heavy_ot_start:
                if emp_id == OT_VERY_HIGH:
                    checkin = gen_checkin_flex_heavy(9 * 60 + 31, 2)
                    checkout = int(random.gauss(24 * 60 + 30, 15))
                elif emp_id == OT_HIGH:
                    checkin = gen_checkin_flex_heavy(9 * 60 + 36, 3)
                    checkout = int(random.gauss(22 * 60 + 50, 12))
                elif emp_id == OT_CAUTION:
                    checkin = gen_checkin_flex_heavy(9 * 60 + 50, 3)
                    checkout = int(random.gauss(21 * 60 + 45, 12))

            checkin = clamp_min(checkin)
            checkout = clamp_min(max(checkout, checkin + 60))

            late_min = max(0, checkin - LATE_THRESHOLD_MIN)
            ot_min = 0
            if FLEX_START_MIN <= checkin <= FLEX_END_MIN:
                normal_end = checkin + STANDARD_SPAN_MIN
                if checkout > normal_end:
                    ot_min = checkout - normal_end

            rows.append({
                "사번": emp_id,
                "이름": name,
                "부서": dept,
                "거래일자": d.isoformat(),
                "출근시각": fmt_time(checkin),
                "퇴근시각": fmt_time(checkout),
                "지각여부": "Y" if late_min > 0 else "N",
                "초과근무시간": round(ot_min / 60, 2),
            })
    return rows


def main():
    import csv

    rows = build_rows()
    with open(OUTPUT, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "사번", "이름", "부서", "거래일자", "출근시각", "퇴근시각", "지각여부", "초과근무시간",
        ])
        writer.writeheader()
        writer.writerows(rows)
    print(f"생성 완료: {OUTPUT} ({len(rows)}행)")


if __name__ == "__main__":
    main()
