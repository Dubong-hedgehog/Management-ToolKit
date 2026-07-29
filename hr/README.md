# hr — 인사/총무

## attendance_dashboard.py

출퇴근 원본 기록을 읽어 부서별 지각률, 평균/총 초과근무시간을 집계하고,
지각률 막대그래프와 월별 초과근무 추이 그래프를 생성합니다.

```bash
python hr/attendance_dashboard.py
python hr/attendance_dashboard.py --input path/to/실제출퇴근기록.csv
```

**입력 컬럼**: `사번`, `이름`, `부서`, `날짜`, `지각여부`(Y/N), `초과근무시간`
**출력**: `output/attendance_summary.xlsx`, `output/late_by_dept.png`, `output/overtime_trend.png`

원본 출퇴근 시스템 데이터에 `지각여부`/`초과근무시간` 컬럼이 없다면,
`출근시각`/`퇴근시각` 원본 컬럼으로부터 계산하는 전처리 단계를 앞에 추가해서
사용하면 됩니다.
