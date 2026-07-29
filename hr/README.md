# hr — 인사/총무

## attendance_dashboard.py

일별 출퇴근 기록을 읽어 부서별 지각률·평균 초과근무시간을 집계하고, 원하는
기간 단위(년/반기/분기/월/주)로 **직전기 대비 / 전년 동기 대비**를 비교합니다.
데이터소스와 기간비교 구조는 `finance/income_statement_generator.py`와 동일한
공통 모듈(`common/sheet_io.py`, `common/period_utils.py`)을 사용합니다.

```bash
# 기본: 가장 최근 '월' 기준
python hr/attendance_dashboard.py

# 기간 단위를 바꿔서 조회
python hr/attendance_dashboard.py --period-type quarter --period 2026-Q2
python hr/attendance_dashboard.py --period-type week    --period 2026-W18
```

**입력 컬럼**: `사번`, `이름`, `부서`, `거래일자`, `지각여부`(Y/N), `초과근무시간`
**출력**: `output/attendance_summary.xlsx` (부서×지표 기간비교 표), `output/late_by_dept.png`, `output/overtime_trend.png`

### 데이터 소스: 구글 스프레드시트 또는 로컬 파일

프로젝트 루트의 `.env`에 구글 시트 정보가 채워져 있으면 자동으로 그 시트에서
읽고, 없으면 `sample_data/sample_attendance.csv`를 사용합니다. finance와
`.env`를 공유하므로, 출퇴근 기록을 다른 시트에 관리한다면 `GOOGLE_WORKSHEET_NAME`을
스크립트 실행 시점에 맞게 바꾸거나 `common/sheet_io.py`의 함수를 직접 인자와
함께 호출하도록 조정하면 됩니다.

### 지표

- **지각률**: 지각건수 / 근무일수
- **평균초과근무시간**: 총초과근무시간 / 근무일수

원본 출퇴근 시스템 데이터에 `지각여부`/`초과근무시간` 컬럼이 없고
`출근시각`/`퇴근시각` 원본만 있다면, 그걸로 자동 계산하는 전처리 단계를
`load_attendance()` 앞에 추가해서 쓰면 됩니다.

### 참고

주 단위 비교는 특정 주에 휴가/공휴일이 몰려 있으면 표본이 적어 지각률·초과근무
수치가 크게 흔들릴 수 있습니다. 안정적인 추세 파악에는 월 단위 이상을
권장합니다.
