# procurement — 구매/계약

## contract_expiry_tracker.py

계약 현황 데이터를 읽어 만료까지 남은 일수(D-day)를 계산하고, 임박한 계약을
색상으로 강조한 엑셀 리포트와 D-day 막대그래프를 생성합니다.

```bash
python procurement/contract_expiry_tracker.py
python procurement/contract_expiry_tracker.py --as-of 2026-09-01 --warn-days 60
```

**입력 컬럼**: `계약번호`, `거래처`, `계약구분`, `시작일`, `종료일`, `계약금액`, `담당자`
**출력**: `output/contract_expiry_report.xlsx`, `output/expiry_dday_chart.png`

`--warn-days`로 "임박"으로 판단할 기준 일수를 조정할 수 있고, `--as-of`로
특정 시점 기준 시뮬레이션도 가능합니다 (예: 다음 분기 시작일 기준으로 미리 확인).
