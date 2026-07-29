# finance — 회계/재무

## income_statement_generator.py

월별 거래 내역(계정과목, 금액)을 읽어 매출/매출원가/판관비/영업이익을 계산하고,
간이 손익계산서 엑셀과 매출·영업이익 추이 차트를 생성합니다.

```bash
python finance/income_statement_generator.py
python finance/income_statement_generator.py --input path/to/실제거래내역.xlsx
```

**입력 컬럼**: `거래월`, `계정과목`, `구분`(수익/비용), `금액`
**출력**: `output/income_statement.xlsx`, `output/revenue_trend.png`

계정과목 이름이 회사마다 다르면 스크립트 상단의 `REVENUE_ACCOUNTS`,
`COGS_ACCOUNTS`, `OPEX_ACCOUNTS`만 실제 계정과목명으로 바꿔주면 됩니다.
