# finance — 회계/재무

## income_statement_generator.py

거래 건별(전표 단위) 데이터를 읽어 **K-IFRS 다단계 손익계산서**(홈택스/DART에서
보는 것과 같은 형식: 매출액 → 매출총이익 → 영업이익 → 법인세비용차감전순이익 →
당기순이익)를 만들고, 원하는 기간 단위로 **직전기 대비 / 전년 동기 대비**를
자동 비교합니다.

```bash
# 기본: 데이터에 있는 가장 최근 '월'을 직전월·전년동월과 비교
python finance/income_statement_generator.py

# 기간 단위를 바꿔서 조회 (년 / 반기 / 분기 / 월 / 주)
python finance/income_statement_generator.py --period-type quarter --period 2026-Q2
python finance/income_statement_generator.py --period-type half    --period 2026-H1
python finance/income_statement_generator.py --period-type week    --period 2026-W18
python finance/income_statement_generator.py --period-type year    --period 2026
```

**입력 컬럼**: `거래일자`, `계정과목`, `금액` (전표 단위 — 한 줄이 거래 한 건)
**출력**: `output/income_statement.xlsx` (손익계산서 + 부가세 참고 시트), `output/trend_chart.png`

### 데이터 소스: 구글 스프레드시트 또는 로컬 파일

`.env` 파일에 구글 시트 정보를 채워두면 자동으로 스프레드시트에서 실시간으로
읽어옵니다. 설정이 없으면 `sample_data/sample_transactions.csv`(로컬 파일)를
사용합니다 — 즉 아무 설정 없이 클론만 해도 바로 실행됩니다.

```bash
cp ../.env.example ../.env
# .env를 열어 GOOGLE_SHEET_ID / GOOGLE_SERVICE_ACCOUNT_FILE 채우기
```

회사나 구글 계정이 바뀌어도 `.env` 값만 새로 채우면 되고, 코드는 건드릴 필요가
없습니다.

### 기간 비교 규칙

- **같은 단위끼리만 비교**합니다 (월 vs 월, 분기 vs 분기 ...). 주 단위와 월
  단위처럼 길이가 다른 기간은 비교하지 않습니다.
- **직전기**: 선택한 기간 바로 앞 기간 (예: 2026-05 → 2026-04)
- **전년 동기**: 1년 전 같은 기간 (예: 2026-05 → 2025-05). 주 단위는 ISO
  주차 번호가 아니라 정확히 364일(52주) 전 날짜를 기준으로 계산합니다
  (요일까지 맞아떨어지게).
- 비교 대상 기간에 데이터가 아예 없으면(예: 데이터가 2025-01부터 시작하는데
  2024년과 비교하려는 경우) "데이터 없음"으로 표시됩니다.
- 증감률은 직전 기간 값이 음수에서 양수로(또는 그 반대로) 바뀌는 구간에서는
  숫자가 매우 커지거나 왜곡될 수 있습니다. 이런 경우 증감률보다 증감액을
  우선 참고하세요.

### 법인세비용 / 부가세 추정 (중요)

- **법인세비용**은 이 기간 손익을 연 환산(annualize)해서 국세청 누진세율
  구간(지방소득세 포함, `common/tax_utils.py`)에 대입한 **추정치**입니다.
  세무조정, 이월결손금, 세액공제·감면은 반영되지 않으므로 실제 신고세액과
  다를 수 있습니다. 세율 자체가 개정되면 `common/tax_utils.py`의
  `CORPORATE_TAX_BRACKETS`만 갱신하면 됩니다.
- **부가가치세**는 손익계산서 항목이 아니라서(예수금 성격) 별도
  `부가세_참고` 시트로 분리했습니다. 매출세액은 매출액의 10%, 매입세액은
  인건비를 제외한 매입성 비용의 10%로 단순 추정한 참고값이며, 실제 세금계산서
  발행 기준과는 차이가 있을 수 있습니다.

### 회사마다 다른 값 맞추기

`finance/income_statement_generator.py` 상단의 계정과목 매핑만 실제 회사
계정과목명으로 바꾸면 됩니다.

```python
REVENUE_ACCOUNTS = {"매출"}
COGS_ACCOUNTS = {"매출원가"}
OPEX_ACCOUNTS = {"급여", "임차료", "광고선전비", "소모품비", "지급수수료"}
NON_OP_INCOME_ACCOUNTS = {"이자수익"}
NON_OP_EXPENSE_ACCOUNTS = {"이자비용", "잡손실"}
```
