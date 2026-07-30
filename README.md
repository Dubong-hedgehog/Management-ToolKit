# biz-support-toolkit

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/Dubong-hedgehog/Management-ToolKit/blob/main/notebooks/finance_demo.ipynb)   < -- 조작예시

경영지원팀(회계·재무 / 인사 / 총무 / 구매·계약)에서 반복적으로 발생하는 업무를
자동화하는 Python 스크립트 모음입니다.
실제 사내 프로세스를 일반화된 로직으로 만들고 데이터는 샘플 데이터를 붙여놓은 상태로
**클론 후 바로 실행하면 결과(엑셀 리포트, 차트)를 눈으로 확인**할 수 있도록 만들었습니다.

> 모든 데이터는 샘플 데이터로 실 회사의 데이터와 관계 없습니다.
> 계정과목/컬럼명 등 회사 맞춤형으로 변경하여 운영 가능.


## 예제

| 카테고리 | 스크립트 | 하는 일 |
|---|---|---|
| [`finance/`](finance) | `income_statement_generator.py` | 월별 거래 내역 → 기간(년/반기/분기/월/주) 자유 비교 손익계산서 + 매출/영업이익 추이 차트 |
| [`finance/`](finance) | `financial_statements_pdf.py` | 거래내역 + 잔액 데이터 → K-GAAP 공식 서식 손익계산서/재무상태표 PDF |
| [`hr/`](hr) | `attendance_dashboard.py` | 출퇴근 기록 → 부서별 지각률/초과근무 기간 비교 + 그래프 |
| [`hr/`](hr) | `attendance_pattern_analysis.py` | 출퇴근 기록 → 직원별 지각/야근 이상패턴 탐지 + 근로복지공단 기준 야근 위험단계 |
| [`procurement/`](procurement) | `contract_wbs_tracker.py` | 계약 현황 → 계약전~종료 WBS 진행률 + 지연/만료임박(D-30/D-60) 알림 |
| [`procurement/`](procurement) | `vendor_purchase_analysis.py` | 구매내역 → 구매처 ABC등급/집중도(HHI) + 단가 급등 이상탐지 |
| [`general_affairs/`](general_affairs) | `asset_lifecycle_tracker.py` | 비품/고정자산/법인차량 → 취득~폐기 생애주기 + 감가상각 + 보험갱신 임박 알림 |
| [`general_affairs/`](general_affairs) | `safety_training_tracker.py` | 교육 이수기록 → 법정의무교육(산업안전보건법) 이수율 + 기한경과 미이수자 알림 |
| [`general_affairs/`](general_affairs) | `general_affairs_ledger.py` | 시설계약/문서·인장/행사·복리후생 → 통합 관리대장 + 계약만료·문서폐기 알림 |

### 실행 결과 미리보기

**회계/재무 — 매출·영업이익 추이**

![income statement trend](docs/screenshots/finance_revenue_trend.png)

**인사 — 직원별 지각·야근 현황 및 위험단계**

![attendance pattern overview](docs/screenshots/hr_person_overview.png)

**구매/계약 — 진행중 계약 현황 및 만료 임박(D-30/D-60)**

![contract gantt](docs/screenshots/procurement_contract_gantt.png)

**총무 — 자산 카테고리별 장부가액 및 상태 분포**

![asset lifecycle status](docs/screenshots/general_affairs_asset_status.png)

## 시작하기

```bash
git clone <이 저장소 주소>
cd Management ToolKit
pip install -r requirements.txt

# 예제 1: 손익계산서 (기간 자유 비교)
python finance/income_statement_generator.py

# 예제 2: 근태 대시보드 (부서별 기간 비교)
python hr/attendance_dashboard.py

# 예제 3: 근태 이상패턴 분석 (개인/부서 이상탐지 + 야근 위험단계)
python hr/attendance_pattern_analysis.py

# 예제 4: 계약 WBS 트래커 (진행률/지연/만료임박 알림)
python procurement/contract_wbs_tracker.py

# 예제 5: 구매처 패턴 분석 (ABC/HHI/단가급등)
python procurement/vendor_purchase_analysis.py

# 예제 6: 자산 생애주기 관리 (비품/고정자산/법인차량 + 보험갱신 알림)
python general_affairs/asset_lifecycle_tracker.py

# 예제 7: 법정의무교육 이수현황 (개인/부서 이수율 + 미이수 알림)
python general_affairs/safety_training_tracker.py

# 예제 8: 총무 통합 관리대장 (시설계약/문서관리/행사·복리후생)
python general_affairs/general_affairs_ledger.py
```

각 스크립트는 `sample_data/` 안의 가짜 데이터로 바로 실행되며, 결과는 같은
카테고리 폴더의 `output/`에 저장됩니다. 실제 데이터로 돌리고 싶다면 `--input`
옵션으로 파일 경로만 바꿔주면 됩니다 (컬럼 이름은 각 스크립트 상단 docstring
참고).

## 폴더 구조

```
biz-support-toolkit/
├── common/              # 카테고리 전반에서 재사용하는 유틸
│   ├── excel_io.py       # 엑셀 읽기/쓰기 (자동필터 포함)
│   ├── chart_style.py    # 차트 스타일/한글 폰트/저장(dpi 고정)
│   ├── format_utils.py   # 숫자/퍼센트 포맷팅
│   ├── period_utils.py   # 년/반기/분기/월/주 기간 비교 로직
│   ├── tax_utils.py      # 법인세/부가세 추정 계산
│   ├── sheet_io.py       # 구글시트 연동(.env, 미설정시 로컬 CSV로 자동 대체)
│   ├── pdf_statement.py  # K-GAAP 서식 PDF 렌더러
│   └── notify_utils.py   # 이메일/슬랙/팀즈 알림(.env, 미설정시 콘솔+로그로 대체)
├── finance/             # 회계·재무 업무
│   ├── sample_data/
│   ├── output/           # 실행 시 생성 (git에는 포함 안 함)
│   ├── income_statement_generator.py
│   └── financial_statements_pdf.py
├── hr/                  # 인사 업무
│   ├── sample_data/
│   ├── output/
│   ├── attendance_dashboard.py
│   └── attendance_pattern_analysis.py
├── procurement/         # 구매·계약 업무
│   ├── sample_data/       # contracts.csv / wbs_tasks.csv / purchases.csv
│   ├── output/
│   ├── contract_wbs_tracker.py
│   └── vendor_purchase_analysis.py
├── general_affairs/     # 총무 업무
│   ├── sample_data/       # assets.csv / training_records.csv / facility_contracts.csv 등
│   ├── output/
│   ├── asset_lifecycle_tracker.py
│   ├── safety_training_tracker.py
│   └── general_affairs_ledger.py
├── notebooks/            # Colab에서 바로 열어 실행해보는 인터랙티브 데모
│   └── finance_demo.ipynb
├── docs/screenshots/    # README용 결과 미리보기 이미지 (고정 보관)
├── CONVENTIONS.md       # 새 스크립트/카테고리 추가 시 지키는 규칙
└── requirements.txt
```
