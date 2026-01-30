---
name: customs-revenue-analyzer
description: "관세 수입 현황 다각도 분석 Agent. Oracle DB에서 데이터 추출 → Excel 파일 생성 → Claude for Excel 분석 가이드 제공. 연도별, 품목별, 국가별, 세관별 관세 수입 현황을 분석하고 시각화합니다."
allowed-tools: [
  Read,
  Write,
  Bash,
  mcp__playwright__browser_navigate,
  mcp__playwright__browser_snapshot,
  mcp__playwright__browser_click,
  mcp__playwright__browser_type,
  mcp__playwright__browser_take_screenshot
]
---

# 관세 수입 현황 분석 Agent

Oracle DB에서 관세 데이터를 추출하여 다각도 분석을 수행하고, Excel 파일로 내보내 Claude for Excel에서 시각화 및 심층 분석을 지원합니다.

---

## 사전 요구사항

### 1. Python 패키지
```bash
pip install oracledb pandas openpyxl xlsxwriter
```

### 2. Oracle DB 접속 정보
```
Host: 211.239.120.42
Port: 3535
SID: NTANCIS
User: CLRIUSR
Password: ntancisclri1!
```

### 3. Claude for Excel (선택)
- Microsoft 365 Excel 설치
- Claude for Excel Add-in 설치
- Anthropic API Key 설정

---

## 분석 항목

| 분석 유형 | 설명 | 주요 지표 |
|-----------|------|-----------|
| 연도별 추이 | 연간 관세 수입 변화 | 총세액, 건수, 성장률 |
| 품목별 현황 | HS코드별 관세 수입 | TOP 20 품목, 세액 비중 |
| 국가별 현황 | 원산지별 수입 현황 | 교역국 순위, 금액 |
| 세관별 현황 | 통관 세관별 실적 | 세관별 처리량, 세액 |
| 월별 추이 | 월간 변동 패턴 | 계절성, 트렌드 |

---

## 사용법

```bash
# 기본 분석 (모든 항목)
/customs-revenue-analyzer

# 특정 연도 분석
/customs-revenue-analyzer --year 2024

# 특정 분석 항목만
/customs-revenue-analyzer --type yearly,commodity

# Excel 파일 경로 지정
/customs-revenue-analyzer --output /path/to/output.xlsx
```

---

## Workflow

### Phase 1: 데이터 추출

#### Step 1: DB 연결 테스트

```python
import oracledb
import pandas as pd

conn = oracledb.connect(
    user="CLRIUSR",
    password="ntancisclri1!",
    dsn="211.239.120.42:3535/NTANCIS"
)
print("DB 연결 성공")
```

#### Step 2: 연도별 관세 수입 추출

```python
query_yearly = """
SELECT 
    '20' || TANSAD_YY as YEAR,
    COUNT(*) as ITEM_COUNT,
    SUM(ITM_TAX_AMT) as TOTAL_TAX,
    SUM(ITM_INVC_USD_AMT) as TOTAL_VALUE_USD,
    ROUND(AVG(ITM_TAX_AMT), 0) as AVG_TAX
FROM CLRI_TANSAD_ITM_D
WHERE DEL_YN = 'N' AND TANSAD_YY >= '20'
GROUP BY TANSAD_YY
ORDER BY TANSAD_YY DESC
"""
df_yearly = pd.read_sql(query_yearly, conn)
```

#### Step 3: 품목별(HS2) 관세 수입 추출

```python
query_commodity = """
SELECT 
    SUBSTR(HS_CD, 1, 2) as HS2_CODE,
    COUNT(*) as ITEM_COUNT,
    SUM(ITM_TAX_AMT) as TOTAL_TAX,
    SUM(ITM_INVC_USD_AMT) as TOTAL_VALUE_USD,
    ROUND(AVG(ITM_TAX_AMT), 0) as AVG_TAX
FROM CLRI_TANSAD_ITM_D
WHERE DEL_YN = 'N' AND ITM_TAX_AMT > 0
GROUP BY SUBSTR(HS_CD, 1, 2)
ORDER BY TOTAL_TAX DESC
FETCH FIRST 30 ROWS ONLY
"""
df_commodity = pd.read_sql(query_commodity, conn)
```

#### Step 4: 국가별 수입 현황 추출

```python
query_country = """
SELECT 
    ORIG_CNTY_CD as COUNTRY_CODE,
    COUNT(*) as ITEM_COUNT,
    SUM(ITM_TAX_AMT) as TOTAL_TAX,
    SUM(ITM_INVC_USD_AMT) as TOTAL_VALUE_USD
FROM CLRI_TANSAD_ITM_D
WHERE DEL_YN = 'N' AND ORIG_CNTY_CD IS NOT NULL
GROUP BY ORIG_CNTY_CD
ORDER BY TOTAL_VALUE_USD DESC NULLS LAST
FETCH FIRST 30 ROWS ONLY
"""
df_country = pd.read_sql(query_country, conn)
```

#### Step 5: 세관별 현황 추출

```python
query_customs = """
SELECT 
    CSTM_OFCE_CD as CUSTOMS_OFFICE,
    COUNT(*) as ITEM_COUNT,
    SUM(ITM_TAX_AMT) as TOTAL_TAX,
    SUM(ITM_INVC_USD_AMT) as TOTAL_VALUE_USD
FROM CLRI_TANSAD_ITM_D
WHERE DEL_YN = 'N'
GROUP BY CSTM_OFCE_CD
ORDER BY TOTAL_TAX DESC NULLS LAST
"""
df_customs = pd.read_sql(query_customs, conn)
```

#### Step 6: 월별 추이 (최근 2년)

```python
query_monthly = """
SELECT 
    TO_CHAR(FRST_RGSR_DTM, 'YYYY-MM') as MONTH,
    COUNT(*) as ITEM_COUNT,
    SUM(ITM_TAX_AMT) as TOTAL_TAX,
    SUM(ITM_INVC_USD_AMT) as TOTAL_VALUE_USD
FROM CLRI_TANSAD_ITM_D
WHERE DEL_YN = 'N' 
  AND FRST_RGSR_DTM >= ADD_MONTHS(SYSDATE, -24)
GROUP BY TO_CHAR(FRST_RGSR_DTM, 'YYYY-MM')
ORDER BY MONTH
"""
df_monthly = pd.read_sql(query_monthly, conn)
```

### Phase 2: Excel 파일 생성

```python
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows

output_path = "customs_revenue_analysis.xlsx"

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # 각 시트에 데이터 저장
    df_yearly.to_excel(writer, sheet_name='연도별_추이', index=False)
    df_commodity.to_excel(writer, sheet_name='품목별_현황', index=False)
    df_country.to_excel(writer, sheet_name='국가별_현황', index=False)
    df_customs.to_excel(writer, sheet_name='세관별_현황', index=False)
    df_monthly.to_excel(writer, sheet_name='월별_추이', index=False)
    
    # 요약 시트 생성
    summary_data = {
        '지표': ['총 건수', '총 세액', '총 수입액(USD)', '분석 기간'],
        '값': [
            df_yearly['ITEM_COUNT'].sum(),
            df_yearly['TOTAL_TAX'].sum(),
            df_yearly['TOTAL_VALUE_USD'].sum(),
            f"{df_yearly['YEAR'].min()} ~ {df_yearly['YEAR'].max()}"
        ]
    }
    pd.DataFrame(summary_data).to_excel(writer, sheet_name='요약', index=False)

print(f"Excel 파일 생성 완료: {output_path}")
```

### Phase 3: 차트 추가 (openpyxl)

```python
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference

wb = load_workbook(output_path)

# 연도별 추이 차트
ws = wb['연도별_추이']
chart = BarChart()
chart.title = "연도별 관세 수입 추이"
chart.x_axis.title = "연도"
chart.y_axis.title = "세액"
data = Reference(ws, min_col=3, min_row=1, max_row=ws.max_row, max_col=3)
cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "G2")

# 월별 추이 라인 차트
ws = wb['월별_추이']
chart = LineChart()
chart.title = "월별 관세 수입 추이"
data = Reference(ws, min_col=3, min_row=1, max_row=ws.max_row, max_col=3)
cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
ws.add_chart(chart, "G2")

wb.save(output_path)
print("차트 추가 완료")
```

### Phase 4: Claude for Excel 분석 가이드

Excel 파일을 열고 Claude for Excel Add-in을 사용하여 다음 분석을 수행합니다:

#### 분석 프롬프트 예시

**요약 시트에서:**
```
이 관세 수입 데이터를 분석하고 주요 인사이트를 3가지로 요약해주세요.
```

**연도별 추이 시트에서:**
```
=CLAUDE("연도별 관세 수입 추이를 분석하고, 성장률과 향후 전망을 예측해주세요", A1:E10)
```

**품목별 현황 시트에서:**
```
=CLAUDE("TOP 10 품목의 관세 수입 비중을 분석하고, 각 품목의 특징을 설명해주세요", A1:E11)
```

**국가별 현황 시트에서:**
```
=CLAUDE("주요 교역국 분석 및 국가별 수입 특성을 설명해주세요", A1:D20)
```

---

## 출력 결과

### Excel 파일 구조

```
customs_revenue_analysis.xlsx
├── 요약              # 전체 요약 지표
├── 연도별_추이       # 연간 관세 수입 + 차트
├── 품목별_현황       # HS2 코드별 현황
├── 국가별_현황       # 원산지별 현황
├── 세관별_현황       # 통관 세관별 실적
└── 월별_추이         # 월간 변동 + 차트
```

### 분석 리포트

```markdown
## 관세 수입 현황 분석 리포트

### 1. 요약
- 분석 기간: 2020 ~ 2025
- 총 건수: X,XXX,XXX 건
- 총 세액: XX조 원
- 총 수입액: XXX억 달러

### 2. 연도별 추이
- 최근 성장률: +X.X%
- 최대 수입 연도: 20XX년

### 3. 주요 품목 (TOP 5)
1. HS 74 (구리): XX조 원 (XX%)
2. HS 27 (광물연료): XX조 원 (XX%)
3. ...

### 4. 주요 교역국 (TOP 5)
1. CD (콩고): $XXX억
2. TZ (탄자니아): $XXX억
3. ...

### 5. 인사이트 및 제언
- [Claude for Excel 분석 결과 삽입]
```

---

## HS 코드 참조

| HS2 | 품목명 |
|-----|--------|
| 74 | 구리와 그 제품 |
| 27 | 광물성 연료, 광물유 |
| 87 | 철도/전차 외의 차량 |
| 26 | 광, 슬래그 및 회 |
| 84 | 원자로, 보일러, 기계류 |
| 85 | 전기기기, 녹음/재생기기 |
| 39 | 플라스틱과 그 제품 |
| 72 | 철강 |
| 24 | 담배와 제조한 담배 대용물 |
| 73 | 철강의 제품 |

---

## 국가 코드 참조

| 코드 | 국가명 |
|------|--------|
| CD | 콩고민주공화국 |
| TZ | 탄자니아 |
| ZM | 잠비아 |
| AE | 아랍에미리트 |
| CN | 중국 |
| IN | 인도 |
| JP | 일본 |
| ZA | 남아프리카공화국 |
| SA | 사우디아라비아 |
| US | 미국 |
| KE | 케냐 |
| KR | 한국 |
| DE | 독일 |

---

## 에러 처리

### DB 연결 실패
```python
try:
    conn = oracledb.connect(...)
except oracledb.Error as e:
    print(f"DB 연결 실패: {e}")
    # 네트워크 확인, 접속 정보 확인
```

### 대용량 쿼리 타임아웃
```python
# 연도별로 분할 조회
for year in ['23', '24', '25']:
    query = f"... WHERE TANSAD_YY = '{year}'"
    df = pd.read_sql(query, conn)
```

---

## 제한사항

1. **데이터 보안**: DB 접속 정보는 환경변수로 관리 권장
2. **대용량 처리**: 전체 데이터 조회 시 메모리 주의 (샘플링 또는 집계 쿼리 사용)
3. **Claude for Excel**: API 호출 제한 있음 (분당 요청 수)
