---
name: customs-anomaly-detector
description: "관세 이상 탐지 Agent. Oracle DB에서 과소신고, 단가 이상, HS코드 변경 등 리스크 패턴을 탐지하고 고위험 건을 식별합니다. Excel 리포트와 Claude for Excel 분석을 지원합니다."
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

# 관세 이상 탐지 Agent

Oracle DB에서 관세 신고 데이터의 이상 패턴을 탐지하여 과소신고, 분류 오류, 가격 조작 등 리스크를 식별합니다.

---

## 탐지 항목

| 리스크 유형 | 설명 | 탐지 기준 |
|-------------|------|-----------|
| 과소신고 | 신고가 < 심사가 | 심사가가 신고가의 130% 이상 |
| 과대신고 | 신고가 > 심사가 | 신고가가 심사가의 130% 이상 |
| HS코드 변경 | 분류 오류 | 신고 HS ≠ 심사 HS |
| 단가 이상 | 동일품목 가격 이상 | 표준편차 > 평균 * 100% |
| 고위험 조합 | 품목+국가 리스크 | 복합 리스크 점수 |

---

## 사용법

```bash
# 기본 분석
/customs-anomaly-detector

# 특정 연도 분석
/customs-anomaly-detector --year 2024

# 임계값 조정
/customs-anomaly-detector --threshold 1.5

# 출력 파일 지정
/customs-anomaly-detector --output anomaly_report.xlsx
```

---

## 리스크 점수 산정

```
리스크 점수 = (과소신고율 × 3) + (HS변경율 × 2) + (단가편차율 × 1)

등급:
- HIGH (80+): 즉시 점검 필요
- MEDIUM (50-79): 정밀 심사 권장
- LOW (30-49): 모니터링
- NORMAL (0-29): 정상
```

---

## Workflow

### Phase 1: 과소신고 탐지

```sql
SELECT 
    ASSD_HS_CD,
    ORIG_CNTY_CD,
    IMPPN_TIN,
    COUNT(*) as CNT,
    AVG((ASSD_UT_USD_VAL - DCLD_UT_USD_VAL) / NULLIF(DCLD_UT_USD_VAL, 0) * 100) as AVG_DIFF_PCT,
    SUM(ASSD_INVC_USD_AMT - DCLD_INVC_USD_AMT) as TOTAL_DIFF_USD
FROM CLRI_TANSAD_UT_PRC_M
WHERE DEL_YN = 'N'
  AND DCLD_UT_USD_VAL > 0
  AND ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * 1.3  -- 30% 이상 차이
GROUP BY ASSD_HS_CD, ORIG_CNTY_CD, IMPPN_TIN
HAVING COUNT(*) >= 5
ORDER BY TOTAL_DIFF_USD DESC
```

### Phase 2: 단가 이상치 탐지

```sql
SELECT 
    ASSD_HS_CD,
    COUNT(*) as CNT,
    AVG(ASSD_UT_USD_VAL) as AVG_PRICE,
    STDDEV(ASSD_UT_USD_VAL) as STD_PRICE,
    MIN(ASSD_UT_USD_VAL) as MIN_PRICE,
    MAX(ASSD_UT_USD_VAL) as MAX_PRICE
FROM CLRI_TANSAD_UT_PRC_M
WHERE DEL_YN = 'N' AND ASSD_UT_USD_VAL > 0
GROUP BY ASSD_HS_CD
HAVING COUNT(*) >= 50 
   AND STDDEV(ASSD_UT_USD_VAL) > AVG(ASSD_UT_USD_VAL)
ORDER BY STD_PRICE DESC
```

### Phase 3: HS코드 변경 탐지

```sql
SELECT 
    DCLD_HS_CD as DECLARED_HS,
    ASSD_HS_CD as ASSESSED_HS,
    COUNT(*) as CNT,
    SUM(ASSD_INVC_USD_AMT) as TOTAL_VALUE
FROM CLRI_TANSAD_UT_PRC_M
WHERE DEL_YN = 'N'
  AND DCLD_HS_CD IS NOT NULL
  AND ASSD_HS_CD IS NOT NULL
  AND DCLD_HS_CD != ASSD_HS_CD
GROUP BY DCLD_HS_CD, ASSD_HS_CD
HAVING COUNT(*) >= 10
ORDER BY CNT DESC
```

### Phase 4: 종합 리스크 분석

```sql
WITH risk_data AS (
    SELECT 
        SUBSTR(ASSD_HS_CD, 1, 4) as HS4,
        ORIG_CNTY_CD,
        IMPPN_TIN,
        CASE WHEN DCLD_HS_CD != ASSD_HS_CD THEN 1 ELSE 0 END as HS_CHANGED,
        CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * 1.5 AND DCLD_UT_USD_VAL > 0 THEN 1 ELSE 0 END as UNDERVALUED,
        ASSD_INVC_USD_AMT
    FROM CLRI_TANSAD_UT_PRC_M
    WHERE DEL_YN = 'N'
)
SELECT 
    HS4,
    ORIG_CNTY_CD,
    COUNT(*) as TOTAL_CNT,
    SUM(HS_CHANGED) as HS_CHANGE_CNT,
    SUM(UNDERVALUED) as UNDERVALUE_CNT,
    ROUND(SUM(UNDERVALUED) * 100.0 / COUNT(*), 1) as UNDERVALUE_RATE,
    SUM(ASSD_INVC_USD_AMT) as TOTAL_VALUE,
    -- 리스크 점수
    ROUND(SUM(UNDERVALUED) * 3.0 / COUNT(*) * 100 + SUM(HS_CHANGED) * 2.0 / COUNT(*) * 100, 1) as RISK_SCORE
FROM risk_data
WHERE ASSD_HS_CD IS NOT NULL
GROUP BY HS4, ORIG_CNTY_CD
HAVING SUM(UNDERVALUED) >= 10 OR SUM(HS_CHANGED) >= 10
ORDER BY RISK_SCORE DESC
```

### Phase 5: 고위험 업체 식별

```sql
SELECT 
    IMPPN_TIN,
    IMPPN_NM,
    COUNT(*) as TOTAL_CNT,
    SUM(CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * 1.3 THEN 1 ELSE 0 END) as UNDERVALUE_CNT,
    SUM(CASE WHEN DCLD_HS_CD != ASSD_HS_CD THEN 1 ELSE 0 END) as HS_CHANGE_CNT,
    SUM(ASSD_INVC_USD_AMT) as TOTAL_VALUE
FROM CLRI_TANSAD_UT_PRC_M
WHERE DEL_YN = 'N'
GROUP BY IMPPN_TIN, IMPPN_NM
HAVING COUNT(*) >= 20
   AND (SUM(CASE WHEN ASSD_UT_USD_VAL > DCLD_UT_USD_VAL * 1.3 THEN 1 ELSE 0 END) >= 5
        OR SUM(CASE WHEN DCLD_HS_CD != ASSD_HS_CD THEN 1 ELSE 0 END) >= 5)
ORDER BY UNDERVALUE_CNT DESC
```

---

## 출력 결과

### Excel 파일 구조

```
customs_anomaly_report.xlsx
├── 요약                    # 전체 리스크 요약
├── 과소신고_의심           # 과소신고 상세
├── 단가_이상               # 단가 이상치
├── HS코드_변경             # 분류 변경 건
├── 품목국가_리스크         # 품목+국가 조합 리스크
├── 고위험_업체             # 업체별 리스크
└── Claude_분석_가이드      # 분석 프롬프트
```

### 리스크 대시보드

```markdown
## 🚨 이상 탐지 리포트

### 요약
- 분석 기간: 2023-2024
- 분석 건수: X,XXX,XXX 건
- 이상 탐지 건수: XXX,XXX 건 (X.X%)

### 과소신고 의심
- 총 건수: XXX,XXX 건
- 추정 탈루 세액: $XXX,XXX,XXX

### 고위험 품목-국가 조합
| 순위 | 품목 | 국가 | 과소신고율 | 리스크 점수 |
|------|------|------|------------|-------------|
| 1 | 8518 | CN | 51.8% | 156.2 |
| 2 | 8528 | CN | 53.7% | 161.4 |

### 고위험 업체 (TOP 10)
- [업체 리스트]
```

---

## Claude for Excel 분석 프롬프트

| 시트 | 프롬프트 |
|------|----------|
| 과소신고_의심 | "이 과소신고 데이터에서 패턴을 분석하고, 의도적 탈세와 단순 오류를 구분할 수 있는 기준을 제시해주세요" |
| 품목국가_리스크 | "고위험 품목-국가 조합의 특성을 분석하고, 우선 점검 대상을 선정해주세요" |
| 고위험_업체 | "업체별 리스크 패턴을 분석하고, 조사 우선순위를 제안해주세요" |

---

## 에러 처리

### 대용량 쿼리 타임아웃
```python
# 연도별 분할 조회
for year in ['23', '24']:
    query = f"... WHERE TANSAD_YY = '{year}'"
```

### 메모리 부족
```python
# 청크 단위 처리
chunksize = 100000
for chunk in pd.read_sql(query, conn, chunksize=chunksize):
    process(chunk)
```
