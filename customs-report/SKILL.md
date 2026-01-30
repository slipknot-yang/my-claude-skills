---
name: customs-report
description: "WCO PMM 기반 관세 분석 보고서 생성 Skill. 경영진 대시보드, KPI 스코어카드, 파레토 분석, 리스크 매트릭스 등 전문 수준의 Excel 보고서를 생성합니다."
version: "1.0.0"
author: "slipknot-yang"
allowed-tools: [
  Read,
  Write,
  Bash
]
---

# 관세 분석 보고서 Skill

WCO PMM(Performance Measurement Model) 프레임워크 기반의 전문 관세 분석 보고서를 생성합니다.

---

## 기능 요약

| 보고서 | 설명 | 시트 구성 |
|--------|------|-----------|
| **관세 수입 현황** | 세수 실적 종합 분석 | 표지, 경영진 대시보드, 연도별, 파레토, 국가별, 월별 |
| **이상 탐지** | 과소신고/리스크 분석 | 표지, 리스크 대시보드, 과소신고, HS-국가, 고위험 업체 |

---

## 사용법

```bash
# 기본 실행 (한국어, 전체 보고서)
/customs-report

# 영어 버전
/customs-report --lang en

# 특정 보고서만
/customs-report --type revenue    # 수입현황만
/customs-report --type anomaly    # 이상탐지만

# 연도 지정
/customs-report --year 2024

# 출력 경로 지정
/customs-report --output /path/to/reports/
```

---

## 사전 요구사항

### Python 패키지
```bash
pip install oracledb pandas openpyxl
```

### DB 접속 정보
```
Host: 211.239.120.42:3535
SID: NTANCIS
User: CLRIUSR
Password: ntancisclri1!
```

---

## Workflow

### Step 1: 환경 확인

```python
# 필수 패키지 확인
import oracledb
import pandas as pd
from openpyxl import Workbook

# DB 연결 테스트
conn = oracledb.connect(
    user="CLRIUSR",
    password="ntancisclri1!",
    dsn="211.239.120.42:3535/NTANCIS"
)
print("DB 연결 성공")
conn.close()
```

### Step 2: 보고서 생성 실행

**한국어 버전:**
```bash
cd /path/to/my-claude-skills
python generate_reports_kr.py
```

**영어 버전:**
```bash
cd /path/to/my-claude-skills
python generate_reports.py
```

### Step 3: 결과 확인

생성된 파일:
- `관세수입현황_보고서_KR.xlsx` (한국어)
- `이상탐지_보고서_KR.xlsx` (한국어)
- `관세수입현황_보고서.xlsx` (영어)
- `이상탐지_보고서.xlsx` (영어)

---

## 보고서 구조

### 관세 수입 현황 보고서

```
├── 표지
│   ├── 관세청 로고
│   ├── 제목/부제목
│   └── KPI 카드 4개 (신고건수, 세수, 수입액, 기간)
│
├── 경영진 대시보드
│   ├── KPI 카드 (건수, 세수, 수입액, 성장률)
│   ├── WCO PMM KPI 스코어카드
│   ├── 리스크 매트릭스 (5x5)
│   └── 주요 발견사항 및 권고사항
│
├── 연도별 추이
│   ├── 데이터 테이블 (연도, 건수, 세액, 성장률)
│   └── 콤보 차트 (막대: 세액, 선: 성장률)
│
├── 파레토 분석
│   ├── HS류별 관세 수입 TOP 20
│   ├── 비중/누적비중/구간 (A/B/C)
│   ├── 히트맵 조건부 서식
│   └── HHI 집중도 지수
│
├── 국가별 현황
│   ├── 원산지 국가 TOP 20
│   ├── 데이터바 시각화
│   └── HHI 집중도 지수
│
├── 월별 추이
│   ├── 최근 36개월 데이터
│   └── 라인 차트
│
├── 분석방법론
│   ├── 데이터 출처
│   ├── KPI 계산 방법론
│   └── 참조 프레임워크
│
└── 용어정의 (한/영)
```

### 이상 탐지 보고서

```
├── 표지
│   └── KPI 카드 (과소신고, 탈루액, 고위험조합, 고위험업체)
│
├── 리스크 대시보드
│   ├── KPI 카드 4개
│   ├── 리스크 매트릭스
│   └── 연도별 과소신고 추이
│
├── 과소신고 분석
│   ├── 연도별 과소신고 통계
│   ├── 히트맵 조건부 서식
│   └── 바 차트
│
├── 품목국가 리스크
│   ├── HS코드 × 국가 조합 TOP 30
│   ├── 리스크 점수 계산
│   └── 히트맵 시각화
│
├── 고위험 업체
│   ├── 과소신고 다발 업체 TOP 30
│   └── 히트맵 시각화
│
├── 품목분류 오류
│   └── 연도별 HS코드 분류오류율
│
├── 분석방법론
└── 용어정의
```

---

## WCO PMM KPI 목록

### 1. Trade Facilitation (무역원활화)
| 코드 | KPI | 단위 | 목표 |
|------|-----|------|------|
| TF001 | 평균 통관시간 | hours | < 12 |
| TF002 | 사전신고처리율 | % | > 50 |
| TF003 | 녹색통로비율 | % | > 85 |
| TF004 | 전자신고율 | % | > 99 |

### 2. Revenue Collection (세수확보)
| 코드 | KPI | 단위 | 목표 |
|------|-----|------|------|
| RC001 | 징수효율 | % | > 99 |
| RC002 | 세액심사정확도 | % | > 95 |
| RC003 | 사후심사커버리지 | % | > 10 |
| RC004 | 세수증가율 | % | > 5 |

### 3. Risk Management (위험관리)
| 코드 | KPI | 단위 | 목표 |
|------|-----|------|------|
| RM001 | 선별율 | % | < 10 |
| RM002 | 적발율 | % | > 30 |
| RM003 | 신고정확도 | % | > 95 |
| RM004 | 과소신고탐지율 | % | > 50 |
| RM005 | 품목분류오류율 | % | < 2 |

### 4. Organizational (조직발전)
| 코드 | KPI | 단위 | 목표 |
|------|-----|------|------|
| OD001 | HHI 집중도 | index | < 1000 |
| OD002 | 월간변동성 | % | < 10 |
| OD003 | 성장일관성 | score | > 4 |

---

## 모듈 구조

```
my-claude-skills/
├── kpi_calculator.py       # WCO PMM KPI 계산 모듈
│   ├── KPICalculator 클래스
│   ├── calc_revenue_by_period()
│   ├── calc_yoy_growth()
│   ├── calc_undervaluation_stats()
│   ├── calc_hhi_by_dimension()
│   ├── calc_pareto_analysis()
│   └── calc_executive_summary()
│
├── visualizations.py       # 시각화 모듈
│   ├── ColorPalette 클래스
│   ├── StyleManager 클래스
│   ├── add_kpi_card()
│   ├── add_risk_matrix()
│   ├── add_scorecard_table()
│   ├── add_heatmap_formatting()
│   ├── add_pareto_chart()
│   └── write_styled_dataframe()
│
├── generate_reports.py      # 영어 보고서 생성기
├── generate_reports_kr.py   # 한국어 보고서 생성기
└── create_styled_reports.py # 기본 보고서 생성기
```

---

## 커스터마이징

### 연도 필터 변경
`kpi_calculator.py`에서 SQL 쿼리의 `TANSAD_YY >= '23'` 조건 수정

### 색상 팔레트 변경
`visualizations.py`의 `ColorPalette` 클래스 수정

### KPI 임계값 변경
`kpi_calculator.py`의 `KPI_DEFINITIONS` 딕셔너리에서 `benchmark`, `target` 값 수정

### 새로운 시트 추가
`generate_reports_kr.py`의 `create_revenue_report()` 또는 `create_anomaly_report()` 메서드에 시트 추가 로직 작성

---

## 에러 처리

### DB 연결 실패
```
❌ DB 연결 실패: ORA-12170: TNS:Connect timeout occurred
```
→ 네트워크 연결 확인, VPN 확인

### 패키지 미설치
```
ModuleNotFoundError: No module named 'oracledb'
```
→ `pip install oracledb pandas openpyxl`

### 메모리 부족
→ `kpi_calculator.py`에서 `FETCH FIRST N ROWS ONLY` 조정

---

## 참조 문서

- [WCO Performance Measurement Model](https://www.wcoomd.org/)
- [WCO Customs Risk Management Compendium](https://www.wcoomd.org/)
- [UN Comtrade Database](https://comtrade.un.org/)
- [관세청 관세연감](https://www.customs.go.kr/)

---

## 변경 이력

| 버전 | 날짜 | 변경 내용 |
|------|------|-----------|
| 1.0.0 | 2026-01-30 | 초기 버전 - 관세 분석 보고서 생성 Skill |
