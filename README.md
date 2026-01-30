# My Claude Skills

Claude Code에서 사용할 수 있는 커스텀 스킬 모음입니다.

## 스킬 목록

| 스킬 | 설명 | 사용법 |
|------|------|--------|
| [notebooklm](./notebooklm/) | Google NotebookLM 자동화 | `/notebooklm <주제>` |
| [customs-revenue-analyzer](./customs-revenue-analyzer/) | 관세 수입 현황 다각도 분석 | `python analyze_customs_revenue.py` |
| [customs-anomaly-detector](./customs-anomaly-detector/) | 관세 이상 탐지 (과소신고, 분류오류) | `python detect_anomalies.py` |

---

## 설치 방법

### 방법 1: 전체 스킬 설치

```bash
git clone https://github.com/slipknot-yang/my-claude-skills.git
cp -r my-claude-skills/* .claude/skills/
```

### 방법 2: 개별 스킬 설치

```bash
# NotebookLM 스킬
mkdir -p .claude/skills/notebooklm
curl -o .claude/skills/notebooklm/SKILL.md \
  https://raw.githubusercontent.com/slipknot-yang/my-claude-skills/main/notebooklm/SKILL.md

# 관세 분석 스킬
mkdir -p .claude/skills/customs-revenue-analyzer
curl -o .claude/skills/customs-revenue-analyzer/SKILL.md \
  https://raw.githubusercontent.com/slipknot-yang/my-claude-skills/main/customs-revenue-analyzer/SKILL.md
```

---

## 스킬 상세

### 1. NotebookLM 자동화

Google NotebookLM을 자동화하여 웹 리서치부터 9가지 콘텐츠 생성까지 한 번에 처리합니다.

**주요 기능:**
- 웹 리서치 및 소스 수집
- 배치 소스 추가
- 9가지 콘텐츠 병렬 생성 (오디오, 동영상, 슬라이드, 인포그래픽 등)

```bash
/notebooklm Claude Code 베스트 프랙티스 --outputs all
```

---

### 2. 관세 수입 현황 분석 Agent

Oracle DB에서 관세 데이터를 추출하여 다각도 분석을 수행하고, Excel + Claude for Excel로 시각화합니다.

**분석 항목:**
- 연도별 추이 (성장률, 트렌드)
- 품목별 현황 (HS 코드별 세액/비중)
- 국가별 현황 (교역국 순위)
- 세관별 현황
- 월별 추이

**분석 결과 예시:**
```
📦 총 건수: 9,239,728 건
💰 총 세액: 149조
💵 총 수입액: $3,200억

🏆 TOP 5 품목:
1. HS 74 (구리): 76조 (34.1%)
2. HS 27 (광물연료): 42조 (18.8%)
3. HS 87 (차량): 26조 (12.0%)

🌍 TOP 5 교역국:
1. CD (콩고): $1,831억
2. TZ (탄자니아): $742억
3. ZM (잠비아): $534억
```

**사용법:**
```bash
cd customs-revenue-analyzer
pip install oracledb pandas openpyxl
python analyze_customs_revenue.py --output customs_revenue_analysis.xlsx
```

---

### 3. 관세 이상 탐지 Agent

Oracle DB에서 관세 신고 데이터의 이상 패턴을 탐지하여 리스크를 식별합니다.

**탐지 항목:**
| 리스크 유형 | 탐지 기준 |
|-------------|-----------|
| 과소신고 | 심사가 > 신고가 130% |
| 단가 이상 | 표준편차 > 평균 |
| HS코드 변경 | 신고 HS ≠ 심사 HS |
| 종합 리스크 | 복합 점수 |
| 고위험 업체 | 과소신고 다발 |

**탐지 결과 예시:**
```
🚨 이상 탐지 알림 요약

📌 과소신고 의심 TOP 3:
1. 8504 (전기 컨버터) + CN → 2,688건, 1015% 차이, $45억 차액
2. 6402 (신발) + CN → 4,991건, 557% 차이, $36억 차액
3. 7318 (볼트/너트) + CN → 19건, 7887% 차이, $28억 차액

🏢 고위험 업체 TOP 3:
1. OPPO AGENCIES → 45,555건 (56.6% 과소신고), $4.6억
2. TOSH LOGISTICS → 9,714건 (86.2%), $1.2억
3. SGS CARGO → 7,521건 (80.4%), $0.8억
```

**사용법:**
```bash
cd customs-anomaly-detector
pip install oracledb pandas openpyxl
python detect_anomalies.py --output customs_anomaly_report.xlsx
```

---

## Claude for Excel 연동

각 Agent는 Excel 파일을 생성하며, `Claude_분석_가이드` 시트에서 분석 프롬프트를 확인할 수 있습니다.

```excel
=CLAUDE("연도별 관세 수입 추이를 분석하고 향후 전망을 예측해주세요", A1:G10)
=CLAUDE("이 과소신고 데이터에서 의도적 탈세 패턴을 분석해주세요", A1:H50)
```

---

## 사전 요구사항

### NotebookLM 스킬
- Google 계정 (Playwright 브라우저에서 로그인)
- NotebookLM 접근 권한
- Playwright MCP 설정

### 관세 분석 스킬
- Python 3.8+
- Oracle DB 접속 정보
- 패키지: `oracledb`, `pandas`, `openpyxl`

---

## 기여하기

이슈와 PR 환영합니다!

## 라이선스

MIT License

---

**Created with Claude Code**
