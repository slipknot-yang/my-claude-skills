# 관세 수입 현황 분석 Agent

Oracle DB에서 관세 데이터를 추출하여 다각도 분석을 수행하고, Excel 파일로 내보내 Claude for Excel에서 심층 분석을 지원합니다.

## 빠른 시작

```bash
# 패키지 설치
pip install oracledb pandas openpyxl

# 분석 실행
python analyze_customs_revenue.py

# 결과 파일: customs_revenue_analysis.xlsx
```

## 분석 항목

| 시트 | 내용 |
|------|------|
| 요약 | 전체 요약 지표 |
| 연도별_추이 | 연간 관세 수입 + 성장률 + 차트 |
| 품목별_현황 | HS2 코드별 세액/비중 |
| 국가별_현황 | 원산지별 수입액/비중 |
| 세관별_현황 | 통관 세관별 실적 |
| 월별_추이 | 월간 변동 + 차트 |
| Claude_분석_가이드 | Claude for Excel 프롬프트 |

## 분석 결과 예시 (2020-2026)

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

## Claude for Excel 연동

Excel 파일의 'Claude_분석_가이드' 시트에서 프롬프트를 확인하세요:

```excel
=CLAUDE("연도별 관세 수입 추이를 분석하고 향후 전망을 예측해주세요", A1:G10)
```

## 파일 구조

```
customs-revenue-analyzer/
├── README.md
├── SKILL.md                         # Claude Code 스킬 정의
├── analyze_customs_revenue.py       # 분석 스크립트
└── customs_revenue_analysis.xlsx    # 결과 파일
```
