# Customs Report Skill

WCO PMM 기반 관세 분석 보고서 생성 Skill

## 개요

Oracle DB에서 관세 데이터를 추출하여 WCO(세계관세기구) PMM(Performance Measurement Model) 프레임워크 기반의 전문 수준 Excel 보고서를 생성합니다.

## 주요 기능

- **경영진 대시보드**: KPI 카드, 스코어카드, 리스크 매트릭스
- **WCO PMM KPI**: 16개 핵심성과지표 자동 계산
- **고급 시각화**: 파레토 차트, 히트맵, 콤보 차트
- **이상 탐지**: 과소신고, 품목분류 오류, 고위험 업체 분석

## 사용법

```bash
# Claude Code에서
/customs-report

# 또는 직접 실행
python generate_reports_kr.py   # 한국어
python generate_reports.py      # 영어
```

## 생성되는 보고서

| 보고서 | 파일명 | 내용 |
|--------|--------|------|
| 관세 수입 현황 | `관세수입현황_보고서_KR.xlsx` | 세수 실적, 추이, 파레토 분석 |
| 이상 탐지 | `이상탐지_보고서_KR.xlsx` | 과소신고, 리스크, 고위험 업체 |

## 요구사항

```bash
pip install oracledb pandas openpyxl
```

## 모듈 구성

```
my-claude-skills/
├── kpi_calculator.py        # KPI 계산 모듈 (16개 WCO PMM KPI)
├── visualizations.py        # 시각화 모듈 (차트, 히트맵, 리스크 매트릭스)
├── generate_reports_kr.py   # 한국어 보고서 생성기
├── generate_reports.py      # 영어 보고서 생성기
└── customs-report/
    ├── SKILL.md             # Skill 정의
    └── README.md            # 이 파일
```

## 참조

- [WCO Performance Measurement Model](https://www.wcoomd.org/)
- [WCO Customs Risk Management Compendium](https://www.wcoomd.org/)
- [UN Comtrade Database](https://comtrade.un.org/)

## 라이선스

MIT License
