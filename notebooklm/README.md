# NotebookLM 자동화 스킬

Google NotebookLM을 자동화하여 웹 리서치부터 9가지 콘텐츠 생성까지 한 번에 처리합니다.

## 빠른 시작

```bash
# 스킬 설치
mkdir -p .claude/skills/notebooklm
curl -o .claude/skills/notebooklm/SKILL.md \
  https://raw.githubusercontent.com/slipknot-yang/my-claude-skills/main/notebooklm/SKILL.md

# 사용
/notebooklm Claude Code 베스트 프랙티스
```

## 사전 요구사항

1. **Google 계정** - Playwright 브라우저에서 로그인 필요
2. **NotebookLM 접근** - https://notebooklm.google.com
3. **MCP 서버** - Playwright, Websearch (선택)

## 생성 가능한 콘텐츠

| 콘텐츠 | 설명 | 시간 |
|--------|------|------|
| AI 오디오 오버뷰 | 팟캐스트 스타일 대화 | 5-30분 |
| 동영상 개요 | AI 내레이션 영상 | 3-15분 |
| 마인드맵 | 시각적 개념 지도 | 1-3분 |
| 보고서 | 구조화된 문서 | 1-3분 |
| 플래시카드 | 학습용 카드 | 1-2분 |
| 퀴즈 | 객관식 문제 | 1-2분 |
| 인포그래픽 | 시각적 요약 | 2-5분 |
| 슬라이드 자료 | 프레젠테이션 | 2-5분 |
| 데이터 표 | 구조화된 데이터 | 1-2분 |

## 사용 예시

```bash
# 기본 (모든 콘텐츠 생성)
/notebooklm 인공지능 윤리

# 소스 개수 지정
/notebooklm 블록체인 기술 --sources 15

# 특정 콘텐츠만
/notebooklm 데이터 분석 --outputs audio,slides

# 오디오 스타일 지정
/notebooklm 스타트업 전략 --style debate
```

## 파일 구조

```
notebooklm/
├── README.md      # 이 파일
└── SKILL.md       # 스킬 정의 (Claude Code가 읽는 파일)
```

---

자세한 내용은 [SKILL.md](./SKILL.md)를 참고하세요.
