# My Claude Skills

Claude Code에서 사용할 수 있는 커스텀 스킬 모음입니다.

## 스킬 목록

| 스킬 | 설명 | 사용법 |
|------|------|--------|
| [notebooklm](./notebooklm/) | Google NotebookLM 자동화 | `/notebooklm <주제>` |

---

## 설치 방법

### 방법 1: 전체 스킬 설치

```bash
# 프로젝트 루트에서 실행
git clone https://github.com/slipknot-yang/my-claude-skills.git
cp -r my-claude-skills/* .claude/skills/
```

### 방법 2: 개별 스킬 설치

```bash
# 원하는 스킬만 다운로드 (예: notebooklm)
mkdir -p .claude/skills/notebooklm
curl -o .claude/skills/notebooklm/SKILL.md \
  https://raw.githubusercontent.com/slipknot-yang/my-claude-skills/main/notebooklm/SKILL.md
```

---

## 스킬 상세

### NotebookLM 자동화 스킬

Google NotebookLM을 자동화하여 웹 리서치부터 콘텐츠 생성까지 한 번에 처리합니다.

#### 주요 기능

- **웹 리서치**: 주제에 맞는 고품질 소스 자동 수집
- **배치 소스 추가**: 여러 URL을 한 번에 NotebookLM에 추가
- **병렬 콘텐츠 생성**: 9가지 콘텐츠 유형 동시 생성
  - AI 오디오 오버뷰 (팟캐스트 스타일)
  - 동영상 개요
  - 마인드맵
  - 보고서
  - 플래시카드
  - 퀴즈
  - 인포그래픽
  - 슬라이드 자료
  - 데이터 표

#### 사전 요구사항

1. **Google 계정**: Playwright 브라우저에서 Google 로그인 필요
2. **NotebookLM 접근**: https://notebooklm.google.com 가입 완료
3. **Claude Code**: Playwright MCP 설정 완료

#### 사용법

```bash
# 기본 사용법 - 주제 연구
/notebooklm Claude Code 베스트 프랙티스

# 소스 개수 지정
/notebooklm 양자 컴퓨팅 --sources 15

# 특정 콘텐츠만 생성
/notebooklm 머신러닝 기초 --outputs audio,video,slides

# 모든 콘텐츠 생성 (기본값)
/notebooklm 기후변화 해결책 --outputs all
```

#### 사용 가능한 옵션

| 옵션 | 기본값 | 설명 |
|------|--------|------|
| `주제` | 필수 | 연구할 주제 |
| `--outputs` | all | 콘텐츠 유형: audio, video, slides, infographic, mindmap, report, flashcards, quiz, table, all |
| `--sources` | 8-12 | 수집할 소스 개수 (최대 50) |
| `--style` | deep-dive | 오디오 스타일: deep-dive, summary, critique, debate |

#### 콘텐츠 생성 시간

| 콘텐츠 유형 | 예상 시간 |
|-------------|-----------|
| 마인드맵 | 1-3분 |
| 보고서 | 1-3분 |
| 플래시카드 | 1-2분 |
| 퀴즈 | 1-2분 |
| 데이터 표 | 1-2분 |
| 인포그래픽 | 2-5분 |
| 슬라이드 자료 | 2-5분 |
| 동영상 개요 | 3-15분 |
| AI 오디오 오버뷰 | 5-30분 |

#### 작동 방식

```
1. 웹 리서치
   └─ 공식 문서, 블로그, YouTube 검색
   └─ 8-12개 고품질 소스 수집

2. NotebookLM 자동화
   └─ 새 노트북 생성
   └─ 소스 일괄 추가
   └─ 소스 개수 확인

3. 콘텐츠 생성
   └─ 9가지 콘텐츠 버튼 연속 클릭
   └─ 모든 콘텐츠 병렬 생성
   └─ 짧은 콘텐츠 1-3분, 긴 콘텐츠 5-30분

4. 결과 반환
   └─ 노트북 URL 제공
   └─ 스크린샷 저장
   └─ 생성 상태 보고
```

---

## 문제 해결

### 인증 오류
NotebookLM에서 로그인 페이지가 나타나면 Playwright 브라우저 창에서 수동으로 Google 로그인 후 "계속"이라고 말씀해주세요.

### 소스 추가 실패
- URL 접근 가능 여부 확인
- NotebookLM이 빠른 추가를 제한할 수 있음
- 30초 대기 후 재시도

### 콘텐츠 생성 멈춤
- 오디오/동영상은 최대 30분 소요될 수 있음
- 브라우저에서 직접 노트북 URL 확인
- 짧은 콘텐츠는 5분 내 완료되어야 함

---

## 기여하기

이슈와 PR 환영합니다!

## 라이선스

MIT License

---

**Created with Claude Code**
