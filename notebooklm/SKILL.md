---
name: notebooklm
description: "Automate NotebookLM research workflow: research a topic, gather sources, add to NotebookLM, generate ALL content types (audio, video, slides, infographic, etc.) in parallel. REQUIRES: Google account logged in via Playwright browser. Use when asked to 'research in NotebookLM', 'create NotebookLM notebook about X', or 'make a podcast about Y'."
allowed-tools: [
  Read,
  Write,
  Bash,
  WebFetch,
  mcp__websearch__web_search_exa,
  mcp__google_search__search,
  mcp__playwright__browser_navigate,
  mcp__playwright__browser_snapshot,
  mcp__playwright__browser_click,
  mcp__playwright__browser_type,
  mcp__playwright__browser_take_screenshot,
  mcp__playwright__browser_wait_for,
  mcp__playwright__browser_evaluate,
  mcp__playwright__browser_press_key
]
---

# NotebookLM Research Automation Skill

Automate the creation of Google NotebookLM notebooks with curated web sources and **parallel generation of ALL content types**.

---

## Prerequisites (REQUIRED)

### 1. Google Account
- NotebookLM requires a Google account
- Must be logged into Google in the browser used by Playwright

### 2. NotebookLM Access
- Visit https://notebooklm.google.com and sign up
- Ensure you can create notebooks manually first

### 3. Browser Session Setup

**Option A: Manual Login (Recommended for first use)**
```
1. Run the skill
2. When prompted "Authentication required", log in manually in the browser window
3. After login, say "continue" to proceed
```

**Option B: Pre-authenticated Browser Profile**
```bash
# Save authenticated session after manual login
# Configure Playwright MCP to use persistent browser context
# In your MCP config, set userDataDir to preserve login state
```

### 4. MCP Servers Required
- `playwright` - Browser automation
- `websearch` or `google_search` - Web research (optional but recommended)

---

## Available Content Types

NotebookLM Studio provides **9 content types**:

| Content Type | Korean UI | Description | Generation Time |
|--------------|-----------|-------------|-----------------|
| AI Audio Overview | AI 오디오 오버뷰 | Podcast-style deep dive conversation | 5-30 min |
| Video Overview | 동영상 개요 | AI-narrated explainer video | 3-15 min |
| Mind Map | 마인드맵 | Visual concept mapping | 1-3 min |
| Report | 보고서 | Structured document | 1-3 min |
| Flashcards | 플래시카드 | Study cards for learning | 1-2 min |
| Quiz | 퀴즈 | Multiple choice questions | 1-2 min |
| Infographic | 인포그래픽 | Visual summary graphic | 2-5 min |
| Slides | 슬라이드 자료 | Presentation slides | 2-5 min |
| Data Table | 데이터 표 | Structured data extraction | 1-2 min |

---

## Input

You will receive a **topic** and optional **output types**. Examples:

```
/notebooklm Claude Code documentation
/notebooklm quantum computing --outputs all
/notebooklm climate change solutions 2026 --sources 10 --outputs audio,video,slides
```

### Arguments

| Argument | Default | Description |
|----------|---------|-------------|
| `topic` | required | Research topic |
| `--outputs` | all | Comma-separated: audio, video, slides, infographic, mindmap, report, flashcards, quiz, table, all |
| `--sources` | 8-12 | Number of sources to gather (max 50) |
| `--style` | deep-dive | Audio style: deep-dive, summary, critique, debate |

---

## Workflow

### Phase 1: Web Research

1. **Search for reliable sources** using web search:

```
Use mcp__websearch__web_search_exa or mcp__google_search__search to find:
- Official documentation (prioritize)
- Academic papers
- Reputable news articles
- Technical blogs from recognized sources
- YouTube tutorials/explainers

Target: 8-12 high-quality URLs for comprehensive coverage
```

2. **Source Quality Criteria**:
   - **Prefer**: Official docs, academic papers, reputable news, technical blogs, YouTube tutorials
   - **Avoid**: Opinion blogs, outdated content (>2 years), paywalled sites
   - **Verify**: Use WebFetch to confirm content is accessible

3. **Compile source list**:
```markdown
## Sources for NotebookLM

### Official Documentation
1. [Title] - URL

### Technical Articles
2. [Title] - URL

### YouTube Videos
3. [Title] - URL
```

### Phase 2: NotebookLM Automation

#### Step 1: Navigate to NotebookLM

```javascript
mcp__playwright__browser_navigate({ url: "https://notebooklm.google.com" })
mcp__playwright__browser_wait_for({ time: 3 })
mcp__playwright__browser_snapshot({})
```

#### Step 2: Check Authentication

After snapshot, check if logged in:
- If login page shown: **STOP and inform user**
  ```
  "NotebookLM requires Google authentication.
   Please log in manually in the browser window, then tell me to continue."
  ```
- If logged in (shows notebook list): Continue to Step 3

#### Step 3: Create New Notebook

```javascript
// Look for "새로 만들기" or "새 노트 만들기" button
mcp__playwright__browser_click({
  element: "Create new notebook button",
  ref: "<ref_from_snapshot>"  // e.g., e96 or e105
})

mcp__playwright__browser_wait_for({ time: 2 })
mcp__playwright__browser_snapshot({})
```

#### Step 4: Add Sources (Batch Method - RECOMMENDED)

NotebookLM accepts multiple URLs at once. Use batch input for efficiency:

```javascript
// 1. Click "소스 추가" (Add sources) button
mcp__playwright__browser_click({
  element: "Add sources",
  ref: "<ref>"  // Usually has icon "add" 
})

mcp__playwright__browser_wait_for({ time: 2 })
mcp__playwright__browser_snapshot({})

// 2. Click "웹사이트" (Website) option
mcp__playwright__browser_click({
  element: "Website",
  ref: "<ref>"
})

// 3. Enter ALL URLs at once (newline-separated)
mcp__playwright__browser_type({
  ref: "<url_input_ref>",
  text: "https://url1.com\nhttps://url2.com\nhttps://url3.com\nhttps://url4.com"
})

// 4. Click "삽입" (Insert) button
mcp__playwright__browser_click({
  element: "Insert",
  ref: "<ref>"
})

// 5. Wait for all sources to process
mcp__playwright__browser_wait_for({ time: 10 })
```

**For YouTube videos**: Select "YouTube" option instead of "Website"

#### Step 5: Set Notebook Title

```javascript
// Click on title textbox (usually at top)
mcp__playwright__browser_click({
  element: "Notebook title",
  ref: "<textbox_ref>"
})

// Clear and type new title
mcp__playwright__browser_type({
  ref: "<ref>",
  text: "{topic} {emoji}",  // e.g., "Claude Code Best Practices"
  submit: false
})
```

#### Step 6: Verify Sources Added

```javascript
mcp__playwright__browser_snapshot({})
// Verify: Source panel shows expected count (e.g., "소스 11개")
// All checkboxes should be checked by default
```

### Phase 3: Generate ALL Content Types (PARALLEL)

**CRITICAL**: Click each content type button in quick succession to start parallel generation.

The Studio panel (스튜디오) is on the right side. Click each content button:

```javascript
// 1. AI Audio Overview (takes longest, start first)
mcp__playwright__browser_click({
  element: "AI 오디오 오버뷰",
  ref: "<ref>"  // e.g., e818
})
mcp__playwright__browser_wait_for({ time: 1 })

// 2. Video Overview
mcp__playwright__browser_click({
  element: "동영상 개요",
  ref: "<ref>"  // e.g., e830
})
mcp__playwright__browser_wait_for({ time: 1 })

// 3. Infographic
mcp__playwright__browser_click({
  element: "인포그래픽",
  ref: "<ref>"  // e.g., e880
})
mcp__playwright__browser_wait_for({ time: 1 })

// 4. Slides
mcp__playwright__browser_click({
  element: "슬라이드 자료",
  ref: "<ref>"  // e.g., e892
})
mcp__playwright__browser_wait_for({ time: 1 })

// 5. Mind Map (optional but useful)
mcp__playwright__browser_click({
  element: "마인드맵",
  ref: "<ref>"  // e.g., e842
})
mcp__playwright__browser_wait_for({ time: 1 })

// 6. Report
mcp__playwright__browser_click({
  element: "보고서",
  ref: "<ref>"  // e.g., e849
})
```

**Generation Status Check**:
After clicking, the Studio panel shows generation status:
- "생성 중..." = Generating
- "기반소스 N개" = Based on N sources

All content types generate in parallel. Short-form content (mindmap, report, flashcards) completes in 1-3 minutes. Audio/Video take 5-30 minutes.

### Phase 4: Monitor and Capture Results

#### Wait for Quick Content (2-3 minutes)

```javascript
mcp__playwright__browser_wait_for({ time: 120 })  // 2 minutes
mcp__playwright__browser_snapshot({})

// Check which content is ready:
// - Completed items show preview
// - In-progress shows "생성 중..."
```

#### Capture Intermediate Screenshot

```javascript
mcp__playwright__browser_take_screenshot({
  type: "png",
  filename: "notebooklm-generating.png",
  fullPage: true
})
```

#### Get Notebook URL

```javascript
mcp__playwright__browser_evaluate({
  function: "() => window.location.href"
})
```

### Phase 5: Return Results

**Report to user**:

```markdown
## NotebookLM Research Complete

**Topic**: {topic}
**Notebook URL**: {url}

### Sources Added ({count} total)
| Type | Title | URL |
|------|-------|-----|
| Official Doc | ... | ... |
| Blog | ... | ... |
| YouTube | ... | ... |

### Content Generation Status

| Content Type | Status |
|--------------|--------|
| AI Audio Overview | Generating (5-30 min) |
| Video Overview | Generating (3-15 min) |
| Mind Map | Ready |
| Report | Ready |
| Infographic | Generating |
| Slides | Generating |
| Flashcards | Ready |
| Quiz | Ready |

### Screenshot
See: notebooklm-generating.png

### Next Steps
1. Open notebook URL to view completed content
2. Audio/Video will be ready in 5-30 minutes
3. Use chat interface to ask questions about sources
4. Download content or share notebook with team
```

---

## NotebookLM UI Reference (Korean)

### Main Navigation

| Element | Korean | Typical Ref Pattern |
|---------|--------|---------------------|
| Home | NotebookLM 홈페이지 | link |
| Create Notebook | 새로 만들기, 노트북 만들기 | button with "add" icon |
| Settings | 설정 | button with "settings" icon |

### Source Panel (Left)

| Element | Korean | Description |
|---------|--------|-------------|
| Sources Header | 출처 | Section heading |
| Add Sources | 소스 추가 | Button with "add" icon |
| Select All | 모든 소스 선택 | Checkbox |
| Source Count | 소스 N개 | Status indicator |
| Collapse Panel | 소스 패널 접기 | Button with "dock_to_right" icon |

### Chat Panel (Center)

| Element | Korean | Description |
|---------|--------|-------------|
| Chat Header | 채팅 | Section heading |
| Query Input | 쿼리 상자 | Textbox |
| Submit | 제출 | Button with "arrow_forward" icon |
| Suggested Questions | - | Auto-generated prompts |

### Studio Panel (Right) - CONTENT GENERATION

| Element | Korean | Icon | Ref Example |
|---------|--------|------|-------------|
| Studio Header | 스튜디오 | - | heading |
| AI Audio Overview | AI 오디오 오버뷰 | audio_magic_eraser | e818 |
| Video Overview | 동영상 개요 | subscriptions | e830 |
| Mind Map | 마인드맵 | flowchart | e842 |
| Report | 보고서 | auto_tab_group | e849 |
| Flashcards | 플래시카드 | cards_star | e856 |
| Quiz | 퀴즈 | quiz | e868 |
| Infographic | 인포그래픽 | stacked_bar_chart | e880 |
| Slides | 슬라이드 자료 | tablet | e892 |
| Data Table | 데이터 표 | table_view | e904 |
| Customize | 맞춤설정 | edit | On each content type |
| Collapse Panel | 스튜디오 패널 접기 | dock_to_left | e810 |

### Add Source Dialog

| Element | Korean | Description |
|---------|--------|-------------|
| Website | 웹사이트, 웹 | URL input option |
| YouTube | YouTube | Video URL option |
| Fast Research | Fast Research | Quick source finder |
| Deep Research | Deep Research | Comprehensive research |
| URL Input | - | Textarea for URLs |
| Insert | 삽입 | Submit button |

### Generation Status

| Status | Korean | Meaning |
|--------|--------|---------|
| Generating | 생성 중... | In progress |
| Based on N sources | 기반소스 N개 | Source count |
| Starting | 생성 시작 중입니다 | Just started |
| Ready | (shows preview) | Complete |
| Customize | 맞춤설정 | Edit options |

---

## Error Handling

### Authentication Required
```
User action needed: Please log into Google in the browser window.
After logging in, say "continue" to proceed.
```

### Source Addition Failed
- Retry once after 5 seconds
- If still fails, skip and note in report
- Continue with remaining sources

### Rate Limiting
- Add 2-3 second delays between batch operations
- If rate limited, pause 30 seconds and retry

### Content Generation Issues
- If a content type fails, note in report
- Other content types will continue generating
- User can manually regenerate failed items

### Generation Timeout
- Quick content (mindmap, report, flashcards): 3 minutes max
- Medium content (infographic, slides): 5 minutes max
- Long content (audio, video): 30 minutes max
- After timeout, report status and notebook URL for user to check later

---

## Best Practices

### Source Selection
1. **8-12 sources** is optimal for comprehensive coverage without overwhelming
2. **Mix source types**: official docs + blogs + videos
3. **Verify accessibility** before adding
4. **Include YouTube** when available - adds multimedia variety

### Content Generation
1. **Start audio first** - it takes longest
2. **Click all content types quickly** - they generate in parallel
3. **Don't wait for completion** - return URL and let user check later
4. **Take screenshots** at key stages for documentation

### Performance Tips
1. Use **batch URL input** (newlines) instead of one-by-one
2. **Parallel content generation** - click all types in sequence
3. **Don't poll excessively** - check status every 60 seconds max
4. **Return early** - provide URL, user can check completion

---

## Example Session

```
User: /notebooklm Claude Code best practices

1. Web Search: Find 10-12 sources
   - Claude Code official docs (3-4)
   - Technical blogs (4-5)
   - YouTube tutorials (2-3)

2. Navigate to NotebookLM
   - Verify Google login
   - Create new notebook

3. Add Sources (batch)
   - Paste all URLs at once
   - Wait for processing
   - Verify count

4. Set Title
   - "Claude Code Best Practices"

5. Generate Content (parallel)
   - Click: Audio, Video, Infographic, Slides, MindMap, Report
   - All start generating simultaneously

6. Capture & Report
   - Screenshot current state
   - Return notebook URL
   - List generation status

7. Done
   - User can access notebook
   - Content completes in background
```

---

## Limitations

1. **Authentication**: Requires manual Google login (no API available)
2. **UI Changes**: NotebookLM UI may change; selectors need periodic updates
3. **Processing Time**: Audio/Video can take 5-30 minutes
4. **Source Limits**: Maximum 50 sources per notebook
5. **Rate Limits**: May be throttled with rapid operations
6. **Language**: UI language depends on user's Google account settings
