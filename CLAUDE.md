# PPT Generator

문서 유형별 설정 파일을 기반으로 PowerPoint를 자동 생성하는 프로젝트.
Phase 1(콘텐츠 생성) → Phase 1.5(구조 확인) → Phase 1.6(HTML 프리뷰) → Phase 2(PPT 빌드) 순서로 동작한다.

## 파이프라인 흐름
```
/ppt-plan  → Phase 1 (초안 생성) → Phase 1.5 (구조 요약 출력 + 사용자 컨펌)
/ppt-preview → Phase 1.6 (HTML 프리뷰 생성 → 브라우저 확인 + 사용자 컨펌)
/ppt-build  → Phase 2 (PPTX 최종 생성)
```
- `/ppt` 는 모든 단계를 순서대로 실행하되, 1.5/1.6 각각 사용자 승인을 기다린다.
- 금융 문서(im, report)는 Phase 1.5에서 반드시 사용자 컨펌 후 진행한다.

## 기술 스택
- Python 3.11+
- python-pptx (슬라이드 마스터 포함 PPT 생성)
- pdfplumber / pymupdf (참고 PDF 스타일 추출)
- Anthropic SDK (콘텐츠 생성)
- FastAPI (추후 원격 API 서버)

## 핵심 디렉토리
- `types/` — 문서 유형별 설정 JSON (fidelity, 소스 경로 등)
- `references/{type}/` — 참고 자료 (PDF 및 PPTX 모두 수용, 유형별·용도별 하위 폴더)
  - `best_practices/` — 양식+내용 모두 훌륭한 자료 → style + logic 양쪽 추출
    - PPTX: 폰트명·EMU 좌표·셀 스타일·슬라이드 마스터까지 정밀 추출 (최고 fidelity)
    - PDF: 색상·레이아웃 패턴 추출
  - `templates/` — 레이아웃·디자인 참고 전용 → style만 추출
    - PPTX 권장: 정확한 도형 좌표·표 구조 확보 가능
    - PDF 가능: 색상·폰트 근사 추출
  - `narratives/` — 내용·논리 참고 전용 (세로형, 해외 양식 등) → logic만 추출
    - PDF 권장 (PPTX도 가능하나 레이아웃 정보는 무시)
  - 파일 수: 유형별 3~5개 권장 (많을수록 공통 스타일 수렴 정확도 향상)
  - style-analyst가 확장자(.pptx/.pdf) 자동 감지 후 적합한 파서 적용
- `src/` — 파이프라인 구현 코드
- `outputs/` — 생성된 .pptx 저장

## 아티팩트 규약
- Phase 1 산출물: `outputs/draft_{type}_{slug}.json`
- Phase 1.6 산출물: `outputs/preview_{type}_{slug}.html`
- Phase 2 산출물: `outputs/{type}_{slug}.pptx`
- 스타일 추출 결과: `outputs/style_{type}.json`
- 슬라이드 배치 설계: `outputs/layout_{type}_{slug}.json`

## 참고 자료 캐시 시스템 (Reference Distillation)
참고 파일이 많아질수록 매번 원본을 파싱하는 것은 비효율적입니다. 두 단계 캐시로 경량화합니다.

### 1단계: 파일별 증류 캐시
```
references/{type}/{subfolder}/.cache/{filename}.json   (~2-5 KB)
```
- `/ppt-add-ref` 실행 시 **ref-distiller** 에이전트가 즉시 생성
- 파일 hash 기반 유효성 검사 (파일 변경 시 자동 갱신)
- `type_usage_hints` 포함: **같은 파일이라도 타입마다 배울 점이 다름**
  - 예) `references/im/templates/` 파일이 `startup` 타입에서도 참조될 때, `startup` 전용 활용 지침 내장
  - style-analyst, logic-analyst, content-planner, content-writer 모두 캐시만 읽음 (원본 로드 없음)

### 2단계: 통합 캐시 (기존)
```
outputs/style_{type}.json    — 스타일 통합 캐시 (style-analyst 생성)
outputs/logic_{type}.json    — 로직 통합 캐시 (logic-analyst 생성, 신규)
```
- 파일별 캐시를 병합한 결과물
- 새 파일 추가 시 무효화 → 다음 파이프라인 실행 시 재병합

### 캐시 활용 에이전트
| 에이전트 | 읽는 캐시 | 원본 로드 |
|---|---|---|
| ref-distiller | 없음 (캐시 생성 주체) | 항상 |
| style-analyst | `outputs/style_{type}.json` → `.cache/*.json` | 캐시 미스 시만 |
| logic-analyst | `outputs/logic_{type}.json` → `.cache/*.json` | 캐시 미스 시만 |
| content-planner | `outputs/logic_{type}.json` → `.cache/*.json` | 캐시 미스 시만 |
| content-writer | `outputs/logic_{type}.json` → `.cache/*.json` | 캐시 미스 시만 |

## 컬러 팔레트
- 기본 팔레트: BCG Forest (`types/` 파일의 `default_color_palette` 참조)
- 사용자가 이미지로 팔레트를 첨부하면 style-analyst가 색상을 추출해 `default_color_palette`를 덮어씀
- `accept_palette_image: true` 인 유형은 이미지 첨부를 허용함

### 색상 해상도 우선순위 (builder & HTML preview 공통)
1. `style_{type}.json` → colors (참고 파일 정밀 추출, 최고 우선)
2. `draft.meta.color_palette` (세션별 — content-planner가 선택)
3. `types/{type}.json` → `default_color_palette` (유형 기본)
4. BCG Forest 하드코딩 (최후 수단)

pptxgen_builder.js는 draft 로드 후 자동으로 이 체인을 적용한다.
`--style` 미지정 시 `outputs/style_{type}.json`을 자동 탐색한다.

### BCG Forest 기본값
| 역할 | 이름 | HEX |
|------|------|-----|
| primary | Deep Forest | #1D3C2F |
| accent | Emerald | #00876A |
| support | Gold | #F2C94C |
| neutral | Warm White | #F5F5F2 |
| text | Near Black | #1A1A1A |

## HTML 프리뷰 렌더링 규칙 (html-preview 에이전트 준수)

### 슬라이드 구조 순서
모든 콘텐츠 슬라이드(`title_slide` · `section_divider` · `toc_slide` 제외)는 아래 순서로 렌더링한다:
1. **slide-header-bar** (primary 배경, 전체폭): `section_number · section_name` (작게) + `slide title` (크게)
2. **head-message box** (인셋): 핵심 인사이트 문장 — 콘텐츠 영역 좌우 여백에 맞춰 들여쓰기, header-bar와 여백을 두고 배치, 좌측 support(gold) 강조 border + 연한 배경
3. **콘텐츠 영역**: 레이아웃별 내용

### 색상 해상도 (pptxgen_builder와 동일 체인 적용)
CSS `:root` 변수를 아래 우선순위로 결정한다. 수치 하드코딩 금지 — 반드시 체인에서 읽어야 한다.
1. `style_{type}.json` → colors (최고 우선)
2. `draft.meta.color_palette`
3. `types/{type}.json` → `default_color_palette`
4. BCG Forest 값 (최후 수단)

### 특별 슬라이드 예외
- `toc_slide` · `title_slide`: 독립 디자인 허용, slide-header-bar 강제 불필요
- `kpi_metrics` 등 section_tag/slide_title 없는 슬라이드: head-message를 제목 영역 아래로 이동 (맨 위 금지)
- `section_divider`: 다크 배경 독립 레이아웃, head-message 없음

## 표준 레이아웃 타입
모든 에이전트(content-planner, logic-analyst, html-preview, ppt-builder)는 아래 이름만 사용한다.
**타입 선택 우선순위**: 수치 데이터가 있으면 표/차트 우선 → 비교 구조면 2열/4분할 → 단순 설명이면 content_text 최후 수단.

### 단일 영역 레이아웃
| 타입 | 용도 |
|------|------|
| `title_slide` | 표지 (제목·부제·날짜·작성기관) |
| `toc_slide` | 목차. report 유형에서 표지 다음 슬라이드로 필수. `sections` 배열 (번호·제목·서브섹션) |
| `section_divider` | 섹션 구분 (대형 섹션 번호 + 제목, 다크 배경) |
| `content_text` | 헤드메시지 + 텍스트 본문. 수치 없는 정성 설명 전용, 남용 금지 |
| `content_chart` | 헤드메시지 + 차트 전체 (`chart_type`: bar / line / waterfall / pie) + key_points |
| `table_slide` | 헤드메시지 + 전체폭 표 (4행×4열 이상). 비교·순위 데이터 |
| `wide_table` | 헤드메시지 + 전체폭 얕은 비교표 (2~5행 × 5~15열). 시장 데이터·시나리오 매트릭스·민감도 분석. SRCIG 빈출 패턴 |
| `kpi_metrics` | 헤드메시지 + KPI 카드 4개 (수치·전년比·단위 필수) |
| `process_flow` | 헤드메시지 + 수평/수직 프로세스 다이어그램 (4~6 스텝). 투자 프로세스·규제 일정·딜 플로우 |
| `roadmap_timeline` | 헤드메시지 + 가로 타임라인 (5노드, 연도·이벤트·설명) |
| `closing_slide` | 마무리 (제언 요약·연락처) |

### 2열 분할 레이아웃 (좌우 50:50)
| 타입 | 용도 |
|------|------|
| `two_col_text_table` | 헤드메시지 + 좌: 분석 텍스트(5~7 bullet) + 우: 데이터 표. **SRCIG 최빈 레이아웃** |
| `two_col_text_chart` | 헤드메시지 + 좌: 분석 텍스트(4~6 bullet) + 우: 차트. 수치 근거 분석 슬라이드 |
| `two_col_chart_text` | 헤드메시지 + 좌: 차트(시각 주역) + 우: 해석 텍스트(3~5 bullet). 차트가 핵심 메시지일 때 |
| `two_column_compare` | 헤드메시지 + 좌우 대비 (장단점·전후·A vs B). 각 열 4~6 bullet |
| `table_chart_combo` | 헤드메시지 + 좌: 표 + 우: 차트 혼합 |

### 복합 분할 레이아웃 (SRCIG 핵심 패턴)
| 타입 | 용도 |
|------|------|
| `composite_split` | 헤드메시지 + 한쪽은 단일 콘텐츠, 반대쪽은 상하 2분할. SRCIG asymmetric/single_left/single_right 패턴 구현. `main_zone` + `sub_zone_top` + `sub_zone_bottom` 필드 사용 |
| `four_quadrant` | 헤드메시지 + 2×2 4분할. 각 셀 독립 콘텐츠 (텍스트/미니표/수치/다이어그램). 물류·호텔 PPTX 핵심 레이아웃. `cells[4]` 배열 사용 |

### 다중 카드/요약 레이아웃
| 타입 | 용도 |
|------|------|
| `three_column_summary` | 헤드메시지 + 3열 요약 카드 (각 카드: 제목·bullet 3~4개·수치) |
| `image_gallery` | 헤드메시지 + 이미지 그리드 2×3 (6칸, 캡션 포함) |

### 복합 레이아웃 JSON 스키마

#### composite_split
```json
{
  "layout": "composite_split",
  "head_message": "...",
  "main_zone": {
    "position": "left",
    "content_type": "chart",
    "chart": {"chart_type": "bar", "data": [...]}
  },
  "sub_zone_top": {
    "position": "right_top",
    "content_type": "bullets",
    "bullets": ["..."]
  },
  "sub_zone_bottom": {
    "position": "right_bottom",
    "content_type": "table",
    "table": {"headers": [...], "rows": [...]}
  }
}
```
`content_type`: `"bullets"` | `"chart"` | `"table"` | `"callout"` | `"process"`

#### four_quadrant
```json
{
  "layout": "four_quadrant",
  "head_message": "...",
  "cells": [
    {"position": "top_left",     "label": "...", "content_type": "bullets", "bullets": [...]},
    {"position": "top_right",    "label": "...", "content_type": "table",   "table": {...}},
    {"position": "bottom_left",  "label": "...", "content_type": "chart",   "chart": {...}},
    {"position": "bottom_right", "label": "...", "content_type": "callout", "value": "...", "description": "..."}
  ]
}
```

#### toc_slide
```json
{
  "layout": "toc_slide",
  "sections": [
    {"number": "01", "title": "섹션명", "subsections": ["하위항목1", "하위항목2"]}
  ]
}
```

#### process_flow
```json
{
  "layout": "process_flow",
  "head_message": "...",
  "flow_type": "horizontal",
  "steps": [
    {"number": 1, "title": "...", "items": ["...", "..."], "highlight": false}
  ]
}
```

## 콘텐츠 밀도 기준 (참고 PDF 수준 준수)
content-planner·content-writer는 아래 기준을 **반드시** 충족해야 한다.

### 슬라이드별 최소 데이터 요건
| 레이아웃 | bullet 수 | 수치 포함 | 표 행×열 | 비고 |
|---|---|---|---|---|
| `content_text` | 5~7개 | bullet당 1개 이상 | - | 수치 없으면 해당 레이아웃 사용 금지 |
| `two_col_text_table` | 좌 5~7개 | 좌 bullet당 1개 | 우 4행×3열 이상 | SRCIG 최빈 레이아웃 |
| `two_col_text_chart` | 좌 4~6개 | 좌 bullet당 1개 | - | 차트 데이터 최소 5포인트 |
| `two_col_chart_text` | 우 3~5개 | 우 bullet당 1개 | - | 차트 데이터 최소 5포인트 |
| `two_column_compare` | 각 열 4~6개 | 각 열 3개 이상 | - | 좌우 항목 1:1 대응 권장 |
| `table_slide` | key_points 3~4개 | 표 셀 70% 이상 수치 | 5행×4열 이상 | |
| `wide_table` | key_points 2~3개 | 표 셀 80% 이상 수치 | 2~5행×5열 이상 | 열 헤더 명확히 |
| `content_chart` | key_points 3~4개 | 차트 데이터 6개 이상 | - | |
| `three_column_summary` | 각 카드 3~4개 | 카드당 1개 이상 | - | |
| `kpi_metrics` | - | KPI 4개 모두 수치 | - | 전년比·단위 필수 |
| `composite_split` | main+sub 합계 6~10개 | main에 수치 3개 이상 | sub 표 있으면 3행×3열 이상 | main_zone 콘텐츠가 핵심 |
| `four_quadrant` | 각 셀 2~4개 | 셀당 1개 이상 | 셀 내 표 3행×3열 이상 | 4셀 모두 채울 것 |
| `process_flow` | 각 스텝 2~3 항목 | 스텝당 1개 이상 | - | 4~6 스텝, 스텝간 인과관계 명시 |

### Note 박스 (note_box 필드)
참고 PDF의 핵심 패턴: 매 슬라이드 하단 또는 우측에 보충 설명/출처/주의사항 박스.
슬라이드 JSON에 선택적으로 추가:
```json
{
  "note_box": {
    "type": "source",
    "content": "출처: GWEC 2025, IEA Renewables 2025"
  }
}
```
`type`: `"source"` | `"definition"` | `"caution"` | `"additional"`

### 데이터 포인트 (data_points 배열)
슬라이드에서 강조할 핵심 수치를 별도 배열로 명시. ppt-builder가 callout 박스로 처리:
```json
{
  "data_points": [
    {"label": "2024 누적 설치", "value": "83GW", "delta": "+26% YoY"},
    {"label": "2034 전망(GWEC)", "value": "441GW"}
  ]
}
```

### content-writer 리서치 기준
- 슬라이드 1장 작성 전 반드시 해당 주제에 대해 **최소 2회 이상** 웹 서치
- 수치는 반드시 출처(기관명+연도) 명기
- 시사점·제언 섹션은 **문제 → 원인(수치 근거) → 해결책 → 기대 효과(수치 목표)** 4단 구조 준수
- 글로벌 vs 국내 비교가 가능한 경우 반드시 병기

## 코드 규약
- 모든 수치(좌표, 폰트 크기)는 PowerPoint EMU 단위 사용 (pt → EMU: ×12700)
- 슬라이드 크기 기본값: 12192000 × 6858000 EMU (와이드 16:9, 960×540pt, HTML 프리뷰 기준)
- 규격 변경 시 `src/build/pptxgen_builder.js` 상단 `CFG.emuW / CFG.emuH` 만 수정 (폰트 pt값은 고정)
- JSON 출력 시 한국어 키 사용 금지, 영문 snake_case 사용
- 타입 설정 파일은 `types/` 에만 수정, 파이프라인 코드는 건드리지 말 것

## MCP 서버 (.mcp.json)
- `fetch` — 웹 페이지 콘텐츠 가져오기 (content-writer 리서치용)
- `memory` — 세션 간 지식 그래프 유지 (시장 데이터, 기업 정보 저장)
- `sequential-thinking` — 복잡한 다단계 추론 (IRR·SWOT·투자 논리)

## 분석 스킬
- `/swot [주제]` — SWOT 분석 슬라이드 생성 (swot-analyst 에이전트)
  - `--draft [경로]` : 기존 draft JSON에 삽입
  - `--strategy` : SO/WO/ST/WT 전략 매트릭스 슬라이드 추가
- `/irr` — IRR·NPV·재무 분석 슬라이드 생성 (financial-analyst 에이전트)
  - `--draft [경로]` : 기존 draft JSON에 삽입
  - `--sensitivity` : 민감도 분석 매트릭스 추가
- `/layout-review [draft_경로]` — 슬라이드 레이아웃 검수 (layout-reviewer 에이전트)
  - `--fix` : 자동 수정 모드

## 명령어
- `python src/core/run.py --type report --topic "주제"` — 전체 실행
- `python src/parsers/extract_style.py references/report/` — 스타일 추출
- `python src/core/validate.py outputs/draft_*.json` — 초안 검증
