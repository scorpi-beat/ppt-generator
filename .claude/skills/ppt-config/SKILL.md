---
name: ppt-config
description: 새로운 PPT 유형을 등록하거나 기존 유형의 fidelity와 참고 소스를 수정한다.
argument-hint: "[유형명] (없으면 인터랙티브 모드)"
user-invocable: true
allowed-tools: Read, Write, Glob
model: sonnet
---

# PPT 유형 설정

**호출 형식**: `/ppt-config [유형명]`

예시:
- `/ppt-config` — 모든 유형 목록 보기
- `/ppt-config im` — IM 설정 보기/수정
- `/ppt-config esg` — 새 유형 'esg' 등록

## 현재 등록된 유형
!`ls types/*.json 2>/dev/null | xargs -I{} basename {} .json || echo "등록된 유형 없음"`

## 유형 설정 파일 구조 (types/{type}.json)
```json
{
  "name": "유형 표시명",
  "description": "이 유형에 대한 설명",
  "style": {
    "sources": ["references/{type}/templates/", "references/{type}/best_practices/"],
    "fidelity": 0.8,
    "fallback": "blend",
    "note": "스타일 소스 관련 메모"
  },
  "logic": {
    "sources": ["references/{type}/narratives/", "references/{type}/best_practices/"],
    "fidelity": 0.5,
    "fallback": "ai",
    "genre": "report",
    "narrative_arc": "evidence_first",
    "note": "논리 소스 관련 메모"
  },
  "content": {
    "mode": "ai_with_data",
    "web_search": false,
    "tone": "objective_analytical",
    "language": "ko",
    "require_review": false
  },
  "default_color_palette": {
    "name": "BCG Forest",
    "primary": "#1D3C2F",
    "accent": "#00876A",
    "support": "#F2C94C",
    "neutral": "#F5F5F2",
    "text": "#1A1A1A"
  },
  "accept_palette_image": true,
  "slide_defaults": {
    "min_slides": 10,
    "max_slides": 25,
    "default_slides": 15
  },
  "slide_master": "src/templates/{type}_master.pptx"
}
```

> `accept_palette_image: false`로 설정하면 사용자가 팔레트 이미지를 첨부해도 무시하고 회사 공식 색상을 유지합니다. 브랜드 컬러가 엄격히 지정된 유형에 사용합니다.

## 참고 폴더 구조
새 유형 등록 시 아래 폴더를 생성합니다:
```
references/{type}/
├── best_practices/   ← 양식 + 내용 모두 훌륭한 자료 (style + logic 양쪽 활용)
├── templates/        ← 레이아웃·디자인 참고 전용 (style만)
└── narratives/       ← 내용·논리 참고 전용, 세로형/해외 양식 등 (logic만)
```

## fidelity 가이드
| 값 | 의미 |
|---|---|
| 0.9~1.0 | 참고 자료를 거의 그대로 재현 |
| 0.6~0.8 | 참고 자료 기반, 창조적 조합 허용 |
| 0.3~0.5 | 참고 자료 참고만, 자유롭게 재구성 |
| 0.0~0.2 | 참고 자료 무시, AI 또는 웹 검색 전적 활용 |

## fallback 옵션
- `blend`: 여러 소스를 혼합
- `ai`: AI 기본값 사용
- `web_search`: 웹 검색으로 보완
- `first`: 첫 번째 소스만 사용

인수 없이 호출하면 등록된 모든 유형의 현재 설정을 표 형식으로 보여줍니다.
특정 유형명을 인수로 주면 해당 설정을 보여주고 수정 여부를 묻습니다.
존재하지 않는 유형명이면 새 설정 파일과 참고 폴더를 대화형으로 생성합니다.
