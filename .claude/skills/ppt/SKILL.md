---
name: ppt
description: PPT 전체 파이프라인 실행. Phase 1(초안) → Phase 1.5(구조 확인) → Phase 1.6(HTML 프리뷰) → Phase 2(PPTX)를 순서대로 실행한다. 각 단계에서 사용자 승인을 기다린다. /ppt [유형] [주제] 형식으로 호출.
argument-hint: "[유형: report|im|startup] [주제]"
user-invocable: true
allowed-tools: Read, Write, Bash, Agent
model: opus
context: fork
---

# PPT 생성 (전체 파이프라인)

**호출 형식**: `/ppt [유형] [주제]`

예시:
- `/ppt im 부산 냉동 물류센터 투자제안`
- `/ppt report 2026년 서울 오피스 시장 동향`
- `/ppt startup AI 기반 법률 문서 검토 서비스`

컬러 팔레트 이미지를 함께 첨부하면 해당 색상이 전체 파이프라인에 적용됩니다.

## 현재 등록된 유형
!`ls types/ 2>/dev/null || echo "types/ 폴더가 없습니다. /ppt-config로 유형을 먼저 등록하세요."`

## 파이프라인 흐름

```
Phase 1   → 콘텐츠 초안 생성 (content-planner → content-writer)
Phase 1.5 → 구조 요약 출력 + 사용자 컨펌 ⏸
Phase 1.6 → HTML 프리뷰 생성 + 사용자 컨펌 ⏸
Phase 2   → PPTX 최종 생성 (style-analyst + logic-analyst → ppt-builder)
```

각 ⏸ 지점에서 사용자 승인을 받아야 다음 단계로 진행합니다.
금융 문서(im, report)는 Phase 1.5에서 수치 항목을 반드시 검토합니다.

## 실행

인수를 파싱합니다:
- `$ARGUMENTS[0]` → 유형 (report / im / startup / 기타 등록된 유형)
- `$ARGUMENTS[1..]` → 주제 (나머지 전체)

유형이 없거나 `types/{유형}.json`이 존재하지 않으면 사용자에게 등록된 유형 목록을 보여주고 선택을 요청합니다.

**orchestrator** 에이전트를 다음 컨텍스트로 호출합니다:

```
type: $ARGUMENTS[0]
topic: $ARGUMENTS[1...]
phase: both
input_completeness: auto-detect
palette_image: (첨부 이미지가 있으면 전달)
```

orchestrator는 입력 완성도를 자동 감지하여 필요한 단계만 실행합니다.
Phase 1.5 확인 후 html-preview 에이전트를 호출하여 HTML 프리뷰를 생성합니다.
사용자 최종 승인 후 Phase 2를 실행합니다.
