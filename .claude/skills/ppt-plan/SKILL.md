---
name: ppt-plan
description: Phase 1만 실행. 주제를 받아 콘텐츠 초안(content_draft.json)을 생성하고, Phase 1.5 구조 요약을 출력한 뒤 사용자 검토를 기다린다. PPT는 생성하지 않는다.
argument-hint: "[유형] [주제]"
user-invocable: true
allowed-tools: Read, Write, WebSearch, WebFetch, Agent
model: opus
context: fork
---

# Phase 1: 콘텐츠 초안 생성

**호출 형식**: `/ppt-plan [유형] [주제]`

예시:
- `/ppt-plan im 강남 오피스 빌딩 매입 투자제안`
- `/ppt-plan report 2026 대체투자 시장 전망`

## 실행 흐름

1. **orchestrator**를 `phase: 1_only`로 호출합니다.
2. orchestrator → content-planner → content-writer 순서로 실행됩니다.
3. **Phase 1.5**: 완료 후 orchestrator가 슬라이드 구조 요약 테이블을 직접 출력합니다:
   - 슬라이드 번호 / 제목 / 핵심 메시지 / 레이아웃 유형 / 데이터 신뢰도
   - [추정] 태그 항목 목록 (수치 확인 필요 항목)
4. 다음 단계 안내를 출력합니다:
   ```
   초안 파일: outputs/draft_{type}_{slug}.json

   다음 단계를 선택하세요:
   ① 수정 요청 → 변경 사항을 알려주세요
   ② HTML 프리뷰 → /ppt-preview outputs/draft_{type}_{slug}.json
   ③ 바로 PPT 생성 → /ppt-build outputs/draft_{type}_{slug}.json
   ```

## 용도
- 내용·논리 흐름을 먼저 확인하고 싶을 때
- 초안을 수동으로 편집하고 PPT를 나중에 만들 때
- 여러 초안을 비교하고 하나만 PPT로 만들 때
