---
name: ppt-build
description: Phase 2만 실행. 검토 완료된 content_draft.json 파일을 받아 PPTX를 생성한다. /ppt-build [draft_파일경로] 형식으로 호출.
argument-hint: "[draft_json_경로]"
user-invocable: true
allowed-tools: Read, Write, Bash, Agent
model: sonnet
context: fork
---

# Phase 2: PPT 생성

**호출 형식**: `/ppt-build [draft_json_경로]`

예시:
- `/ppt-build outputs/draft_im_부산냉동물류센터.json`
- `/ppt-build outputs/draft_report_2026오피스시장.json`

## 현재 초안 파일 목록
!`ls outputs/draft_*.json 2>/dev/null || echo "초안 파일이 없습니다. /ppt-plan을 먼저 실행하세요."`

## 권장 흐름
HTML 프리뷰를 먼저 확인한 뒤 이 명령을 실행하는 것을 권장합니다:
```
/ppt-preview outputs/draft_*.json   ← 시각적 확인
/ppt-build outputs/draft_*.json     ← 최종 PPTX 생성
```
단, 구성이 이미 확정된 경우 `/ppt-preview` 없이 바로 실행해도 됩니다.

## 실행 흐름

1. `$ARGUMENTS[0]` 경로의 draft JSON 파일을 로드합니다.
2. draft의 `type` 필드로 `types/{type}.json`을 로드합니다.
3. **orchestrator**를 `phase: 2_only`로 호출합니다.
4. orchestrator → (style-analyst || logic-analyst 병렬) → ppt-builder 순서로 실행됩니다.
   - 컬러 팔레트: 사용자가 이미지 첨부 시 추출, 없으면 `type_config.default_color_palette` 사용
5. 완료 후 출력합니다:
   - 생성된 .pptx 파일 경로
   - 총 슬라이드 수
   - 적용된 컬러 팔레트 이름과 색상 목록
   - 적용된 스타일 소스 (어떤 참고 PDF를 사용했는지)

## 용도
- `/ppt-plan` + `/ppt-preview`로 확인한 초안을 최종 PPT로 변환
- 초안 JSON을 직접 수정한 후 재생성
- 동일 초안으로 다른 유형의 스타일 적용 테스트
