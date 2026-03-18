---
name: ppt-preview
description: Phase 1.6. draft JSON을 받아 HTML 슬라이드 프리뷰를 생성한다. 브라우저에서 열어 레이아웃·컬러·내용을 확인한 뒤 /ppt-build로 넘어간다. 컬러 팔레트 이미지 첨부 시 해당 색상을 적용한다.
argument-hint: "[draft_json_경로] [--palette 이미지경로(선택)]"
user-invocable: true
allowed-tools: Read, Write, Bash, Agent
model: sonnet
context: fork
---

# Phase 1.6: HTML 슬라이드 프리뷰

**호출 형식**: `/ppt-preview [draft_json_경로]`

예시:
- `/ppt-preview outputs/draft_im_강남오피스.json`
- `/ppt-preview outputs/draft_report_2026오피스시장.json`

컬러 팔레트 이미지를 첨부하면 해당 색상이 프리뷰에 자동 적용됩니다.

## 현재 초안 파일 목록
!`ls outputs/draft_*.json 2>/dev/null || echo "초안 파일이 없습니다. /ppt-plan을 먼저 실행하세요."`

## 실행 흐름

1. `$ARGUMENTS[0]` 경로의 draft JSON 파일을 로드합니다.
2. draft의 `type` 필드로 `types/{type}.json`을 로드합니다.
3. 컬러 팔레트를 결정합니다:
   - 사용자가 이미지를 첨부한 경우 → 이미지에서 5가지 주요 색상 추출
   - 이미지 없음 → `type_config.default_color_palette` (BCG Forest) 사용
4. **html-preview** 에이전트를 호출합니다:
   ```
   입력: { draft_path, style_path (있으면), palette_override (이미지 추출 시) }
   출력: outputs/preview_{type}_{slug}.html
   ```
5. 완료 후 출력합니다:
   ```
   HTML 프리뷰 생성 완료: outputs/preview_{type}_{slug}.html
   적용된 팔레트: {팔레트 이름} — primary:{색상} / accent:{색상} / ...

   브라우저에서 열어 확인 후:
   ① 수정 요청 → 변경 사항을 알려주세요 (draft JSON 수정 후 재실행 가능)
   ② 확정 → /ppt-build outputs/draft_{type}_{slug}.json
   ```

## 용도
- `/ppt-plan` 이후 시각적 레이아웃을 확인할 때
- 컬러 팔레트 이미지를 바꿔가며 색상 테스트를 할 때
- draft JSON을 수정한 후 결과를 빠르게 확인할 때
