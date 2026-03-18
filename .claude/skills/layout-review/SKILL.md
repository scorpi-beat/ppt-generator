# /layout-review — 슬라이드 레이아웃 검수

## 트리거
사용자가 `/layout-review [draft_경로]`를 입력할 때 실행.
또는 ppt-plan 완료 후 사용자가 검수를 요청할 때.

## 목적
content_draft.json을 검수하여 슬라이드 흐름, 헤드메시지 품질, 레이아웃 다양성, 데이터 충분성을 점검하고 수정 권고안을 제시한다.

## 사용법
```
/layout-review outputs/draft_report_xxx.json
/layout-review outputs/draft_im_xxx.json --fix   # 자동 수정 모드
```

## 실행 절차

1. **파일 읽기**: draft JSON 및 타입 설정(types/{type}.json) 로드
2. **에이전트 호출**: `layout-reviewer` 에이전트 실행
   - 검수 항목: 내러티브 흐름, 헤드메시지 품질, 레이아웃 다양성, 데이터 충분성, 슬라이드 수
3. **리뷰 리포트 출력**:
   - 점수 (/100)
   - 통과/경고/오류 항목 목록
   - 슬라이드별 수정 권고 표
4. **`--fix` 모드**: 자동 수정 후 `draft_*_reviewed.json` 저장

## 검수 기준 요약
| 항목 | 기준 |
|------|------|
| 헤드메시지 | 40자 이내 인사이트 문장 |
| 레이아웃 | 동일 타입 3연속 금지 |
| 데이터 | 주장에 수치/근거 포함 |
| 추정치 | 전체 슬라이드의 50% 이하 |
| 슬라이드 수 | 타입 설정의 min~max 범위 |
| 구성 | title_slide + closing_slide 필수 |

## 파이프라인 연계
검수 통과 후: `/ppt-preview [draft_경로]` 또는 `/ppt-build [draft_경로]` 실행 권장
