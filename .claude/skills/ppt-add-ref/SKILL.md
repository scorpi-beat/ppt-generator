---
name: ppt-add-ref
description: 특정 유형에 참고 자료(PDF)를 추가하고 스타일/논리 소스 목록을 업데이트한다.
argument-hint: "[유형] [파일경로] [용도: style|logic|both]"
user-invocable: true
allowed-tools: Read, Write, Glob, Bash
model: sonnet
---

# 참고 자료 추가

**호출 형식**: `/ppt-add-ref [유형] [파일경로] [용도]`

예시:
- `/ppt-add-ref im references/im/best_practices/gs_im_2024.pdf both`
- `/ppt-add-ref report references/report/templates/annual_report.pdf style`
- `/ppt-add-ref im references/im/narratives/overseas_im_vertical.pdf logic`

## 참고 폴더 분류 기준

파일을 추가하기 전에 아래 기준에 따라 **물리적 폴더**를 먼저 결정하세요:

| 폴더 | 어떤 자료? | 용도 파라미터 |
|------|-----------|------------|
| `best_practices/` | 양식 + 내용 모두 훌륭한 자료 | `both` |
| `templates/` | 레이아웃·디자인만 참고할 자료 | `style` |
| `narratives/` | 내용·논리는 좋지만 레이아웃 비율이 다른 자료 (세로형, 해외 양식 등) | `logic` |

> **주의**: `narratives/`에 있는 파일은 style-analyst가 읽지 않습니다. 레이아웃 좌표 추출이 필요한 자료는 `templates/` 또는 `best_practices/`에 넣으세요.

## 현재 참고 자료 현황
!`find references/ -name "*.pdf" 2>/dev/null | sort || echo "참고 자료 없음"`

## 실행 내용
1. 파일 존재 여부 확인
2. `types/{유형}.json` 업데이트:
   - `style`: `style.sources` 배열에 경로 추가
   - `logic`: `logic.sources` 배열에 경로 추가
   - `both`: `style.sources`와 `logic.sources` **두 배열 모두**에 경로 추가
3. **증류 캐시 즉시 생성** (ref-distiller 에이전트 호출):
   - 파일별 `.cache/{filename}.json` 생성
   - 이 시점에 한 번만 파싱 → 이후 모든 파이프라인은 캐시만 읽음
   - `type_usage_hints` 생성: 다른 타입이 이 파일 참조 시 활용 지침 포함
4. **통합 캐시 무효화**: `outputs/style_{유형}.json` 및 `outputs/logic_{유형}.json` 삭제
   - 다음 파이프라인 실행 시 새 파일이 반영된 상태로 재병합
5. 업데이트 결과 출력:
   - 추가된 파일명, 캐시 저장 경로
   - 추출된 색상·폰트 요약 (style 용도인 경우)
   - 추출된 논리 패턴 요약 (logic 용도인 경우)
   - cross-type hints 생성 여부

## 증분 업데이트 원칙
파일을 추가할 때마다 **전체 재파싱 없이** 해당 파일만 처리합니다:
- 기존 `.cache/` 파일은 그대로 유지
- 새 파일의 캐시만 생성
- 통합 캐시(`style_{type}.json`)는 다음 PPT 생성 시 자동 재병합
