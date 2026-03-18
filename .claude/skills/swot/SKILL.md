# /swot — SWOT 분석 슬라이드 생성

## 트리거
사용자가 `/swot [주제]` 또는 `/swot [주제] --draft [경로]`를 입력할 때 실행.

## 목적
특정 대상에 대한 SWOT 분석을 수행하고, content_draft.json에 삽입 가능한 슬라이드 블록을 생성한다.

## 사용법
```
/swot [주제]
/swot [주제] --draft outputs/draft_{type}_{topic}.json
/swot [주제] --strategy   # SO/WO/ST/WT 전략 매트릭스 슬라이드 추가
```

## 실행 절차

1. **주제 파악**: 분석 대상(기업, 사업, 부동산, 시장)과 맥락 확인
2. **에이전트 호출**: `swot-analyst` 에이전트 실행
   - 인자: subject, context (있으면), draft_path (--draft 제공 시), include_strategy (--strategy 제공 시)
3. **결과 출력**:
   - 슬라이드 블록 (JSON) 출력
   - `--draft` 제공 시: 해당 draft JSON에 SWOT 슬라이드 삽입 후 저장 확인
4. **사용자 확인**: "이 SWOT 분석을 draft에 삽입할까요?" 질문

## 출력 예시
```
## SWOT 분석 완료: [주제]

### 생성된 슬라이드
- 슬라이드 1: [내부 역량] Strengths vs Weaknesses (two_column_compare)
- 슬라이드 2: [외부 환경] Opportunities vs Threats (two_column_compare)

draft 파일에 삽입하려면 경로를 알려주세요.
또는 /ppt-build 로 바로 PPT를 생성할 수 있습니다.
```
