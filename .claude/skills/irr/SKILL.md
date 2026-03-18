# /irr — IRR·NPV·재무 분석 슬라이드 생성

## 트리거
사용자가 `/irr` 또는 `/irr --draft [경로]`를 입력할 때 실행.

## 목적
투자 재무 지표(IRR, NPV, Cap Rate, DSCR 등)를 계산하고, 표·차트 슬라이드로 변환한다.
부동산·대체투자 IM 및 리포트에 특화.

## 사용법
```
/irr                                    # 대화형: 수치 입력 안내
/irr --draft outputs/draft_im_xxx.json  # 기존 draft에 재무 슬라이드 추가
/irr --type waterfall                   # 현금흐름 워터폴 차트 생성
/irr --sensitivity                      # IRR 민감도 분석 매트릭스 추가
```

## 실행 절차

1. **입력 수집** (대화형):
   ```
   Claude: 재무 분석에 필요한 수치를 입력해주세요.

   [필수]
   - 초기 투자금 (억원):
   - 연도별 현금흐름 (쉼표 구분, 억원):
   - 예상 보유 기간 (년):
   - 출구 Cap Rate 또는 매도가 (억원):

   [선택]
   - 목표 수익률/할인율 (%):
   - 레버리지 비율 (%):
   ```

2. **에이전트 호출**: `financial-analyst` 에이전트 실행
   - Python(numpy_financial)으로 IRR/NPV 정확 계산
   - 보수적/기본/낙관 3개 시나리오 생성

3. **슬라이드 생성**:
   - IRR·NPV 요약 표 슬라이드 (table_slide)
   - 현금흐름 워터폴 차트 (content_chart)
   - 민감도 분석 매트릭스 (table_slide) — `--sensitivity` 시

4. **결과 출력 및 draft 삽입 확인**

## 출력 예시
```
## 재무 분석 완료

기본 시나리오: IRR 15.2% | NPV 38억원 | MOIC 1.9x

생성된 슬라이드 3장:
- [표] IRR·NPV 3개 시나리오 비교
- [차트] 연도별 현금흐름 워터폴
- [표] Cap Rate × 임차율 민감도 매트릭스

draft에 삽입하시겠습니까? (Y/N)
```
