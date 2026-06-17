# OpenAI 질의 컨텍스트 기준

## 목적
지원 Agent의 OpenAI 답변 품질을 높이기 위한 질문 컨텍스트 구조를 정의한다.

## 현재 입력 구조
- `question`
- `answerFocus`
- `ignoreFields`
- `product`
- `comparisonSummary`
- `reference`
- `fileInsights`

## 우선순위
1. `product.comparisonSummary`
2. `product.limits.generalAsIs/generalToBe`
3. `product.limits.simpleAsIs/simpleToBe`
4. `product.reference.individualSum`
5. `fileInsights`

## 답변 규칙
- 일반/간편 단일건만 최종 한도로 본다.
- 인별합산은 별도 참고값이다.
- 질문이 특약명 중심이면 해당 특약 item만 답한다.
- 질문이 변경안 정리면 변경 전/후를 먼저 말한다.

## 프롬프트 의도
- 모델이 5000 같은 참고값에 끌리지 않도록 한다.
- 모델이 상품 그룹 전체가 아니라 특약 item을 읽도록 한다.
- 실무자가 바로 검토할 수 있게 판단 근거를 짧게 만든다.

