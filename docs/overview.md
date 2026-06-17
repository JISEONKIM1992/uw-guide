# 신계약지원 Agent 개요

## 현재 목표
가이드라인과 RD 원본을 업로드해 `시스템반영_RD.xlsx`를 생성하고, 저장된 상품을 기준으로 비교 질의를 수행하는 지원 화면을 운영한다.

## 현재 구현 범위
- 파일 업로드
- `xlsx/csv` 요약 생성
- 변환 결과 다운로드
- 상품 목록 누적
- OpenAI 질의/응답
- payload 미리보기

## 핵심 원칙
- 질문 답변은 일반/간편 단일건을 우선 기준으로 본다.
- 인별합산은 참고값으로만 사용한다.
- 질문 키워드에 가장 잘 맞는 특약 item을 우선 찾는다.
- API 키는 프런트에 두지 않고 로컬 프록시에서만 사용한다.

## 주요 파일
- `support-agent-app.html`
- `support-agent-openai-proxy.mjs`
- `index.html`
- `docs/screen_spec.md`

## 문서 역할
- `docs/screen_spec.md`: 화면 기획
- `docs/overview.md`: 전체 맥락 요약
- `data/catalog.md`: 파일/산출물 목록
- `tests/test_plan.md`: 검증 시나리오
- `skills/agent_rules.md`: 운영 규칙
- `ml/model_context.md`: 질문/응답 기준

