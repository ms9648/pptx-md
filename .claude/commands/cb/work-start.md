---
description: 보드의 Todo 이슈를 표준 사이클로 자율 실행
---
pm-orchestration 스킬의 "이슈 단위 실행 루프"를 따른다.
1. 보드에서 Status=Todo 이슈를 Priority 순으로 읽는다. 인자: $ARGUMENTS (개수, 기본 1)
2. 각 이슈에 대해 디스패치 표준에 따라 담당 에이전트를 호출한다.
3. 보고 수신 → PM 검증 절차(4단계) → 상태 전이 → 다음 단계 에이전트 호출.
4. NEEDS_DECISION / ESCALATE / 사람 게이트 도달 시 즉시 멈추고 사람에게 보고한다.
5. 종료 시 /cb:status 양식으로 결과를 요약한다.
