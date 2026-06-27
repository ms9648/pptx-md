---
description: 스프린트 계획 — 마일스톤 잔여 이슈를 우선순위/의존성으로 정렬해 Todo 승격
---

인자: $ARGUMENTS (마일스톤 이름 또는 번호)

platform-ops 스킬(§7)의 gh CLI 표준 조작을 사용하여 보드를 반영한다.

## 절차

1. **보드 현황 읽기**
   ```bash
   gh project item-list 3 --owner ms9648 --format json \
     | jq '[.items[] | select(.status == "Backlog") | {title, number: .content.number, priority: .priority}]'
   ```

2. **마일스톤 필터링**
   지정된 마일스톤($ARGUMENTS)에 속하는 Backlog 이슈만 추린다.

3. **우선순위·의존성 정렬**
   - Priority P0 → P1 → P2 순
   - 의존관계 있는 이슈는 선행 이슈 완료 후 Todo 승격

4. **이번 스프린트 Todo 승격안 제시**
   사람에게 승격할 이슈 목록을 보여주고 확인 받는다.

5. **[사람 확인] 보드 반영**
   승인된 이슈들을 Backlog → Todo 로 상태 전이한다:
   ```bash
   gh project item-edit --id <ITEM_ID> --project-id PVT_kwHOApzjFM4BbyNH \
     --field-id PVTSSF_lAHOApzjFM4BbyNHzhWgMgU --single-select-option-id e6d4212e
   ```

6. **결과 보고**
   Todo 로 승격된 이슈 목록과 이번 스프린트 범위를 요약 보고한다.
