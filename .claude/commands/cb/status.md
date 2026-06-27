---
description: 프로젝트 현황 보고 — pm-orchestration §6 양식으로 보고 후 트래킹 이슈에 박제
---

인자: $ARGUMENTS (선택 — 마일스톤 이름)

## 절차

1. **보드 현황 읽기**
   ```bash
   gh project item-list 3 --owner ms9648 --format json \
     | jq '[.items[] | {title, status: .status, number: .content.number}]'
   ```

2. **아래 양식으로 현황 보고**

   ## 프로젝트 현황 — <날짜>
   - 마일스톤: <명> (<완료>/<전체> 이슈)
   - 이번 보고 기간 완료: #n, #n …
   - 진행 중: #n(담당 에이전트, 단계), …
   - 차단/대기: #n — 사유, 필요한 결정
   - 리스크: …
   - 다음 액션: …

3. **트래킹 이슈에 박제**
   ```bash
   gh issue comment <트래킹이슈번호> --body-file /tmp/status-report.md
   ```
