#!/usr/bin/env bash
# 서브에이전트 컨텍스트에서 보드 상태 변경/이슈 close 시도를 차단한다.
# (PreToolUse 훅: stdin 으로 도구 호출 JSON 이 들어온다)
input=$(cat)
cmd=$(echo "$input" | jq -r '.tool_input.command // empty')

# Done 전이·이슈 close 시에만 확인 — 나머지 상태 전이(Backlog/Todo 등)는 allow.
# 훅의 목적: 검증 없는 Done 처리 방지(G3). 계획 작업에 불필요한 마찰 없음.
# Done option-id 는 부트스트랩 후 platform-ops SKILL.md "Status 옵션 ID" 표에서 확인.
if echo "$cmd" | grep -q 'gh issue close'; then
  echo '{"hookSpecificOutput":{"hookEventName":"PreToolUse","permissionDecision":"ask","permissionDecisionReason":"이슈 close는 Done 전이와 동시에 수행됩니다. AC 최종 대조(pm-orchestration §4)를 완료했습니까?"}}'
  exit 0
fi

if echo "$cmd" | grep -q 'gh project item-edit' && echo "$cmd" | grep -q 'f14a6349'; then
  echo '{"hookSpecificOutput":{"hookEventName":"PreToolUse","permissionDecision":"ask","permissionDecisionReason":"Done 전이를 시도합니다. AC 최종 대조(pm-orchestration §4)를 완료했습니까?"}}'
  exit 0
fi

exit 0
