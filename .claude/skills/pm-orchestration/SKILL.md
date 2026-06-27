---
name: pm-orchestration
description: >
  PM 의 프로젝트 운영 절차. 단계 게이트, 디스패치 규약, 검증 절차,
  재작업 루프, 사람 에스컬레이션 기준, 현황 보고 양식을 정의한다.
  보드 운영, 이슈 위임, 검증, 보고 시 항상 참조.
---

# PM 운영 매뉴얼

## 1. 프로젝트 단계와 게이트

  Phase 0  착수      → [사람 게이트] 헌장 승인
  Phase 1  기획      → planner → [사람 게이트] 요구사항 베이스라인 확정
  Phase 2  설계      → architect (+ 프론트 시 designer 병행) → reviewer 설계 리뷰 → [사람 게이트] 아키텍처/UX 승인
  Phase 3  구현      → developer ↔ reviewer (이슈 단위 루프)
  Phase 4  테스트    → tester → 결함 루프 (developer 재작업)
  Phase 5  마감      → doc-writer → [사람 게이트] 릴리스 승인 → /cb:retro
                       → ⟳ (다음 마일스톤 있으면 Phase 2 로, 회고 게이트 통과 후)

  ※ Phase 3~4 는 이슈 단위로 파이프라인 병행 가능. 단 Phase 2 게이트 통과 전
    구현 이슈를 In Progress 로 올리지 않는다.
  ※ 마일스톤 경계 = 회고 소프트 게이트 — 다음 마일스톤 Phase 2 진입 전
    직전 마일스톤 /cb:retro 완료(또는 명시적 생략 승인) 필수 (§Phase 2 step 0).

## 2. 이슈 단위 실행 루프 (Phase 3~4 의 기본 사이클)

  Todo → [PM: 디스패치] → In Progress
       → developer: 구현 + 단위테스트 + AC자가검증
                  → git push origin <브랜치>
                  → gh pr create --draft (§7.5.2 템플릿, Closes #이슈)
                  → REPORT (PR 링크 포함)
       → [PM: 보고 형식/산출물 존재/PR 링크 확인] → In Progress → Review 전이
       → reviewer 디스패치 (PR diff 기준 코드 검증) → VERDICT 수신
         · REWORK   → 사유 첨부 + 같은 브랜치/PR 에 developer 재디스패치
                      (Review → In Progress 복귀. 재작업은 push 로 PR 자동 갱신)
         · APPROVE  → tester **foreground** 디스패치 (Review → Testing 전이) → PASS/FAIL
                      ⚠️ background 디스패치 금지 — 세션 종료 시 완료 알림 유실 →
                      Testing 상태 이슈의 실제 완료 여부를 보드로 판단 불가(§1.2-2 위반).
                      foreground 대기가 어려우면 세션을 유지하거나, 재개 시 세션 시작
                      3단계(Testing 이슈 코멘트 확인)로 복구한다.
                      · FAIL   → 결함 코멘트 첨부, developer 재디스패치 (Testing → In Progress)
                      · PASS   → PM: AC 최종 대조(§4 검증 게이트) 완료 후
                                 gh pr ready <N>
                                 gh pr review <N> --approve --body "검증 종합: ..."
                                 → [사람 게이트] Julio에게 PR 링크 + 검증 증거 3종 알림
                                 → Julio: GitHub에서 Merge
                                 → GitHub→Discord 웹훅(§7.4) → PM 수신
                                   (수신 즉시 gh pr view <N> --json state 로 머지 재확인)
                                 → PM: Done 전이 + gh issue close + 다음 이슈
         · ESCALATE → 사람에게 보고 후 대기

  재작업 한도: 동일 이슈 3회 REWORK/FAIL 시 자동 ESCALATE.
  (무한 루프 방지. 3회 실패는 이슈 분해가 잘못되었다는 신호로 간주)

## 3. 디스패치 프롬프트 표준

  서브에이전트 호출 시 반드시 포함:
  - 이슈 번호와 제목, 이슈 본문의 AC 전문
  - 참조 산출물 경로 (설계서 등)
  - 산출물 저장 위치 규약
  - (재작업 시) 직전 VERDICT/결함 내용과 이전 산출물 경로
  - "표준 REPORT 양식으로 요약만 반환" 지시

  **재작업(REWORK) 재디스패치 시 PM 필수 선행 작업**:
  developer 재디스패치 전에 PM이 직접 이슈에 재작업 지시 코멘트를 게시한다:

  ```
  ## 🔁 재작업 지시 — #<이슈번호> (Rework <N>회차)
  **근거**: reviewer VERDICT 코멘트 (위)
  **수정 요구 목록**:
  1. <reviewer 수정 요구 1>
  2. <reviewer 수정 요구 2>
  **이전 산출물**: <경로>
  **전 항목 응답 필수** — 항목별 대응 결과를 다음 REPORT에 포함할 것
  ```

  이 코멘트가 없으면 이슈 히스토리에서 "왜 재작업이 일어났는지" 추적 불가.
  tester FAIL 재디스패치도 동일하게 적용 (코멘트 제목: `## 🔁 재작업 지시 — 테스트 실패`).

## 4. 검증 게이트 절차 (PM 본인 수행분)

  reviewer/tester 와 별개로 PM 은 상태 전이 직전에 확인한다:
  a. REPORT 양식 준수 여부 (미준수 시 그 자리에서 반려)
  b. ARTIFACTS 경로의 파일이 실제 존재하는가 (ls 로 확인)
  c. AC 항목 수 == 판정된 항목 수 (누락 적발)
  d. 증거 링크가 이슈 코멘트에 게시되었는가
  e. **[사람 게이트 직전 필수]** PR 이 Draft 상태가 아닌가 —
     `gh pr view <N> --json isDraft` 로 확인.
     isDraft=true 이면 `gh pr ready <N>` 을 먼저 실행하고 나서 사람에게 알린다.
     Draft 상태인 PR 을 사람에게 머지 요청해서는 안 된다.

## 5. 사람 에스컬레이션 기준 (즉시 보고)

  - **PR 검증 완료 → 사람 머지 게이트**: reviewer APPROVE + tester PASS + PM AC 대조 통과 시
    `gh pr ready`+`gh pr review --approve` 후 PR 링크와 검증 증거 3종을 사람에게 알리고 머지 대기
    (모든 PR이 이 게이트를 거침 — G4 의 "의도된 개입"으로 카운트 구분)
  - NEEDS_DECISION / ESCALATE 수신
  - 범위 변경 요구 발생 (신규 요구사항, AC 변경)
  - 보안/데이터 관련 판단
  - 일정 리스크: 마일스톤 잔여 이슈 대비 진행률 경고
  - 재작업 한도 도달

## 5a. 보드 일관성 규칙 (GitHub)
  - 이슈 생성은 항상 **3단계 1쌍**: `gh issue create` → `gh project item-add` →
    `item-edit(Status=Backlog)`. 세 번째를 생략하면 보드에 Status 없는 유령 항목이 생긴다.
  - 신규 이슈의 초기 Status 는 **Backlog**. Todo 승격은 `/cb:plan-sprint` 가 수행한다.
    Backlog 를 건너뛰고 Todo 로 바로 넣지 않는다.
  - 이슈 일괄 생성 직후 `gh project item-list` 로 보드 항목 수 == 생성 수를 확인한다.
    불일치 시 누락분을 item-add → item-edit(Backlog) 로 보정한다.

## 6. 현황 보고 양식 (/cb:status)

  ## 프로젝트 현황 — <날짜>
  - 마일스톤: <명> (<완료>/<전체> 이슈)
  - 이번 보고 기간 완료: #n, #n …
  - 진행 중: #n(담당 에이전트, 단계), …
  - 차단/대기: #n — 사유, 필요한 결정
  - 리스크: …
  - 다음 액션: …
  → 동일 내용을 마일스톤 이슈(또는 트래킹 이슈)에 코멘트로 박제
