---
description: 특정 이슈에 reviewer+tester 를 즉시 투입해 검증만 수행
---

인자: $ARGUMENTS (이슈 번호, 예: #14)

docs/GUIDE.md 의 §8 /cb:verify 흐름과 pm-orchestration §4 검증 게이트 절차를 따른다.

## 절차

1. **이슈 정보 수집**
   ```bash
   gh issue view $ARGUMENTS --json title,body,comments
   ```
   - AC 전문 확인
   - developer 의 AC-증거 매핑 코멘트 확인
   - 참조 설계 문서 경로 확인

2. **reviewer 디스패치 (MODE=code)**
   reviewer 에이전트를 호출해 코드 검증 체크리스트 7개 항목 전수 검토.
   VERDICT 를 이슈 코멘트에 게시하도록 지시.

3. **VERDICT 수신 및 PM 검증 게이트 (§4)**
   a. REPORT 양식 준수 여부 확인
   b. ARTIFACTS 경로 파일 존재 확인
   c. AC 항목 수 == 판정 항목 수 확인
   d. 이슈 코멘트에 VERDICT 게시 확인

4. **결과에 따른 분기**
   - REWORK → 사유 첨부 + developer 재디스패치 안내
   - APPROVE → tester foreground 디스패치 (Testing 전이)
   - ESCALATE → 사람에게 즉시 보고

5. **tester PASS 시 최종 처리**
   ```bash
   gh pr ready <PR번호>
   gh pr review <PR번호> --approve --body "검증 종합: ..."
   ```
   → 사람에게 PR 링크 + 검증 증거 3종 알림 (머지 게이트)
