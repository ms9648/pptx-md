---
name: learned-m1-single-pr-milestone
description: >
  마일스톤 내 FR이 2개 이하이고 서로 강한 의존 관계(한 FR의 산출물이 다른 FR의 전제)일 때,
  단일 PR에 전체 구현을 묶되 이슈별 커밋을 분리하는 패턴.
  pptx-md #8+#9 에서 비의도적으로 발생하여 결과적으로 유효함이 검증됨.
---

## 문제

마일스톤 범위가 좁고 FR 간 의존이 강할 때(예: FR-A의 pyproject.toml이 FR-B의 ci.yml의 전제),
이슈를 분리하면 developer가 FR-B 산출물을 FR-A PR에 자연스럽게 포함시키게 된다.
이 경우 FR-B 이슈가 코드 변경 없이 형식적인 루프만 거치는 오버헤드가 발생한다.

## 검증된 해법 (절차)

architect 단계:
1. 단일 ARCH 문서로 두 FR을 함께 설계 (WBS는 이슈 경계로 분리 기술)
2. "단일 PR 처리 여부"를 ARCH 문서 WBS 절에 명시

developer 단계:
1. 단일 feature 브랜치에서 구현하되, 커밋은 이슈별로 분리
   - `feat(#N): FR-A 구현 내용`
   - `feat(#M): FR-B 구현 내용`
2. PR 본문에 `Closes #N` + `Closes #M` 모두 기재
3. ARTIFACTS 절에 이슈별 산출물 파일 경로 명시

PM 단계:
1. 단일 PR 머지 후 두 이슈 모두 Done 전이 + close

## 주의점

- FR이 3개 이상이거나 의존 관계가 복잡하면 이 패턴 적용 금지 (추적성 붕괴 위험)
- 반드시 커밋을 이슈별로 분리해야 `git log --follow` 추적 가능
- 이슈 경계를 넘는 구현은 PM이 사전 인지해야 함 (비의도적 포함은 회고에서 발견됨)

## 증거

- 이슈 #8(FR-01), #9(FR-02), PR #10 — pptx-md M1
- 커밋 `12a9bb4`(FR-01 구현)에 ci.yml(FR-02 범위) 포함, 결과적으로 AC 전 항목 충족
