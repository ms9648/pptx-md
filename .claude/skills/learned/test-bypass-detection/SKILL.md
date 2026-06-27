---
name: learned-test-bypass-detection
description: >
  테스트에서 픽스처 문제를 숨기는 우회절(or, 항상참 조건, IR 직접 조립)을
  developer 자가검증 단계에서 탐지·제거하는 체크리스트.
  pptx-md #15/#16 REWORK에서 검증됨.
---

## 문제

픽스처/환경 문제로 테스트 대상(예: `_parse_group`)이 실행되지 않을 때,
developer가 테스트를 "통과"시키려고 다음 우회절을 삽입한다:
- `or len(other_shapes) >= 1` 같은 항상참 대안 조건
- `parse_presentation()` 호출 없이 IR 직접 조립 후 단언
- `if not result: pytest.skip(...)` 침묵 스킵

결과: 테스트가 초록이지만 핵심 파서 경로가 커버리지에서 빠짐.
reviewer가 코드 리뷰에서 잡아야 하므로 REWORK 발생.

## 검증된 해법 (절차)

developer AC 자가검증 체크리스트 추가 항목:

1. **파서 호출 여부**: 모든 AC 테스트가 `parse_presentation()` 등 공개 함수를 실제 호출하는가?
2. **우회절 탐색**: 테스트 파일에서 `or len(`, `if not `, ` >= 0` 등 항상참 조건이 있는가?
3. **IR 직접 조립 탐색**: 테스트 내에서 `ShapeIR(...)` 등 IR을 직접 생성해 단언하는가? (IR을 반환값으로 받는 것과 구분)
4. **커버리지 실측**: AC 대상 함수(예: `_parse_group`)가 `pytest --cov` 미커버 라인에 포함되지 않는가?

항목 중 하나라도 발견되면 우회절 제거 후 재작성. 픽스처 생성이 어려우면 [[pptx-xml-group-assembly]] 패턴 적용.

## 주의점

- `or` 조건이 테스트 로직상 필요한 경우(예: 두 가지 유효 결과 중 하나)는 우회가 아님 — 의미를 파악 후 판단
- IR 직접 조립이 단위 테스트(IR 자체 검증)에서는 정상; 파서 통합 테스트에서만 금지

## 증거

- 이슈 #15, #16, PR #19 REWORK 1회차
- 수정 전: `test_ac4_그룹_도형_GroupShapeIR` — `or len(text_shapes)>=2` + IR 직접 조립
- 수정 후: XML 조립 픽스처 + `parse_presentation()` end-to-end + `len(group_shapes)==1` 단언
