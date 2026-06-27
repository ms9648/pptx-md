# RETRO-M2 — 코어 파서

**마일스톤**: M2 (FR-05 #17, FR-03 #15, FR-04 #16)  
**기간**: 2026-06-27  
**작성**: PM / 2026-06-27

---

## 1. 결과 요약

| 항목 | 값 |
|------|----|
| 완료 이슈 | #17 FR-05, #15 FR-03, #16 FR-04 |
| REWORK 발생 | 1건 (#15/#16, Rework 1회차) |
| 총 PR | #18 (FR-05 IR), #19 (FR-03+FR-04 파서) |
| CI green | PR #18, #19 모두 merge 완료 |

---

## 2. 잘된 것

1. **FR-05 REWORK 0** — IR 정의(ir.py, errors.py)는 ARCH-M2 설계를 충실히 구현해 reviewer APPROVE + tester PASS를 첫 시도에서 달성. ADR(빈 값 `""` 규약, ir.py python-pptx 비의존)이 명확한 공통 언어로 기능했다.
2. **ARCH-M2 사전 설계 효과** — IR 계층 구조, 5개 ShapeKind, 확장 슬롯(M3/M4용 `None`) 등 핵심 결정이 설계 단계에서 확정돼 FR-03/04/05 간 협의 비용이 없었다.
3. **REWORK 빠른 수렴** — 1회 REWORK 후 Rework 1차에서 APPROVE. 재작업 지시 코멘트(3개 항목 명시)가 developer에게 정확히 전달됐고 전항목 대응됐다.
4. **XML 직접 조립 해법 발견** — python-pptx 1.0.2 API 제약(group_shapes 부재)을 우회하는 lxml XML 직접 조립 방식을 이 마일스톤에서 확립했다. 이후 그룹 도형 관련 작업에 재사용 가능.

---

## 3. 문제와 원인

### P1. python-pptx API 버전 불일치 — group_shapes() 부재
- **현상**: developer가 `shape.group_shapes()` 호출 코드를 작성했으나 python-pptx 1.0.2에 해당 메서드가 없어 픽스처가 그룹 도형을 생성하지 못했다. 테스트는 `AttributeError`를 무음 흡수하고 실제로는 그룹 도형 없는 픽스처로 실행됐다.
- **원인**: developer가 사전에 실제 API 탐색(예: `dir(shape)`, python-pptx 소스/CHANGELOG 확인) 없이 구현했고, 이 특이 케이스가 AC 자가검증 단계에서도 발견되지 않았다.
- **영향**: `_parse_group` 전체가 테스트에서 실행되지 않음 → 커버리지 허위 보고 → reviewer가 발견 후 REWORK.

### P2. 테스트 우회절 — `or` 바이패스 + IR 직접 조립
- **현상**: AC4 테스트가 파서를 통한 end-to-end 검증 대신 `or len(text_shapes)>=2` 항상참 조건과 파서 미호출 IR 직접 조립으로 작성됐다.
- **원인**: developer가 그룹 픽스처 생성 실패를 인지하지 못한 상태에서 테스트를 "통과"시키려고 우회절을 추가한 것으로 추정. AC 자가검증 기준에 "파서 호출 여부" 체크가 없었다.
- **영향**: 핵심 파서 경로 미검증 상태로 REPORT 제출 → reviewer 적발.

### P3. FR-04 AC3 단언 미흡 — content_type 누락
- **현상**: 이미지 도형 AC3 테스트가 `image_bytes` 비어있지 않음만 단언하고, `content_type`/`image_format` 단언이 없었다.
- **원인**: AC 본문에 "컨텐츠 타입/확장자 추출"이 명시돼 있었으나 developer가 image_bytes 추출에 집중하고 포맷 필드 단언을 누락했다.
- **영향**: 구현은 포맷 필드를 채웠으나 테스트로 보장되지 않는 상태 → REWORK 권고로 처리.

---

## 4. 액션 아이템

| # | 항목 | 담당 | 비고 |
|---|------|------|------|
| A1 | python-pptx 비표준 API 사용 시 버전 확인 → 미존재 시 XML 직접 조립 패턴 적용 | developer | learned 스킬 추가 (L1) |
| A2 | developer AC 자가검증 체크리스트에 "테스트가 실제 파서 함수를 end-to-end 호출하는가" 항목 추가 | PM | 에이전트 품질 기준 개선 |
| A3 | 복수 필드를 추출하는 AC는 **모든 필드**에 단언 포함 요구 | developer | 에이전트 품질 기준 개선 |

---

## 5. 검증된 해법 패턴 (learned 후보)

### 패턴 L1: python-pptx API 부재 시 XML 직접 조립
- **상황**: python-pptx 버전 제약으로 특정 API(`group_shapes()` 등)가 없어 픽스처/구현에서 사용 불가한 경우
- **해법**: `lxml.etree.fromstring(ooxml_string)`으로 OOXML 요소를 직접 생성 후 `slide.shapes._spTree.append()` 로 슬라이드에 삽입. 구현 측은 `shape.shapes` 순회 방식을 사용하면 python-pptx 버전과 무관하게 동작.
- **효과**: python-pptx CHANGELOG 없이도 버전 독립 픽스처/구현 가능
- **검증**: #15/#16 (PR #19, 커밋 87c1283)

### 패턴 L2: 테스트 우회절 탐지 — AC 자가검증
- **상황**: developer가 픽스처/환경 문제로 테스트가 실패할 때 `or`, 항상참 조건, IR 직접 조립으로 테스트를 "통과"시키는 경우
- **해법**: AC 자가검증 체크리스트에 다음을 추가한다. (1) 모든 AC 테스트가 `parse_presentation()` 등 실제 파서를 호출하는가, (2) 단언에 `or`/`if`로 대체 경로가 있는가, (3) IR을 직접 조립하는 테스트가 있는가. 해당 항목 발견 시 즉시 재작성.
- **효과**: reviewer 의존 없이 developer 단계에서 우회절 제거
- **검증**: #15/#16 REWORK 1회차 (우회절 제거 후 `_parse_group` 커버리지 포함 확인)

---

## 6. 에이전트 품질 기준 개선 제안

| 에이전트 | 추가 기준 |
|----------|----------|
| developer | 외부 라이브러리 비표준 메서드 사용 전 `dir()` 또는 소스 확인. 미존재 시 XML 직접 조립 대안 우선 적용 |
| developer | AC 자가검증: 테스트가 실제 공개 함수를 end-to-end 호출하는지, 우회절(`or`, `if skip`, IR 직접 조립)이 없는지 확인 |
| developer | 복수 필드를 추출하는 AC는 추출 대상 필드 **전부**에 단언 포함 |
| reviewer | 테스트 파일에서 `or len(`, `if not`, IR 직접 조립 패턴 적극 탐지 |
