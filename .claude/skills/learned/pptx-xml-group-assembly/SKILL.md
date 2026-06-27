---
name: learned-pptx-xml-group-assembly
description: >
  python-pptx에서 특정 API(group_shapes() 등)가 버전 제약으로 없을 때,
  lxml XML 직접 조립으로 그룹/복합 도형 픽스처를 만드는 패턴.
  pptx-md #15/#16 에서 검증됨.
---

## 문제

python-pptx 1.0.2 이하에서 `shape.group_shapes()` 메서드가 없다.
이 메서드를 호출하면 `AttributeError`가 발생하며, 무음 흡수(`try/except`) 시
픽스처가 빈 채로 생성돼 테스트가 의미 없이 통과한다.

## 검증된 해법 (절차)

픽스처 생성 측:
1. python-pptx `Presentation`, `Slide` 객체는 그대로 사용
2. 그룹 도형을 OOXML로 직접 조립:

```python
from lxml import etree

grp_xml = """<p:grpSp xmlns:p="..." xmlns:a="..." xmlns:r="...">
  <p:grpSpPr>...</p:grpSpPr>
  <p:sp>...</p:sp>   <!-- 자식 도형 -->
</p:grpSp>"""

sp_tree = slide.shapes._spTree
sp_tree.append(etree.fromstring(grp_xml))
```

3. 중첩 그룹도 동일하게 outer > inner > sp 구조로 XML 조립 후 append

구현 측:
- `shape.shapes` 순회(python-pptx 표준)를 사용하면 버전 무관하게 동작
- `shape.shape_type == MSO_SHAPE_TYPE.GROUP` 판별 후 `shape.shapes` 재귀

## 주의점

- `_spTree`는 내부 API — 공식 지원 아님. python-pptx 메이저 업그레이드 시 재확인 필요
- OOXML 네임스페이스(`xmlns:p`, `xmlns:a`, `xmlns:r`)를 정확히 명시해야 파싱됨
- 중첩 그룹의 경우 `<p:grpSpPr>`의 `<p:xfrm>` 변환 행렬 없어도 파싱은 됨

## 증거

- 이슈 #15, #16, PR #19, 커밋 87c1283
- `tests/conftest.py` `pptx_with_group` 픽스처 (L194-250)
