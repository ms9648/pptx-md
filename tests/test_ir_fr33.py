"""Tests for SlideIR.reconstructed_md slot — FR-33 (issue #95, ADR-627, INV-2).

WBS-O AC2: SlideIR.reconstructed_md 슬롯(default None, INV-2 하위호환 —
기존 호출·테스트 무손상).

Test naming convention: test_ac<N>_<description>.
"""

from __future__ import annotations

import dataclasses

from pptx_md.ir import ParagraphIR, ShapeKind, SlideIR, TextShapeIR


def test_ac2_reconstructed_md_defaults_to_none() -> None:
    """AC2: 기존 호출부처럼 reconstructed_md를 지정하지 않으면 default는 None이다."""
    slide = SlideIR(index=0, title="제목")
    assert slide.reconstructed_md is None


def test_ac2_existing_positional_and_keyword_construction_still_works() -> None:
    """AC2 (INV-2): 기존 SlideIR(index=..., title=..., shapes=[...]) 호출이 무손상."""
    shape = TextShapeIR(
        shape_id=1,
        name="body",
        kind=ShapeKind.TEXT,
        paragraphs=[ParagraphIR(text="hello", level=0)],
    )
    slide = SlideIR(index=2, title="t", notes="n", shapes=[shape])
    assert slide.index == 2
    assert slide.title == "t"
    assert slide.notes == "n"
    assert slide.shapes == [shape]
    assert slide.reconstructed_md is None


def test_ac2_reconstructed_md_can_be_set_explicitly() -> None:
    """AC2: reconstructed_md는 명시적으로 문자열 값을 채울 수 있다(슬롯 충전 대상)."""
    slide = SlideIR(index=0, reconstructed_md="## Reconstructed\n\nbody")
    assert slide.reconstructed_md == "## Reconstructed\n\nbody"


def test_ac2_in_place_mutation_after_construction() -> None:
    """AC2: reconstruct_slides가 사용하는 in-place 슬롯 충전 패턴이 가능하다."""
    slide = SlideIR(index=0)
    assert slide.reconstructed_md is None
    slide.reconstructed_md = "## filled"
    assert slide.reconstructed_md == "## filled"


def test_ac2_field_is_declared_on_dataclass() -> None:
    """AC2: reconstructed_md가 실제 dataclass 필드로 선언되어 있다(타입 str | None)."""
    fields = {f.name: f for f in dataclasses.fields(SlideIR)}
    assert "reconstructed_md" in fields
    assert fields["reconstructed_md"].default is None
