"""Tests for FR-12: Map-Reduce pipeline (assemble_document).

Covers AC1–AC8 of issue #34.  All tests use synthetic IR fixtures
(no python-pptx dependency — ADR-206/211).

AC mapping:
    AC1  two slides -> both blocks in single str
    AC2  slides in index=2,0,1 order -> output in index 0,1,2 order
    AC3  exactly one separator between adjacent blocks; none at start/end
    AC4  description=None IR -> no exception, no VLM SDK import
    AC5  same IR twice -> identical output (determinism)
    AC6  slides=[] -> no exception, empty string returned
    AC7  slide containing OtherShapeIR -> whole document still produced
    AC8  20-slide IR (VLM excluded) -> completed within 5 seconds
"""

from __future__ import annotations

import time

from pptx_md.assembler import assemble_document
from pptx_md.ir import (
    ImageClass,
    ImageShapeIR,
    OtherShapeIR,
    ParagraphIR,
    PresentationIR,
    ShapeKind,
    SlideIR,
    TableShapeIR,
    TextShapeIR,
)

# ---------------------------------------------------------------------------
# Helpers — IR factory functions (shared across tests)
# ---------------------------------------------------------------------------

_SEPARATOR = "\n\n---\n\n"


def _text_shape(
    paragraphs: list[ParagraphIR],
    shape_id: int = 10,
    is_title: bool = False,
) -> TextShapeIR:
    return TextShapeIR(
        shape_id=shape_id,
        name="text",
        kind=ShapeKind.TEXT,
        paragraphs=paragraphs,
        is_title=is_title,
    )


def _para(text: str, level: int = 0) -> ParagraphIR:
    return ParagraphIR(text=text, level=level)


def _image_shape(
    alt_text: str = "",
    classification: ImageClass | None = None,
    description: str | None = None,
    shape_id: int = 30,
) -> ImageShapeIR:
    return ImageShapeIR(
        shape_id=shape_id,
        name="image",
        kind=ShapeKind.IMAGE,
        image_bytes=b"",
        image_format="",
        image_ext="",
        alt_text=alt_text,
        classification=classification,
        description=description,
    )


def _other_shape(fallback_text: str = "", shape_id: int = 40) -> OtherShapeIR:
    return OtherShapeIR(
        shape_id=shape_id,
        name="other",
        kind=ShapeKind.OTHER,
        fallback_text=fallback_text,
    )


def _table_shape(
    rows: list[list[str]],
    shape_id: int = 20,
) -> TableShapeIR:
    n_rows = len(rows)
    n_cols = len(rows[0]) if rows else 0
    return TableShapeIR(
        shape_id=shape_id,
        name="table",
        kind=ShapeKind.TABLE,
        rows=rows,
        n_rows=n_rows,
        n_cols=n_cols,
    )


def _make_presentation(slides: list[SlideIR]) -> PresentationIR:
    return PresentationIR(source_path="test.pptx", slides=slides)


# ---------------------------------------------------------------------------
# AC1: slides=[SlideIR(index=0,...), SlideIR(index=1,...)]
#      -> both slide blocks in a single str
# ---------------------------------------------------------------------------


class TestAc1TwoSlidesSingleStr:
    """ac1_두슬라이드_단일str_반환"""

    def test_ac1_both_slide_titles_present(self) -> None:
        """Both slide titles must appear in the assembled document."""
        slides = [
            SlideIR(index=0, title="Slide Zero", shapes=[]),
            SlideIR(index=1, title="Slide One", shapes=[]),
        ]
        pres = _make_presentation(slides)
        doc = assemble_document(pres)
        assert "Slide Zero" in doc
        assert "Slide One" in doc

    def test_ac1_result_is_single_str(self) -> None:
        """Return value must be a single str."""
        slides = [
            SlideIR(index=0, title="A", shapes=[]),
            SlideIR(index=1, title="B", shapes=[]),
        ]
        doc = assemble_document(_make_presentation(slides))
        assert isinstance(doc, str)

    def test_ac1_slide_content_present(self) -> None:
        """Shape content from both slides must appear in the document."""
        shape0 = _text_shape([_para("Content of slide 0")], shape_id=1)
        shape1 = _text_shape([_para("Content of slide 1")], shape_id=2)
        slides = [
            SlideIR(index=0, title="S0", shapes=[shape0]),
            SlideIR(index=1, title="S1", shapes=[shape1]),
        ]
        doc = assemble_document(_make_presentation(slides))
        assert "Content of slide 0" in doc
        assert "Content of slide 1" in doc


# ---------------------------------------------------------------------------
# AC2: slides in index=2,0,1 order -> output in index 0,1,2 ascending order
# ---------------------------------------------------------------------------


class TestAc2IndexAscendingOrder:
    """ac2_index_오름차순_병합"""

    def test_ac2_out_of_order_slides_sorted_by_index(self) -> None:
        """Slides stored as index=2,0,1 must appear in 0,1,2 order in output."""
        slides = [
            SlideIR(index=2, title="Slide Two", shapes=[]),
            SlideIR(index=0, title="Slide Zero", shapes=[]),
            SlideIR(index=1, title="Slide One", shapes=[]),
        ]
        doc = assemble_document(_make_presentation(slides))
        pos_zero = doc.index("Slide Zero")
        pos_one = doc.index("Slide One")
        pos_two = doc.index("Slide Two")
        assert pos_zero < pos_one < pos_two

    def test_ac2_reversed_order_normalized(self) -> None:
        """Slides stored in descending index order must be re-sorted ascending."""
        slides = [
            SlideIR(index=4, title="Last", shapes=[]),
            SlideIR(index=3, title="Fourth", shapes=[]),
            SlideIR(index=2, title="Third", shapes=[]),
            SlideIR(index=1, title="Second", shapes=[]),
            SlideIR(index=0, title="First", shapes=[]),
        ]
        doc = assemble_document(_make_presentation(slides))
        titles = ("First", "Second", "Third", "Fourth", "Last")
        positions = [doc.index(title) for title in titles]
        assert positions == sorted(positions)

    def test_ac2_single_slide_unaffected(self) -> None:
        """Single slide must appear unchanged regardless of its index value."""
        slides = [SlideIR(index=5, title="Only Slide", shapes=[])]
        doc = assemble_document(_make_presentation(slides))
        assert "Only Slide" in doc


# ---------------------------------------------------------------------------
# AC3: exactly one separator between adjacent blocks; none at start/end
# ---------------------------------------------------------------------------


class TestAc3DeterministicSeparator:
    """ac3_구분자_정확히1회_시작끝없음"""

    def test_ac3_separator_between_two_slides(self) -> None:
        """Exactly one separator must appear between two slide blocks."""
        slides = [
            SlideIR(index=0, title="Alpha", shapes=[]),
            SlideIR(index=1, title="Beta", shapes=[]),
        ]
        doc = assemble_document(_make_presentation(slides))
        assert doc.count(_SEPARATOR) == 1

    def test_ac3_separator_count_equals_n_minus_1(self) -> None:
        """N slides must produce exactly N-1 separators."""
        n = 5
        slides = [SlideIR(index=i, title=f"S{i}", shapes=[]) for i in range(n)]
        doc = assemble_document(_make_presentation(slides))
        assert doc.count(_SEPARATOR) == n - 1

    def test_ac3_no_leading_separator(self) -> None:
        """Document must not start with the separator."""
        slides = [
            SlideIR(index=0, title="First", shapes=[]),
            SlideIR(index=1, title="Second", shapes=[]),
        ]
        doc = assemble_document(_make_presentation(slides))
        assert not doc.startswith(_SEPARATOR)
        assert not doc.startswith("\n\n---")

    def test_ac3_no_trailing_separator(self) -> None:
        """Document must not end with the separator."""
        slides = [
            SlideIR(index=0, title="First", shapes=[]),
            SlideIR(index=1, title="Second", shapes=[]),
        ]
        doc = assemble_document(_make_presentation(slides))
        assert not doc.endswith(_SEPARATOR)
        assert not doc.endswith("---\n\n")

    def test_ac3_separator_is_deterministic_string(self) -> None:
        """The separator is the module constant '\\n\\n---\\n\\n' (ADR-218)."""
        slides = [
            SlideIR(index=0, title="P", shapes=[]),
            SlideIR(index=1, title="Q", shapes=[]),
        ]
        doc = assemble_document(_make_presentation(slides))
        # Find the gap between the two slide blocks
        p_end = doc.index("## P") + len("## P")
        q_start = doc.index("## Q")
        gap = doc[p_end:q_start]
        assert gap == _SEPARATOR


# ---------------------------------------------------------------------------
# AC4: description=None IR -> no exception; VLM SDK not imported
# ---------------------------------------------------------------------------


class TestAc4DescriptionNoneNoVlm:
    """ac4_description_None_예외없음_VLM_import없음"""

    def test_ac4_image_description_none_no_exception(self) -> None:
        """ImageShapeIR with description=None must not raise."""
        shape = _image_shape(alt_text="chart", description=None)
        slides = [SlideIR(index=0, title="S", shapes=[shape])]
        doc = assemble_document(_make_presentation(slides))
        assert isinstance(doc, str)

    def test_ac4_all_description_none_no_exception(self) -> None:
        """Multiple images with description=None produce output without error."""
        shapes = [_image_shape(description=None, shape_id=i + 30) for i in range(5)]
        slides = [SlideIR(index=0, title="NoDesc", shapes=shapes)]
        doc = assemble_document(_make_presentation(slides))
        assert isinstance(doc, str)

    def test_ac4_no_vlm_sdk_imported_in_assembler(self) -> None:
        """assembler.py must not import anthropic, openai, or python-pptx (NFR-08)."""
        import importlib
        import sys

        # Reload assembler fresh to capture its imports
        mod_name = "pptx_md.assembler"
        if mod_name in sys.modules:
            mod = sys.modules[mod_name]
        else:
            mod = importlib.import_module(mod_name)

        # The assembler itself must not expose these as attributes
        forbidden = {"anthropic", "openai", "pptx"}
        for name in forbidden:
            assert not hasattr(
                mod, name
            ), f"assembler imported forbidden module: {name}"

    def test_ac4_image_placeholder_when_no_description_no_alt(self) -> None:
        """description=None, alt_text='' -> placeholder output, no exception."""
        shape = _image_shape(alt_text="", description=None, classification=None)
        slides = [SlideIR(index=0, title="T", shapes=[shape])]
        doc = assemble_document(_make_presentation(slides))
        assert "_[image]_" in doc


# ---------------------------------------------------------------------------
# AC5: same IR twice -> identical output (determinism)
# ---------------------------------------------------------------------------


class TestAc5Determinism:
    """ac5_동일IR_2회_동일출력"""

    def test_ac5_two_calls_identical_output(self) -> None:
        """assemble_document called twice with same IR must return equal strings."""
        shape = _text_shape([_para("deterministic"), _para("content", level=1)])
        slides = [
            SlideIR(index=0, title="D0", shapes=[shape]),
            SlideIR(index=1, title="D1", shapes=[_image_shape(alt_text="img")]),
        ]
        pres = _make_presentation(slides)
        doc1 = assemble_document(pres)
        doc2 = assemble_document(pres)
        assert doc1 == doc2

    def test_ac5_out_of_order_slides_deterministic(self) -> None:
        """Determinism holds even when input slides are out of index order."""
        slides = [
            SlideIR(index=2, title="C", shapes=[]),
            SlideIR(index=0, title="A", shapes=[]),
            SlideIR(index=1, title="B", shapes=[]),
        ]
        pres = _make_presentation(slides)
        doc1 = assemble_document(pres)
        doc2 = assemble_document(pres)
        assert doc1 == doc2

    def test_ac5_ir_not_mutated_after_call(self) -> None:
        """IR must not be mutated by assembler (read-only, ADR-218)."""
        shape = _text_shape([_para("immutable")], shape_id=99)
        slide = SlideIR(index=0, title="Orig", shapes=[shape])
        pres = _make_presentation([slide])
        _ = assemble_document(pres)
        # Verify IR is unchanged
        assert slide.title == "Orig"
        assert shape.paragraphs[0].text == "immutable"


# ---------------------------------------------------------------------------
# AC6: slides=[] -> no exception, empty string returned
# ---------------------------------------------------------------------------


class TestAc6EmptySlidesEmptyStr:
    """ac6_빈슬라이드_리스트_빈문자열"""

    def test_ac6_empty_slides_no_exception(self) -> None:
        """assemble_document(slides=[]) must not raise."""
        pres = _make_presentation([])
        result = assemble_document(pres)
        assert isinstance(result, str)

    def test_ac6_empty_slides_returns_empty_string(self) -> None:
        """assemble_document(slides=[]) must return '' (empty string)."""
        pres = _make_presentation([])
        result = assemble_document(pres)
        assert result == ""

    def test_ac6_empty_slides_no_separator(self) -> None:
        """Empty document must not contain the separator."""
        pres = _make_presentation([])
        result = assemble_document(pres)
        assert _SEPARATOR not in result


# ---------------------------------------------------------------------------
# AC7: slide with OtherShapeIR -> whole document still produced
# ---------------------------------------------------------------------------


class TestAc7OtherShapeIrFullDocument:
    """ac7_OtherShapeIR_포함_전체문서_생성"""

    def test_ac7_othershapeir_does_not_abort_document(self) -> None:
        """A slide containing OtherShapeIR must not abort full document assembly."""
        other = _other_shape(fallback_text="fallback content")
        normal_text = _text_shape([_para("Normal paragraph")], shape_id=11)
        slides = [
            SlideIR(index=0, title="Has Other", shapes=[other, normal_text]),
            SlideIR(index=1, title="After Other", shapes=[]),
        ]
        doc = assemble_document(_make_presentation(slides))
        assert "Has Other" in doc
        assert "After Other" in doc

    def test_ac7_empty_fallback_othershapeir_whole_document(self) -> None:
        """OtherShapeIR with empty fallback_text -> shape omitted, doc continues."""
        other_empty = _other_shape(fallback_text="", shape_id=40)
        other_text = _other_shape(fallback_text="visible fallback", shape_id=41)
        slides = [
            SlideIR(
                index=0,
                title="Mixed Other",
                shapes=[other_empty, other_text],
            ),
            SlideIR(index=1, title="Next Slide", shapes=[]),
        ]
        doc = assemble_document(_make_presentation(slides))
        assert "visible fallback" in doc
        assert "Next Slide" in doc

    def test_ac7_only_othershapeir_slide_does_not_abort(self) -> None:
        """Slide consisting solely of OtherShapeIR must not abort assembly."""
        other = _other_shape(fallback_text="sole other shape")
        slides = [
            SlideIR(index=0, title="Only Other", shapes=[other]),
        ]
        doc = assemble_document(_make_presentation(slides))
        assert isinstance(doc, str)
        assert "Only Other" in doc

    def test_ac7_multiple_slides_all_produced(self) -> None:
        """All slides must appear in output even when some contain OtherShapeIR."""
        slides = []
        for i in range(4):
            shapes = [_other_shape(fallback_text=f"other-{i}")] if i % 2 == 0 else []
            slides.append(SlideIR(index=i, title=f"Slide{i}", shapes=shapes))
        doc = assemble_document(_make_presentation(slides))
        for i in range(4):
            assert f"Slide{i}" in doc


# ---------------------------------------------------------------------------
# AC8 (NFR-01): 20-slide IR transformation (VLM excluded) within 5 seconds
# ---------------------------------------------------------------------------


class TestAc8PerformanceTwentySlides:
    """ac8_20슬라이드_5초이내"""

    def _make_rich_slide(self, index: int) -> SlideIR:
        """Create a slide with representative shapes for performance testing."""
        shapes: list = []
        # Several text shapes with multiple paragraphs
        for j in range(3):
            paragraphs = [
                _para(f"Slide {index} body paragraph {j} line 0", level=0),
                _para(f"Slide {index} body paragraph {j} line 1", level=1),
                _para(f"Slide {index} body paragraph {j} line 2", level=2),
            ]
            shapes.append(_text_shape(paragraphs, shape_id=index * 100 + j + 10))
        # Small table
        table_rows = [
            [f"H{index}-1", f"H{index}-2", f"H{index}-3"],
            [f"R{index}-1", f"R{index}-2", f"R{index}-3"],
            [f"R{index}-4", f"R{index}-5", f"R{index}-6"],
        ]
        shapes.append(_table_shape(table_rows, shape_id=index * 100 + 50))
        # Image with description=None (VLM excluded path)
        shapes.append(
            _image_shape(
                alt_text=f"slide-{index}-image",
                classification=None,
                description=None,
                shape_id=index * 100 + 60,
            )
        )
        # OtherShapeIR with fallback
        shapes.append(
            _other_shape(
                fallback_text=f"other-{index}",
                shape_id=index * 100 + 70,
            )
        )
        return SlideIR(index=index, title=f"Performance Slide {index}", shapes=shapes)

    def test_ac8_twenty_slides_within_five_seconds(self) -> None:
        """20-slide IR assembly (VLM excluded) must complete in < 5 seconds (NFR-01)."""
        slides = [self._make_rich_slide(i) for i in range(20)]
        pres = _make_presentation(slides)

        start = time.perf_counter()
        doc = assemble_document(pres)
        elapsed = time.perf_counter() - start

        assert (
            elapsed < 5.0
        ), f"assemble_document took {elapsed:.3f}s for 20 slides (limit: 5s)"
        assert isinstance(doc, str)
        # Sanity: all 20 slides must appear
        for i in range(20):
            assert f"Performance Slide {i}" in doc
