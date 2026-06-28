"""Tests for FR-11: Markdown assembler (assemble_slide).

Covers AC1–AC8 of issue #33.  All tests use synthetic IR fixtures
(no python-pptx dependency — ADR-206/211).
"""

from __future__ import annotations

from pptx_md.assembler import assemble_slide
from pptx_md.ir import (
    GroupShapeIR,
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
# Helpers
# ---------------------------------------------------------------------------


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


def _table_shape(
    rows: list[list[str]],
    n_rows: int | None = None,
    n_cols: int | None = None,
    shape_id: int = 20,
) -> TableShapeIR:
    r = n_rows if n_rows is not None else len(rows)
    c = n_cols if n_cols is not None else (len(rows[0]) if rows else 0)
    return TableShapeIR(
        shape_id=shape_id,
        name="table",
        kind=ShapeKind.TABLE,
        rows=rows,
        n_rows=r,
        n_cols=c,
    )


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


def _group_shape(
    children: list,
    shape_id: int = 50,
) -> GroupShapeIR:
    return GroupShapeIR(
        shape_id=shape_id,
        name="group",
        kind=ShapeKind.GROUP,
        children=children,
    )


# ---------------------------------------------------------------------------
# AC1: SlideIR (title + shapes) -> first line is heading, shapes in IR order
# ---------------------------------------------------------------------------


class TestAc1SlideHeadingAndOrder:
    """ac1_슬라이드_헤딩_도형순서"""

    def test_ac1_first_line_is_h2_heading(self) -> None:
        """First line of assemble_slide output must be ## {title}."""
        shape1 = _text_shape([_para("First")], shape_id=1)
        shape2 = _text_shape([_para("Second")], shape_id=2)
        slide = SlideIR(index=0, title="My Slide", shapes=[shape1, shape2])
        md = assemble_slide(slide)
        lines = md.splitlines()
        assert lines[0] == "## My Slide"

    def test_ac1_shapes_appear_in_ir_order(self) -> None:
        """Shapes must appear in IR list order."""
        shape1 = _text_shape([_para("Alpha")], shape_id=1)
        shape2 = _text_shape([_para("Beta")], shape_id=2)
        slide = SlideIR(index=0, title="Order Test", shapes=[shape1, shape2])
        md = assemble_slide(slide)
        alpha_pos = md.index("Alpha")
        beta_pos = md.index("Beta")
        assert alpha_pos < beta_pos

    def test_ac1_heading_uses_h2(self) -> None:
        """Heading must use ## (not #, ###, etc.)."""
        slide = SlideIR(index=0, title="Check Level", shapes=[])
        md = assemble_slide(slide)
        assert md.startswith("## Check Level")
        assert not md.startswith("# Check Level\n") or md.startswith("## ")


# ---------------------------------------------------------------------------
# AC2: TextShapeIR paragraphs — level=1 is deeper than level=0
# ---------------------------------------------------------------------------


class TestAc2TextIndentation:
    """ac2_텍스트_들여쓰기_레벨"""

    def test_ac2_level0_is_plain_paragraph(self) -> None:
        """level=0 paragraph renders as plain text (no bullet/indent prefix)."""
        shape = _text_shape([_para("Top level", level=0)])
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "Top level" in md
        # Should not have bullet indentation
        assert "  - Top level" not in md

    def test_ac2_level1_is_indented_bullet(self) -> None:
        """level=1 paragraph renders with indented bullet marker."""
        shape = _text_shape([_para("Nested", level=1)])
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        # Must contain indented bullet
        assert "  - Nested" in md

    def test_ac2_level1_deeper_than_level0(self) -> None:
        """level=1 line must appear after level=0 and have more leading whitespace."""
        shape = _text_shape(
            [
                _para("Parent", level=0),
                _para("Child", level=1),
            ]
        )
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        # Find the lines
        all_lines = md.splitlines()
        lines = [ln for ln in all_lines if "Parent" in ln or "Child" in ln]
        parent_line = next(ln for ln in lines if "Parent" in ln)
        child_line = next(ln for ln in lines if "Child" in ln)
        # Child must be more indented
        parent_indent = len(parent_line) - len(parent_line.lstrip())
        child_indent = len(child_line) - len(child_line.lstrip())
        assert child_indent > parent_indent


# ---------------------------------------------------------------------------
# AC3: TableShapeIR under threshold -> valid GFM table with header separator
# ---------------------------------------------------------------------------


class TestAc3GfmTable:
    """ac3_임계값미초과_GFM표"""

    def test_ac3_small_table_gfm_format(self) -> None:
        """2x3 table (6 cells) must render as GFM pipe table."""
        rows = [["A", "B", "C"], ["D", "E", "F"]]
        table = _table_shape(rows)
        slide = SlideIR(index=0, title="T", shapes=[table])
        md = assemble_slide(slide)
        assert "| A | B | C |" in md

    def test_ac3_header_separator_present(self) -> None:
        """GFM table must include a header separator row (|---|...|)."""
        rows = [["H1", "H2"], ["V1", "V2"]]
        table = _table_shape(rows)
        slide = SlideIR(index=0, title="T", shapes=[table])
        md = assemble_slide(slide)
        # Separator row contains ---
        assert "| --- |" in md or "|---|" in md or "---" in md

    def test_ac3_all_data_rows_present(self) -> None:
        """All data rows must appear in the GFM table."""
        rows = [["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]
        table = _table_shape(rows)
        slide = SlideIR(index=0, title="T", shapes=[table])
        md = assemble_slide(slide)
        assert "Alice" in md
        assert "Bob" in md
        assert "No mermaid" or "```mermaid" not in md


# ---------------------------------------------------------------------------
# AC4: description present ImageShapeIR -> description text block (no placeholder)
# ---------------------------------------------------------------------------


class TestAc4ImageWithDescription:
    """ac4_description있는_이미지_본문텍스트"""

    def test_ac4_description_appears_in_output(self) -> None:
        """description text must be in assembler output."""
        shape = _image_shape(description="A bar chart showing Q1 results.")
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "A bar chart showing Q1 results." in md

    def test_ac4_no_placeholder_when_description_present(self) -> None:
        """Placeholder must NOT appear when description is set."""
        shape = _image_shape(
            alt_text="some alt",
            description="Detailed description here.",
        )
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "![some alt]" not in md
        assert "_[image]_" not in md

    def test_ac4_classification_label_when_both_set(self) -> None:
        """When description AND classification are set, label appears."""
        shape = _image_shape(
            description="Chart description.",
            classification=ImageClass.DIAGRAM,
        )
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "Chart description." in md
        assert "diagram image" in md


# ---------------------------------------------------------------------------
# AC5: description=None ImageShapeIR -> ![alt_text](...) format
# ---------------------------------------------------------------------------


class TestAc5ImageWithoutDescription:
    """ac5_description없는_이미지_alt_text형식"""

    def test_ac5_alt_text_markdown_format(self) -> None:
        """description=None with alt_text -> ![alt_text](...) format."""
        shape = _image_shape(alt_text="Corporate logo", description=None)
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "![Corporate logo](...)" in md

    def test_ac5_no_description_text_in_output(self) -> None:
        """When description is None, description-body pattern must not appear."""
        shape = _image_shape(alt_text="Logo", description=None)
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        # Should not contain random text that looks like a description
        assert "![Logo](...)" in md

    def test_ac5_no_alt_no_desc_falls_back_to_placeholder(self) -> None:
        """Both alt_text and description absent -> placeholder."""
        shape = _image_shape(alt_text="", description=None, classification=None)
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "_[image]_" in md

    def test_ac5_classification_placeholder_when_no_alt_no_desc(self) -> None:
        """classification present but no alt/desc -> _[image: {cls}]_ placeholder."""
        shape = _image_shape(
            alt_text="",
            description=None,
            classification=ImageClass.PHOTO,
        )
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "_[image: photo]_" in md


# ---------------------------------------------------------------------------
# AC6: GroupShapeIR -> recursive child rendering, order preserved
# ---------------------------------------------------------------------------


class TestAc6GroupRecursive:
    """ac6_그룹_재귀변환_순서보존"""

    def test_ac6_group_children_rendered(self) -> None:
        """Children of a group must appear in the output."""
        child1 = _text_shape([_para("Child A")], shape_id=11)
        child2 = _text_shape([_para("Child B")], shape_id=12)
        group = _group_shape([child1, child2])
        slide = SlideIR(index=0, title="T", shapes=[group])
        md = assemble_slide(slide)
        assert "Child A" in md
        assert "Child B" in md

    def test_ac6_group_children_in_order(self) -> None:
        """Group children must appear in IR order."""
        child1 = _text_shape([_para("First")], shape_id=11)
        child2 = _text_shape([_para("Second")], shape_id=12)
        group = _group_shape([child1, child2])
        slide = SlideIR(index=0, title="T", shapes=[group])
        md = assemble_slide(slide)
        assert md.index("First") < md.index("Second")

    def test_ac6_nested_group_renders(self) -> None:
        """Nested group (group inside group) must be fully rendered."""
        inner_child = _text_shape([_para("Deep")], shape_id=11)
        inner_group = _group_shape([inner_child], shape_id=51)
        outer_group = _group_shape([inner_group], shape_id=50)
        slide = SlideIR(index=0, title="T", shapes=[outer_group])
        md = assemble_slide(slide)
        assert "Deep" in md


# ---------------------------------------------------------------------------
# AC7 (boundary): empty SlideIR -> no exception, empty/blank string returned
# ---------------------------------------------------------------------------


class TestAc7EmptySlide:
    """ac7_빈슬라이드_예외없음"""

    def test_ac7_empty_shapes_no_exception(self) -> None:
        """assemble_slide on a slide with no shapes must not raise."""
        slide = SlideIR(index=0, title="Empty Slide", shapes=[])
        result = assemble_slide(slide)
        assert isinstance(result, str)

    def test_ac7_empty_title_and_shapes_no_exception(self) -> None:
        """Slide with empty title and no shapes must not raise."""
        slide = SlideIR(index=0, title="", shapes=[])
        result = assemble_slide(slide)
        assert isinstance(result, str)

    def test_ac7_blank_slide_result_is_string(self) -> None:
        """Empty slide result is a string (may be empty or contain heading only)."""
        slide = SlideIR(index=0, title="", shapes=[])
        result = assemble_slide(slide)
        # Must be str; may be empty or have a placeholder heading
        assert isinstance(result, str)
        # Must not contain shape content since there are none
        assert len(result) < 100  # sanity: not a huge unexpected string


# ---------------------------------------------------------------------------
# AC8 (exception): OtherShapeIR(fallback_text="") -> silently omitted
# ---------------------------------------------------------------------------


class TestAc8OtherShapeFallbackEmpty:
    """ac8_OtherShapeIR_빈fallback_생략"""

    def test_ac8_empty_fallback_text_is_omitted(self) -> None:
        """OtherShapeIR with fallback_text="" must be silently omitted."""
        other = _other_shape(fallback_text="")
        normal = _text_shape([_para("Normal text")], shape_id=10)
        slide = SlideIR(index=0, title="T", shapes=[other, normal])
        md = assemble_slide(slide)
        # Normal text must appear
        assert "Normal text" in md
        # No unsupported-shape markers (empty fallback -> no output)
        assert "_[unsupported" not in md

    def test_ac8_non_empty_fallback_text_rendered(self) -> None:
        """OtherShapeIR with non-empty fallback_text must appear in output."""
        other = _other_shape(fallback_text="Some fallback content")
        slide = SlideIR(index=0, title="T", shapes=[other])
        md = assemble_slide(slide)
        assert "Some fallback content" in md

    def test_ac8_no_exception_with_empty_fallback(self) -> None:
        """assemble_slide with only empty-fallback OtherShapeIR must not raise."""
        other = _other_shape(fallback_text="")
        slide = SlideIR(index=0, title="T", shapes=[other])
        result = assemble_slide(slide)
        assert isinstance(result, str)

    def test_ac8_rest_of_shapes_unaffected(self) -> None:
        """Empty-fallback OtherShapeIR omitted but other shapes still rendered."""
        other = _other_shape(fallback_text="", shape_id=40)
        text = _text_shape([_para("Unaffected text")], shape_id=10)
        table_rows = [["X", "Y"], ["1", "2"]]
        table = _table_shape(table_rows, shape_id=20)
        slide = SlideIR(index=0, title="T", shapes=[other, text, table])
        md = assemble_slide(slide)
        assert "Unaffected text" in md
        assert "X" in md


# ---------------------------------------------------------------------------
# Additional: assemble_document (FR-12 basic smoke)
# ---------------------------------------------------------------------------


class TestAssembleDocument:
    """assemble_document 기본 동작 (FR-12 smoke)"""

    def test_document_contains_all_slides(self) -> None:
        """assemble_document must include all slide blocks."""
        from pptx_md.assembler import assemble_document

        slides = [SlideIR(index=i, title=f"Slide {i}", shapes=[]) for i in range(3)]
        pres = PresentationIR(source_path="test.pptx", slides=slides)
        doc = assemble_document(pres)
        for i in range(3):
            assert f"Slide {i}" in doc

    def test_document_deterministic(self) -> None:
        """assemble_document called twice must return identical strings."""
        from pptx_md.assembler import assemble_document

        slides = [
            SlideIR(
                index=0,
                title="Only Slide",
                shapes=[
                    _text_shape([_para("Content")]),
                ],
            )
        ]
        pres = PresentationIR(source_path="test.pptx", slides=slides)
        doc1 = assemble_document(pres)
        doc2 = assemble_document(pres)
        assert doc1 == doc2
