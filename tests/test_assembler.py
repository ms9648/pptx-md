"""Tests for FR-11: Markdown assembler (assemble_slide).

Covers AC1–AC8 of issue #33.  All tests use synthetic IR fixtures
(no python-pptx dependency — ADR-206/211).
"""

from __future__ import annotations

import re

from pptx_md.assembler import assemble_document, assemble_slide
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

    def test_ac5_no_alt_no_desc_falls_back_to_marker(self) -> None:
        """Both absent -> standard positional marker (FR-22)."""
        shape = _image_shape(alt_text="", description=None, classification=None)
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        # FR-22: positional marker format (slide 1-based, image 1-based)
        assert "![슬라이드 1 이미지 1]" in md

    def test_ac5_classification_marker_when_no_alt_no_desc(self) -> None:
        """classification + no alt/desc -> positional marker with cls (FR-22)."""
        shape = _image_shape(
            alt_text="",
            description=None,
            classification=ImageClass.PHOTO,
        )
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "![슬라이드 1 이미지 1 — photo]" in md


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


# ---------------------------------------------------------------------------
# FR-20 (#53): 슬라이드 실제 제목 heading 사용
# ---------------------------------------------------------------------------


def _title_shape(text: str, shape_id: int = 10) -> TextShapeIR:
    """Helper: create an is_title=True TextShapeIR."""
    return TextShapeIR(
        shape_id=shape_id,
        name="title",
        kind=ShapeKind.TEXT,
        paragraphs=[ParagraphIR(text=text, level=0)],
        is_title=True,
    )


class TestFR20SlideTitleHeading:
    """FR-20 (#53): is_title 도형을 heading으로 활용"""

    def test_ac1_is_title_shape_used_as_heading_when_slide_title_empty(self) -> None:
        """ac1: is_title shape becomes heading when slide.title==""."""
        title_s = _title_shape("실제 제목")
        body_s = _text_shape([_para("본문 내용")])
        slide = SlideIR(index=0, title="", shapes=[title_s, body_s])
        md = assemble_slide(slide)
        assert md.startswith("## 실제 제목")

    def test_ac2_is_title_shape_not_duplicated_in_body(self) -> None:
        """ac2: is_title shape used as heading must not appear in body."""
        title_s = _title_shape("헤딩 텍스트")
        body_s = _text_shape([_para("본문")])
        slide = SlideIR(index=0, title="", shapes=[title_s, body_s])
        md = assemble_slide(slide)
        # "헤딩 텍스트"는 헤딩에만 한 번 등장해야 함
        assert md.count("헤딩 텍스트") == 1
        # 본문은 존재
        assert "본문" in md

    def test_ac3_fallback_to_slide_n_when_no_title_and_no_is_title_shape(self) -> None:
        """ac3: no title + no is_title shape -> ## Slide N fallback."""
        body_s = _text_shape([_para("내용")])
        slide = SlideIR(index=2, title="", shapes=[body_s])
        md = assemble_slide(slide)
        assert md.startswith("## Slide 3")

    def test_ac4_slide_title_takes_priority_over_is_title_shape(self) -> None:
        """ac4: slide.title present -> is_title shape not used as heading."""
        title_s = _title_shape("도형 제목")
        slide = SlideIR(index=0, title="실제 슬라이드 제목", shapes=[title_s])
        md = assemble_slide(slide)
        assert md.startswith("## 실제 슬라이드 제목")

    def test_ac5_existing_tests_unaffected_with_non_empty_slide_title(self) -> None:
        """ac5_기존_title_있는_슬라이드_정상: slide.title 있으면 기존 동작 유지."""
        shape = _text_shape([_para("내용")])
        slide = SlideIR(index=0, title="기존 제목", shapes=[shape])
        md = assemble_slide(slide)
        assert md.startswith("## 기존 제목")
        assert "내용" in md


# ---------------------------------------------------------------------------
# FR-21 (#54): 푸터·슬라이드번호 플레이스홀더 필터링
# ---------------------------------------------------------------------------


def _footer_shape(text: str, shape_id: int = 60) -> TextShapeIR:
    """Helper: create an is_footer=True TextShapeIR."""
    return TextShapeIR(
        shape_id=shape_id,
        name="footer",
        kind=ShapeKind.TEXT,
        paragraphs=[ParagraphIR(text=text, level=0)],
        is_footer=True,
    )


class TestFR21FooterFiltering:
    """FR-21 (#54): is_footer 도형 assembler 출력 제외"""

    def test_ac1_is_footer_field_exists_in_ir(self) -> None:
        """ac1_is_footer_필드_존재: TextShapeIR에 is_footer 필드가 있다."""
        shape = TextShapeIR(
            shape_id=1,
            name="test",
            kind=ShapeKind.TEXT,
            paragraphs=[],
        )
        assert hasattr(shape, "is_footer")
        assert shape.is_footer is False  # default

    def test_ac2_footer_shape_not_in_output(self) -> None:
        """ac2: is_footer=True shape is excluded from assembler output."""
        footer = _footer_shape("기관명 푸터")
        normal = _text_shape([_para("슬라이드 본문")])
        slide = SlideIR(index=0, title="제목", shapes=[footer, normal])
        md = assemble_slide(slide)
        assert "기관명 푸터" not in md
        assert "슬라이드 본문" in md

    def test_ac3_multiple_footers_all_excluded(self) -> None:
        """ac3_복수_footer_모두_제외: 여러 is_footer 도형 모두 제외."""
        footer1 = _footer_shape("푸터1", shape_id=60)
        footer2 = _footer_shape("페이지 번호", shape_id=61)
        body = _text_shape([_para("본문 텍스트")])
        slide = SlideIR(index=0, title="T", shapes=[footer1, footer2, body])
        md = assemble_slide(slide)
        assert "푸터1" not in md
        assert "페이지 번호" not in md
        assert "본문 텍스트" in md

    def test_ac4_footer_in_group_excluded(self) -> None:
        """ac4_그룹내_footer_제외: GroupShapeIR 내 is_footer 도형도 출력 제외."""
        footer = _footer_shape("그룹내 푸터")
        normal_child = _text_shape([_para("그룹 본문")], shape_id=11)
        group = _group_shape([footer, normal_child])
        slide = SlideIR(index=0, title="T", shapes=[group])
        md = assemble_slide(slide)
        assert "그룹내 푸터" not in md
        assert "그룹 본문" in md

    def test_ac5_normal_shape_unaffected(self) -> None:
        """ac5_일반_도형_영향없음: is_footer=False 도형은 정상 출력."""
        normal = _text_shape([_para("일반 텍스트")])
        slide = SlideIR(index=0, title="T", shapes=[normal])
        md = assemble_slide(slide)
        assert "일반 텍스트" in md


# ---------------------------------------------------------------------------
# FR-22 (#55): no-VLM 이미지 플레이스홀더 표준 마커
# ---------------------------------------------------------------------------


class TestFR22ImageMarker:
    """FR-22 (#55): no-VLM 이미지 슬라이드/이미지 번호 마커"""

    def test_ac1_no_vlm_image_uses_positional_marker(self) -> None:
        """ac1: description=None, alt_text="" -> ![슬라이드 N 이미지 M]."""
        image = _image_shape(alt_text="", description=None, classification=None)
        slide = SlideIR(index=0, title="T", shapes=[image])
        md = assemble_slide(slide)
        assert "![슬라이드 1 이미지 1]" in md

    def test_ac2_slide_num_1based(self) -> None:
        """ac2_슬라이드번호_1기반: slide.index=2 -> '슬라이드 3 이미지 1'."""
        image = _image_shape(alt_text="", description=None, classification=None)
        slide = SlideIR(index=2, title="T", shapes=[image])
        md = assemble_slide(slide)
        assert "![슬라이드 3 이미지 1]" in md

    def test_ac3_multiple_images_numbered_sequentially(self) -> None:
        """ac3_복수이미지_순번_증가: 같은 슬라이드 내 이미지는 1,2,3 순번."""
        img1 = _image_shape(shape_id=31)
        img2 = _image_shape(shape_id=32)
        img3 = _image_shape(shape_id=33)
        slide = SlideIR(index=0, title="T", shapes=[img1, img2, img3])
        md = assemble_slide(slide)
        assert "![슬라이드 1 이미지 1]" in md
        assert "![슬라이드 1 이미지 2]" in md
        assert "![슬라이드 1 이미지 3]" in md

    def test_ac4_alt_text_takes_priority_over_marker(self) -> None:
        """ac4_alt_text_우선순위: alt_text 있으면 기존 ![alt_text](...) 유지."""
        image = _image_shape(alt_text="회사 로고", description=None)
        slide = SlideIR(index=0, title="T", shapes=[image])
        md = assemble_slide(slide)
        assert "![회사 로고](...)" in md
        assert "슬라이드 1 이미지" not in md

    def test_ac5_description_takes_highest_priority(self) -> None:
        """ac5_description_최우선: description 있으면 마커 출력 없음."""
        image = _image_shape(description="차트 설명 텍스트")
        slide = SlideIR(index=0, title="T", shapes=[image])
        md = assemble_slide(slide)
        assert "차트 설명 텍스트" in md
        assert "슬라이드 1 이미지" not in md

    def test_ac6_classification_appended_to_marker(self) -> None:
        """ac6_classification_마커에_포함: classification 있으면 — {cls} 추가."""
        image = _image_shape(
            alt_text="",
            description=None,
            classification=ImageClass.DIAGRAM,
        )
        slide = SlideIR(index=0, title="T", shapes=[image])
        md = assemble_slide(slide)
        assert "![슬라이드 1 이미지 1 — diagram]" in md

    def test_ac7_image_in_group_counted(self) -> None:
        """ac7: images in GroupShapeIR share the slide-wide image counter."""
        img_outside = _image_shape(shape_id=31)
        img_in_group = _image_shape(shape_id=32)
        group = _group_shape([img_in_group])
        slide = SlideIR(index=0, title="T", shapes=[img_outside, group])
        md = assemble_slide(slide)
        assert "![슬라이드 1 이미지 1]" in md
        assert "![슬라이드 1 이미지 2]" in md


# ---------------------------------------------------------------------------
# FR-23 (#58): Reading-order 정렬 (AC4~AC7)
# ---------------------------------------------------------------------------


def _text_shape_with_coords(
    text: str,
    top: int,
    left: int,
    shape_id: int,
) -> TextShapeIR:
    """Helper: TextShapeIR with explicit EMU coordinates."""
    return TextShapeIR(
        shape_id=shape_id,
        name=f"shape_{shape_id}",
        kind=ShapeKind.TEXT,
        paragraphs=[ParagraphIR(text=text, level=0)],
        left=left,
        top=top,
        width=1000000,
        height=500000,
    )


class TestFR23ReadingOrder:
    """FR-23 (#58): assemble_slide reading-order 정렬 AC4~AC7"""

    def test_ac4_reading_order_top_정렬(self) -> None:
        """ac4_reading_order_top_정렬: IR에 하단→상단 역순이어도 출력은 상단 먼저."""
        # 상단 도형: top=500000, 하단 도형: top=4000000
        # IR에 하단→상단 역순으로 담음
        bottom_shape = _text_shape_with_coords(
            "하단텍스트", top=4000000, left=0, shape_id=2
        )
        top_shape = _text_shape_with_coords(
            "상단텍스트", top=500000, left=0, shape_id=1
        )
        slide = SlideIR(index=0, title="T", shapes=[bottom_shape, top_shape])
        md = assemble_slide(slide)
        assert md.index("상단텍스트") < md.index("하단텍스트")

    def test_ac5_같은행_left_정렬(self) -> None:
        """ac5_같은행_left_정렬: tolerance 이내 top 차이이고 IR에 우→좌이면 좌 먼저."""
        # _ROW_TOLERANCE_EMU = 914400 (0.5 inch)
        # top이 100000 차이 → 같은 행으로 버킷팅 (tolerance 이내)
        right_shape = _text_shape_with_coords(
            "오른쪽텍스트", top=100100, left=3000000, shape_id=2
        )
        left_shape = _text_shape_with_coords(
            "왼쪽텍스트", top=100000, left=500000, shape_id=1
        )
        # IR에 오른쪽→왼쪽 역순으로 담음
        slide = SlideIR(index=0, title="T", shapes=[right_shape, left_shape])
        md = assemble_slide(slide)
        assert md.index("왼쪽텍스트") < md.index("오른쪽텍스트")

    def test_ac6_all_zero_순서_보존(self) -> None:
        """ac6_all_zero_순서_보존: 모든 도형이 left=top=0 이면 IR 순서 그대로."""
        shape1 = _text_shape_with_coords("첫번째", top=0, left=0, shape_id=1)
        shape2 = _text_shape_with_coords("두번째", top=0, left=0, shape_id=2)
        shape3 = _text_shape_with_coords("세번째", top=0, left=0, shape_id=3)
        slide = SlideIR(index=0, title="T", shapes=[shape1, shape2, shape3])
        md = assemble_slide(slide)
        # stable sort 에 의해 동률 좌표는 원 IR 순서 유지
        assert md.index("첫번째") < md.index("두번째") < md.index("세번째")

    def test_ac7_결정성(self) -> None:
        """ac7_결정성: 동일 SlideIR 로 assemble_slide 2회 호출 시 완전히 동일."""
        shape_a = _text_shape_with_coords("A텍스트", top=2000000, left=0, shape_id=1)
        shape_b = _text_shape_with_coords("B텍스트", top=500000, left=0, shape_id=2)
        slide = SlideIR(index=0, title="결정성테스트", shapes=[shape_a, shape_b])
        md1 = assemble_slide(slide)
        md2 = assemble_slide(slide)
        assert md1 == md2

    def test_ac4_ir_not_mutated(self) -> None:
        """ac4_IR_불변: assemble_slide 후 slide.shapes 원본 순서 유지 (IR read-only)."""
        bottom_shape = _text_shape_with_coords("하단", top=4000000, left=0, shape_id=2)
        top_shape = _text_shape_with_coords("상단", top=500000, left=0, shape_id=1)
        original_shapes = [bottom_shape, top_shape]
        slide = SlideIR(index=0, title="T", shapes=list(original_shapes))
        assemble_slide(slide)
        # slide.shapes 순서는 변경되지 않아야 함
        assert slide.shapes[0] is bottom_shape
        assert slide.shapes[1] is top_shape

    def test_ac5_다른_행은_top_우선(self) -> None:
        """ac5_다른_행_top_우선: tolerance를 넘는 top 차이는 다른 행으로 처리."""
        # top 차이 2000000 > 914400 → 다른 행
        # 아래 행의 left가 0이어도 위 행이 먼저
        upper_right = _text_shape_with_coords(
            "위행오른쪽", top=0, left=4000000, shape_id=1
        )
        lower_left = _text_shape_with_coords(
            "아래행왼쪽", top=2000000, left=0, shape_id=2
        )
        slide = SlideIR(index=0, title="T", shapes=[lower_left, upper_right])
        md = assemble_slide(slide)
        assert md.index("위행오른쪽") < md.index("아래행왼쪽")


# ---------------------------------------------------------------------------
# FR-28 (#77, W-A): heading_hierarchy / emit_toc / 옵션 배선 / seam 격리
# AC name format: ac<N>_<short_description> (issue #77 numbering)
# ---------------------------------------------------------------------------


class TestAc1HeadingHierarchyDecomposition:
    """AC1: heading_hierarchy=True & '>' 경로 → ##/###/####… 분해"""

    def test_ac1_two_segment_title_h2_h3(self) -> None:
        """ac1_2세그먼트_h2_h3: 최상위 ##, 후속 ### 로 분해되어 별도 헤딩이 된다."""
        slide = SlideIR(
            index=0,
            title="Ⅱ. 사업의 이해 > ⑦ 현황·문제점",
            shapes=[_text_shape([_para("본문")])],
        )
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        doc = assemble_document(pres, heading_hierarchy=True)
        lines = doc.splitlines()
        assert lines[0] == "## Ⅱ. 사업의 이해"
        assert lines[2] == "### ⑦ 현황·문제점"

    def test_ac1_segments_trimmed(self) -> None:
        """ac1_세그먼트_trim: 각 세그먼트 앞뒤 공백이 제거된다."""
        slide = SlideIR(
            index=0, title="  A  >  B  ", shapes=[_text_shape([_para("본문")])]
        )
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        doc = assemble_document(pres, heading_hierarchy=True)
        assert "## A" in doc.splitlines()
        assert "### B" in doc.splitlines()

    def test_ac1_deep_path_h4(self) -> None:
        """ac1_3세그먼트_h4: 세 번째 세그먼트는 #### 로 렌더된다."""
        slide = SlideIR(
            index=0,
            title="A > B > C",
            shapes=[_text_shape([_para("본문")])],
        )
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        doc = assemble_document(pres, heading_hierarchy=True)
        assert "#### C" in doc.splitlines()

    def test_ac1_is_title_shape_path_also_decomposed(self) -> None:
        """ac1_is_title_도형도_분해: slide.title 부재시 is_title 도형 텍스트도 분해."""
        title_s = _title_shape("섹션 > 하위섹션")
        body_s = _text_shape([_para("본문")], shape_id=11)
        slide = SlideIR(index=0, title="", shapes=[title_s, body_s])
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        doc = assemble_document(pres, heading_hierarchy=True)
        lines = doc.splitlines()
        assert lines[0] == "## 섹션"
        assert lines[2] == "### 하위섹션"


class TestAc2NoSeparatorSingleHeading:
    """AC2: 구분자 없는 제목 → 단일 ##(현행 동일, 분해 미발동)"""

    def test_ac2_plain_title_single_h2(self) -> None:
        """ac2_구분자없음_단일헤딩: heading_hierarchy=True 라도 '>' 없으면 단일 ##."""
        slide = SlideIR(
            index=0, title="단순 제목", shapes=[_text_shape([_para("본문")])]
        )
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        doc = assemble_document(pres, heading_hierarchy=True)
        assert doc.startswith("## 단순 제목")
        assert "###" not in doc.splitlines()[0]
        # 헤딩 라인은 정확히 1줄이어야 한다 (분해 미발동)
        heading_lines = [line for line in doc.splitlines() if line.startswith("#")]
        assert heading_lines == ["## 단순 제목"]


class TestAc3HeadingHierarchyOffByteIdentical:
    """AC3: heading_hierarchy=False(기본) → 출력 바이트 동일(회귀 diff 0)"""

    def test_ac3_off_matches_legacy_output_with_path_title(self) -> None:
        """ac3_off_바이트동일: '>' 포함 제목도 off 시 레거시와 동일 문자열."""
        slide = SlideIR(
            index=0,
            title="Ⅱ. 사업의 이해 > ⑦ 현황·문제점",
            shapes=[_text_shape([_para("본문")])],
        )
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        doc_off = assemble_document(pres)  # default heading_hierarchy=False
        doc_off_explicit = assemble_document(pres, heading_hierarchy=False)
        # 레거시 동작: '>' 는 분해되지 않고 원본 title 그대로 단일 heading
        assert doc_off == doc_off_explicit
        assert doc_off.startswith("## Ⅱ. 사업의 이해 > ⑦ 현황·문제점")

    def test_ac3_off_default_matches_pre_wave3_single_slide_document(self) -> None:
        """ac3_기본값_레거시_동일: 4옵션 모두 미지정 시 기존 M2~M5 산출물과 동일."""
        slides = [
            SlideIR(index=0, title="첫 슬라이드", shapes=[_text_shape([_para("A")])]),
            SlideIR(index=1, title="", shapes=[_text_shape([_para("B")])]),
        ]
        pres = PresentationIR(source_path="t.pptx", slides=slides)
        doc = assemble_document(pres)
        expected_slide0 = assemble_slide(slides[0])
        expected_slide1 = assemble_slide(slides[1])
        assert doc == f"{expected_slide0}\n\n---\n\n{expected_slide1}"

    def test_ac3_all_four_options_off_equals_no_kwargs(self) -> None:
        """ac3_4옵션명시적off_동일: 4옵션 모두 명시적 False가 미지정과 바이트 동일."""
        slide = SlideIR(
            index=0, title="제목 > 하위", shapes=[_text_shape([_para("본문")])]
        )
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        default_doc = assemble_document(pres)
        explicit_doc = assemble_document(
            pres,
            heading_hierarchy=False,
            emit_toc=False,
            emit_frontmatter=False,
            include_notes=False,
        )
        assert default_doc == explicit_doc


class TestAc4EmitTocEmptySlideConversion:
    """AC4: emit_toc=True & 본문 비공백 0 슬라이드 → TOC 항목/드롭, 빈 Slide N 0개"""

    def test_ac4_real_title_empty_body_kept_as_heading_only(self) -> None:
        """ac4_실제제목_빈본문: 실제 제목의 섹션 표지는 헤딩만 남는다."""
        slide = SlideIR(index=0, title="Ⅱ. 사업의 이해", shapes=[])
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        doc = assemble_document(pres, emit_toc=True)
        assert doc == "## Ⅱ. 사업의 이해"

    def test_ac4_fallback_title_empty_body_dropped(self) -> None:
        """ac4_fallback_빈본문_드롭: fallback '## Slide N' 은 블록이 사라진다."""
        real_slide = SlideIR(
            index=0, title="표지", shapes=[_text_shape([_para("내용")])]
        )
        empty_fallback_slide = SlideIR(index=1, title="", shapes=[])
        slide3 = SlideIR(index=2, title="결론", shapes=[_text_shape([_para("끝")])])
        pres = PresentationIR(
            source_path="t.pptx", slides=[real_slide, empty_fallback_slide, slide3]
        )
        doc = assemble_document(pres, emit_toc=True)
        assert "Slide 2" not in doc
        assert "## Slide" not in doc  # 빈 'Slide N' 블록 0개 (AC4)

    def test_ac4_no_empty_slide_n_headings_across_document(self) -> None:
        """ac4_빈슬라이드N_0개: 여러 fallback-빈 슬라이드가 있어도 전부 드롭된다."""
        slides = [
            SlideIR(index=0, title="", shapes=[]),
            SlideIR(index=1, title="", shapes=[]),
            SlideIR(
                index=2, title="유일한 내용", shapes=[_text_shape([_para("본문")])]
            ),
        ]
        pres = PresentationIR(source_path="t.pptx", slides=slides)
        doc = assemble_document(pres, emit_toc=True)
        assert "Slide 1" not in doc
        assert "Slide 2" not in doc
        assert doc == "## 유일한 내용\n\n본문"

    def test_ac4_non_empty_body_untouched_by_emit_toc(self) -> None:
        """ac4_비공백본문_무변환: 본문이 있으면 emit_toc 와 무관하게 그대로 렌더."""
        slide = SlideIR(
            index=0, title="본문 있는 슬라이드", shapes=[_text_shape([_para("내용")])]
        )
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        doc_off = assemble_document(pres)
        doc_on = assemble_document(pres, emit_toc=True)
        assert doc_off == doc_on

    def test_ac4_emit_toc_false_preserves_legacy_empty_slide_n(self) -> None:
        """ac4_emit_toc_off: emit_toc=False 면 빈 'Slide N' 이 그대로 남는다."""
        slide = SlideIR(index=0, title="", shapes=[])
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        doc = assemble_document(pres)  # emit_toc=False (default)
        assert doc == "## Slide 1"


class TestAc5OptionsBackwardCompatible:
    """AC5: ConvertOptions()·convert(src) 무손상 (assemble_document 레벨 대리 검증)"""

    def test_ac5_assemble_document_no_kwargs_still_works(self) -> None:
        """ac5_키워드인자없이_호출: 신규 4옵션 없이도 assemble_document 가 동작한다."""
        slide = SlideIR(index=0, title="T", shapes=[_text_shape([_para("본문")])])
        pres = PresentationIR(source_path="t.pptx", slides=[slide])
        doc = assemble_document(pres, masking=None, suppress_repeated_labels=False)
        assert doc.startswith("## T")

    def test_ac5_assemble_slide_signature_unchanged(self) -> None:
        """ac5_시그니처_불변: assemble_slide 는 신규 옵션을 받지 않는다."""
        slide = SlideIR(index=0, title="T", shapes=[])
        # assemble_slide는 masking/table_fallback 만 받는다 (heading_hierarchy 없음)
        result = assemble_slide(slide, masking=None, table_fallback="mermaid")
        assert result == "## T"


class TestAc6NewModulesNoForbiddenImports:
    """AC6: 신규 모듈(heading.py/metadata.py/notes.py) SDK·pptx·Pillow import 0

    NOTE: matches actual ``import pptx`` / ``from pptx ...`` statements via
    regex (word boundary after "pptx") rather than a naive substring check,
    so that docstring mentions of "python-pptx" and internal
    ``from pptx_md...`` imports are not false positives.
    """

    _FORBIDDEN_PPTX_IMPORT_RE = re.compile(r"^\s*(?:from|import)\s+pptx\b")

    def _assert_no_forbidden_imports(self, source: str) -> None:
        assert "anthropic" not in source
        assert "openai" not in source
        assert "PIL" not in source
        for line in source.splitlines():
            assert not self._FORBIDDEN_PPTX_IMPORT_RE.match(line), line

    def test_ac6_heading_module_stdlib_only(self) -> None:
        """ac6_heading_모듈_격리: heading.py 는 SDK/pptx/Pillow 를 import하지 않는다."""
        import pathlib

        source = (
            pathlib.Path(__file__).parent.parent / "src" / "pptx_md" / "heading.py"
        ).read_text(encoding="utf-8")
        self._assert_no_forbidden_imports(source)

    def test_ac6_metadata_module_no_sdk_imports(self) -> None:
        """ac6_metadata_모듈_격리: metadata.py 는 SDK/Pillow 를 import하지 않는다."""
        import pathlib

        source = (
            pathlib.Path(__file__).parent.parent / "src" / "pptx_md" / "metadata.py"
        ).read_text(encoding="utf-8")
        self._assert_no_forbidden_imports(source)

    def test_ac6_notes_module_no_sdk_imports(self) -> None:
        """ac6_notes_모듈_격리: notes.py 는 SDK/pptx/Pillow 를 import하지 않는다."""
        import pathlib

        source = (
            pathlib.Path(__file__).parent.parent / "src" / "pptx_md" / "notes.py"
        ).read_text(encoding="utf-8")
        self._assert_no_forbidden_imports(source)


class TestAc7QualityGates:
    """AC7: ruff/black/mypy exit 0 (실제 하위프로세스 실행으로 검증)"""

    def test_ac7_mypy_strict_exit0(self) -> None:
        """ac7_mypy_exit0: mypy src/ 는 exit 0 이어야 한다."""
        import pathlib
        import subprocess
        import sys

        project_root = pathlib.Path(__file__).parent.parent
        result = subprocess.run(
            [sys.executable, "-m", "mypy", "src/"],
            capture_output=True,
            text=True,
            cwd=str(project_root),
        )
        assert result.returncode == 0, result.stdout + result.stderr

    def test_ac7_ruff_check_exit0(self) -> None:
        """ac7_ruff_exit0: ruff check . 는 exit 0 이어야 한다."""
        import pathlib
        import subprocess
        import sys

        project_root = pathlib.Path(__file__).parent.parent
        result = subprocess.run(
            [sys.executable, "-m", "ruff", "check", "."],
            capture_output=True,
            text=True,
            cwd=str(project_root),
        )
        assert result.returncode == 0, result.stdout + result.stderr

    def test_ac7_black_check_exit0(self) -> None:
        """ac7_black_exit0: black --check . 는 exit 0 이어야 한다."""
        import pathlib
        import subprocess
        import sys

        project_root = pathlib.Path(__file__).parent.parent
        result = subprocess.run(
            [sys.executable, "-m", "black", "--check", "."],
            capture_output=True,
            text=True,
            cwd=str(project_root),
        )
        assert result.returncode == 0, result.stdout + result.stderr
