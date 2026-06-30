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
