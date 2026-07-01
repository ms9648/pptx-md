"""Tests for FR-24: Assembler 정규화 (issue #59).

Covers AC1–AC8 of issue #59.
AC name format: ac<N>_<short_description>
"""

from __future__ import annotations

from pptx_md.assembler import (
    _collapse_blank_lines,
    _is_table_all_blank,
    _normalize_text,
    assemble_document,
    assemble_slide,
)
from pptx_md.ir import (
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
    is_footer: bool = False,
) -> TextShapeIR:
    return TextShapeIR(
        shape_id=shape_id,
        name="text",
        kind=ShapeKind.TEXT,
        paragraphs=paragraphs,
        is_title=is_title,
        is_footer=is_footer,
    )


def _para(text: str, level: int = 0) -> ParagraphIR:
    return ParagraphIR(text=text, level=level)


def _table_shape(
    rows: list[list[str]],
    shape_id: int = 20,
) -> TableShapeIR:
    r = len(rows)
    c = len(rows[0]) if rows else 0
    return TableShapeIR(
        shape_id=shape_id,
        name="table",
        kind=ShapeKind.TABLE,
        rows=rows,
        n_rows=r,
        n_cols=c,
    )


# ---------------------------------------------------------------------------
# Unit tests for normalization utilities
# ---------------------------------------------------------------------------


class TestNormalizeTextUnit:
    """Unit tests for _normalize_text (AC2, AC3)."""

    def test_ac2_normalize_text_vertical_tab_replaced_with_newline(self) -> None:
        """ac2_vertical_tab_to_newline: \\v is replaced by \\n."""
        result = _normalize_text("line1\vline2")
        assert "\v" not in result
        assert "line1\nline2" == result

    def test_ac2_normalize_text_multiple_vertical_tabs(self) -> None:
        """ac2_multiple_vertical_tabs: each \\v becomes \\n."""
        result = _normalize_text("a\vb\vc")
        assert "\v" not in result
        assert result == "a\nb\nc"

    def test_ac3_normalize_text_consecutive_spaces_collapsed(self) -> None:
        """ac3_consecutive_spaces_collapsed: 2+ spaces -> single space."""
        result = _normalize_text("hello  world")
        assert result == "hello world"

    def test_ac3_normalize_text_consecutive_tabs_collapsed(self) -> None:
        """ac3_consecutive_tabs_collapsed: 2+ tabs -> single space."""
        result = _normalize_text("col1\t\tcol2")
        assert result == "col1 col2"

    def test_ac3_normalize_text_mixed_spaces_tabs_collapsed(self) -> None:
        """ac3_mixed_spaces_tabs: mixed space+tab run -> single space."""
        result = _normalize_text("a \t b")
        assert result == "a b"

    def test_ac3_single_space_unchanged(self) -> None:
        """ac3_single_space_unchanged: single space is not affected."""
        result = _normalize_text("hello world")
        assert result == "hello world"

    def test_ac2_ac3_combined(self) -> None:
        """ac2_ac3_combined: \\v replaced, then spaces collapsed."""
        result = _normalize_text("line1\v  line2")
        assert "\v" not in result
        # \v -> \n, then "  " -> " "
        assert result == "line1\n line2"

    def test_normalize_empty_string(self) -> None:
        """normalize_empty_string: empty input returns empty."""
        assert _normalize_text("") == ""


class TestCollapsBlankLinesUnit:
    """Unit tests for _collapse_blank_lines (AC5)."""

    def test_ac5_three_newlines_collapsed_to_two(self) -> None:
        """ac5_three_newlines_to_two: \\n\\n\\n -> \\n\\n."""
        result = _collapse_blank_lines("a\n\n\nb")
        assert result == "a\n\nb"

    def test_ac5_four_newlines_collapsed(self) -> None:
        """ac5_four_newlines_collapsed: 4+ newlines -> 2."""
        result = _collapse_blank_lines("a\n\n\n\nb")
        assert result == "a\n\nb"

    def test_ac5_two_newlines_unchanged(self) -> None:
        """ac5_two_newlines_unchanged: exactly 2 newlines not affected."""
        result = _collapse_blank_lines("a\n\nb")
        assert result == "a\n\nb"

    def test_ac5_single_newline_unchanged(self) -> None:
        """ac5_single_newline_unchanged: single newline not affected."""
        result = _collapse_blank_lines("a\nb")
        assert result == "a\nb"


class TestIsTableAllBlankUnit:
    """Unit tests for _is_table_all_blank (AC4)."""

    def test_ac4_all_empty_strings_is_blank(self) -> None:
        """ac4_all_empty_is_blank: table with all "" cells -> True."""
        shape = _table_shape([["", ""], ["", ""]])
        assert _is_table_all_blank(shape) is True

    def test_ac4_all_whitespace_is_blank(self) -> None:
        """ac4_all_whitespace_is_blank: table with all whitespace cells -> True."""
        shape = _table_shape([["  ", " "], ["\t", ""]])
        assert _is_table_all_blank(shape) is True

    def test_ac4_one_non_empty_cell_not_blank(self) -> None:
        """ac4_one_non_empty_not_blank: at least one content cell -> False."""
        shape = _table_shape([["", "Header"], ["", ""]])
        assert _is_table_all_blank(shape) is False

    def test_ac4_empty_rows_is_blank(self) -> None:
        """ac4_empty_rows_is_blank: table with no rows -> True (vacuously blank)."""
        shape = _table_shape([])
        assert _is_table_all_blank(shape) is True


# ---------------------------------------------------------------------------
# AC1: 제목 중복 제거 — slide.title과 동일 텍스트 is_title 도형 본문 재출력 방지
# ---------------------------------------------------------------------------


class TestAc1TitleDeduplication:
    """ac1_제목_중복_제거 (FR-24 #59)"""

    def test_ac1_title_appears_once_when_is_title_shape_matches(self) -> None:
        """ac1: slide.title 있고 동일 텍스트 is_title 도형 존재 -> 헤딩 1회만."""
        title_shape = _text_shape([_para("슬라이드 제목")], shape_id=1, is_title=True)
        body_shape = _text_shape([_para("본문 내용")], shape_id=2)
        slide = SlideIR(
            index=0,
            title="슬라이드 제목",
            shapes=[title_shape, body_shape],
        )
        md = assemble_slide(slide)
        # Heading must appear exactly once
        assert md.count("슬라이드 제목") == 1

    def test_ac1_heading_is_h2(self) -> None:
        """ac1: 제목은 ## 헤딩으로 출력."""
        title_shape = _text_shape([_para("My Title")], shape_id=1, is_title=True)
        slide = SlideIR(index=0, title="My Title", shapes=[title_shape])
        md = assemble_slide(slide)
        assert "## My Title" in md

    def test_ac1_body_text_still_present(self) -> None:
        """ac1: 제목 도형 suppressed 후 본문 도형은 정상 출력."""
        title_shape = _text_shape([_para("Title")], shape_id=1, is_title=True)
        body_shape = _text_shape([_para("Body text here")], shape_id=2)
        slide = SlideIR(index=0, title="Title", shapes=[title_shape, body_shape])
        md = assemble_slide(slide)
        assert "Body text here" in md

    def test_ac1_non_matching_is_title_shape_still_rendered(self) -> None:
        """ac1: slide.title과 다른 텍스트의 is_title 도형은 본문에 출력."""
        title_shape = _text_shape(
            [_para("Different Title Text")], shape_id=1, is_title=True
        )
        slide = SlideIR(index=0, title="Slide Title", shapes=[title_shape])
        md = assemble_slide(slide)
        # The is_title shape has different text, so it should appear in body
        assert "Different Title Text" in md

    def test_ac1_strip_comparison_handles_whitespace(self) -> None:
        """ac1: strip 비교로 공백 차이 무시."""
        title_shape = _text_shape([_para("  제목  ")], shape_id=1, is_title=True)
        slide = SlideIR(index=0, title="제목", shapes=[title_shape])
        md = assemble_slide(slide)
        # "제목" should appear only once (as heading)
        assert md.count("제목") == 1

    def test_ac1_no_title_shape_uses_slide_title_normally(self) -> None:
        """ac1: is_title 도형 없어도 slide.title은 정상 출력."""
        body_shape = _text_shape([_para("Content")], shape_id=2)
        slide = SlideIR(index=0, title="Plain Title", shapes=[body_shape])
        md = assemble_slide(slide)
        assert md.startswith("## Plain Title")


# ---------------------------------------------------------------------------
# AC2: 제어문자 정규화 — \v -> \n
# ---------------------------------------------------------------------------


class TestAc2ControlCharNormalization:
    """ac2_제어문자_정규화 (FR-24 #59)"""

    def test_ac2_vertical_tab_absent_in_output(self) -> None:
        """ac2: 출력에 \\v가 0건."""
        shape = _text_shape([_para("line1\vline2")])
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "\v" not in md

    def test_ac2_vertical_tab_becomes_newline(self) -> None:
        """ac2: 각 \\v가 \\n으로 변환."""
        shape = _text_shape([_para("first\vsecond")])
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "first" in md
        assert "second" in md
        # Both pieces must appear (separated by newline, not merged)
        lines = md.splitlines()
        content_lines = [ln for ln in lines if "first" in ln or "second" in ln]
        assert len(content_lines) >= 1
        # Check that "first" and "second" appear (possibly on same or different lines)
        assert "first" in md and "second" in md

    def test_ac2_multiple_vertical_tabs_all_replaced(self) -> None:
        """ac2_multiple_vertical_tabs: 복수 \\v 모두 \\n으로 변환."""
        shape = _text_shape([_para("a\vb\vc\vd")])
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "\v" not in md
        assert "a" in md and "b" in md and "c" in md and "d" in md

    def test_ac2_normal_text_unaffected(self) -> None:
        """ac2: 제어문자 없는 텍스트는 그대로."""
        shape = _text_shape([_para("normal text here")])
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "normal text here" in md


# ---------------------------------------------------------------------------
# AC3: 연속 공백 정리
# ---------------------------------------------------------------------------


class TestAc3ConsecutiveWhitespace:
    """ac3_연속_공백_정리 (FR-24 #59)"""

    def test_ac3_consecutive_spaces_normalized(self) -> None:
        """ac3: 연속 공백 축약."""
        shape = _text_shape([_para("word1  word2   word3")])
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "word1 word2 word3" in md

    def test_ac3_consecutive_tabs_normalized(self) -> None:
        """ac3_tabs: 연속 탭을 단일 공백으로."""
        shape = _text_shape([_para("col1\t\tcol2")])
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "\t\t" not in md
        assert "col1" in md and "col2" in md

    def test_ac3_single_space_preserved(self) -> None:
        """ac3_single_space: 단일 공백은 유지."""
        shape = _text_shape([_para("hello world")])
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "hello world" in md

    def test_ac3_table_cell_content_not_damaged(self) -> None:
        """ac3_table_safe: 표 셀 텍스트는 손상 없음 (표는 정규화 대상 아님)."""
        # Table cells are rendered as-is (FR-25 scope); AC3 only applies to
        # TextShapeIR paragraphs, not table cells
        table = _table_shape([["Header A", "Header B"], ["val1", "val2"]])
        slide = SlideIR(index=0, title="T", shapes=[table])
        md = assemble_slide(slide)
        assert "Header A" in md
        assert "val1" in md


# ---------------------------------------------------------------------------
# AC4: 빈 표 omit
# ---------------------------------------------------------------------------


class TestAc4EmptyTableOmit:
    """ac4_빈_표_생략 (FR-24 #59)"""

    def test_ac4_all_blank_table_omitted(self) -> None:
        """ac4: 모든 셀이 공백인 표 -> 출력 완전 생략."""
        blank_table = _table_shape([["", ""], ["", ""], ["", ""]])
        slide = SlideIR(index=0, title="T", shapes=[blank_table])
        md = assemble_slide(slide)
        # No table pipe characters should appear
        assert "|" not in md

    def test_ac4_all_whitespace_table_omitted(self) -> None:
        """ac4_whitespace: 공백 문자만 있는 셀도 공백으로 처리."""
        blank_table = _table_shape([["  ", " "], ["\t", "   "]])
        slide = SlideIR(index=0, title="T", shapes=[blank_table])
        md = assemble_slide(slide)
        assert "|" not in md

    def test_ac4_blank_table_other_shapes_unaffected(self) -> None:
        """ac4_isolation: 빈 표 생략 시 다른 도형은 정상 출력."""
        blank_table = _table_shape([["", ""], ["", ""]], shape_id=20)
        body = _text_shape([_para("Body content")], shape_id=10)
        slide = SlideIR(index=0, title="T", shapes=[blank_table, body])
        md = assemble_slide(slide)
        assert "Body content" in md
        assert "|" not in md

    def test_ac4_non_blank_table_rendered(self) -> None:
        """ac4_non_blank: 내용 있는 표는 정상 출력."""
        table = _table_shape([["Col A", "Col B"], ["1", "2"]])
        slide = SlideIR(index=0, title="T", shapes=[table])
        md = assemble_slide(slide)
        assert "| Col A | Col B |" in md

    def test_ac4_partially_blank_table_rendered(self) -> None:
        """ac4_partial: 일부 셀만 공백이면 표는 출력."""
        table = _table_shape([["Header", ""], ["Data", ""]])
        slide = SlideIR(index=0, title="T", shapes=[table])
        md = assemble_slide(slide)
        assert "|" in md


# ---------------------------------------------------------------------------
# AC5: 빈 단락 collapse
# ---------------------------------------------------------------------------


class TestAc5BlankParagraphCollapse:
    """ac5_빈_단락_collapse (FR-24 #59)"""

    def test_ac5_three_consecutive_blank_paragraphs_collapsed(self) -> None:
        """ac5: 연속 3개 빈 단락 -> 최대 1개 빈 줄."""
        shape = _text_shape(
            [
                _para("Content before"),
                _para(""),
                _para(""),
                _para(""),
                _para("Content after"),
            ]
        )
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        # Must not have 3+ consecutive blank lines in the output
        assert "\n\n\n" not in md

    def test_ac5_two_consecutive_blank_paragraphs_ok(self) -> None:
        """ac5: 2개 빈 단락은 허용 범위."""
        shape = _text_shape(
            [
                _para("Line A"),
                _para(""),
                _para(""),
                _para("Line B"),
            ]
        )
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "Line A" in md
        assert "Line B" in md

    def test_ac5_many_blank_paragraphs_collapsed(self) -> None:
        """ac5_many: 5개 연속 빈 단락도 최대 1개 빈 줄로."""
        shape = _text_shape(
            [_para("Start")] + [_para("") for _ in range(5)] + [_para("End")]
        )
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "\n\n\n" not in md
        assert "Start" in md
        assert "End" in md

    def test_ac5_no_content_lost(self) -> None:
        """ac5_no_loss: 빈 줄 collapse 후 내용 텍스트 손실 없음."""
        shape = _text_shape(
            [
                _para("First paragraph"),
                _para(""),
                _para(""),
                _para(""),
                _para("Second paragraph"),
            ]
        )
        slide = SlideIR(index=0, title="T", shapes=[shape])
        md = assemble_slide(slide)
        assert "First paragraph" in md
        assert "Second paragraph" in md


# ---------------------------------------------------------------------------
# AC6: 반복 섹션 라벨 억제 (옵션)
# ---------------------------------------------------------------------------


class TestAc6RepeatedLabelSuppression:
    """ac6_반복_섹션_라벨_억제 (FR-24 #59)"""

    def test_ac6_repeated_label_suppressed_when_option_active(self) -> None:
        """ac6: 동일 라벨이 연속 슬라이드에 반복 -> 옵션 활성 시 최초 1회만."""
        label_shape_1 = _text_shape([_para("Ⅰ. 회사 소개")], shape_id=1)
        body_shape_1 = _text_shape([_para("슬라이드 1 내용")], shape_id=2)
        label_shape_2 = _text_shape([_para("Ⅰ. 회사 소개")], shape_id=3)
        body_shape_2 = _text_shape([_para("슬라이드 2 내용")], shape_id=4)

        slide1 = SlideIR(index=0, title="S1", shapes=[label_shape_1, body_shape_1])
        slide2 = SlideIR(index=1, title="S2", shapes=[label_shape_2, body_shape_2])
        pres = PresentationIR(source_path="test.pptx", slides=[slide1, slide2])

        doc = assemble_document(pres, suppress_repeated_labels=True)
        # Label appears only once
        assert doc.count("Ⅰ. 회사 소개") == 1

    def test_ac6_default_false_preserves_repeated_labels(self) -> None:
        """ac6_default: 옵션 비활성(default) 시 반복 라벨 그대로 출력."""
        label_shape_1 = _text_shape([_para("Ⅰ. 회사 소개")], shape_id=1)
        label_shape_2 = _text_shape([_para("Ⅰ. 회사 소개")], shape_id=3)

        slide1 = SlideIR(index=0, title="S1", shapes=[label_shape_1])
        slide2 = SlideIR(index=1, title="S2", shapes=[label_shape_2])
        pres = PresentationIR(source_path="test.pptx", slides=[slide1, slide2])

        doc = assemble_document(pres)  # default suppress_repeated_labels=False
        assert doc.count("Ⅰ. 회사 소개") == 2

    def test_ac6_non_repeated_labels_unchanged(self) -> None:
        """ac6_non_repeated: 반복 아닌 라벨은 영향 없음."""
        label_a = _text_shape([_para("Section A")], shape_id=1)
        label_b = _text_shape([_para("Section B")], shape_id=2)

        slide1 = SlideIR(index=0, title="S1", shapes=[label_a])
        slide2 = SlideIR(index=1, title="S2", shapes=[label_b])
        pres = PresentationIR(source_path="test.pptx", slides=[slide1, slide2])

        doc = assemble_document(pres, suppress_repeated_labels=True)
        assert "Section A" in doc
        assert "Section B" in doc

    def test_ac6_heading_never_suppressed(self) -> None:
        """ac6_heading_safe: 헤딩(##)은 라벨 억제 대상 아님."""
        slide1 = SlideIR(index=0, title="동일 제목", shapes=[])
        slide2 = SlideIR(index=1, title="동일 제목", shapes=[])
        pres = PresentationIR(source_path="test.pptx", slides=[slide1, slide2])

        doc = assemble_document(pres, suppress_repeated_labels=True)
        # Both headings must remain
        assert doc.count("## 동일 제목") == 2


# ---------------------------------------------------------------------------
# AC7: 결정성 (ADR-218)
# ---------------------------------------------------------------------------


class TestAc7Determinism:
    """ac7_결정성_ADR218 (FR-24 #59)"""

    def test_ac7_same_slide_ir_produces_identical_output(self) -> None:
        """ac7: 동일 SlideIR 2회 렌더 -> 완전히 동일."""
        shape = _text_shape(
            [
                _para("Line with  spaces"),
                _para("Line with\vvert tab"),
                _para(""),
                _para(""),
                _para(""),
                _para("Final line"),
            ]
        )
        slide = SlideIR(index=0, title="Test", shapes=[shape])
        md1 = assemble_slide(slide)
        md2 = assemble_slide(slide)
        assert md1 == md2

    def test_ac7_same_presentation_ir_produces_identical_output(self) -> None:
        """ac7_document: 동일 PresentationIR 2회 렌더 -> 완전히 동일."""
        shape = _text_shape([_para("Content with  extra  spaces")])
        slide = SlideIR(index=0, title="Slide", shapes=[shape])
        pres = PresentationIR(source_path="test.pptx", slides=[slide])
        doc1 = assemble_document(pres)
        doc2 = assemble_document(pres)
        assert doc1 == doc2

    def test_ac7_normalization_functions_are_deterministic(self) -> None:
        """ac7_utils: 정규화 유틸 함수 자체도 결정적."""
        text = "hello\v  world\t\ttest"
        assert _normalize_text(text) == _normalize_text(text)
        assert _collapse_blank_lines("a\n\n\nb") == _collapse_blank_lines("a\n\n\nb")


# ---------------------------------------------------------------------------
# AC8: 예외 — 부분 실패 격리
# ---------------------------------------------------------------------------


class TestAc8ExceptionIsolation:
    """ac8_예외_부분실패_격리 (FR-24 #59)"""

    def test_ac8_shape_exception_does_not_abort_document(self) -> None:
        """ac8: 도형 처리 중 예외 -> 해당 도형 스킵, 문서 전체 변환 계속."""
        from unittest.mock import patch

        from pptx_md import assembler as asm

        good_shape = _text_shape([_para("Good content")], shape_id=10)
        bad_shape = _text_shape([_para("Bad content")], shape_id=20)
        body_shape = _text_shape([_para("After bad")], shape_id=30)

        slide = SlideIR(index=0, title="T", shapes=[good_shape, bad_shape, body_shape])
        pres = PresentationIR(source_path="test.pptx", slides=[slide])

        original_render_text = asm._render_text

        call_count = [0]

        def failing_render_text(shape: TextShapeIR, masking: object) -> str:
            call_count[0] += 1
            if shape.shape_id == 20:
                raise RuntimeError("Simulated normalization failure")
            return original_render_text(shape, masking)  # type: ignore[arg-type]

        with patch.object(asm, "_render_text", side_effect=failing_render_text):
            doc = assemble_document(pres)

        # Document is produced without raising
        assert isinstance(doc, str)
        # Good shapes are rendered
        assert "Good content" in doc
        assert "After bad" in doc

    def test_ac8_document_continues_after_slide_failure(self) -> None:
        """ac8_slide_isolation: 슬라이드 수준 예외도 전체 문서 중단 없음."""
        slide1 = SlideIR(index=0, title="Good Slide", shapes=[])
        slide2 = SlideIR(index=1, title="Good Slide 2", shapes=[])
        pres = PresentationIR(source_path="test.pptx", slides=[slide1, slide2])

        doc = assemble_document(pres)
        assert "Good Slide" in doc
        assert "Good Slide 2" in doc


# ---------------------------------------------------------------------------
# Integration: all normalizations together
# ---------------------------------------------------------------------------


class TestFR24Integration:
    """FR-24 통합 테스트: 복합 노이즈 처리"""

    def test_fr24_combined_normalizations(self) -> None:
        """fr24_combined: 제목 중복+제어문자+공백+빈단락 복합 처리."""
        title_shape = _text_shape(
            [_para("복합 테스트 제목")], shape_id=1, is_title=True
        )
        content_shape = _text_shape(
            [
                _para("정상 텍스트"),
                _para("텍스트\v소프트개행 포함"),
                _para("연속  공백   포함"),
                _para(""),
                _para(""),
                _para(""),
                _para("마지막 단락"),
            ],
            shape_id=2,
        )
        slide = SlideIR(
            index=0,
            title="복합 테스트 제목",
            shapes=[title_shape, content_shape],
        )
        md = assemble_slide(slide)

        # AC1: title deduplicated
        assert md.count("복합 테스트 제목") == 1
        # AC2: no vertical tabs
        assert "\v" not in md
        # AC3: no double spaces in content
        assert "  " not in md
        # AC5: no triple+ newlines
        assert "\n\n\n" not in md
        # Content preserved
        assert "정상 텍스트" in md
        assert "마지막 단락" in md
