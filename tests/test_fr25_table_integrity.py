"""Tests for FR-25: 표 무결성 (병합 셀·깨진 파이프·Mermaid fallback 프리픽스).

Covers AC1–AC8 of issue #60.  All tests use synthetic IR fixtures
(no python-pptx dependency — ADR-206/211).

AC naming convention: ac<N>_<desc>
"""

from __future__ import annotations

from pptx_md.assembler import assemble_slide
from pptx_md.ir import ShapeKind, SlideIR, TableShapeIR
from pptx_md.mermaid import is_complex_table, table_to_mermaid

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_table(
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


def _slide_with_table(table: TableShapeIR, title: str = "T") -> SlideIR:
    return SlideIR(index=0, title=title, shapes=[table])


# ---------------------------------------------------------------------------
# AC1: 셀 내 줄바꿈 — \n → <br>, GFM 행이 깨지지 않는다
# ---------------------------------------------------------------------------


class TestAc1CellNewline:
    """ac1_셀_줄바꿈_br_치환"""

    def test_ac1_newline_replaced_with_br(self) -> None:
        """ac1: Cell text with \\n is replaced with <br> in GFM output."""
        rows = [["Header"], ["line1\nline2"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        assert "<br>" in md
        assert "\nline2" not in md  # raw newline must not leak into pipe table

    def test_ac1_vertical_tab_replaced_with_br(self) -> None:
        """ac1: Cell text with \\v (vertical tab) is also replaced with <br>."""
        rows = [["H"], ["a\vb"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        assert "<br>" in md
        assert "\x0b" not in md

    def test_ac1_row_count_preserved(self) -> None:
        """ac1: A 3-row table still renders as exactly 3 pipe-table rows (+ separator)."""
        rows = [["H1", "H2"], ["a\nb", "c"], ["d", "e"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        # GFM table: header | separator | data1 | data2 = 4 pipe lines
        pipe_lines = [ln for ln in md.splitlines() if ln.startswith("|")]
        # header + separator + 2 data rows = 4
        assert len(pipe_lines) == 4

    def test_ac1_no_newline_cell_unchanged(self) -> None:
        """ac1: Cells without newlines are not modified."""
        rows = [["A", "B"], ["x", "y"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        assert "| x | y |" in md

    def test_ac1_multiple_newlines_all_replaced(self) -> None:
        """ac1: Multiple \\n in one cell are each replaced with <br>."""
        rows = [["H"], ["a\nb\nc"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        # "a<br>b<br>c" expected
        assert "a<br>b<br>c" in md


# ---------------------------------------------------------------------------
# AC2: 파이프 이스케이프 — 리터럴 | → \|
# ---------------------------------------------------------------------------


class TestAc2PipeEscape:
    """ac2_파이프_이스케이프"""

    def test_ac2_pipe_in_cell_escaped(self) -> None:
        """ac2: Literal | in cell text is escaped as \\| in GFM output."""
        rows = [["Header"], ["a|b"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        assert r"\|" in md

    def test_ac2_raw_pipe_not_present_as_column_separator(self) -> None:
        """ac2: After escaping, the cell text a|b appears as a\\|b, not splitting columns."""
        rows = [["H"], ["val|ue"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        # The escaped form must be present
        assert r"val\|ue" in md

    def test_ac2_multiple_pipes_all_escaped(self) -> None:
        """ac2: Multiple | in one cell are all escaped."""
        rows = [["H"], ["a|b|c"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        assert r"a\|b\|c" in md

    def test_ac2_pipe_in_header_escaped(self) -> None:
        """ac2: Pipe in header cell is also escaped."""
        rows = [["A|B", "C"], ["x", "y"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        assert r"A\|B" in md


# ---------------------------------------------------------------------------
# AC3: Mermaid 프리픽스 강제 — fence 안 모든 라인이 %%로 시작
# ---------------------------------------------------------------------------


class TestAc3MermaidPrefixEnforced:
    """ac3_mermaid_프리픽스_강제"""

    def _get_fence_content_lines(self, mermaid_str: str) -> list[str]:
        """Return lines inside the ```mermaid ... ``` fence (not fence markers)."""
        lines = mermaid_str.splitlines()
        # Strip opening ```mermaid and closing ```
        inner = [ln for ln in lines if ln not in ("```mermaid", "```")]
        return inner

    def test_ac3_all_content_lines_start_with_percent_percent(self) -> None:
        """ac3: Every content line in mermaid block starts with %%."""
        rows = [["H1", "H2"], ["a", "b"], ["c", "d"]]
        table = _make_table(rows, n_rows=3, n_cols=2)
        # Force complex via large cell count
        big_rows = [["h" + str(c) for c in range(9)] for _ in range(6)]
        big_table = _make_table(big_rows, n_rows=6, n_cols=9)
        result = table_to_mermaid(big_table)
        inner = self._get_fence_content_lines(result)
        assert inner, "mermaid block must have content lines"
        for line in inner:
            assert line.startswith("%%"), f"Line does not start with %%: {line!r}"

    def test_ac3_header_line_starts_with_percent_percent_headers(self) -> None:
        """ac3: Headers line format: %% headers: ..."""
        rows = [["Col1", "Col2"]] + [["v", "w"]] * 10
        table = _make_table(rows, n_rows=11, n_cols=2)
        result = table_to_mermaid(table)
        assert "%% headers: Col1 | Col2" in result

    def test_ac3_row_lines_start_with_percent_percent_row(self) -> None:
        """ac3: Data row lines format: %% row N: ..."""
        rows = [["H"]] + [[f"v{i}"] for i in range(51)]
        table = _make_table(rows, n_rows=52, n_cols=1)
        result = table_to_mermaid(table)
        assert "%% row 1: v0" in result
        assert "%% row 2: v1" in result

    def test_ac3_table_dimension_line_starts_with_percent_percent(self) -> None:
        """ac3: First content line is %% table: NxM."""
        rows = [["A", "B"]] * 6
        table = _make_table(rows, n_rows=6, n_cols=2)
        # 12 cells <= 50 and 2 cols — not complex by size; force via merged cells
        merged_rows = [["병합", "병합", "x"]] * 10  # 10*3=30 cells, col adj dup
        merged_table = _make_table(merged_rows, n_rows=10, n_cols=3)
        result = table_to_mermaid(merged_table)
        lines = result.splitlines()
        # line 0 = "```mermaid", line 1 = "%% table: ..."
        assert lines[1].startswith("%% table:")

    def test_ac3_no_bare_data_line_without_prefix(self) -> None:
        """ac3: No content line is present without %% prefix."""
        big_rows = [["cell_" + str(c) for c in range(9)] for _ in range(6)]
        big_table = _make_table(big_rows, n_rows=6, n_cols=9)
        result = table_to_mermaid(big_table)
        inner = self._get_fence_content_lines(result)
        bare_lines = [ln for ln in inner if not ln.startswith("%%")]
        assert bare_lines == [], f"Bare lines found: {bare_lines}"


# ---------------------------------------------------------------------------
# AC4: 빈 열/전공백 표 omit
# ---------------------------------------------------------------------------


class TestAc4BlankTableOmit:
    """ac4_빈표_omit"""

    def test_ac4_all_whitespace_cells_omit_gfm(self) -> None:
        """ac4: Table where all cells are whitespace -> omitted from GFM output."""
        rows = [["  ", "  "], [" ", "\t"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        # Should not contain any pipe-table lines
        assert "|" not in md.replace("## T", "")  # heading may not have pipes

    def test_ac4_empty_rows_omit(self) -> None:
        """ac4: Table with no rows -> omitted (returns empty string)."""
        table = _make_table([], n_rows=0, n_cols=0)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        # Only heading should remain
        assert "## T" in md
        pipe_lines = [ln for ln in md.splitlines() if "|" in ln]
        assert pipe_lines == []

    def test_ac4_mermaid_blank_table_returns_empty_string(self) -> None:
        """ac4: table_to_mermaid on all-whitespace table returns ''."""
        rows = [["   "], ["  "]]
        table = _make_table(rows)
        result = table_to_mermaid(table)
        assert result == ""

    def test_ac4_mermaid_empty_table_returns_empty_string(self) -> None:
        """ac4: table_to_mermaid on 0-row table returns ''."""
        table = _make_table([], n_rows=0, n_cols=0)
        result = table_to_mermaid(table)
        assert result == ""

    def test_ac4_table_with_at_least_one_nonempty_cell_not_omitted(self) -> None:
        """ac4: Table with at least one non-empty cell must NOT be omitted."""
        rows = [["  ", "content"], [" ", "  "]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        assert "content" in md


# ---------------------------------------------------------------------------
# AC5: 병합 셀 보존 — 텍스트 손실 없이 보존, fallback 경로로 직렬화
# ---------------------------------------------------------------------------


class TestAc5MergedCellPreservation:
    """ac5_병합셀_보존"""

    def test_ac5_merged_cell_triggers_complex(self) -> None:
        """ac5: Row with adjacent identical non-empty cells -> is_complex_table=True."""
        rows = [["병합", "병합", "x"]]
        table = _make_table(rows, n_rows=1, n_cols=3)
        assert is_complex_table(table)

    def test_ac5_merged_cell_text_in_mermaid_output(self) -> None:
        """ac5: All cell texts (including repeated merged text) appear in mermaid output."""
        rows = [["병합", "병합", "x"], ["a", "b", "c"]]
        table = _make_table(rows, n_rows=2, n_cols=3)
        result = table_to_mermaid(table)
        assert "병합" in result
        assert "x" in result
        assert "a" in result
        assert "b" in result
        assert "c" in result

    def test_ac5_merged_cell_routes_through_fallback(self) -> None:
        """ac5: Merged-cell table renders via mermaid fallback, not GFM pipe table."""
        rows = [["병합", "병합", "x"], ["a", "b", "c"]]
        table = _make_table(rows, n_rows=2, n_cols=3)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        assert "```mermaid" in md
        # No GFM pipe-table header row (unescaped column separator outside mermaid)
        # The content is inside the mermaid block — verify cell presence
        assert "병합" in md

    def test_ac5_no_cell_text_lost(self) -> None:
        """ac5: Zero cell texts are lost when serialising merged-cell table."""
        rows = [["A", "A", "B"], ["C", "D", "D"]]
        table = _make_table(rows, n_rows=2, n_cols=3)
        result = table_to_mermaid(table)
        for cell_text in ["A", "B", "C", "D"]:
            assert cell_text in result, f"Cell text {cell_text!r} missing"


# ---------------------------------------------------------------------------
# AC6: fallback 옵션 — table-data / html / default mermaid
# ---------------------------------------------------------------------------


class TestAc6FallbackOptions:
    """ac6_fallback_옵션"""

    def _complex_table(self) -> TableShapeIR:
        """Return a table that triggers is_complex_table=True (merged cells)."""
        rows = [["병합", "병합", "x"], ["a", "b", "c"]]
        return _make_table(rows, n_rows=2, n_cols=3)

    def test_ac6_default_mermaid_format(self) -> None:
        """ac6: Default (no option) -> mermaid fence."""
        table = self._complex_table()
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        assert "```mermaid" in md

    def test_ac6_table_data_format(self) -> None:
        """ac6: table_fallback='table-data' -> ```table-data fence."""
        table = self._complex_table()
        slide = _slide_with_table(table)
        md = assemble_slide(slide, table_fallback="table-data")
        assert "```table-data" in md
        assert "```mermaid" not in md

    def test_ac6_html_format(self) -> None:
        """ac6: table_fallback='html' -> HTML <table> block."""
        table = self._complex_table()
        slide = _slide_with_table(table)
        md = assemble_slide(slide, table_fallback="html")
        assert "<table>" in md
        assert "```mermaid" not in md

    def test_ac6_html_contains_all_cell_texts(self) -> None:
        """ac6: HTML fallback contains all cell texts."""
        table = self._complex_table()
        slide = _slide_with_table(table)
        md = assemble_slide(slide, table_fallback="html")
        for text in ["병합", "x", "a", "b", "c"]:
            assert text in md, f"Cell text {text!r} missing from HTML output"

    def test_ac6_table_data_contains_all_cell_texts(self) -> None:
        """ac6: table-data fallback contains all cell texts."""
        table = self._complex_table()
        slide = _slide_with_table(table)
        md = assemble_slide(slide, table_fallback="table-data")
        for text in ["병합", "x", "a", "b", "c"]:
            assert text in md, f"Cell text {text!r} missing from table-data output"

    def test_ac6_simple_table_unaffected_by_fallback_option(self) -> None:
        """ac6: Simple (non-complex) table always renders as GFM regardless of fallback option."""
        rows = [["H1", "H2"], ["a", "b"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        # Even if table_fallback='html', simple tables go through GFM path
        md = assemble_slide(slide, table_fallback="html")
        assert "| H1 | H2 |" in md
        assert "<table>" not in md

    def test_ac6_html_escapes_special_chars(self) -> None:
        """ac6: HTML fallback escapes < > & in cell texts."""
        rows = [["H"], ["<tag> & more"]]
        # Force complex by using merged cells
        rows_complex = [["<tag>", "<tag>", "&"]] + [["a", "b", "c"]] * 2
        table = _make_table(rows_complex, n_rows=3, n_cols=3)
        slide = _slide_with_table(table)
        md = assemble_slide(slide, table_fallback="html")
        assert "&lt;tag&gt;" in md
        assert "&amp;" in md


# ---------------------------------------------------------------------------
# AC7: 결정성 — 동일 IR 2회 호출 동일 출력
# ---------------------------------------------------------------------------


class TestAc7Determinism:
    """ac7_결정성"""

    def test_ac7_render_table_gfm_deterministic(self) -> None:
        """ac7: _render_table (GFM path) called twice returns identical output."""
        rows = [["H1", "H2"], ["a", "b"], ["c", "d"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md1 = assemble_slide(slide)
        md2 = assemble_slide(slide)
        assert md1 == md2

    def test_ac7_render_table_mermaid_deterministic(self) -> None:
        """ac7: table_to_mermaid called twice returns identical output."""
        rows = [["H1", "H2"]] + [["v" + str(i), "w" + str(i)] for i in range(26)]
        table = _make_table(rows, n_rows=27, n_cols=2)
        r1 = table_to_mermaid(table)
        r2 = table_to_mermaid(table)
        assert r1 == r2

    def test_ac7_assemble_slide_deterministic_with_special_chars(self) -> None:
        """ac7: assemble_slide with special-char table cells is deterministic."""
        rows = [["H|ead", "Col\n2"], ["val|ue", "line1\nline2"]]
        table = _make_table(rows)
        slide = _slide_with_table(table)
        md1 = assemble_slide(slide)
        md2 = assemble_slide(slide)
        assert md1 == md2


# ---------------------------------------------------------------------------
# AC8: 예외 없음 — 비정형 표 안전 처리
# ---------------------------------------------------------------------------


class TestAc8NoRaise:
    """ac8_no_raise_비정형표"""

    def test_ac8_jagged_rows_no_raise(self) -> None:
        """ac8: Table with rows of unequal column counts does not raise."""
        rows = [["H1", "H2", "H3"], ["only_one"], ["a", "b"]]
        table = _make_table(rows, n_rows=3, n_cols=3)
        slide = _slide_with_table(table)
        result = assemble_slide(slide)
        assert isinstance(result, str)

    def test_ac8_jagged_rows_header_preserved(self) -> None:
        """ac8: Header row cells are preserved even with jagged data rows."""
        rows = [["H1", "H2"], ["only_one"]]
        table = _make_table(rows, n_rows=2, n_cols=2)
        slide = _slide_with_table(table)
        md = assemble_slide(slide)
        assert "H1" in md
        assert "H2" in md

    def test_ac8_empty_rows_list_no_raise(self) -> None:
        """ac8: Empty rows list -> no exception, returns empty string."""
        table = _make_table([], n_rows=0, n_cols=0)
        slide = _slide_with_table(table)
        result = assemble_slide(slide)
        assert isinstance(result, str)

    def test_ac8_table_to_mermaid_no_raise_blank(self) -> None:
        """ac8: table_to_mermaid on all-blank table -> no exception, returns ''."""
        rows = [["  "], [" "]]
        table = _make_table(rows)
        result = table_to_mermaid(table)
        assert result == ""

    def test_ac8_document_not_aborted_by_bad_table(self) -> None:
        """ac8: One odd table in a document does not abort other slides."""
        from pptx_md.assembler import assemble_document
        from pptx_md.ir import PresentationIR, TextShapeIR, ParagraphIR

        odd_table = _make_table([["a"], [], ["b"]], n_rows=3, n_cols=1)
        normal_text = TextShapeIR(
            shape_id=1,
            name="text",
            kind=ShapeKind.TEXT,
            paragraphs=[ParagraphIR(text="normal content", level=0)],
        )
        slide1 = SlideIR(index=0, title="S1", shapes=[odd_table])
        slide2 = SlideIR(index=1, title="S2", shapes=[normal_text])
        pres = PresentationIR(source_path="test.pptx", slides=[slide1, slide2])
        doc = assemble_document(pres)
        assert isinstance(doc, str)
        assert "normal content" in doc
