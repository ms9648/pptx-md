"""Tests for FR-13: Mermaid fallback for complex tables.

Covers AC1–AC8 of issue #35.  All tests use synthetic TableShapeIR fixtures
(no python-pptx dependency — ADR-211).
"""

from __future__ import annotations

from pptx_md.ir import ShapeKind, TableShapeIR
from pptx_md.mermaid import (
    MAX_TABLE_CELLS,
    MAX_TABLE_COLS,
    is_complex_table,
    table_to_mermaid,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_table(
    rows: list[list[str]],
    n_rows: int | None = None,
    n_cols: int | None = None,
    shape_id: int = 1,
) -> TableShapeIR:
    """Construct a TableShapeIR from rows data."""
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


def _make_rows(n_rows: int, n_cols: int, prefix: str = "cell") -> list[list[str]]:
    """Generate a n_rows x n_cols grid of unique cell texts."""
    return [[f"{prefix}_{r}_{c}" for c in range(n_cols)] for r in range(n_rows)]


# ---------------------------------------------------------------------------
# AC1: cell count not exceeded (6 cells) -> GFM table, Mermaid NOT triggered
# ---------------------------------------------------------------------------


class TestAc1SmallTableNoMermaid:
    """ac1_소표_GFM표_mermaid미발동"""

    def test_ac1_2x3_table_is_not_complex(self) -> None:
        """6 cells (2x3) is below MAX_TABLE_CELLS=50 — should NOT be complex."""
        rows = [["A", "B", "C"], ["D", "E", "F"]]
        table = _make_table(rows)
        assert not is_complex_table(table)

    def test_ac1_assemble_gives_gfm_table(self) -> None:
        """6-cell table rendered by assembler produces GFM pipe table, not mermaid."""
        from pptx_md.assembler import assemble_slide
        from pptx_md.ir import SlideIR

        rows = [["A", "B", "C"], ["D", "E", "F"]]
        table = _make_table(rows)
        slide = SlideIR(index=0, title="T", shapes=[table])
        md = assemble_slide(slide)
        assert "```mermaid" not in md
        assert "| A | B | C |" in md
        assert "|" in md


# ---------------------------------------------------------------------------
# AC2: n_rows=11, n_cols=5 (55 cells > 50) -> mermaid block, no GFM table marker
# ---------------------------------------------------------------------------


class TestAc2LargeTableMermaid:
    """ac2_55셀_mermaid발동"""

    def test_ac2_55_cells_is_complex(self) -> None:
        """11x5=55 cells > MAX_TABLE_CELLS(50) -> is_complex_table=True."""
        rows = _make_rows(11, 5)
        table = _make_table(rows, n_rows=11, n_cols=5)
        assert is_complex_table(table)

    def test_ac2_mermaid_block_produced(self) -> None:
        """55-cell table renders as ```mermaid block."""
        rows = _make_rows(11, 5)
        table = _make_table(rows, n_rows=11, n_cols=5)
        result = table_to_mermaid(table)
        assert result.startswith("```mermaid")
        assert result.strip().endswith("```")

    def test_ac2_no_gfm_pipe_table(self) -> None:
        """Mermaid output must NOT contain GFM pipe-table rows."""
        rows = _make_rows(11, 5)
        table = _make_table(rows, n_rows=11, n_cols=5)
        result = table_to_mermaid(table)
        # GFM tables have lines like "| x | y |" without %% prefix
        gfm_lines = [
            line
            for line in result.splitlines()
            if line.startswith("|") and not line.startswith("%%")
        ]
        assert gfm_lines == []


# ---------------------------------------------------------------------------
# AC3: n_cols=9 (> MAX_TABLE_COLS=8) -> Mermaid fallback triggered
# ---------------------------------------------------------------------------


class TestAc3WideTableMermaid:
    """ac3_9열_mermaid발동"""

    def test_ac3_9_cols_is_complex(self) -> None:
        """9 columns > MAX_TABLE_COLS(8) -> is_complex_table=True."""
        rows = _make_rows(2, 9)
        table = _make_table(rows, n_rows=2, n_cols=9)
        assert is_complex_table(table)

    def test_ac3_8_cols_is_not_complex_by_cols(self) -> None:
        """8 columns == MAX_TABLE_COLS -> NOT triggered by column condition alone."""
        rows = _make_rows(2, 8)
        table = _make_table(rows, n_rows=2, n_cols=8)
        # 2*8=16 cells <= 50, 8 cols not > 8 -> not complex
        assert not is_complex_table(table)


# ---------------------------------------------------------------------------
# AC4: consecutive identical non-empty rows -> merge heuristic -> Mermaid
# ---------------------------------------------------------------------------


class TestAc4RepeatedRowsMermaid:
    """ac4_행내인접반복셀_mermaid발동"""

    def test_ac4_intrarow_adjacent_duplicate_nonempty_triggers_fallback(self) -> None:
        """Row with adjacent identical non-empty cells -> is_complex_table=True.

        AC4 example: ["병합", "병합", "x"] has two adjacent "병합" cells.
        """
        rows = [["병합", "병합", "x"]]
        table = _make_table(rows, n_rows=1, n_cols=3)
        assert is_complex_table(table)

    def test_ac4_no_adjacent_duplicates_not_triggered(self) -> None:
        """Row with no adjacent duplicate cells -> NOT triggered by heuristic."""
        rows = [["a", "b", "c"]]
        table = _make_table(rows, n_rows=1, n_cols=3)
        # 1*3=3 cells, 3 cols, no adjacent duplicates
        assert not is_complex_table(table)

    def test_ac4_empty_cells_not_counted_as_duplicate(self) -> None:
        """Adjacent empty cells do NOT trigger fallback (empty cells excluded)."""
        rows = [["", "", "x"]]
        table = _make_table(rows, n_rows=1, n_cols=3)
        # empty cells are excluded from intra-row repetition check
        assert not is_complex_table(table)

    def test_ac4_multirow_one_row_has_duplicate_triggers(self) -> None:
        """Multi-row table where one row has adjacent duplicate -> True."""
        rows = [
            ["Header1", "Header2", "Header3"],
            ["병합", "병합", "x"],
        ]
        table = _make_table(rows, n_rows=2, n_cols=3)
        assert is_complex_table(table)

    def test_ac4_all_distinct_multirow_not_triggered(self) -> None:
        """Multi-row table with all distinct adjacent cells -> NOT triggered."""
        rows = [
            ["Header1", "Header2"],
            ["A", "B"],
            ["C", "D"],
        ]
        table = _make_table(rows, n_rows=3, n_cols=2)
        # 3*2=6 cells, 2 cols, no intra-row adjacent duplicates
        assert not is_complex_table(table)


# ---------------------------------------------------------------------------
# AC5: all cell texts present in Mermaid block (no data loss)
# ---------------------------------------------------------------------------


class TestAc5CellTextCompleteness:
    """ac5_모든셀텍스트_mermaid포함"""

    def test_ac5_all_cells_in_mermaid_output(self) -> None:
        """Every non-empty cell text must appear in the Mermaid block."""
        rows = _make_rows(11, 5)  # 55 cells
        table = _make_table(rows, n_rows=11, n_cols=5)
        result = table_to_mermaid(table)
        for r in rows:
            for cell in r:
                assert cell in result, f"Cell text '{cell}' missing from Mermaid output"

    def test_ac5_mermaid_block_fenced_correctly(self) -> None:
        """Mermaid block must start with ```mermaid and end with ```."""
        rows = _make_rows(11, 5)
        table = _make_table(rows, n_rows=11, n_cols=5)
        result = table_to_mermaid(table)
        lines = result.splitlines()
        assert lines[0] == "```mermaid"
        assert lines[-1] == "```"


# ---------------------------------------------------------------------------
# AC6 (boundary): empty TableShapeIR -> no fallback, no exception
# ---------------------------------------------------------------------------


class TestAc6EmptyTableNoFallback:
    """ac6_빈테이블_fallback미발동_예외없음"""

    def test_ac6_empty_table_not_complex(self) -> None:
        """Empty table (0 rows, 0 cols) -> is_complex_table=False."""
        table = _make_table([], n_rows=0, n_cols=0)
        assert not is_complex_table(table)

    def test_ac6_empty_table_mermaid_no_exception(self) -> None:
        """table_to_mermaid on empty table must not raise and returns "" (FR-25 AC4)."""
        table = _make_table([], n_rows=0, n_cols=0)
        result = table_to_mermaid(table)
        # FR-25 AC4: empty table -> omit (return empty string, not a mermaid block)
        assert result == ""

    def test_ac6_assembler_empty_table_no_exception(self) -> None:
        """assemble_slide with an empty table must not raise."""
        from pptx_md.assembler import assemble_slide
        from pptx_md.ir import SlideIR

        table = _make_table([], n_rows=0, n_cols=0)
        slide = SlideIR(index=0, title="Empty", shapes=[table])
        result = assemble_slide(slide)
        assert isinstance(result, str)


# ---------------------------------------------------------------------------
# AC7: determinism — same input twice -> same output
# ---------------------------------------------------------------------------


class TestAc7Determinism:
    """ac7_결정성_동일입력_동일출력"""

    def test_ac7_is_complex_table_deterministic(self) -> None:
        """is_complex_table called twice with same input returns same result."""
        rows = _make_rows(11, 5)
        table = _make_table(rows, n_rows=11, n_cols=5)
        result1 = is_complex_table(table)
        result2 = is_complex_table(table)
        assert result1 == result2

    def test_ac7_table_to_mermaid_deterministic(self) -> None:
        """table_to_mermaid called twice with same input returns identical strings."""
        rows = _make_rows(11, 5)
        table = _make_table(rows, n_rows=11, n_cols=5)
        result1 = table_to_mermaid(table)
        result2 = table_to_mermaid(table)
        assert result1 == result2

    def test_ac7_small_table_gfm_deterministic(self) -> None:
        """GFM table rendered twice is identical."""
        from pptx_md.assembler import assemble_slide
        from pptx_md.ir import SlideIR

        rows = [["A", "B"], ["C", "D"]]
        table = _make_table(rows)
        slide = SlideIR(index=0, title="T", shapes=[table])
        md1 = assemble_slide(slide)
        md2 = assemble_slide(slide)
        assert md1 == md2


# ---------------------------------------------------------------------------
# AC8: named constants MAX_TABLE_CELLS and MAX_TABLE_COLS exist
# ---------------------------------------------------------------------------


class TestAc8NamedConstants:
    """ac8_명명상수_존재"""

    def test_ac8_max_table_cells_constant_exists(self) -> None:
        """MAX_TABLE_CELLS must be importable and equal 50."""
        assert MAX_TABLE_CELLS == 50

    def test_ac8_max_table_cols_constant_exists(self) -> None:
        """MAX_TABLE_COLS must be importable and equal 8."""
        assert MAX_TABLE_COLS == 8

    def test_ac8_constants_are_int(self) -> None:
        """Both threshold constants must be integers."""
        assert isinstance(MAX_TABLE_CELLS, int)
        assert isinstance(MAX_TABLE_COLS, int)

    def test_ac8_boundary_cells_exactly_50_not_complex(self) -> None:
        """Exactly MAX_TABLE_CELLS (50) cells -> NOT complex (strictly >)."""
        # 5x10 = 50 cells, 10 cols > MAX_TABLE_COLS(8) -> col condition triggers
        # Use 10x5 = 50 cells, 5 cols -> only cell condition; 50 is NOT > 50
        rows = _make_rows(10, 5)
        table = _make_table(rows, n_rows=10, n_cols=5)
        # 50 cells, 5 cols — neither condition triggered
        assert not is_complex_table(table)

    def test_ac8_boundary_cells_51_is_complex(self) -> None:
        """51 cells (> MAX_TABLE_CELLS) -> complex."""
        rows = _make_rows(51, 1)
        table = _make_table(rows, n_rows=51, n_cols=1)
        assert is_complex_table(table)
