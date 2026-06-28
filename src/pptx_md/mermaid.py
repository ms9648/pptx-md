"""Mermaid fallback serialisation for complex tables (FR-13, ADR-219).

When a table exceeds the Markdown-table complexity threshold, this module
serialises it as a ```mermaid fenced code block containing a deterministic
header-row text structure (ADR-219 option C: not a formal graph diagram,
optimised for LLM consumption with zero data loss).

Public interface:
    MAX_TABLE_CELLS: int  -- named constant (FR-13 AC8)
    MAX_TABLE_COLS:  int  -- named constant (FR-13 AC8)
    is_complex_table(table: TableShapeIR) -> bool
    table_to_mermaid(table: TableShapeIR) -> str

Design constraints:
    - Deterministic: same input -> same output (FR-13 AC7, ADR-219).
    - Pure functions, no side effects.
    - No raise: all errors are handled gracefully.
    - Imports pptx_md.ir for TableShapeIR (per issue #35 task scope §1).
"""

from __future__ import annotations

import logging

from pptx_md.ir import TableShapeIR

__all__ = [
    "MAX_TABLE_CELLS",
    "MAX_TABLE_COLS",
    "is_complex_table",
    "table_to_mermaid",
]

_logger = logging.getLogger("pptx_md.mermaid")

# ---------------------------------------------------------------------------
# Named threshold constants (FR-13 AC8, ADR-219)
# ---------------------------------------------------------------------------
MAX_TABLE_CELLS: int = 50
MAX_TABLE_COLS: int = 8


def is_complex_table(table: TableShapeIR) -> bool:
    """Decide whether *table* exceeds the Markdown-table complexity threshold.

    Returns True if any of the three conditions hold (ADR-219):

    1. Total cell count (n_rows * n_cols) > MAX_TABLE_CELLS (50)
    2. Column count (n_cols) > MAX_TABLE_COLS (8)
    3. Any non-empty row text sequence appears in two or more consecutive
       rows at the same column position — heuristic for merged cells that
       python-pptx expands by repeating the text (ARCH-M2 §11 debt).

    Deterministic: no randomness, no external state (ADR-219).

    Args:
        table: A TableShapeIR instance (read-only).

    Returns:
        True when the table should be serialised as a Mermaid block.
    """
    # Condition 1: total cell count
    total_cells = table.n_rows * table.n_cols
    if total_cells > MAX_TABLE_CELLS:
        _logger.debug(
            "is_complex_table: shape_id=%d total_cells=%d > %d",
            table.shape_id,
            total_cells,
            MAX_TABLE_CELLS,
        )
        return True

    # Condition 2: column count
    if table.n_cols > MAX_TABLE_COLS:
        _logger.debug(
            "is_complex_table: shape_id=%d n_cols=%d > %d",
            table.shape_id,
            table.n_cols,
            MAX_TABLE_COLS,
        )
        return True

    # Condition 3: consecutive rows with identical non-empty text in any column
    # (merged-cell heuristic — python-pptx repeats merged cell text)
    if _has_repeated_consecutive_rows(table):
        _logger.debug(
            "is_complex_table: shape_id=%d repeated-consecutive-rows detected",
            table.shape_id,
        )
        return True

    return False


def _has_repeated_consecutive_rows(table: TableShapeIR) -> bool:
    """Return True if any non-empty row appears identically in two consecutive rows.

    A "row" here means the tuple of all cell texts in that row.  We check
    whether the same non-empty tuple appears in row[i] and row[i+1].
    """
    if table.n_rows < 2:
        return False

    rows = table.rows
    for i in range(len(rows) - 1):
        row_a = rows[i]
        row_b = rows[i + 1]
        # Rows must be equal, non-empty (at least one non-blank cell)
        if row_a == row_b and any(cell.strip() for cell in row_a):
            return True
    return False


def table_to_mermaid(table: TableShapeIR) -> str:
    """Serialise *table* as a fenced ```mermaid block (FR-13, ADR-219 option C).

    The block contains a deterministic header-row text structure:

        ```mermaid
        %% table: {n_rows}x{n_cols}
        %% headers: col1 | col2 | ...
        %% row 1: val1 | val2 | ...
        %% row 2: val1 | val2 | ...
        ```

    This is NOT a formal Mermaid diagram (ADR-219 C); the ```mermaid fence is
    used so that downstream LLM pipelines can identify and handle the block.
    All cell texts are included verbatim (zero data loss).

    Deterministic: identical TableShapeIR -> identical output string (FR-13 AC7).

    Args:
        table: A TableShapeIR instance (read-only).

    Returns:
        A string beginning with ```mermaid and ending with ``` (including
        surrounding newlines).
    """
    rows = table.rows
    n_rows = table.n_rows
    n_cols = table.n_cols

    lines: list[str] = []
    lines.append("```mermaid")
    lines.append(f"%% table: {n_rows}x{n_cols}")

    if rows:
        # First row treated as headers
        header_cells = rows[0] if rows else []
        header_line = " | ".join(header_cells)
        lines.append(f"%% headers: {header_line}")

        # Subsequent rows
        for row_idx, row in enumerate(rows[1:], start=1):
            row_line = " | ".join(row)
            lines.append(f"%% row {row_idx}: {row_line}")
    else:
        lines.append("%% (empty table)")

    lines.append("```")
    return "\n".join(lines)
