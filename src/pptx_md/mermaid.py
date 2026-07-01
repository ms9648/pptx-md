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

    # Condition 3: intra-row adjacent repeated non-empty cells
    # (merged-cell heuristic — python-pptx repeats merged cell text within a row)
    if _has_intrarow_repeated_cells(table):
        _logger.debug(
            "is_complex_table: shape_id=%d intra-row repeated cells detected",
            table.shape_id,
        )
        return True

    return False


def _has_intrarow_repeated_cells(table: TableShapeIR) -> bool:
    """Return True if any row has adjacent cells with the same non-empty text.

    Checks each row for consecutive cell pairs where both cells share an
    identical non-empty string — a heuristic for horizontally merged cells
    that python-pptx expands by repeating the text (ARCH-M2 §11 debt).

    Examples:
        ["병합", "병합", "x"] -> True  (adjacent duplicate non-empty cells)
        ["a", "b", "c"]       -> False (no adjacent duplicates)
        ["", "", "x"]         -> False (empty cells are excluded)
    """
    for row in table.rows:
        for j in range(len(row) - 1):
            cell_a = row[j].strip()
            cell_b = row[j + 1].strip()
            if cell_a and cell_a == cell_b:
                return True
    return False


def _is_blank_table(table: TableShapeIR) -> bool:
    """Return True if the table has no rows, or all cells are empty/whitespace.

    Used to decide whether to omit the table entirely (FR-25 AC4).

    Args:
        table: A TableShapeIR instance (read-only).

    Returns:
        True when the table should be omitted from output.
    """
    if not table.rows:
        return True
    for row in table.rows:
        for cell in row:
            if cell.strip():
                return False
    return True


def table_to_mermaid(table: TableShapeIR) -> str:
    """Serialise *table* as a fenced ```mermaid block (FR-13, ADR-219 option C).

    The block contains a deterministic header-row text structure.
    Every content line inside the fence begins with ``%%`` (FR-25 AC3):

        ```mermaid
        %% table: {n_rows}x{n_cols}
        %% headers: col1 | col2 | ...
        %% row 1: val1 | val2 | ...
        %% row 2: val1 | val2 | ...
        ```

    Returns an empty string when *table* is blank (all cells whitespace or no
    rows — FR-25 AC4).

    This is NOT a formal Mermaid diagram (ADR-219 C); the ```mermaid fence is
    used so that downstream LLM pipelines can identify and handle the block.
    All cell texts are included verbatim (zero data loss, FR-25 AC5/AC7).

    Deterministic: identical TableShapeIR -> identical output string (FR-13 AC7).

    Args:
        table: A TableShapeIR instance (read-only).

    Returns:
        A string beginning with ```mermaid and ending with ``` (including
        surrounding newlines), or "" for blank/empty tables (FR-25 AC4).
    """
    # AC4: omit blank/empty tables entirely — no mermaid block produced
    if _is_blank_table(table):
        return ""

    rows = table.rows
    n_rows = table.n_rows
    n_cols = table.n_cols

    lines: list[str] = []
    lines.append("```mermaid")
    lines.append(f"%% table: {n_rows}x{n_cols}")

    # First row treated as headers (AC3: %% prefix enforced on every content line)
    header_cells = rows[0]
    header_line = " | ".join(header_cells)
    lines.append(f"%% headers: {header_line}")

    # Subsequent rows (AC3: every data line starts with %%)
    for row_idx, row in enumerate(rows[1:], start=1):
        row_line = " | ".join(row)
        lines.append(f"%% row {row_idx}: {row_line}")

    lines.append("```")
    return "\n".join(lines)
