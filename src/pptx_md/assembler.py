"""Markdown assembler — converts a SlideIR / PresentationIR into Markdown (FR-11/12).

Public interface:
    assemble_slide(slide: SlideIR, *, masking: MaskingOptions | None = None) -> str
    assemble_document(
        presentation: PresentationIR,
        *,
        masking: MaskingOptions | None = None,
        suppress_repeated_labels: bool = False,
    ) -> str

Design constraints (ADR-218):
    - Deterministic: same IR -> same Markdown (no randomness, no set/dict iteration).
    - IR is read-only: assembler never mutates input objects.
    - Partial-failure isolation: one shape failure does not abort the slide/document.
    - No VLM/python-pptx/Pillow imports (NFR-08).
    - Uses mermaid.is_complex_table() for FR-13 Mermaid fallback.
    - masking=None means no masking (opt-in, FR-15).

FR-24 assembler normalizations (AC1–AC8):
    - slide.title duplicated by is_title shape is suppressed from body (AC1).
    - Vertical-tab \\v in paragraph text is replaced with \\n (AC2).
    - Consecutive spaces/tabs in paragraph text are collapsed to one space (AC3).
    - Blank tables (all cells whitespace) are omitted from output (AC4).
    - Runs of 3+ consecutive newlines are collapsed to 2 (AC5).
    - suppress_repeated_labels option removes consecutive duplicate labels (AC6).

FR-25 table integrity (AC1–AC8):
    - Cells with \\n/\\v are normalised to <br> before GFM rendering (AC1).
    - Literal | in cells is escaped as \\| to prevent column splitting (AC2).
    - Blank tables (all cells whitespace) are omitted from output (AC4).
    - table_to_mermaid fallback options: "table-data" fence or "html" <table> (AC6).
    - AC7: determinism preserved — no randomness added.
    - AC8: no raise — non-conforming tables handled gracefully.
"""

from __future__ import annotations

import logging
import re
from typing import Literal

from pptx_md.ir import (
    GroupShapeIR,
    ImageShapeIR,
    OtherShapeIR,
    PresentationIR,
    ShapeIR,
    SlideIR,
    TableShapeIR,
    TextShapeIR,
)
from pptx_md.masking import MaskingOptions, mask_text
from pptx_md.mermaid import (
    _is_blank_table,
    _normalise_cell_text,
    is_complex_table,
    table_to_mermaid,
)

__all__ = ["assemble_slide", "assemble_document"]

# Literal type for the complex-table fallback format option (FR-25 AC6)
TableFallbackFormat = Literal["mermaid", "table-data", "html"]

_logger = logging.getLogger("pptx_md.assembler")

# ---------------------------------------------------------------------------
# Module-level named constants (ADR-218 — determinism, §3.2 heading rule)
# ---------------------------------------------------------------------------
_SLIDE_HEADING_PREFIX: str = "##"
_INDENT_UNIT: str = "  "  # two spaces per indent level
_BULLET_MARKER: str = "- "
_SLIDE_SEPARATOR: str = "\n\n---\n\n"

# Pre-compiled regex patterns for normalization (ADR-218: deterministic, pure)
_RE_VERTICAL_TAB: re.Pattern[str] = re.compile(r"\v")
_RE_CONSECUTIVE_SPACES: re.Pattern[str] = re.compile(r"[ \t]{2,}")
_RE_CONSECUTIVE_BLANK_LINES: re.Pattern[str] = re.compile(r"\n{3,}")

# FR-23: Reading-order sort tolerance (EMU).
# Shapes whose `top` values differ by less than this are treated as the same
# visual row and sorted by `left` within that row.
# 914400 EMU = 0.5 inch ≈ 1.27 cm — a practical "same row" threshold for
# PowerPoint slides (approximately 13% of standard slide height 6858000 EMU).
# Using integer floor-division bucketing ensures determinism (ADR-218).
_ROW_TOLERANCE_EMU: int = 914400


# ---------------------------------------------------------------------------
# Normalization utilities (FR-24: pure functions, input-immutable)
# ---------------------------------------------------------------------------


def _normalize_text(text: str) -> str:
    """Normalize a paragraph text string (AC2, AC3 — FR-24).

    Transformations applied in order:
    1. Replace vertical-tab U+000B (PPTX soft line-break) with newline (AC2).
    2. Collapse consecutive spaces/tabs to a single space (AC3).

    Pure function: does not mutate input.  Deterministic (ADR-218).

    Args:
        text: Raw paragraph text from ParagraphIR.

    Returns:
        Normalized text string.
    """
    # AC2: vertical tab -> newline
    text = _RE_VERTICAL_TAB.sub("\n", text)
    # AC3: consecutive spaces/tabs -> single space
    text = _RE_CONSECUTIVE_SPACES.sub(" ", text)
    return text


def _collapse_blank_lines(text: str) -> str:
    """Collapse runs of 3+ consecutive newlines down to 2 (AC5 — FR-24).

    Pure function. Deterministic (ADR-218).

    Args:
        text: Multi-line text (already assembled from paragraphs).

    Returns:
        Text with at most one blank line between content lines.
    """
    return _RE_CONSECUTIVE_BLANK_LINES.sub("\n\n", text)


def _is_table_all_blank(shape: TableShapeIR) -> bool:
    """Return True when every cell text in the table is empty/whitespace (AC4 — FR-24).

    Delegates to mermaid._is_blank_table which handles both no-rows and all-whitespace
    cases.  Kept here for backward compatibility (test_assembler_fr24.py imports it).

    Pure function.  Deterministic (ADR-218).

    Args:
        shape: A TableShapeIR instance (read-only).

    Returns:
        True if all cells are blank or table has no rows, False otherwise.
    """
    return _is_blank_table(shape)


# ---------------------------------------------------------------------------
# Public functions
# ---------------------------------------------------------------------------


def assemble_slide(
    slide: SlideIR,
    *,
    masking: MaskingOptions | None = None,
    table_fallback: TableFallbackFormat = "mermaid",
) -> str:
    """Render one SlideIR into a deterministic Markdown block (FR-11).

    Read-only: does not mutate the IR.  Same IR -> same Markdown (ADR-218).
    masking=None means no masking (opt-in, FR-15).

    Args:
        slide: A SlideIR instance (read-only).
        masking: Optional MaskingOptions instance (FR-15).
        table_fallback: Fallback format for complex/merged-cell tables (FR-25 AC6).
            One of "mermaid" (default, backward-compatible), "table-data", or "html".

    Returns:
        A Markdown string representing this slide.  Empty/blank slides return
        an empty string or heading-only string without raising (FR-11 AC7).
    """
    return _render_slide(slide, masking, table_fallback=table_fallback)


def assemble_document(
    presentation: PresentationIR,
    *,
    masking: MaskingOptions | None = None,
    suppress_repeated_labels: bool = False,
    table_fallback: TableFallbackFormat = "mermaid",
) -> str:
    """Map-Reduce assembly of the whole presentation into one document (FR-12).

    Map: assemble_slide per slide (sequential, ADR-220).
    Reduce: join with a slide separator, in IR list order.
    Deterministic and order-preserving (ADR-220).

    Args:
        presentation: A PresentationIR instance (read-only).
        masking: Optional MaskingOptions instance (FR-15).
        suppress_repeated_labels: When True, deduplicate consecutive identical
            section labels across slide blocks (AC6 — FR-24).  Default False
            preserves existing behaviour.
        table_fallback: Fallback format for complex/merged-cell tables (FR-25 AC6).
            One of "mermaid" (default), "table-data", or "html".

    Returns:
        A single Markdown document string.
    """
    # sort slides by index (ascending) regardless of IR list order (ADR-220)
    sorted_slides = sorted(presentation.slides, key=lambda s: s.index)

    slide_blocks: list[str] = []
    for slide in sorted_slides:
        try:
            block = _render_slide(slide, masking, table_fallback=table_fallback)
        except Exception as exc:  # noqa: BLE001 — slide-level isolation
            _logger.warning(
                "assemble_document: slide index=%d failed, skipping. "
                "shape_count=%d error_type=%s",
                slide.index,
                len(slide.shapes),
                type(exc).__name__,
            )
            block = f"{_SLIDE_HEADING_PREFIX} Slide {slide.index + 1}"
        slide_blocks.append(block)

    if suppress_repeated_labels:
        slide_blocks = _suppress_repeated_labels(slide_blocks)

    return _SLIDE_SEPARATOR.join(slide_blocks)


# ---------------------------------------------------------------------------
# Private render functions
# ---------------------------------------------------------------------------


def _reading_order_key(shape: ShapeIR) -> tuple[int, int]:
    """Return a sort key for reading-order traversal (FR-23).

    Shapes are first grouped into horizontal rows by floor-dividing `top` by
    _ROW_TOLERANCE_EMU (bucket index), then sorted left-to-right within each
    row by `left`.  Python's sorted() is stable, so shapes with identical keys
    retain their original IR order (AC6, AC7, ADR-218).

    Args:
        shape: Any ShapeIR instance (uses left/top fields added by FR-23).

    Returns:
        (row_bucket, left) — both non-negative ints when coordinates are valid.
    """
    row_bucket = shape.top // _ROW_TOLERANCE_EMU if _ROW_TOLERANCE_EMU > 0 else 0
    return (row_bucket, shape.left)


def _render_slide(
    slide: SlideIR,
    masking: MaskingOptions | None,
    *,
    table_fallback: TableFallbackFormat = "mermaid",
) -> str:
    """Render a single SlideIR to Markdown."""
    parts: list[str] = []
    slide_num = slide.index + 1

    # FR-20: Heading — use slide.title first; fall back to first is_title shape;
    # last resort: ## Slide N
    # AC1 (FR-24): when slide.title is set, also skip any is_title=True shape
    # whose text matches slide.title (strip comparison) to prevent duplicate output.
    # NOTE: title shape search uses original IR order (heading search is not
    #       position-dependent; we want the designated title placeholder).
    title_shape_used: TextShapeIR | None = None
    if slide.title:
        parts.append(f"{_SLIDE_HEADING_PREFIX} {slide.title}")
        # AC1: find a matching is_title shape to suppress from body rendering
        for shape in slide.shapes:
            if (
                isinstance(shape, TextShapeIR)
                and shape.is_title
                and shape.paragraphs
                and shape.paragraphs[0].text.strip() == slide.title.strip()
            ):
                title_shape_used = shape
                break
    else:
        # Search for the first is_title=True shape to use as heading
        for shape in slide.shapes:
            if isinstance(shape, TextShapeIR) and shape.is_title and shape.paragraphs:
                heading_text = shape.paragraphs[0].text.strip()
                if heading_text:
                    parts.append(f"{_SLIDE_HEADING_PREFIX} {heading_text}")
                    title_shape_used = shape
                    break
        if title_shape_used is None:
            # Fallback: no title available
            parts.append(f"{_SLIDE_HEADING_PREFIX} Slide {slide_num}")

    # FR-22: image counter (mutable list used as a reference for group recursion)
    img_counter: list[int] = [0]

    # FR-23: sort shapes by reading order (top row first, then left-to-right).
    # sorted() is stable: equal-key shapes keep their original IR order (AC6/AC7).
    # IR is read-only — we create a new list, never mutate slide.shapes (ADR-218).
    ordered_shapes: list[ShapeIR] = sorted(slide.shapes, key=_reading_order_key)

    for shape in ordered_shapes:
        # FR-20 / AC1: skip the shape already used as heading
        if shape is title_shape_used:
            continue
        # FR-21: skip footer/slide-number/date placeholders
        if isinstance(shape, TextShapeIR) and shape.is_footer:
            continue
        rendered = _render_shape(
            shape,
            masking,
            slide_num=slide_num,
            img_counter=img_counter,
            table_fallback=table_fallback,
        )
        if rendered:
            parts.append(rendered)

    return "\n\n".join(parts)


def _render_shape(
    shape: ShapeIR,
    masking: MaskingOptions | None,
    *,
    slide_num: int = 0,
    img_counter: list[int] | None = None,
    table_fallback: TableFallbackFormat = "mermaid",
) -> str:
    """Dispatch to the appropriate renderer; isolate per-shape failures (ADR-220).

    Args:
        shape: Any ShapeIR subclass (read-only).
        masking: Optional masking config (FR-15).
        slide_num: 1-based slide number for image markers (FR-22).
        img_counter: Mutable [int] shared across the slide for image numbering.
        table_fallback: Fallback format for complex tables (FR-25 AC6).
    """
    if img_counter is None:
        img_counter = [0]
    try:
        if isinstance(shape, TextShapeIR):
            return _render_text(shape, masking)
        if isinstance(shape, TableShapeIR):
            return _render_table(shape, fallback_format=table_fallback)
        if isinstance(shape, ImageShapeIR):
            img_counter[0] += 1
            return _render_image(
                shape, masking, slide_num=slide_num, img_num=img_counter[0]
            )
        if isinstance(shape, GroupShapeIR):
            return _render_group(
                shape,
                masking,
                slide_num=slide_num,
                img_counter=img_counter,
                table_fallback=table_fallback,
            )
        if isinstance(shape, OtherShapeIR):
            return _render_other(shape, masking)
        # Unknown subclass — best-effort skip
        _logger.warning(
            "_render_shape: unknown shape kind=%s shape_id=%d — skipping",
            getattr(shape, "kind", "?"),
            shape.shape_id,
        )
        return ""
    except Exception as exc:  # noqa: BLE001 — shape-level isolation
        _logger.warning(
            "_render_shape: shape_id=%d kind=%s render failed error_type=%s — skipping",
            shape.shape_id,
            getattr(shape, "kind", "?"),
            type(exc).__name__,
        )
        return ""


def _render_text(shape: TextShapeIR, masking: MaskingOptions | None) -> str:
    """Render a TextShapeIR to Markdown paragraphs / indented bullets.

    Applies FR-24 normalizations:
    - AC2: \\v -> \\n via _normalize_text
    - AC3: consecutive spaces/tabs collapsed via _normalize_text
    - AC5: consecutive blank lines collapsed to max 1 blank line
    """
    lines: list[str] = []
    for para in shape.paragraphs:
        text = para.text
        # FR-24 AC2+AC3: normalize control characters and whitespace
        text = _normalize_text(text)
        # Apply masking if provided (FR-15)
        if masking is not None:
            text = mask_text(text, masking)

        if para.level > 0:
            # Indented bullet item (level=1 deeper than level=0)
            indent = _INDENT_UNIT * para.level
            lines.append(f"{indent}{_BULLET_MARKER}{text}")
        else:
            # Top-level paragraph
            lines.append(text)

    result = "\n".join(lines)
    # AC5: collapse 3+ consecutive newlines to at most 2 (one blank line)
    result = _collapse_blank_lines(result)
    return result


def _sanitise_cell(text: str) -> str:
    """Sanitise a table cell for GFM rendering (FR-25 AC1, AC2).

    Applies two transformations in order:
    1. AC1: replace \\n and \\v (vertical-tab) with <br> so that multiline
       cell text stays within a single pipe-table row.
    2. AC2: escape literal ``|`` as ``\\|`` so that the cell boundary is not
       misinterpreted as a column separator.

    Args:
        text: Raw cell text from TableShapeIR (read-only).

    Returns:
        Sanitised cell string safe for inclusion between pipe delimiters.
    """
    # AC1: normalise newlines — \v (vertical tab, \x0b) and \n both become <br>
    sanitised = text.replace("\v", "<br>").replace("\n", "<br>")
    # AC2: escape literal pipe characters
    sanitised = sanitised.replace("|", r"\|")
    return sanitised


def _table_to_html(shape: TableShapeIR) -> str:
    """Serialise *shape* as an HTML <table> block (FR-25 AC6 — html fallback option).

    Deterministic: same IR -> same HTML string.  No raise.

    Args:
        shape: A TableShapeIR instance (read-only).

    Returns:
        An HTML <table> string, or "" for blank tables.
    """
    if _is_blank_table(shape):
        return ""

    rows = shape.rows
    html_lines: list[str] = ["<table>"]
    for row_idx, row in enumerate(rows):
        html_lines.append("  <tr>")
        tag = "th" if row_idx == 0 else "td"
        for cell in row:
            # Escape HTML special chars minimally (no external deps)
            escaped = (
                cell.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            )
            html_lines.append(f"    <{tag}>{escaped}</{tag}>")
        html_lines.append("  </tr>")
    html_lines.append("</table>")
    return "\n".join(html_lines)


def _table_to_table_data(shape: TableShapeIR) -> str:
    """Serialise *shape* as a ```table-data fenced block (FR-25 AC6).

    Format mirrors the mermaid block but uses a ``table-data`` fence tag so
    that tooling can distinguish it from a real Mermaid diagram:

        ```table-data
        table: {n_rows}x{n_cols}
        headers: col1 | col2 | ...
        row 1: val1 | val2 | ...
        ```

    Deterministic: same IR -> same output.  No raise.

    Args:
        shape: A TableShapeIR instance (read-only).

    Returns:
        A fenced string, or "" for blank tables.
    """
    if _is_blank_table(shape):
        return ""

    rows = shape.rows
    n_rows = shape.n_rows
    n_cols = shape.n_cols

    lines: list[str] = []
    lines.append("```table-data")
    lines.append(f"table: {n_rows}x{n_cols}")
    lines.append("headers: " + " | ".join(_normalise_cell_text(c) for c in rows[0]))
    for row_idx, row in enumerate(rows[1:], start=1):
        lines.append(
            f"row {row_idx}: " + " | ".join(_normalise_cell_text(c) for c in row)
        )
    lines.append("```")
    return "\n".join(lines)


def _render_table(
    shape: TableShapeIR,
    fallback_format: TableFallbackFormat = "mermaid",
) -> str:
    """Render a TableShapeIR as GFM table or complex-table fallback (FR-13, FR-25).

    For GFM rendering (simple tables), cell texts are sanitised:
    - Newlines (\\n, \\v) -> <br>  (AC1)
    - Literal | -> \\|            (AC2)

    Blank tables (all cells whitespace) are omitted regardless of format (AC4).

    For complex tables the fallback format is determined by *fallback_format*:
    - "mermaid"     (default): ```mermaid block (AC6, backward-compatible)
    - "table-data"           : ```table-data block (AC6)
    - "html"                 : HTML <table> block (AC6)

    Deterministic (AC7).  No raise (AC8).

    Args:
        shape: A TableShapeIR instance (read-only).
        fallback_format: Fallback serialisation format for complex tables.

    Returns:
        A Markdown/HTML string, or "" for blank tables.
    """
    # AC4 (FR-24/FR-25): omit blank tables entirely (checked before complexity decision)
    if _is_blank_table(shape):
        return ""

    if is_complex_table(shape):
        # FR-13 / FR-25 AC6: complex-table fallback
        if fallback_format == "html":
            return _table_to_html(shape)
        if fallback_format == "table-data":
            return _table_to_table_data(shape)
        # default: mermaid
        return table_to_mermaid(shape)

    # GFM pipe table
    rows = shape.rows
    if not rows:
        return ""

    header = rows[0]
    # AC1+AC2: sanitise header cells
    sanitised_header = [_sanitise_cell(c) for c in header]
    table_lines: list[str] = []

    # Header row
    table_lines.append("| " + " | ".join(sanitised_header) + " |")
    # Separator row (one --- per column)
    table_lines.append("| " + " | ".join("---" for _ in header) + " |")
    # Data rows
    for row in rows[1:]:
        # Pad/truncate row to match header column count (AC8: robustness)
        padded = list(row) + [""] * max(0, len(header) - len(row))
        padded = padded[: len(header)]
        # AC1+AC2: sanitise each cell
        sanitised_row = [_sanitise_cell(c) for c in padded]
        table_lines.append("| " + " | ".join(sanitised_row) + " |")

    return "\n".join(table_lines)


def _render_image(
    shape: ImageShapeIR,
    masking: MaskingOptions | None,
    *,
    slide_num: int = 0,
    img_num: int = 0,
) -> str:
    """Render an ImageShapeIR according to M3/M4 slot priority (AC4, AC5, FR-22, §3.4).

    Priority (first match):
    1. description (M4) present -> body paragraph (+ classification label)
    2. description absent, alt_text present -> ![alt_text](...)
    3. both absent -> standard marker ![슬라이드 N 이미지 M] (FR-22)
       3a. classification present -> ![슬라이드 N 이미지 M — {classification}]
    """
    description = shape.description
    alt_text = shape.alt_text
    classification = shape.classification

    if description is not None:
        # AC4: description present -> description text block (no placeholder)
        text = description
        if masking is not None:
            text = mask_text(text, masking)
        if classification is not None:
            return f"{text}\n\n_[{classification} image]_"
        return text

    # AC5/FR-22: description=None
    if alt_text:
        # alt_text present -> ![alt_text](...) (priority over no-VLM marker)
        masked_alt = mask_text(alt_text, masking) if masking is not None else alt_text
        return f"![{masked_alt}](...)"

    # FR-22: no description, no alt_text -> standard positional marker
    if classification is not None:
        return f"![슬라이드 {slide_num} 이미지 {img_num} — {classification}]"
    return f"![슬라이드 {slide_num} 이미지 {img_num}]"


def _render_group(
    shape: GroupShapeIR,
    masking: MaskingOptions | None,
    *,
    slide_num: int = 0,
    img_counter: list[int] | None = None,
    table_fallback: TableFallbackFormat = "mermaid",
) -> str:
    """Render a GroupShapeIR by recursively rendering children in order.

    slide_num and img_counter are propagated to children so that image numbers
    are shared across the whole slide (FR-22).  table_fallback is propagated
    to nested table shapes (FR-25 AC6).
    """
    if img_counter is None:
        img_counter = [0]
    child_parts: list[str] = []
    for child in shape.children:
        # FR-21: skip footer shapes inside groups too
        if isinstance(child, TextShapeIR) and child.is_footer:
            continue
        rendered = _render_shape(
            child,
            masking,
            slide_num=slide_num,
            img_counter=img_counter,
            table_fallback=table_fallback,
        )
        if rendered:
            child_parts.append(rendered)
    return "\n\n".join(child_parts)


def _render_other(shape: OtherShapeIR, masking: MaskingOptions | None) -> str:
    """Render an OtherShapeIR: non-empty fallback_text -> paragraph, else skip."""
    if not shape.fallback_text:
        # empty fallback_text -> silently omit
        return ""
    text = shape.fallback_text
    if masking is not None:
        text = mask_text(text, masking)
    return text


def _suppress_repeated_labels(slide_blocks: list[str]) -> list[str]:
    """Remove duplicate consecutive section labels across slide blocks (AC6 — FR-24).

    A "section label" is defined as a non-heading, non-empty first line of a
    slide block that appears identically in the immediately preceding slide
    block.  Only consecutive duplicates are removed (not all occurrences).

    Pure function. Deterministic (ADR-218). Input list is not mutated.

    Args:
        slide_blocks: Ordered list of per-slide Markdown strings.

    Returns:
        New list with repeated consecutive labels suppressed.
    """
    if not slide_blocks:
        return []

    result: list[str] = [slide_blocks[0]]
    prev_labels: set[str] = _extract_labels(slide_blocks[0])

    for block in slide_blocks[1:]:
        lines = block.splitlines()
        # Collect non-heading content lines from this block
        filtered: list[str] = []
        skip_next_blank = False
        for line in lines:
            # Headings (## ...) are never suppressed
            if line.startswith("#"):
                filtered.append(line)
                skip_next_blank = False
                continue
            # Suppress repeated labels (exact strip match)
            if line.strip() and line.strip() in prev_labels:
                skip_next_blank = True
                continue
            # Optionally skip blank line immediately after a suppressed label
            if skip_next_blank and not line.strip():
                skip_next_blank = False
                continue
            skip_next_blank = False
            filtered.append(line)

        result.append("\n".join(filtered))
        prev_labels = _extract_labels(block)

    return result


def _extract_labels(block: str) -> set[str]:
    """Extract non-heading, non-blank content lines as label candidates.

    Used by _suppress_repeated_labels to build the set of labels seen in
    the previous slide.

    Args:
        block: A single slide Markdown string.

    Returns:
        Set of stripped non-heading, non-empty line texts.
    """
    labels: set[str] = set()
    for line in block.splitlines():
        stripped = line.strip()
        if stripped and not stripped.startswith("#"):
            labels.add(stripped)
    return labels
