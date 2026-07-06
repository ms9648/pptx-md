"""Markdown assembler — converts a SlideIR / PresentationIR into Markdown (FR-11/12).

Public interface:
    assemble_slide(slide: SlideIR, *, masking: MaskingOptions | None = None) -> str
    assemble_document(
        presentation: PresentationIR,
        *,
        masking: MaskingOptions | None = None,
        suppress_repeated_labels: bool = False,
        table_fallback: TableFallbackFormat = "mermaid",
        heading_hierarchy: bool = False,
        emit_toc: bool = False,
        emit_frontmatter: bool = False,
        include_notes: bool = False,
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

FR-28 structured output (AC1–AC4, ARCH-Wave3 §3.1/§3.2, issue #77 — opt-in,
all default False so existing behaviour/output is byte-identical, AC3):
    - heading_hierarchy: decompose a '>'-separated slide title into a
      '##'/'###'/'####'... heading chain (heading.render_heading_lines).
    - emit_toc: convert slides whose rendered body is empty besides the
      heading (heading.is_body_empty, D-4) into a TOC-style heading-only
      block (real title) or drop the block entirely (fallback "## Slide N"
      title), so that no empty "## Slide N" block is ever emitted (AC4).
    - emit_frontmatter / include_notes: seam calls into metadata.py /
      notes.py (owned by issues #78/#79).  Both modules are currently
      minimal no-op stubs, so these two options have no visible effect
      yet — the seam only attaches non-empty output (§3.3/§3.6).
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from typing import Literal

from pptx_md.heading import is_body_empty, render_heading_lines
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
from pptx_md.metadata import (
    build_frontmatter,
    build_slide_comment,
    derive_has_diagram,
    derive_section,
)
from pptx_md.notes import render_notes_block

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
# FR-28 structured render result (ARCH-Wave3 §3.1, ADR-615/616/621)
# ---------------------------------------------------------------------------


@dataclass
class _RenderedSlide:
    """Structured render result for one slide, prior to block assembly.

    Produced by `_render_slide_structured` and consumed by both
    `assemble_slide` (off-path, byte-identical) and `assemble_document`
    (TOC/frontmatter/notes seams, ARCH-Wave3 §3.1).

    Attributes:
        heading_lines: Rendered heading line(s) — a single "## {title}"
            (or "## Slide N" fallback) when heading_hierarchy=False/no
            path separator, or multiple "##"/"###"/... lines when
            decomposed (FR-28 AC1/AC2).
        body_parts: Rendered body parts, heading excluded, in the same
            order as the pre-Wave-3 `_render_slide` parts[1:].
        title_is_fallback: True when no real title was available and the
            positional "## Slide N" fallback heading was used (ARCH-Wave3
            §3.2 — fallback-empty slides are dropped under emit_toc).
        slide: The source SlideIR (read-only; used by metadata/notes seams
            for index/notes/shapes derivation).
    """

    heading_lines: list[str]
    body_parts: list[str]
    title_is_fallback: bool
    slide: SlideIR


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
    # heading_hierarchy is always False here: assemble_slide's public
    # signature is unchanged by FR-28 (only assemble_document gained the
    # 4 new keyword-only options) — render_heading_lines(hierarchy=False)
    # produces the exact pre-Wave-3 "## {title}" line, so this stays
    # byte-identical to the previous single-pass implementation.
    rendered = _render_slide_structured(
        slide, masking, table_fallback=table_fallback, heading_hierarchy=False
    )
    return "\n\n".join(rendered.heading_lines + rendered.body_parts)


def assemble_document(
    presentation: PresentationIR,
    *,
    masking: MaskingOptions | None = None,
    suppress_repeated_labels: bool = False,
    table_fallback: TableFallbackFormat = "mermaid",
    heading_hierarchy: bool = False,
    emit_toc: bool = False,
    emit_frontmatter: bool = False,
    include_notes: bool = False,
) -> str:
    """Map-Reduce assembly of the whole presentation into one document (FR-12).

    Map: assemble_slide per slide (sequential, ADR-220).
    Reduce: join with a slide separator, in IR list order.
    Deterministic and order-preserving (ADR-220).

    FR-28 opt-in structured output (ARCH-Wave3 §3.1/§3.2/§3.3/§3.6, issue
    #77 — all four options default to False, which preserves the exact
    pre-Wave-3 output byte-for-byte, AC3):

    Args:
        presentation: A PresentationIR instance (read-only).
        masking: Optional MaskingOptions instance (FR-15).
        suppress_repeated_labels: When True, deduplicate consecutive identical
            section labels across slide blocks (AC6 — FR-24).  Default False
            preserves existing behaviour.
        table_fallback: Fallback format for complex/merged-cell tables (FR-25 AC6).
            One of "mermaid" (default), "table-data", or "html".
        heading_hierarchy: When True, decompose a '>'-separated slide title
            into a "##"/"###"/"####"... heading chain (FR-28 AC1/AC2).
            Default False preserves the single "## {title}" heading.
        emit_toc: When True, slides whose rendered body is empty besides
            the heading (D-4) are converted to a heading-only TOC block
            (real title) or dropped entirely (fallback "## Slide N"
            title) so that no empty "## Slide N" block is ever emitted
            (FR-28 AC4). Default False preserves existing behaviour.
        emit_frontmatter: When True, attaches a document-level YAML
            frontmatter block and per-slide HTML-comment metadata via the
            metadata.py seam (FR-28 AC5-AC7, owned by issue #78). Default
            False; currently a no-op until #78 fills in metadata.py.
        include_notes: When True, attaches a "> notes" blockquote per
            slide via the notes.py seam (FR-28 AC8/AC9, owned by issue
            #79). Default False; currently a no-op until #79 fills in
            notes.py.

    Returns:
        A single Markdown document string.
    """
    # sort slides by index (ascending) regardless of IR list order (ADR-220)
    sorted_slides = sorted(presentation.slides, key=lambda s: s.index)

    slide_blocks: list[str] = []
    for slide in sorted_slides:
        try:
            rendered = _render_slide_structured(
                slide,
                masking,
                table_fallback=table_fallback,
                heading_hierarchy=heading_hierarchy,
            )
        except Exception as exc:  # noqa: BLE001 — slide-level isolation
            _logger.warning(
                "assemble_document: slide index=%d failed, skipping. "
                "shape_count=%d error_type=%s",
                slide.index,
                len(slide.shapes),
                type(exc).__name__,
            )
            # Preserve the legacy fallback string exactly (bypasses the
            # TOC/frontmatter/notes seams — matches pre-Wave-3 behaviour).
            slide_blocks.append(f"{_SLIDE_HEADING_PREFIX} Slide {slide.index + 1}")
            continue

        block_parts = _build_slide_block_parts(rendered, emit_toc=emit_toc)
        if block_parts is None:
            # ARCH-Wave3 §3.2: emit_toc dropped this fallback-empty slide —
            # no block is emitted at all (FR-28 AC4: 0 empty blocks).
            continue

        if include_notes:
            notes_block = render_notes_block(rendered.slide.notes)
            if notes_block:
                block_parts.append(notes_block)

        block = "\n\n".join(block_parts)

        if emit_frontmatter:
            section = derive_section(rendered.heading_lines)
            has_diagram = derive_has_diagram(rendered.slide)
            comment = build_slide_comment(
                rendered.slide.index + 1, section, has_diagram
            )
            if comment:
                block = f"{comment}\n{block}"

        slide_blocks.append(block)

    if suppress_repeated_labels:
        slide_blocks = _suppress_repeated_labels(slide_blocks)

    document = _SLIDE_SEPARATOR.join(slide_blocks)

    if emit_frontmatter:
        frontmatter = build_frontmatter(presentation)
        if frontmatter:
            document = f"{frontmatter}\n\n{document}"

    return document


def _build_slide_block_parts(
    rendered: _RenderedSlide, *, emit_toc: bool
) -> list[str] | None:
    """Decide the emitted block content for one rendered slide (FR-28 AC4).

    ARCH-Wave3 §3.2 (D-4 판정, ADR-616):
        - body not empty -> normal block (heading + body), unchanged.
        - emit_toc & body empty & real title -> heading-only block (TOC
          entry / section divider).
        - emit_toc & body empty & fallback title ("## Slide N") -> drop
          the block entirely (returns None) so that 0 empty "## Slide N"
          blocks are ever emitted.
        - emit_toc=False -> normal block, unchanged (backward-compatible).

    Args:
        rendered: The structured render result for one slide.
        emit_toc: FR-28 emit_toc option value.

    Returns:
        The list of Markdown parts to join for this slide's block, or
        None when the block should be dropped entirely.
    """
    if not emit_toc or not is_body_empty(rendered.body_parts):
        return list(rendered.heading_lines) + list(rendered.body_parts)
    if rendered.title_is_fallback:
        return None
    return list(rendered.heading_lines)


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


def _render_slide_structured(
    slide: SlideIR,
    masking: MaskingOptions | None,
    *,
    table_fallback: TableFallbackFormat = "mermaid",
    heading_hierarchy: bool = False,
) -> _RenderedSlide:
    """Render a single SlideIR into a structured `_RenderedSlide` (ARCH-Wave3 §3.1).

    Splits the pre-Wave-3 `_render_slide` into a heading-line list and a
    body-part list so that `assemble_document` can apply the FR-28
    TOC/frontmatter/notes seams without re-parsing rendered Markdown.
    `assemble_slide` (off-path) re-joins heading_lines + body_parts with
    "\\n\\n", which is byte-identical to the pre-Wave-3 single-pass output
    when heading_hierarchy=False (AC3).

    Args:
        slide: A SlideIR instance (read-only).
        masking: Optional masking config (FR-15).
        table_fallback: Fallback format for complex tables (FR-25 AC6).
        heading_hierarchy: Decompose a '>'-separated title into multiple
            heading lines (FR-28 AC1/AC2, ADR-615). Default False.

    Returns:
        A `_RenderedSlide` with heading_lines, body_parts,
        title_is_fallback, and the source slide (read-only).
    """
    slide_num = slide.index + 1

    # FR-20: Heading — use slide.title first; fall back to first is_title shape;
    # last resort: ## Slide N
    # AC1 (FR-24): when slide.title is set, also skip any is_title=True shape
    # whose text matches slide.title (strip comparison) to prevent duplicate output.
    # NOTE: title shape search uses original IR order (heading search is not
    #       position-dependent; we want the designated title placeholder).
    title_shape_used: TextShapeIR | None = None
    heading_lines: list[str] = []
    title_is_fallback = False
    if slide.title:
        heading_lines = render_heading_lines(slide.title, hierarchy=heading_hierarchy)
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
                    heading_lines = render_heading_lines(
                        heading_text, hierarchy=heading_hierarchy
                    )
                    title_shape_used = shape
                    break
        if title_shape_used is None:
            # Fallback: no title available. Not passed through
            # render_heading_lines (ARCH-Wave3 §3.1: position fallback is
            # excluded from '>' decomposition — it never contains one).
            heading_lines = [f"{_SLIDE_HEADING_PREFIX} Slide {slide_num}"]
            title_is_fallback = True

    # FR-22: image counter (mutable list used as a reference for group recursion)
    img_counter: list[int] = [0]

    # FR-23: sort shapes by reading order (top row first, then left-to-right).
    # sorted() is stable: equal-key shapes keep their original IR order (AC6/AC7).
    # IR is read-only — we create a new list, never mutate slide.shapes (ADR-218).
    ordered_shapes: list[ShapeIR] = sorted(slide.shapes, key=_reading_order_key)

    body_parts: list[str] = []
    for shape in ordered_shapes:
        # FR-20 / AC1: skip the shape already used as heading
        if shape is title_shape_used:
            continue
        # FR-21: skip footer/slide-number/date placeholders
        if isinstance(shape, TextShapeIR) and shape.is_footer:
            continue
        rendered_shape = _render_shape(
            shape,
            masking,
            slide_num=slide_num,
            img_counter=img_counter,
            table_fallback=table_fallback,
        )
        if rendered_shape:
            body_parts.append(rendered_shape)

    return _RenderedSlide(
        heading_lines=heading_lines,
        body_parts=body_parts,
        title_is_fallback=title_is_fallback,
        slide=slide,
    )


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
