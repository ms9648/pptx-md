"""Markdown assembler — converts a SlideIR / PresentationIR into Markdown (FR-11/12).

Public interface:
    assemble_slide(slide: SlideIR, *, masking: MaskingOptions | None = None) -> str
    assemble_document(
        presentation: PresentationIR, *, masking: MaskingOptions | None = None
    ) -> str

Design constraints (ADR-218):
    - Deterministic: same IR -> same Markdown (no randomness, no set/dict iteration).
    - IR is read-only: assembler never mutates input objects.
    - Partial-failure isolation: one shape failure does not abort the slide/document.
    - No VLM/python-pptx/Pillow imports (NFR-08).
    - Uses mermaid.is_complex_table() for FR-13 Mermaid fallback.
    - masking=None means no masking (opt-in, FR-15).
"""

from __future__ import annotations

import logging

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
from pptx_md.mermaid import is_complex_table, table_to_mermaid

__all__ = ["assemble_slide", "assemble_document"]

_logger = logging.getLogger("pptx_md.assembler")

# ---------------------------------------------------------------------------
# Module-level named constants (ADR-218 — determinism, §3.2 heading rule)
# ---------------------------------------------------------------------------
_SLIDE_HEADING_PREFIX: str = "##"
_INDENT_UNIT: str = "  "  # two spaces per indent level
_BULLET_MARKER: str = "- "
_SLIDE_SEPARATOR: str = "\n\n---\n\n"


# ---------------------------------------------------------------------------
# Public functions
# ---------------------------------------------------------------------------


def assemble_slide(
    slide: SlideIR,
    *,
    masking: MaskingOptions | None = None,
) -> str:
    """Render one SlideIR into a deterministic Markdown block (FR-11).

    Read-only: does not mutate the IR.  Same IR -> same Markdown (ADR-218).
    masking=None means no masking (opt-in, FR-15).

    Args:
        slide: A SlideIR instance (read-only).
        masking: Optional MaskingOptions instance (FR-15).

    Returns:
        A Markdown string representing this slide.  Empty/blank slides return
        an empty string or heading-only string without raising (FR-11 AC7).
    """
    return _render_slide(slide, masking)


def assemble_document(
    presentation: PresentationIR,
    *,
    masking: MaskingOptions | None = None,
) -> str:
    """Map-Reduce assembly of the whole presentation into one document (FR-12).

    Map: assemble_slide per slide (sequential, ADR-220).
    Reduce: join with a slide separator, in IR list order.
    Deterministic and order-preserving (ADR-220).

    Args:
        presentation: A PresentationIR instance (read-only).
        masking: Optional MaskingOptions instance (FR-15).

    Returns:
        A single Markdown document string.
    """
    # AC2: sort slides by index (ascending) regardless of IR list order (ADR-220)
    sorted_slides = sorted(presentation.slides, key=lambda s: s.index)

    slide_blocks: list[str] = []
    for slide in sorted_slides:
        try:
            block = _render_slide(slide, masking)
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

    return _SLIDE_SEPARATOR.join(slide_blocks)


# ---------------------------------------------------------------------------
# Private render functions
# ---------------------------------------------------------------------------


def _render_slide(slide: SlideIR, masking: MaskingOptions | None) -> str:
    """Render a single SlideIR to Markdown."""
    parts: list[str] = []
    slide_num = slide.index + 1

    # FR-20: Heading — use slide.title first; fall back to first is_title shape;
    # last resort: ## Slide N
    title_shape_used: TextShapeIR | None = None
    if slide.title:
        parts.append(f"{_SLIDE_HEADING_PREFIX} {slide.title}")
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

    # Shapes in IR order (AC1: shapes serialised in IR appearance order)
    for shape in slide.shapes:
        # FR-20: skip the shape already used as heading
        if shape is title_shape_used:
            continue
        # FR-21: skip footer/slide-number/date placeholders
        if isinstance(shape, TextShapeIR) and shape.is_footer:
            continue
        rendered = _render_shape(
            shape, masking, slide_num=slide_num, img_counter=img_counter
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
) -> str:
    """Dispatch to the appropriate renderer; isolate per-shape failures (ADR-220)."""
    if img_counter is None:
        img_counter = [0]
    try:
        if isinstance(shape, TextShapeIR):
            return _render_text(shape, masking)
        if isinstance(shape, TableShapeIR):
            return _render_table(shape)
        if isinstance(shape, ImageShapeIR):
            img_counter[0] += 1
            return _render_image(
                shape, masking, slide_num=slide_num, img_num=img_counter[0]
            )
        if isinstance(shape, GroupShapeIR):
            return _render_group(
                shape, masking, slide_num=slide_num, img_counter=img_counter
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
    """Render a TextShapeIR to Markdown paragraphs / indented bullets (AC1, AC2)."""
    lines: list[str] = []
    for para in shape.paragraphs:
        text = para.text
        # Apply masking if provided (FR-15)
        if masking is not None:
            text = mask_text(text, masking)

        if para.level > 0:
            # Indented bullet item (AC2: level=1 deeper than level=0)
            indent = _INDENT_UNIT * para.level
            lines.append(f"{indent}{_BULLET_MARKER}{text}")
        else:
            # Top-level paragraph
            lines.append(text)

    return "\n".join(lines)


def _render_table(shape: TableShapeIR) -> str:
    """Render a TableShapeIR as GFM table or Mermaid fallback (AC3, FR-13)."""
    if is_complex_table(shape):
        # FR-13 Mermaid fallback
        return table_to_mermaid(shape)

    # GFM pipe table (AC3: header separator included)
    rows = shape.rows
    if not rows:
        return ""

    header = rows[0]
    table_lines: list[str] = []

    # Header row
    table_lines.append("| " + " | ".join(header) + " |")
    # Separator row (one --- per column)
    table_lines.append("| " + " | ".join("---" for _ in header) + " |")
    # Data rows
    for row in rows[1:]:
        # Pad/truncate row to match header column count
        padded = list(row) + [""] * (len(header) - len(row))
        padded = padded[: len(header)]
        table_lines.append("| " + " | ".join(padded) + " |")

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
) -> str:
    """Render a GroupShapeIR by recursively rendering children in order (AC6).

    slide_num and img_counter are propagated to children so that image numbers
    are shared across the whole slide (FR-22).
    """
    if img_counter is None:
        img_counter = [0]
    child_parts: list[str] = []
    for child in shape.children:
        # FR-21: skip footer shapes inside groups too
        if isinstance(child, TextShapeIR) and child.is_footer:
            continue
        rendered = _render_shape(
            child, masking, slide_num=slide_num, img_counter=img_counter
        )
        if rendered:
            child_parts.append(rendered)
    return "\n\n".join(child_parts)


def _render_other(shape: OtherShapeIR, masking: MaskingOptions | None) -> str:
    """Render an OtherShapeIR: non-empty fallback_text -> paragraph, else skip (AC8)."""
    if not shape.fallback_text:
        # AC8: empty fallback_text -> silently omit
        return ""
    text = shape.fallback_text
    if masking is not None:
        text = mask_text(text, masking)
    return text
