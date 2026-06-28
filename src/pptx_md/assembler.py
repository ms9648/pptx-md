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
    - masking=None means no masking (opt-in, FR-15 — masking.py integration point).
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
    masking: object | None = None,
) -> str:
    """Render one SlideIR into a deterministic Markdown block (FR-11).

    Read-only: does not mutate the IR.  Same IR -> same Markdown (ADR-218).
    masking=None means no masking (opt-in, FR-15).

    Args:
        slide: A SlideIR instance (read-only).
        masking: Optional MaskingOptions instance (FR-15, currently unused stub).

    Returns:
        A Markdown string representing this slide.  Empty/blank slides return
        an empty string or heading-only string without raising (FR-11 AC7).
    """
    return _render_slide(slide, masking)


def assemble_document(
    presentation: PresentationIR,
    *,
    masking: object | None = None,
) -> str:
    """Map-Reduce assembly of the whole presentation into one document (FR-12).

    Map: assemble_slide per slide (sequential, ADR-220).
    Reduce: join with a slide separator, in IR list order.
    Deterministic and order-preserving (ADR-220).

    Args:
        presentation: A PresentationIR instance (read-only).
        masking: Optional MaskingOptions instance (FR-15, currently unused stub).

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


def _render_slide(slide: SlideIR, masking: object | None) -> str:
    """Render a single SlideIR to Markdown."""
    parts: list[str] = []

    # Heading (AC1: first line is heading)
    if slide.title:
        parts.append(f"{_SLIDE_HEADING_PREFIX} {slide.title}")
    else:
        # Blank title: emit heading with slide index+1 (1-based display)
        parts.append(f"{_SLIDE_HEADING_PREFIX} Slide {slide.index + 1}")

    # Shapes in IR order (AC1: shapes serialised in IR appearance order)
    for shape in slide.shapes:
        rendered = _render_shape(shape, masking)
        if rendered:
            parts.append(rendered)

    return "\n\n".join(parts)


def _render_shape(shape: ShapeIR, masking: object | None) -> str:
    """Dispatch to the appropriate renderer; isolate per-shape failures (ADR-220)."""
    try:
        if isinstance(shape, TextShapeIR):
            return _render_text(shape, masking)
        if isinstance(shape, TableShapeIR):
            return _render_table(shape)
        if isinstance(shape, ImageShapeIR):
            return _render_image(shape, masking)
        if isinstance(shape, GroupShapeIR):
            return _render_group(shape, masking)
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


def _render_text(shape: TextShapeIR, masking: object | None) -> str:
    """Render a TextShapeIR to Markdown paragraphs / indented bullets (AC1, AC2)."""
    lines: list[str] = []
    for para in shape.paragraphs:
        text = para.text
        # Apply masking if provided (FR-15 integration — stub for now)
        if masking is not None:
            text = _apply_masking(text, masking)

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


def _render_image(shape: ImageShapeIR, masking: object | None) -> str:
    """Render an ImageShapeIR according to M3/M4 slot priority (AC4, AC5, §3.4).

    Priority (first match):
    1. description (M4) present -> body paragraph (+ classification label)
    2. description absent, alt_text present -> ![alt_text](...)
    3. both absent, classification (M3) present -> _[image: {classification}]_
    4. all absent -> _[image]_
    """
    description = shape.description
    alt_text = shape.alt_text
    classification = shape.classification

    if description is not None:
        # AC4: description present -> description text block (no placeholder)
        text = description
        if masking is not None:
            text = _apply_masking(text, masking)
        if classification is not None:
            return f"{text}\n\n_[{classification} image]_"
        return text

    # AC5: description=None -> alt_text image syntax
    if alt_text:
        if masking is not None:
            masked_alt = _apply_masking(alt_text, masking)
        else:
            masked_alt = alt_text
        return f"![{masked_alt}](...)"

    if classification is not None:
        return f"_[image: {classification}]_"

    return "_[image]_"


def _render_group(shape: GroupShapeIR, masking: object | None) -> str:
    """Render a GroupShapeIR by recursively rendering children in order (AC6)."""
    child_parts: list[str] = []
    for child in shape.children:
        rendered = _render_shape(child, masking)
        if rendered:
            child_parts.append(rendered)
    return "\n\n".join(child_parts)


def _render_other(shape: OtherShapeIR, masking: object | None) -> str:
    """Render an OtherShapeIR: non-empty fallback_text -> paragraph, else skip (AC8)."""
    if not shape.fallback_text:
        # AC8: empty fallback_text -> silently omit
        return ""
    text = shape.fallback_text
    if masking is not None:
        text = _apply_masking(text, masking)
    return text


# ---------------------------------------------------------------------------
# Masking integration stub (FR-15 — masking.py not yet part of this issue)
# ---------------------------------------------------------------------------


def _apply_masking(text: str, masking: object) -> str:
    """Apply masking if the masking object exposes a mask_text interface.

    Forward-compatible stub for FR-15 (#37).  When masking.py is available and
    ``masking.enabled`` is True, delegates to ``masking.mask_text``.
    Returns *text* unchanged if masking is disabled or masking.py is not installed.
    """
    enabled = getattr(masking, "enabled", False)
    if not enabled:
        return text
    # Delegate to masking.mask_text when masking.py is available (FR-15 #37).
    # importlib.import_module avoids a hard dependency that mypy would flag when
    # masking.py does not yet exist in the type-checked source tree.
    import importlib  # noqa: PLC0415

    try:
        mod = importlib.import_module("pptx_md.masking")
        mask_fn = getattr(mod, "mask_text", None)
        if callable(mask_fn):
            return str(mask_fn(text, masking))
    except ImportError:
        pass
    return text
