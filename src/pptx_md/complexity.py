"""Visual complexity detection over parsed slide IR (FR-31, ADR-624).

Public interface:
    complexity_score(slide) -> int
    is_visually_complex(slide) -> bool

Design (ARCH-v020 §3.2):
    - Pure, deterministic: same SlideIR -> same score/bool, always (INV-1).
    - VLM/rendering is NEVER invoked here; only pptx_md.ir types are consumed
      (no python-pptx, no Pillow, no VLM SDK imports — INV-5 isolation).
    - Score = weighted sum of IR-derived signals; ``is_visually_complex`` is a
      threshold on that score. Weights/threshold are named module constants
      (⟨DEC-2⟩, classifier.py precedent) so they can be recalibrated without
      changing rule structure.
    - Recall-biased (DEC-3/5 quality-first policy): several signals are sized
      so that a single strong indicator (multi-column, connector flow, or a
      chart/SmartArt) alone reaches ``COMPLEXITY_THRESHOLD`` — boundary cases
      resolve to "complex" rather than being missed (over-detection is
      acceptable; cost is capped downstream by FR-34 ``max_vlm_slides``).
"""

from __future__ import annotations

from collections.abc import Sequence

from pptx_md.ir import (
    GroupShapeIR,
    OtherShapeIR,
    ShapeIR,
    SlideIR,
    TableShapeIR,
    TextShapeIR,
    iter_shapes,
)

__all__ = ["complexity_score", "is_visually_complex"]

# ---------------------------------------------------------------------------
# ⟨DEC-2⟩ named constants (ARCH-v020 §3.2, ADR-624).
# Calibration target: S5 (multi-column mapping) / S7 (connector flow) / S12
# (KPI+chart) -> True; pure title+linear-bullet text and empty slides -> False
# (FR-31 AC2~AC5). Tunable without changing scoring structure.
# ---------------------------------------------------------------------------

COMPLEXITY_THRESHOLD: int = 4  # score >= this -> visually complex (AC1, AC6)

# Overall shape-count pressure: many shapes on one slide (beyond a single
# text frame + title) is a weak-but-real signal of layout complexity.
SHAPE_COUNT_THRESHOLD: int = 8  # total flattened shape count above which "busy"
W_SHAPE_COUNT_OVER: int = 1

# AC2 — multi-column: 3+ text shapes aligned in a similar top band ("row"),
# each at a distinct left position (i.e. side-by-side columns, not a stack).
MULTI_COLUMN_MIN_SHAPES: int = 3
W_MULTI_COLUMN: int = 4  # alone reaches COMPLEXITY_THRESHOLD

# AC3 — connector/arrow flow: OtherShapeIR shapes whose mso_shape_type marks
# them as line/connector shapes (python-pptx reports MSO_SHAPE_TYPE.LINE for
# any Connector). ⟨DEC-2: 임계 N⟩ = CONNECTOR_COUNT_THRESHOLD.
CONNECTOR_COUNT_THRESHOLD: int = 2  # N — 2+ connectors signals a flow/diagram
CONNECTOR_MSO_SHAPE_TYPES: frozenset[str] = frozenset({"LINE", "CONNECTOR"})
W_CONNECTOR: int = 4  # alone reaches COMPLEXITY_THRESHOLD

# AC3 — chart/SmartArt GraphicFrame (FR-26 derived mso_shape_type labels).
# A single chart/SmartArt (e.g. S12 KPI donut) is already a strong signal.
CHART_SMARTART_MSO_SHAPE_TYPES: frozenset[str] = frozenset({"CHART", "DIAGRAM"})
W_CHART_SMARTART: int = 4  # alone reaches COMPLEXITY_THRESHOLD

# Structural signals (weaker alone; combine with others to cross threshold).
W_TABLE: int = 2  # any TableShapeIR present
W_GROUP: int = 2  # any GroupShapeIR present

# Row/column clustering tolerances (EMU). 914400 EMU == 1 inch.
_ROW_TOLERANCE_EMU: int = 250_000  # ~0.27in: top delta -> same row/band
_COLUMN_LEFT_TOLERANCE_EMU: int = 100_000  # ~0.11in: min left delta -> distinct column


def complexity_score(slide: SlideIR) -> int:
    """Weighted sum of IR-derived visual-complexity signals (ADR-624).

    Pure and deterministic: depends only on the fields of *slide* (and the
    shapes reachable from it via :func:`iter_shapes`). Never calls VLM or
    rendering code (AC1). Same ``SlideIR`` value always yields the same
    ``int`` (AC5/INV-1) regardless of how many times it is invoked.

    Args:
        slide: Parsed slide IR to score.

    Returns:
        Non-negative integer score; compare against
        :data:`COMPLEXITY_THRESHOLD` via :func:`is_visually_complex`.
    """
    shapes = list(iter_shapes(slide))

    score = 0
    if len(shapes) > SHAPE_COUNT_THRESHOLD:
        score += W_SHAPE_COUNT_OVER
    if _has_multi_column(shapes):
        score += W_MULTI_COLUMN
    if _connector_count(shapes) >= CONNECTOR_COUNT_THRESHOLD:
        score += W_CONNECTOR
    if _has_chart_or_smartart(shapes):
        score += W_CHART_SMARTART
    if any(isinstance(s, TableShapeIR) for s in shapes):
        score += W_TABLE
    if any(isinstance(s, GroupShapeIR) for s in shapes):
        score += W_GROUP
    return score


def is_visually_complex(slide: SlideIR) -> bool:
    """Return True when *slide* should be routed to the render+VLM path.

    ``complexity_score(slide) >= COMPLEXITY_THRESHOLD`` (AC1). Pure,
    deterministic, VLM/render-free (INV-1, INV-5). Recall-biased: ties at
    the threshold resolve to True (AC6).

    Args:
        slide: Parsed slide IR to classify.

    Returns:
        True if the slide is visually complex (render+VLM candidate),
        False if the current text-extraction path is sufficient.
    """
    return complexity_score(slide) >= COMPLEXITY_THRESHOLD


# ---------------------------------------------------------------------------
# Internal signal helpers
# ---------------------------------------------------------------------------


def _connector_count(shapes: Sequence[ShapeIR]) -> int:
    """Count OtherShapeIR shapes tagged as connector/line types (AC3)."""
    return sum(
        1
        for s in shapes
        if isinstance(s, OtherShapeIR) and s.mso_shape_type in CONNECTOR_MSO_SHAPE_TYPES
    )


def _has_chart_or_smartart(shapes: Sequence[ShapeIR]) -> bool:
    """True if any OtherShapeIR is a chart or SmartArt GraphicFrame (AC3)."""
    return any(
        isinstance(s, OtherShapeIR)
        and s.mso_shape_type in CHART_SMARTART_MSO_SHAPE_TYPES
        for s in shapes
    )


def _has_multi_column(shapes: Sequence[ShapeIR]) -> bool:
    """True if 3+ text shapes form a multi-column band (AC2).

    Text shapes are clustered into row "bands" by their EMU ``top``
    coordinate (within :data:`_ROW_TOLERANCE_EMU` of the band's first
    member). A band qualifies as multi-column when it has at least
    :data:`MULTI_COLUMN_MIN_SHAPES` members at distinct ``left`` positions
    (i.e. laid out side-by-side, not stacked at the same horizontal offset).
    """
    text_shapes = [s for s in shapes if isinstance(s, TextShapeIR)]
    if len(text_shapes) < MULTI_COLUMN_MIN_SHAPES:
        return False

    ordered = sorted(text_shapes, key=lambda s: s.top)
    band: list[TextShapeIR] = []
    for shape in ordered:
        if band and shape.top - band[0].top > _ROW_TOLERANCE_EMU:
            if _band_is_multi_column(band):
                return True
            band = []
        band.append(shape)
    return _band_is_multi_column(band)


def _band_is_multi_column(band: Sequence[TextShapeIR]) -> bool:
    """True if *band* (a single row-aligned cluster) has 3+ distinct columns."""
    if len(band) < MULTI_COLUMN_MIN_SHAPES:
        return False

    lefts = sorted(s.left for s in band)
    distinct = 1
    last = lefts[0]
    for left in lefts[1:]:
        if left - last >= _COLUMN_LEFT_TOLERANCE_EMU:
            distinct += 1
            last = left
    return distinct >= MULTI_COLUMN_MIN_SHAPES
