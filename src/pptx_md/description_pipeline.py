"""Description enrichment pipeline — fills ImageShapeIR.description in-place.

FR-10, M4.

Public interface:
    enrich_descriptions(pres: PresentationIR, describer: ImageDescriber | None) -> None

Design:
    - Mutates the IR in-place (ADR-210/215): shape.description = text.
    - _shape_hint maps ImageClass -> prompt hint string (4 classes + None, FR-10).
    - Per-image isolation: one describe() failure does not affect others (ADR-214/204).
    - EMF/WMF shapes are routed through vector.convert_vector_to_png (ADR-216);
      on failure/skip the description remains None.
    - describer=None -> immediate return (AC3, NFR-08 gate).
    - NO VLM SDK imports anywhere in this module (NFR-08, ADR-211).
"""

from __future__ import annotations

import logging

from pptx_md import vector
from pptx_md.describer import ImageDescriber
from pptx_md.ir import ImageClass, ImageShapeIR, PresentationIR, SlideIR, iter_shapes

__all__ = ["enrich_descriptions"]

_logger = logging.getLogger("pptx_md.description_pipeline")

# ---------------------------------------------------------------------------
# shape_hint constants — deterministic FR-10 mapping (ARCH-M4 §3.3)
# ---------------------------------------------------------------------------

_HINTS: dict[ImageClass, str] = {
    ImageClass.TEXT: (
        "This image appears to contain primarily text;"
        " transcribe the text content accurately."
    ),
    ImageClass.DIAGRAM: (
        "This image appears to be a diagram or chart;"
        " describe its structure, components, and relationships."
    ),
    ImageClass.PHOTO: (
        "This image appears to be a photograph;"
        " describe the scene and visible objects."
    ),
    ImageClass.LOGO: (
        "This image appears to be a logo or icon; identify it concisely."
    ),
}


def _shape_hint(classification: ImageClass | None) -> str | None:
    """Map an ImageClass to a prompt hint string (FR-10). Deterministic.

    Args:
        classification: The ``ImageClass`` enum value from M3 classification,
            or ``None`` if the image was not classified.

    Returns:
        A non-empty hint string for a known class, or ``None`` when
        *classification* is ``None`` (no hint — provider uses generic prompt).
    """
    if classification is None:
        return None
    return _HINTS.get(classification)


def enrich_descriptions(
    presentation: PresentationIR,
    describer: ImageDescriber | None,
) -> None:
    """Fill ``ImageShapeIR.description`` for every image in *presentation* (FR-10).

    Mutates the IR in place (ADR-210/215).  For each
    :class:`~pptx_md.ir.ImageShapeIR` found via :func:`~pptx_md.ir.iter_shapes`:

    1. If *describer* is ``None``, return immediately — all descriptions stay
       ``None`` and no SDK is invoked (AC3, NFR-08).
    2. If ``image_ext`` is in :data:`~pptx_md.vector.VECTOR_EXTS` (EMF/WMF),
       attempt :func:`~pptx_md.vector.convert_vector_to_png`.  On skip/failure
       leave ``description=None`` and continue to the next shape (ADR-216).
    3. Build a ``shape_hint`` from ``shape.classification`` via
       :func:`_shape_hint` (FR-10).
    4. Call ``describer.describe(image_bytes, image_ext, shape_hint)`` and
       write the result to ``shape.description`` in-place.
    5. Per-image isolation: any exception is caught, ``description`` stays
       ``None``, a WARNING is logged (meta only — NFR-06), and processing
       continues with the next shape (ADR-214, ADR-204 계승).

    Args:
        presentation: The parsed :class:`~pptx_md.ir.PresentationIR`; mutated
            in place.
        describer: An :class:`~pptx_md.describer.ImageDescriber` provider
            instance, or ``None`` to skip all descriptions.

    Returns:
        ``None`` — in-place mutation only (ADR-210).
    """
    if describer is None:
        _logger.debug("enrich_descriptions: describer=None, skipping all shapes")
        return

    for slide in presentation.slides:
        if not isinstance(slide, SlideIR):
            continue
        for shape in iter_shapes(slide):
            if not isinstance(shape, ImageShapeIR):
                continue
            try:
                _enrich_single(shape, describer)
            except Exception as exc:
                # Per-image isolation (ADR-214/204): log meta only, never PII
                _logger.warning(
                    "enrich_descriptions: failed for shape_id=%d"
                    " hint_class=%r exc_type=%s",
                    shape.shape_id,
                    shape.classification,
                    type(exc).__name__,
                )
                # description stays None


def _enrich_single(shape: ImageShapeIR, describer: ImageDescriber) -> None:
    """Describe one *shape* and write the result to ``shape.description``."""
    ext = shape.image_ext.lower()
    image_bytes = shape.image_bytes
    image_ext = ext

    if ext in vector.VECTOR_EXTS:
        # ADR-216: convert EMF/WMF to PNG for VLM input
        png_bytes = vector.convert_vector_to_png(shape.image_bytes, ext)
        if png_bytes is None:
            # LibreOffice unavailable or conversion failed — graceful skip
            _logger.debug(
                "enrich_descriptions: vector conversion skipped"
                " for shape_id=%d ext=%r -> description=None",
                shape.shape_id,
                ext,
            )
            return
        image_bytes = png_bytes
        image_ext = "png"

    hint = _shape_hint(shape.classification)
    _logger.debug(
        "enrich_descriptions: shape_id=%d image_ext=%r hint_class=%r has_hint=%s",
        shape.shape_id,
        image_ext,
        shape.classification,
        hint is not None,
    )

    # Call provider — any exception propagates to enrich_descriptions try/except
    description = describer.describe(image_bytes, image_ext, hint)

    shape.description = description
    _logger.debug(
        "enrich_descriptions: shape_id=%d -> description filled (len=%d)",
        shape.shape_id,
        len(description) if description else 0,
    )
