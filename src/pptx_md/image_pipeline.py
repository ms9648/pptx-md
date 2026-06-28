"""Image enrichment pipeline — fills ImageShapeIR.classification in-place (FR-06, M3).

Public interface:
    enrich_images(presentation: PresentationIR) -> None

Design:
    - Mutates the IR in-place (ADR-210): shape.classification = result.
    - Per-image isolation: one image failure does not affect others (ADR-204).
    - EMF/WMF shapes are routed through vector.convert_vector_to_png (FR-07);
      on failure/skip the classification remains None.
    - Raster shapes go directly to classifier.classify_image.
    - No VLM SDK imports (NFR-08, ADR-211).
"""

from __future__ import annotations

import logging

# FR-07 vector conversion (imported here so image_pipeline owns the glue logic)
from pptx_md import vector
from pptx_md.classifier import classify_image
from pptx_md.ir import ImageShapeIR, PresentationIR, SlideIR, iter_shapes

__all__ = ["enrich_images"]

_logger = logging.getLogger("pptx_md.image_pipeline")


def enrich_images(presentation: PresentationIR) -> None:
    """Fill ``ImageShapeIR.classification`` for every image in *presentation*.

    Mutates the IR in place (ADR-210).  For each
    :class:`~pptx_md.ir.ImageShapeIR` found via :func:`~pptx_md.ir.iter_shapes`:

    1. If ``image_ext`` is in ``vector.VECTOR_EXTS`` (EMF/WMF) and
       :func:`~pptx_md.vector.convert_vector_to_png` succeeds, classify the
       resulting PNG bytes.  On skip/failure the ``classification`` stays ``None``.
    2. Otherwise classify ``image_bytes`` directly.

    Per-image isolation: one image's failure never affects others (ADR-204).

    Args:
        presentation: The parsed :class:`~pptx_md.ir.PresentationIR`; mutated
            in place.
    """
    for slide in presentation.slides:
        if not isinstance(slide, SlideIR):
            continue
        for shape in iter_shapes(slide):
            if not isinstance(shape, ImageShapeIR):
                continue
            try:
                _enrich_single(shape)
            except Exception as exc:
                # Unexpected failure: keep classification=None, log, continue
                _logger.warning(
                    "enrich_images: unexpected error on shape %d (%r): %s",
                    shape.shape_id,
                    type(exc).__name__,
                    exc,
                )


def _enrich_single(shape: ImageShapeIR) -> None:
    """Attempt to classify one *shape* and write the result in-place."""
    ext = shape.image_ext.lower()

    if ext in vector.VECTOR_EXTS:
        # FR-07 path: attempt LibreOffice vector -> PNG conversion first
        png_bytes = vector.convert_vector_to_png(shape.image_bytes, ext)
        if png_bytes is not None:
            shape.classification = classify_image(png_bytes, image_ext="png")
        else:
            # LibreOffice unavailable or conversion failed — graceful skip
            _logger.debug(
                "enrich_images: vector conversion skipped for shape %d (ext=%r)",
                shape.shape_id,
                ext,
            )
            # classification stays None
        return

    # Raster path
    shape.classification = classify_image(shape.image_bytes, image_ext=ext)
