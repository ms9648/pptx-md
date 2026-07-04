"""Description enrichment pipeline — fills ImageShapeIR.description in-place.

FR-10 (M4), extended by FR-27 (M12, issue #68): hash-based dedup caching,
bounded describe() concurrency, and opt-in Mermaid flowchart rendering for
DIAGRAM images.

Public interface:
    enrich_descriptions(
        pres: PresentationIR,
        describer: ImageDescriber | None,
        *,
        max_workers: int = 4,
        diagram_mermaid: bool = False,
    ) -> None

Design:
    - Mutates the IR in-place (ADR-210/215): shape.description = text.
    - _shape_hint maps ImageClass -> prompt hint string (4 classes + None, FR-10).
    - describer=None -> immediate return (AC3, NFR-08 gate). No pool/cache built.
    - Collection order is deterministic: presentation.slides (index order) ->
      iter_shapes() (DFS, ADR-203) -> ImageShapeIR filter (ARCH-M12 §3.1).
    - Hash dedup (AC6, ADR-609): shapes are grouped by
      sha256(shape.image_bytes) *before* any job submission. Each unique hash
      is described exactly once; duplicate bytes (logos, watermarks) share
      one describe() call. Cache lives for the duration of this call only
      (function-local dict) -- no persistence, no raw bytes stored (NFR-06).
    - Concurrency (AC7/AC8, ADR-608/610): one ThreadPoolExecutor job per
      unique hash. max_workers is clamped to [1, unique_count]. Results are
      collected into a hash->text dict via as_completed (order-independent),
      then written to each shape in *groups insertion order* (=first-seen
      order) -- deterministic regardless of completion order (ADR-218
      inherited). Workers never touch the IR directly (no lock needed).
    - EMF/WMF shapes are routed through vector.convert_vector_to_png (ADR-216);
      on failure/skip the description remains None for that hash group.
    - Diagram Mermaid (AC4/AC5, ADR-611/612): when diagram_mermaid=True and
      the representative shape's classification is ImageClass.DIAGRAM, the
      shape_hint is augmented with mermaid.DIAGRAM_HINT_SUFFIX and the
      describer response is post-processed by
      mermaid.render_diagram_mermaid(). A valid fence replaces the
      description; an unstructured response falls back to the raw text
      (zero content loss).
    - Per-image isolation: one describe()/job failure does not affect others
      (ADR-214/204/610). Failed hashes are simply absent from the result
      dict, leaving the corresponding shapes' description=None.
    - NO VLM SDK imports anywhere in this module (NFR-08, ADR-211). New
      imports for M12 are stdlib only: hashlib, concurrent.futures.
"""

from __future__ import annotations

import hashlib
import logging
from concurrent.futures import Future, ThreadPoolExecutor, as_completed

from pptx_md import mermaid, vector
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


def _collect_image_shapes(presentation: PresentationIR) -> list[ImageShapeIR]:
    """Collect all ``ImageShapeIR`` nodes in deterministic order (ARCH-M12 §3.1).

    Order: ``presentation.slides`` (index order) -> ``iter_shapes()`` (DFS,
    ADR-203) -> ``ImageShapeIR`` filter. Each returned item is a distinct
    object (group nesting included).
    """
    shapes: list[ImageShapeIR] = []
    for slide in presentation.slides:
        if not isinstance(slide, SlideIR):
            continue
        for shape in iter_shapes(slide):
            if isinstance(shape, ImageShapeIR):
                shapes.append(shape)
    return shapes


def _group_by_hash(
    shapes: list[ImageShapeIR],
) -> dict[str, list[ImageShapeIR]]:
    """Group *shapes* by ``sha256(image_bytes)`` digest (AC6, ADR-609).

    key = digest hex of the *original* bytes (pre EMF/WMF conversion); value
    = the list of shapes sharing that digest. Insertion order == first-seen
    order (deterministic, Py3.7+ dict guarantee). Shapes with empty
    ``image_bytes`` (``b""``) are excluded from the grouping. The original
    bytes themselves are never logged (NFR-06).
    """
    groups: dict[str, list[ImageShapeIR]] = {}
    for shape in shapes:
        if not shape.image_bytes:
            continue
        key = hashlib.sha256(shape.image_bytes).hexdigest()
        groups.setdefault(key, []).append(shape)
    return groups


def _describe_one(
    key: str,
    shape: ImageShapeIR,
    describer: ImageDescriber,
    diagram_mermaid: bool,
) -> str | None:
    """Describe one unique image (group representative). Never raises.

    Any failure (vector conversion skip, describe() exception) results in
    ``None`` — per-image isolation is enforced here so that a raised
    exception never escapes into a worker thread's Future (AC9, ADR-214).
    """
    ext = shape.image_ext.lower()
    image_bytes = shape.image_bytes
    image_ext = ext

    if ext in vector.VECTOR_EXTS:
        # ADR-216: convert EMF/WMF to PNG for VLM input
        png_bytes = vector.convert_vector_to_png(shape.image_bytes, ext)
        if png_bytes is None:
            _logger.debug(
                "_describe_one: vector conversion skipped key=%s ext=%r"
                " -> description=None",
                key[:12],
                ext,
            )
            return None
        image_bytes = png_bytes
        image_ext = "png"

    want_mermaid = diagram_mermaid and shape.classification is ImageClass.DIAGRAM
    hint = _shape_hint(shape.classification)
    if want_mermaid:
        base = hint or ""
        hint = (base + " " + mermaid.DIAGRAM_HINT_SUFFIX).strip()

    _logger.debug(
        "_describe_one: key=%s shape_id=%d image_ext=%r hint_class=%r"
        " has_hint=%s want_mermaid=%s",
        key[:12],
        shape.shape_id,
        image_ext,
        shape.classification,
        hint is not None,
        want_mermaid,
    )

    try:
        text = describer.describe(image_bytes, image_ext, hint)
    except Exception as exc:
        _logger.warning(
            "_describe_one: describe failed key=%s shape_id=%d exc_type=%s",
            key[:12],
            shape.shape_id,
            type(exc).__name__,
        )
        return None

    if want_mermaid:
        fenced = mermaid.render_diagram_mermaid(text)
        _logger.debug(
            "_describe_one: key=%s mermaid_rendered=%s", key[:12], fenced is not None
        )
        return fenced if fenced is not None else text
    return text


def enrich_descriptions(
    presentation: PresentationIR,
    describer: ImageDescriber | None,
    *,
    max_workers: int = 4,
    diagram_mermaid: bool = False,
) -> None:
    """Fill ``ImageShapeIR.description`` for every image (FR-10/FR-27).

    Mutates the IR in place (ADR-210/215).

    1. If *describer* is ``None``, return immediately — all descriptions stay
       ``None`` and no SDK is invoked, no cache/pool is built (AC3, NFR-08).
    2. Collect all ``ImageShapeIR`` in deterministic order (§_collect_image_shapes).
    3. Group by ``sha256(image_bytes)`` (AC6, ADR-609) — duplicate bytes
       (e.g. repeated logos) are described exactly once.
    4. Run one job per unique hash on a ``ThreadPoolExecutor`` bounded by
       ``max_workers`` (clamped to ``[1, unique_count]``) (AC7, ADR-608).
    5. Collect results into a ``hash -> text`` dict (order-independent via
       ``as_completed``), then write to each shape in *groups* insertion
       order — deterministic regardless of completion order (AC8, ADR-610).
    6. When ``diagram_mermaid`` is enabled and a shape's classification is
       ``ImageClass.DIAGRAM``, the response is post-processed into a Mermaid
       flowchart fence, falling back to the raw text when unstructured
       (AC4/AC5, ADR-611).
    7. Per-image isolation: any job failure leaves the corresponding shapes'
       ``description`` as ``None``; other images are unaffected and no
       exception propagates (AC9, ADR-214/204/610 계승).

    Args:
        presentation: The parsed :class:`~pptx_md.ir.PresentationIR`; mutated
            in place.
        describer: An :class:`~pptx_md.describer.ImageDescriber` provider
            instance, or ``None`` to skip all descriptions.
        max_workers: Upper bound on concurrent ``describe()`` calls
            (default 4). Clamped to ``[1, unique_image_count]``.
        diagram_mermaid: When ``True``, DIAGRAM images are described with a
            Mermaid-flowchart-requesting hint and post-processed into a
            fenced flowchart block (default ``False`` — full backward
            compatibility with pre-M12 behaviour).

    Returns:
        ``None`` — in-place mutation only (ADR-210).
    """
    if describer is None:
        _logger.debug("enrich_descriptions: describer=None, skipping all shapes")
        return

    shapes = _collect_image_shapes(presentation)
    groups = _group_by_hash(shapes)
    if not groups:
        _logger.debug("enrich_descriptions: no image shapes with bytes -> no-op")
        return

    unique_count = len(groups)
    workers = max(1, min(max_workers, unique_count))
    _logger.debug(
        "enrich_descriptions: unique_images=%d total_images=%d workers=%d"
        " diagram_mermaid=%s",
        unique_count,
        len(shapes),
        workers,
        diagram_mermaid,
    )

    results: dict[str, str] = {}
    with ThreadPoolExecutor(max_workers=workers) as pool:
        future_to_key: dict[Future[str | None], str] = {
            pool.submit(_describe_one, key, grp[0], describer, diagram_mermaid): key
            for key, grp in groups.items()
        }
        for fut in as_completed(future_to_key):
            key = future_to_key[fut]
            try:
                text = fut.result()
            except Exception as exc:  # defensive: unexpected job exception (AC9)
                _logger.warning(
                    "enrich_descriptions: job failed key=%s exc_type=%s",
                    key[:12],
                    type(exc).__name__,
                )
                continue
            if text is not None:
                results[key] = text

    # Reduce (main thread, single assembly — order-independent -> deterministic,
    # ADR-610/218).
    cache_hits = len(shapes) - unique_count
    _logger.debug(
        "enrich_descriptions: describe_calls=%d cache_hits=%d",
        unique_count,
        cache_hits,
    )
    for key, grp in groups.items():
        text = results.get(key)
        if text is not None:
            for shape in grp:
                shape.description = text
