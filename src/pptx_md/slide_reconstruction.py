"""Hybrid render+VLM orchestration stage (FR-33, ADR-627).

Public interface:
    DeckRenderer          -- Protocol matching slide_render.render_deck's shape
    reconstruct_slides(
        presentation: PresentationIR,
        reconstructor: SlideReconstructor | None,
        *,
        renderer: DeckRenderer = slide_render.render_deck,
        complexity_fn: Callable[[SlideIR], bool] = complexity.is_visually_complex,
        max_slides: int | None = None,
        force_render: frozenset[int] = frozenset(),
        force_text: frozenset[int] = frozenset(),
        embed: bool = False,
        render_dpi: int = slide_render.RENDER_DPI_DEFAULT,
        max_workers: int = 4,
    ) -> None

Design (ARCH-v020 §3.4, ADR-627/628 -- description_pipeline.py precedent):
    - Fills ``SlideIR.reconstructed_md`` in-place (ADR-210/215 slot-filling
      pattern, extended to whole-slide reconstruction). None (default)
      means "text path" -- the assembler's off-path stays byte-identical.
    - Early return (AC3/AC5, INV-4/INV-5): ``reconstructor is None`` or
      ``not slide_render_available()`` -> immediate return, no renderer/pool
      built, all slides keep ``reconstructed_md=None`` (text path only).
      ``slide_render_available`` is imported by *name* (not via the
      ``slide_render`` module object) so tests can monkeypatch
      ``pptx_md.slide_reconstruction.slide_render_available`` directly to
      exercise the FakeRenderer/FakeReconstructor path deterministically,
      independent of whether real LibreOffice/pypdfium2 are installed in
      the test environment.
    - Routing (AC1/AC7): candidate = (complexity_fn(slide) or index in
      force_render) and index not in force_text. Pure/deterministic when
      complexity_fn is (INV-1).
    - Cap (FR-34 AC2/AC4, ADR-628): when max_slides is not None, candidates
      are ranked by (complexity_score desc, index asc) and only the top
      max_slides survive; the rest fall back to the text path with a
      warning log. max_slides<=0 -> zero renders, zero VLM calls.
    - Render (FR-30): the injected renderer is called exactly once with the
      final selected index list (soffice-cost amortisation, ADR-623).
    - Cache (FR-34 AC3, ADR-609/628 extension): selected indices are grouped
      by sha256(png) *before* any reconstruct() job is submitted, so
      identical renders are reconstructed exactly once.
    - Concurrency (ADR-608/610 precedent): one ThreadPoolExecutor job per
      unique hash, max_workers clamped to [1, unique_count]. Results are
      collected into a hash->text dict via as_completed (order-independent)
      then written to every slide sharing that hash in *groups* insertion
      order -- deterministic regardless of completion order.
    - Isolation (AC5/INV-4): a render failure (png is None) or a
      reconstruct() failure/exception for one hash leaves every slide in
      that group at reconstructed_md=None (text-path fallback); no
      exception ever propagates out of this function, and other slides are
      unaffected.
    - Embed (FR-34 AC1, ADR-628): when embed=True, a base64 data-URI image
      line is prepended to the reconstructed Markdown fragment for each
      individual slide (per-slide PNG, even when the hash was shared).
    - PII (NFR-06): only counts/indices/hash-prefixes/exception type names
      are logged -- never PNG bytes, prompt text, or reconstructed Markdown.
    - No VLM SDK import anywhere in this module (INV-5): the reconstructor
      is always injected by the caller (DIP, ImageDescriber/description_
      pipeline.py precedent). ``pypdfium2`` is never imported here either --
      it stays isolated inside slide_render.py (lazy import).
"""

from __future__ import annotations

import base64
import hashlib
import logging
import os
from collections.abc import Callable, Sequence
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from typing import Protocol

from pptx_md.assembler import assemble_slide
from pptx_md.complexity import complexity_score, is_visually_complex
from pptx_md.ir import PresentationIR, SlideIR
from pptx_md.slide_describer import SlideContext, SlideReconstructor
from pptx_md.slide_render import RENDER_DPI_DEFAULT, render_deck, slide_render_available

__all__ = ["DeckRenderer", "reconstruct_slides"]

_logger = logging.getLogger("pptx_md.slide_reconstruction")


class DeckRenderer(Protocol):
    """Structural interface matching :func:`slide_render.render_deck` (FR-30).

    Lets tests inject a ``FakeRenderer`` (index -> fixed PNG bytes) without
    depending on real LibreOffice/pypdfium2 (ARCH-v020 §7 test strategy).
    """

    def __call__(
        self,
        pptx_path: str | os.PathLike[str],
        indices: Sequence[int],
        *,
        dpi: int = ...,
    ) -> dict[int, bytes | None]: ...


def _embed_prefix(png: bytes) -> str:
    """Return a base64 data-URI image line for *png* (FR-34 AC1, ADR-628)."""
    b64 = base64.standard_b64encode(png).decode("ascii")
    return f"![slide render](data:image/png;base64,{b64})\n\n"


def reconstruct_slides(
    presentation: PresentationIR,
    reconstructor: SlideReconstructor | None,
    *,
    renderer: DeckRenderer = render_deck,
    complexity_fn: Callable[[SlideIR], bool] = is_visually_complex,
    max_slides: int | None = None,
    force_render: frozenset[int] = frozenset(),
    force_text: frozenset[int] = frozenset(),
    embed: bool = False,
    render_dpi: int = RENDER_DPI_DEFAULT,
    max_workers: int = 4,
) -> None:
    """Fill ``SlideIR.reconstructed_md`` for visually-complex slides (FR-33).

    Mutates *presentation* in place (ADR-210/215/627). Never raises: any
    render/reconstruct failure is isolated to the affected slide(s), which
    keep ``reconstructed_md=None`` (text-path fallback, INV-4).

    1. Early return (AC3/AC5): *reconstructor* is ``None`` or
       ``slide_render_available()`` is ``False`` -> immediate no-op. Safe in
       core-only environments (INV-5) and when the caller never opts in.
    2. Routing (AC1/AC7): a slide is a render+VLM candidate iff
       ``complexity_fn(slide)`` is ``True`` or its index is in
       *force_render*, **unless** its index is in *force_text* (which always
       wins). Pure/deterministic when *complexity_fn* is (INV-1).
    3. Cap (FR-34 AC2/AC4): when *max_slides* is not ``None``, candidates are
       ranked by ``(complexity_score desc, index asc)`` and only the first
       *max_slides* survive; the rest fall back to text with a warning log.
       ``max_slides <= 0`` -> zero renders, zero VLM calls.
    4. Render (FR-30): *renderer* is called exactly once with every selected
       index (soffice-cost amortisation).
    5. Cache (FR-34 AC3): selected indices are grouped by ``sha256(png)``
       before any reconstruct() call -- identical renders are reconstructed
       exactly once.
    6. Reconstruct (FR-32): one ``reconstructor.reconstruct(png, "png",
       SlideContext(...))`` call per unique hash, run on a
       ``ThreadPoolExecutor`` bounded by *max_workers* (clamped to
       ``[1, unique_count]``). Completion order does not affect the result
       (ADR-610 precedent).
    7. Embed (FR-34 AC1): when *embed* is ``True``, a base64 data-URI image
       line (this slide's own render) is prepended to its reconstructed
       Markdown fragment.
    8. Fill/isolate (AC5/INV-4): on success, ``slide.reconstructed_md`` is
       set; on any failure (render ``None``, reconstruct exception/empty),
       the slide's slot is left at ``None`` (text-path fallback) and the
       failure never propagates.

    Args:
        presentation: The parsed :class:`~pptx_md.ir.PresentationIR`; mutated
            in place.
        reconstructor: A :class:`~pptx_md.slide_describer.SlideReconstructor`
            provider instance, or ``None`` to skip the whole stage.
        renderer: Deck-rendering callable (defaults to
            :func:`~pptx_md.slide_render.render_deck`); injectable for tests
            (``FakeRenderer``).
        complexity_fn: Routing predicate (defaults to
            :func:`~pptx_md.complexity.is_visually_complex`); injectable for
            tests.
        max_slides: Upper bound on rendered+reconstructed slides per
            document (``None`` = no cap; ``0`` = render nothing).
        force_render: Slide indices that are always routed to the
            render+VLM path regardless of *complexity_fn* (subject to the
            cap and to *force_text*, which wins on conflict).
        force_text: Slide indices that are always routed to the text path
            regardless of *complexity_fn*/*force_render*.
        embed: When ``True``, prepend a base64 PNG data-URI to each
            reconstructed slide's Markdown fragment (default ``False``).
        render_dpi: Target render DPI forwarded to *renderer*.
        max_workers: Upper bound on concurrent ``reconstruct()`` calls
            (default 4); clamped to ``[1, unique_render_count]``.

    Returns:
        ``None`` -- in-place mutation only (ADR-210/627).
    """
    if reconstructor is None:
        _logger.debug("reconstruct_slides: reconstructor=None -> skip (text-only)")
        return
    if not slide_render_available():
        _logger.debug(
            "reconstruct_slides: render unavailable -> skip (text-only, INV-5)"
        )
        return

    slides_by_index: dict[int, SlideIR] = {s.index: s for s in presentation.slides}

    # --- 2. Routing (AC1/AC7) ---
    routed: list[int] = []
    for slide in presentation.slides:
        idx = slide.index
        if idx in force_text:
            continue
        if idx in force_render or complexity_fn(slide):
            routed.append(idx)

    if not routed:
        _logger.debug("reconstruct_slides: no visually-complex slides routed")
        return

    # --- 3. Cap (FR-34 AC2/AC4) ---
    if max_slides is not None:
        if max_slides <= 0:
            _logger.info(
                "reconstruct_slides: max_vlm_slides=%d -> 0 renders, 0 VLM calls",
                max_slides,
            )
            return
        ranked = sorted(
            routed,
            key=lambda i: (-complexity_score(slides_by_index[i]), i),
        )
        selected = ranked[:max_slides]
        overflow = len(routed) - len(selected)
        if overflow > 0:
            _logger.warning(
                "reconstruct_slides: %d candidate(s) exceeded max_vlm_slides=%d"
                " -> text fallback",
                overflow,
                max_slides,
            )
    else:
        selected = list(routed)

    selected_indices = sorted(selected)

    # --- 4. Render (FR-30, exactly one call) ---
    png_map = renderer(presentation.source_path, selected_indices, dpi=render_dpi)

    # --- 5. sha256 cache groups (FR-34 AC3) ---
    groups: dict[str, list[int]] = {}
    render_failed = 0
    for idx in selected_indices:
        png = png_map.get(idx)
        if png is None:
            render_failed += 1
            continue
        key = hashlib.sha256(png).hexdigest()
        groups.setdefault(key, []).append(idx)

    if not groups:
        _logger.info(
            "reconstruct_slides: routed=%d selected=%d render_failed=%d"
            " -> all text fallback",
            len(routed),
            len(selected_indices),
            render_failed,
        )
        return

    unique_count = len(groups)
    workers = max(1, min(max_workers, unique_count))

    def _reconstruct_one(key: str, rep_idx: int) -> str | None:
        """Reconstruct one unique render (group representative). Never raises."""
        slide = slides_by_index[rep_idx]
        png = png_map[rep_idx]
        assert png is not None  # guaranteed by the grouping loop above
        text_outline = assemble_slide(slide)
        context = SlideContext(
            slide_index=rep_idx, title=slide.title, text_outline=text_outline
        )
        try:
            text = reconstructor.reconstruct(png, "png", context)
        except Exception as exc:
            _logger.warning(
                "reconstruct_slides: reconstruct failed key=%s rep_idx=%d"
                " error_type=%s",
                key[:12],
                rep_idx,
                type(exc).__name__,
            )
            return None
        if not text:
            _logger.warning(
                "reconstruct_slides: empty reconstruction key=%s rep_idx=%d",
                key[:12],
                rep_idx,
            )
            return None
        return text

    results: dict[str, str] = {}
    with ThreadPoolExecutor(max_workers=workers) as pool:
        future_to_key: dict[Future[str | None], str] = {
            pool.submit(_reconstruct_one, key, idxs[0]): key
            for key, idxs in groups.items()
        }
        for fut in as_completed(future_to_key):
            key = future_to_key[fut]
            try:
                text = fut.result()
            except Exception as exc:  # defensive: unexpected job exception
                _logger.warning(
                    "reconstruct_slides: job failed key=%s error_type=%s",
                    key[:12],
                    type(exc).__name__,
                )
                continue
            if text is not None:
                results[key] = text

    # --- 6/7/8: reduce (main thread, order-independent -> deterministic) ---
    reconstruct_failed = 0
    success_count = 0
    for key, idxs in groups.items():
        text = results.get(key)
        if text is None:
            reconstruct_failed += len(idxs)
            continue
        success_count += len(idxs)
        for idx in idxs:
            md = text
            if embed:
                png = png_map[idx]
                if png is not None:
                    md = _embed_prefix(png) + md
            slides_by_index[idx].reconstructed_md = md

    fallback_count = render_failed + reconstruct_failed
    _logger.info(
        "reconstruct_slides: routed=%d selected=%d render_calls=1"
        " reconstruct_calls=%d cache_hits=%d success=%d fallback=%d",
        len(routed),
        len(selected_indices),
        unique_count,
        len(selected_indices) - render_failed - unique_count,
        success_count,
        fallback_count,
    )
