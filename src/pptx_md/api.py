"""Public API orchestrator — convert() + ConvertOptions (FR-16, ADR-601/602).

This module is the single entry point for the pptx-md conversion pipeline.
It assembles the M2–M5 pipeline stages into one function and exposes the
ConvertOptions value object that aggregates all per-stage parameters.

Design constraints:
    - ADR-601: Orchestrator lives here; __init__.py re-exports only.
    - ADR-602: ConvertOptions is a frozen dataclass (immutable, thread-safe).
    - ADR-604: validate=True is logging-only; return type is always str.
    - NFR-08: This module MUST NOT import VLM SDKs directly.
              describer is injected by the caller via ConvertOptions.
    - NFR-06: Log metadata only — never raw text or PII fragments.
"""

from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

# Internal pipeline imports — these are all stdlib+internal modules,
# no VLM SDK is imported here (NFR-08, ADR-601).
from pptx_md.assembler import assemble_document
from pptx_md.description_pipeline import enrich_descriptions
from pptx_md.image_pipeline import enrich_images
from pptx_md.parser import parse_presentation
from pptx_md.validator import validate_markdown

if TYPE_CHECKING:
    from pptx_md.describer import ImageDescriber
    from pptx_md.masking import MaskingOptions

_logger = logging.getLogger("pptx_md.api")

__all__ = ["ConvertOptions", "convert"]


# ---------------------------------------------------------------------------
# ConvertOptions — frozen dataclass (ADR-602)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ConvertOptions:
    """Aggregated options for convert() (FR-16, ADR-602).

    Collects the per-stage parameters of the M2–M5 pipeline into one
    immutable value object so callers configure the whole conversion in a
    single place.

    Attributes:
        describer: VLM image-description provider (ImageDescriber Protocol).
                   None (default) = skip VLM enrichment (NFR-08).
        masking:   PII masking options (MaskingOptions).
                   None (default) = no masking (FR-15 opt-in).
        validate:  When True, validate_markdown() is called after assembly
                   and results are emitted as log warnings.  Return type
                   is always str regardless of this flag (ADR-604).
        describe_max_workers: Upper bound on concurrent describe() calls
                   (FR-27 AC7, ADR-608/613). Default 4; clamped internally
                   to [1, unique_image_count]. Has no effect when
                   describer is None.
        diagram_mermaid: When True, DIAGRAM images are described with a
                   Mermaid-flowchart-requesting hint and the response is
                   rendered as a fenced flowchart block when structured
                   (FR-27 AC4/AC5, ADR-611/613). Default False preserves
                   the pre-M12 plain-text description behaviour.
        heading_hierarchy: When True, decompose a '>'-separated slide
                   title into a "##"/"###"/"####"... heading chain
                   (FR-28 AC1/AC2, ADR-614/615). Default False preserves
                   the single "## {title}" heading (byte-identical, AC3).
        emit_toc:  When True, slides whose rendered body is empty besides
                   the heading are converted to a heading-only TOC block
                   or dropped entirely (FR-28 AC4, ADR-614/616). Default
                   False preserves existing behaviour.
        emit_frontmatter: When True, attaches document-level YAML
                   frontmatter and per-slide HTML-comment metadata
                   (FR-28 AC5-AC7, ADR-614/617). Default False.
        include_notes: When True, attaches a "> notes" blockquote per
                   slide from SlideIR.notes (FR-28 AC8/AC9, ADR-614/618).
                   Default False.
    """

    describer: ImageDescriber | None = None
    masking: MaskingOptions | None = None
    validate: bool = False
    describe_max_workers: int = 4
    diagram_mermaid: bool = False
    heading_hierarchy: bool = False
    emit_toc: bool = False
    emit_frontmatter: bool = False
    include_notes: bool = False


# ---------------------------------------------------------------------------
# convert() — 5-stage pipeline orchestrator (ADR-601)
# ---------------------------------------------------------------------------


def convert(
    source: str | os.PathLike[str],
    *,
    options: ConvertOptions | None = None,
) -> str:
    """Convert a PPTX file into a single Markdown document (FR-16).

    Orchestrates the M2–M5 pipeline:
        parse -> enrich_images -> enrich_descriptions -> assemble (-> validate?).

    Args:
        source:  Path to a .pptx file (str or pathlib.Path).
        options: Optional ConvertOptions; None means defaults
                 (no VLM, no masking, no validation logging).

    Returns:
        The assembled Markdown document as a str.

    Raises:
        ParseError: If the file is missing or not a valid PPTX (file-level
                    fail-fast — M2 ADR-205).  Per-shape/per-image failures
                    are isolated and never raise (M3/M4/M5 graceful policy).
    """
    opts = options if options is not None else ConvertOptions()

    source_path = Path(source)
    _logger.debug("convert: start source=%s", source_path.name)

    # --- Stage 1: parse (ParseError propagates on file-level failure) ---
    pres = parse_presentation(source_path)
    _logger.debug("convert: parsed slides=%d", len(pres.slides))

    # --- Stage 2: enrich images (classification, always runs) ---
    enrich_images(pres)

    # --- Stage 3: enrich descriptions (opt-in via describer, NFR-08) ---
    enrich_descriptions(
        pres,
        opts.describer,
        max_workers=opts.describe_max_workers,
        diagram_mermaid=opts.diagram_mermaid,
    )

    # --- Stage 4: assemble Markdown (masking + FR-28 structured opt-in) ---
    md = assemble_document(
        pres,
        masking=opts.masking,
        heading_hierarchy=opts.heading_hierarchy,
        emit_toc=opts.emit_toc,
        emit_frontmatter=opts.emit_frontmatter,
        include_notes=opts.include_notes,
    )

    # --- Stage 5: validate (opt-in, logging-only — ADR-604) ---
    if opts.validate:
        result = validate_markdown(md)
        if not result.valid:
            _logger.warning(
                "validate_markdown: invalid document (warnings=%d)",
                len(result.warnings),
            )
        elif result.warnings:
            _logger.info(
                "validate_markdown: %d warning(s)",
                len(result.warnings),
            )

    _logger.debug("convert: done md_len=%d", len(md))
    return md
