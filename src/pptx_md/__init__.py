"""pptx-md — Convert PPTX files into LLM-friendly Markdown.

Public API (FR-16, ADR-603):

    from pptx_md import convert, ConvertOptions

    md = convert("deck.pptx")

    # With options:
    from pptx_md import ConvertOptions, MaskingOptions, get_describer
    opts = ConvertOptions(
        describer=get_describer("anthropic", api_key="..."),
        masking=MaskingOptions(enabled=True),
        validate=True,
    )
    md = convert("deck.pptx", options=opts)

Internal submodules (parse_presentation, enrich_images, etc.) are NOT part
of the public API and may change without notice.  Access via
``from pptx_md.parser import parse_presentation`` is possible but outside
SemVer guarantees (ADR-603).
"""

from __future__ import annotations

from pptx_md.api import ConvertOptions, convert
from pptx_md.describer import ImageDescriber
from pptx_md.errors import DescribeError, InstallationError, ParseError, PptxMdError
from pptx_md.masking import MASK_TOKEN, MaskingOptions
from pptx_md.providers import get_describer
from pptx_md.validator import ValidationResult, validate_markdown

__version__: str = "0.1.2"

__all__ = [
    "convert",
    "ConvertOptions",
    "MaskingOptions",
    "MASK_TOKEN",
    "validate_markdown",
    "ValidationResult",
    "ImageDescriber",
    "get_describer",
    "PptxMdError",
    "ParseError",
    "DescribeError",
    "InstallationError",
]
