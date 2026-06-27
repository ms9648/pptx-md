"""pptx-md — Convert PPTX files into LLM-friendly Markdown.

Public API re-exports go here as modules land in later milestones, e.g.:

    # M6 / FR-16:
    # from pptx_md.api import convert, ConvertOptions
    # __all__ = ["convert", "ConvertOptions"]

Until then this module only guarantees that ``import pptx_md`` succeeds
(FR-01 AC2). Internal submodules must not be exposed directly (skill §1).
"""

__all__: list[str] = []
