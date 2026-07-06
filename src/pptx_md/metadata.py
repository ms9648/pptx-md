"""Slide/document metadata contract — MINIMAL STUB (FR-28 AC5-AC7, issue #77).

Ownership (ARCH-Wave3 §2.1/ADR-621, WBS issue B / #78):
    This module is created by issue #77 (W-A) *only* to lock the public
    contract signature that `assembler.py` seams call, so that #78 (W-B)
    can implement the body without touching `assembler.py`.  The function
    bodies below are intentionally no-op stubs (return "" / False) — do
    NOT fill in real frontmatter/comment logic here as part of #77; that
    is #78's scope (ARCH-Wave3 §3.3/§3.4).

    A no-op stub guarantees zero side effects even when
    `emit_frontmatter=True` before #78 lands: `assemble_document` only
    prepends/attaches the seam's output when it is non-empty.

Public interface (contract, ADR-617):
    build_frontmatter(pres: PresentationIR) -> str
    build_slide_comment(slide_index: int, section: str, has_diagram: bool) -> str
    derive_has_diagram(slide: SlideIR) -> bool
    derive_section(heading_lines: list[str]) -> str

INV-5: stdlib + internal `ir` imports only. No VLM SDK / python-pptx / Pillow.
"""

from __future__ import annotations

from pptx_md.ir import PresentationIR, SlideIR

__all__ = [
    "build_frontmatter",
    "build_slide_comment",
    "derive_has_diagram",
    "derive_section",
]


def build_frontmatter(pres: PresentationIR) -> str:
    """STUB — contract only. Real implementation owned by issue #78.

    Args:
        pres: The presentation IR (read-only).

    Returns:
        "" (no-op) until #78 implements the YAML frontmatter block.
    """
    return ""


def build_slide_comment(slide_index: int, section: str, has_diagram: bool) -> str:
    """STUB — contract only. Real implementation owned by issue #78.

    Args:
        slide_index: 1-based slide number.
        section: Current top-level '##' heading text.
        has_diagram: Whether the slide contains a DIAGRAM-classified image.

    Returns:
        "" (no-op) until #78 implements the HTML-comment metadata line.
    """
    return ""


def derive_has_diagram(slide: SlideIR) -> bool:
    """STUB — contract only. Real implementation owned by issue #78.

    Args:
        slide: The slide IR (read-only).

    Returns:
        False (no-op) until #78 implements the DIAGRAM-scan derivation.
    """
    return False


def derive_section(heading_lines: list[str]) -> str:
    """STUB — contract only. Real implementation owned by issue #78.

    Args:
        heading_lines: Rendered heading lines for the slide.

    Returns:
        "" (no-op) until #78 implements the top-level '##' extraction.
    """
    return ""
