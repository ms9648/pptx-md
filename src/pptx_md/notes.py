"""Speaker-notes rendering contract — MINIMAL STUB (FR-28 AC8-AC9, issue #77).

Ownership (ARCH-Wave3 §2.1/ADR-621, WBS issue C / #79):
    This module is created by issue #77 (W-A) *only* to lock the public
    contract signature that `assembler.py` seams call, so that #79 (W-C)
    can implement the body without touching `assembler.py`.  The function
    body below is intentionally a no-op stub (always returns "") — do NOT
    fill in the real '> notes' rendering here as part of #77; that is
    #79's scope (ARCH-Wave3 §3.6).

    A no-op stub guarantees zero side effects even when
    `include_notes=True` before #79 lands: `assemble_document` only
    appends the seam's output when it is non-empty (AC9 semantics).

Public interface (contract, ADR-618):
    render_notes_block(notes: str) -> str

INV-5: stdlib-only. No VLM SDK / python-pptx / Pillow / internal imports needed.
"""

from __future__ import annotations

__all__ = ["render_notes_block"]


def render_notes_block(notes: str) -> str:
    """STUB — contract only. Real implementation owned by issue #79.

    Args:
        notes: Raw SlideIR.notes text.

    Returns:
        "" (no-op) until #79 implements the '> notes' blockquote rendering.
    """
    return ""
