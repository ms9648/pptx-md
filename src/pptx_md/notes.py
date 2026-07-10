"""Speaker-notes blockquote rendering (FR-28 AC8-AC9, issue #79, W-C).

Public interface:
    render_notes_block(notes: str) -> str

Design constraints (ARCH-Wave3 §3.6, ADR-618):
    - Pure function: no mutation, no I/O, deterministic (INV-1).
    - stdlib only — no VLM SDK / python-pptx / Pillow imports (INV-5).
    - Empty/whitespace-only notes render to "" (no blockquote emitted,
      AC2/AC9) so that `include_notes=True` with the default
      `SlideIR.notes == ""` stays byte-identical to the pre-Wave-3
      assembler output.
    - `\\v` (U+000B) is normalized to `\\n` before splitting into lines,
      matching the FR-24 control-character normalization convention
      already applied elsewhere in the assembler (AC3).
"""

from __future__ import annotations

__all__ = ["render_notes_block"]

_NOTES_HEADER: str = "> notes"


def render_notes_block(notes: str) -> str:
    """Render *notes* as a '> notes' blockquote (FR-28 AC8-AC9).

    Behaviour:
        - ``notes`` empty or whitespace-only: returns ``""`` — no
          blockquote is emitted (AC2/AC9). The caller (assembler seam)
          only appends this result as a slide block part when it is
          non-empty, so the off/empty path stays byte-identical.
        - Otherwise: ``\\v`` is normalized to ``\\n`` (FR-24 parity,
          AC3), the text is split into lines, and each line — including
          the fixed ``"> notes"`` header — is emitted with a ``"> "``
          prefix (AC1/AC8).

    Args:
        notes: Raw ``SlideIR.notes`` text.

    Returns:
        The rendered '> notes' blockquote, or ``""`` when *notes* is
        empty/whitespace-only.
    """
    if not notes.strip():
        return ""
    normalized = notes.replace("\v", "\n")
    lines = [_NOTES_HEADER] + [f"> {line}" for line in normalized.split("\n")]
    return "\n".join(lines)
