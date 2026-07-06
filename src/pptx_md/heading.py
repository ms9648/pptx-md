"""Heading-path decomposition and body-emptiness detection (FR-28, issue #77).

Public interface:
    split_heading_path(title: str) -> list[str]
    render_heading_lines(title: str, *, hierarchy: bool) -> list[str]
    is_body_empty(body_parts: list[str]) -> bool

Design constraints (ARCH-Wave3 §3.1/§3.2, ADR-615/616):
    - Pure functions: no mutation, no I/O, deterministic (INV-1).
    - stdlib only — no VLM SDK / python-pptx / Pillow imports (INV-5).
    - hierarchy=False (or <=1 path segment) renders a single '## {title}'
      line using the *original* title text verbatim, so that the
      heading_hierarchy=False path stays byte-identical to the pre-Wave-3
      assembler output (FR-28 AC3).
"""

from __future__ import annotations

__all__ = ["split_heading_path", "render_heading_lines", "is_body_empty"]

# ---------------------------------------------------------------------------
# Named constants (ADR-615)
# ---------------------------------------------------------------------------
_HEADING_PATH_SEP: str = ">"
_MAX_HEADING_LEVEL: int = 6  # ATX max; segments beyond this are clamped to h6


def split_heading_path(title: str) -> list[str]:
    """Split *title* on the '>' path separator into trimmed segments.

    Each segment is stripped of surrounding whitespace; empty segments
    (e.g. from a leading/trailing/doubled separator) are dropped.  Pure,
    deterministic (ADR-615).

    Args:
        title: Raw slide title text, possibly containing '>'-separated
            path segments (e.g. "Ⅱ. 사업의 이해 > ⑦ 현황·문제점").

    Returns:
        Ordered list of trimmed, non-empty path segments.
    """
    return [seg.strip() for seg in title.split(_HEADING_PATH_SEP) if seg.strip()]


def render_heading_lines(title: str, *, hierarchy: bool) -> list[str]:
    """Render *title* as one or more ATX heading lines (FR-28 AC1-AC3).

    Behaviour:
        - ``hierarchy=False``: always a single ``## {title}`` line using
          the original title text verbatim (AC3 — byte-identical to the
          pre-Wave-3 renderer when the feature is off).
        - ``hierarchy=True`` and the title has no '>' separator (0 or 1
          trimmed segments): single ``## {title}`` line, i.e. current
          behaviour is preserved (AC2).
        - ``hierarchy=True`` and 2+ segments: seg[0] -> ``##`` (h2),
          seg[1] -> ``###`` (h3), ... clamped at h6 for any further
          segments (AC1, ADR-615).

    Args:
        title: Raw slide title text.
        hierarchy: Enable '>' path decomposition when True.

    Returns:
        Ordered list of ATX heading line strings (each a separate
        Markdown block when joined with "\\n\\n" by the caller).
    """
    if not hierarchy:
        return [f"## {title}"]
    segs = split_heading_path(title)
    if len(segs) <= 1:
        return [f"## {title}"]
    return [
        f"{'#' * min(2 + i, _MAX_HEADING_LEVEL)} {seg}" for i, seg in enumerate(segs)
    ]


def is_body_empty(body_parts: list[str]) -> bool:
    """Return True when *body_parts* contain no non-whitespace content (D-4).

    D-4 (REQ-wave3 §7): "본문 렌더 결과가 헤딩 라인 외 비공백 0인 슬라이드" —
    the shared empty/section-divider-slide definition used by both FR-28
    AC4 (TOC conversion) and FR-29 AC6 (empty-slide validator warning).

    Args:
        body_parts: Rendered body parts of a slide, excluding heading lines.

    Returns:
        True if the concatenated body has no non-whitespace characters.
    """
    return not "".join(body_parts).strip()
