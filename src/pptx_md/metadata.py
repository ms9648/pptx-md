"""Frontmatter / slide metadata (FR-28 AC5-AC7, issue #78, W-B).

Ownership (ARCH-Wave3 §2.1/ADR-621, WBS issue B / #78):
    This module implements the metadata seam whose contract signature was
    locked by issue #77 (W-A / `assembler.py` seam). `assembler.py` /
    `api.py` are NOT touched here (§2.1 file-ownership isolation).

Public interface (contract, ADR-617):
    build_frontmatter(pres: PresentationIR) -> str
    build_slide_comment(slide_index: int, section: str, has_diagram: bool) -> str
    derive_has_diagram(slide: SlideIR) -> bool
    derive_section(heading_lines: list[str]) -> str

Design constraints (ARCH-Wave3 §3.3/§3.4, ADR-617):
    - Pure functions: no mutation, no I/O, deterministic (INV-1).
    - stdlib + internal `ir` imports only — no VLM SDK / python-pptx /
      Pillow imports (INV-5).
    - IR is unchanged (INV-2): `section`/`has_diagram` are derived purely
      at assemble time from existing IR fields.
    - No PyYAML: frontmatter is assembled via plain string joins (INV-6,
      D-1). `source_path` is intentionally NOT included in the
      frontmatter — it is a local filesystem path and may carry PII.
"""

from __future__ import annotations

from pptx_md.ir import ImageClass, ImageShapeIR, PresentationIR, SlideIR, iter_shapes

__all__ = [
    "build_frontmatter",
    "build_slide_comment",
    "derive_has_diagram",
    "derive_section",
]

# ---------------------------------------------------------------------------
# Named constants (ADR-617)
# ---------------------------------------------------------------------------
_FRONTMATTER_DELIMITER: str = "---"
_HEADING_H2_PREFIX: str = "## "
_HEADING_H3_PREFIX: str = "### "


def build_frontmatter(pres: PresentationIR) -> str:
    """Build the document-level YAML frontmatter block (FR-28 AC5, AC7).

    Emitted once at the very top of the document (assembler seam,
    ARCH-Wave3 §3.3). Uses stdlib string assembly only — no PyYAML
    (D-1/INV-6). Deliberately omits `source_path` (local filesystem path,
    potential PII) — only minimal, deterministic keys are included.

    Args:
        pres: The presentation IR (read-only).

    Returns:
        A `"---\\n...\\n---"` YAML frontmatter block string.
    """
    lines = [
        _FRONTMATTER_DELIMITER,
        "generator: pptx-md",
        f"slide_count: {len(pres.slides)}",
        _FRONTMATTER_DELIMITER,
    ]
    return "\n".join(lines)


def build_slide_comment(slide_index: int, section: str, has_diagram: bool) -> str:
    """Build one slide's HTML-comment metadata line (FR-28 AC5, AC7).

    The comment is an HTML comment only — never a heading/fence/table —
    so it never conflicts with the assembler's `---` slide separator or
    FR-14's heading/fence validator rules (AC7).

    Args:
        slide_index: 1-based slide number.
        section: Current top-level '##' heading text (may contain '\\n'
            or '|', which are sanitised below to avoid breaking the
            single-line comment format).
        has_diagram: Whether the slide contains a DIAGRAM-classified
            image.

    Returns:
        A single-line `<!-- slide_index: N | section: ... | has_diagram:
        bool -->` HTML comment string.
    """
    # Sanitise section so embedded '\n'/'|' cannot break the single-line
    # "key: value | key: value" comment format (ARCH-Wave3 §3.3).
    safe_section = section.replace("\n", " ").replace("|", "/")
    flag = "true" if has_diagram else "false"
    return (
        f"<!-- slide_index: {slide_index} | section: {safe_section} "
        f"| has_diagram: {flag} -->"
    )


def derive_has_diagram(slide: SlideIR) -> bool:
    """Derive whether *slide* contains a DIAGRAM-classified image (FR-28 AC6).

    Scans the slide's shapes (including nested group children, via
    `iter_shapes`'s depth-first traversal) for any `ImageShapeIR` whose
    `classification` is `ImageClass.DIAGRAM`. IR is unchanged (D-3/INV-2);
    this is a pure, deterministic derivation at assemble time.

    Args:
        slide: The slide IR (read-only).

    Returns:
        True if at least one DIAGRAM-classified image shape is present
        (at any nesting depth), False otherwise.
    """
    return any(
        isinstance(shape, ImageShapeIR) and shape.classification is ImageClass.DIAGRAM
        for shape in iter_shapes(slide)
    )


def derive_section(heading_lines: list[str]) -> str:
    """Derive the current top-level '##' (h2) heading text (FR-28 AC5).

    Args:
        heading_lines: Rendered heading lines for the slide (e.g. from
            `heading.render_heading_lines`), in rendered order.

    Returns:
        The text of the first h2 (`## ...`) line found, with surrounding
        whitespace stripped; `""` if no h2 line is present (AC4/AC5 —
        every slide has at least one `##` line in practice, including the
        `## Slide N` fallback, but this returns `""` defensively when none
        is found).
    """
    for line in heading_lines:
        if line.startswith(_HEADING_H2_PREFIX) and not line.startswith(
            _HEADING_H3_PREFIX
        ):
            return line[len(_HEADING_H2_PREFIX) :].strip()
    return ""
