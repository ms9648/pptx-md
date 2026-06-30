"""Intermediate Representation (IR) for parsed PPTX content.

Internal module (skill §1, ARCH-M1 8.3): NOT re-exported from package root.
Pure data structures — MUST NOT import python-pptx (ADR-206).
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass, field
from enum import StrEnum


class ShapeKind(StrEnum):
    """Discriminant tag for shape IR nodes (ADR-201)."""

    TEXT = "text"
    TABLE = "table"
    IMAGE = "image"
    GROUP = "group"
    OTHER = "other"


class ImageClass(StrEnum):
    """Rule-based image classes (filled by M3 / FR-06; v1 parser leaves None).

    Defined in M2 to lock the slot type of ImageShapeIR.classification
    as ``ImageClass | None`` under mypy strict (ADR-205).
    """

    TEXT = "text"
    DIAGRAM = "diagram"
    PHOTO = "photo"
    LOGO = "logo"


@dataclass
class ShapeIR:
    """Abstract base for all shape IR nodes.

    Subclasses: TextShapeIR, TableShapeIR, ImageShapeIR, GroupShapeIR,
    OtherShapeIR.  Direct instantiation is not intended; use a concrete
    subclass.
    """

    shape_id: int
    name: str  # "" if absent (ADR-202)
    kind: ShapeKind


@dataclass
class ParagraphIR:
    """A single paragraph within a text shape.

    Empty paragraphs (text == "") are preserved to retain original line
    break structure; the assembler (M5) decides whether to collapse them
    (§5.4).
    """

    text: str  # "" allowed; empty paragraphs preserved (§5.4)
    level: int  # 0-based indent level (python-pptx paragraph.level)


@dataclass
class TextShapeIR(ShapeIR):
    """IR for a text-frame shape (placeholder or freeform text box)."""

    paragraphs: list[ParagraphIR] = field(default_factory=list)
    is_title: bool = False  # True when the shape originates from a title placeholder
    is_footer: bool = (
        False  # True when shape is a footer/slide-number/date placeholder (FR-21)
    )


@dataclass
class TableShapeIR(ShapeIR):
    """IR for a table shape.

    rows is a 2-D list of cell texts; merged cells are expanded with the
    cell text repeated in each logical cell (v1 limitation — §11 debt).
    """

    rows: list[list[str]] = field(default_factory=list)
    n_rows: int = 0
    n_cols: int = 0


@dataclass
class ImageShapeIR(ShapeIR):
    """IR for a picture shape.

    Holds raw image bytes for M3/M4 consumption.  Extension slots
    (classification, description) are always None in v1; M3/M4 fill them
    without changing the IR schema (ADR-205).
    """

    image_bytes: bytes = b""
    image_format: str = ""  # e.g. "png", "jpeg" (ADR-202: "" not None)
    image_ext: str = ""  # e.g. "png", "emf"  (ADR-202)
    alt_text: str = ""  # PPTX alt text; "" if absent (ADR-202)
    # Extension slots — v1 parser always leaves None (ADR-205)
    classification: ImageClass | None = None  # filled by M3 / FR-06
    description: str | None = None  # filled by M4 / FR-08+


@dataclass
class GroupShapeIR(ShapeIR):
    """IR for a group shape; preserves the tree structure (ADR-203).

    Consumers that need a flat sequence use iter_shapes().
    """

    children: list[ShapeIR] = field(default_factory=list)


@dataclass
class OtherShapeIR(ShapeIR):
    """Fallback IR for unsupported/unrecognised shapes (ADR-204).

    Guarantees that one shape failure never aborts the full conversion.
    """

    mso_shape_type: str = "UNKNOWN"  # original MSO type name or "UNKNOWN"
    fallback_text: str = ""  # best-effort text if the shape has a text_frame


@dataclass
class SlideIR:
    """IR for a single slide."""

    index: int  # 0-based slide index
    title: str = ""  # "" when absent (ADR-202)
    notes: str = ""  # "" when absent (ADR-202)
    shapes: list[ShapeIR] = field(default_factory=list)


@dataclass
class PresentationIR:
    """IR for the entire presentation."""

    source_path: str  # original file path (diagnostic/tracing; not PII)
    slide_width_emu: int = 0  # python-pptx prs.slide_width in EMU
    slide_height_emu: int = 0  # python-pptx prs.slide_height in EMU
    slides: list[SlideIR] = field(default_factory=list)


def iter_shapes(container: SlideIR | GroupShapeIR) -> Iterator[ShapeIR]:
    """Depth-first flattening of the shape tree (groups expanded).

    Yields every ShapeIR in depth-first order, recursively descending into
    GroupShapeIR nodes.  Group nodes themselves are also yielded before
    their children, so callers that want to skip groups can filter by
    ``isinstance(shape, GroupShapeIR)``.

    Args:
        container: A SlideIR or GroupShapeIR whose shapes/children to iterate.

    Yields:
        Each ShapeIR node in depth-first pre-order.
    """
    shapes: list[ShapeIR] = (
        container.shapes if isinstance(container, SlideIR) else container.children
    )
    for shape in shapes:
        yield shape
        if isinstance(shape, GroupShapeIR):
            yield from iter_shapes(shape)
