"""Shared pytest fixtures (skill §5: fixtures concentrated here).

M2: in-memory PPTX generation fixtures for parser tests (ADR-207).
python-pptx is NOT mocked — real API calls only.
"""

from __future__ import annotations

import io
import struct
import zlib
from pathlib import Path

import pytest
from pptx import Presentation as PptxPresentation
from pptx.util import Inches

# ---------------------------------------------------------------------------
# Minimal 1×1 PNG helper (for image fixture)
# ---------------------------------------------------------------------------


def _make_png_bytes() -> bytes:
    """Return a minimal 1x1 white PNG as bytes (no Pillow dependency)."""
    # PNG signature
    sig = b"\x89PNG\r\n\x1a\n"

    def _chunk(name: bytes, data: bytes) -> bytes:
        length = struct.pack(">I", len(data))
        crc = struct.pack(">I", zlib.crc32(name + data) & 0xFFFFFFFF)
        return length + name + data + crc

    ihdr_data = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)  # 1x1 RGB
    ihdr = _chunk(b"IHDR", ihdr_data)

    # IDAT: filtered scanline for 1x1 RGB (filter byte 0 + RGB white)
    raw = b"\x00\xff\xff\xff"
    compressed = zlib.compress(raw)
    idat = _chunk(b"IDAT", compressed)

    iend = _chunk(b"IEND", b"")
    return sig + ihdr + idat + iend


_PNG_BYTES = _make_png_bytes()


# ---------------------------------------------------------------------------
# In-memory PPTX factories
# ---------------------------------------------------------------------------


def _save_to_bytes(prs: PptxPresentation) -> bytes:
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


@pytest.fixture
def pptx_with_text_slide() -> bytes:
    """Single slide with a title placeholder and a body text box.

    Used for: FR-03 AC1/AC2/AC4, FR-04 AC1.
    """
    prs = PptxPresentation()
    layout = prs.slide_layouts[1]  # Title and Content
    slide = prs.slides.add_slide(layout)

    # Set title
    title_shape = slide.shapes.title
    if title_shape is not None:
        title_shape.text = "Test Slide Title"

    # Set body content
    body = slide.placeholders[1]
    tf = body.text_frame
    tf.text = "First paragraph"
    p2 = tf.add_paragraph()
    p2.text = "Second paragraph"
    p2.level = 1

    return _save_to_bytes(prs)


@pytest.fixture
def pptx_with_notes() -> bytes:
    """Single slide with notes text.

    Used for: FR-03 AC3.
    """
    prs = PptxPresentation()
    layout = prs.slide_layouts[6]  # Blank
    slide = prs.slides.add_slide(layout)

    notes_slide = slide.notes_slide
    notes_tf = notes_slide.notes_text_frame
    notes_tf.text = "Speaker notes text"

    return _save_to_bytes(prs)


@pytest.fixture
def pptx_multi_slide() -> bytes:
    """Three slides in order: titled, blank, titled again.

    Used for: FR-03 AC1/AC4 (order, index).
    """
    prs = PptxPresentation()

    for i in range(3):
        layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        title_shape = slide.shapes.title
        if title_shape is not None:
            title_shape.text = f"Slide {i}"

    return _save_to_bytes(prs)


@pytest.fixture
def pptx_no_title_slide() -> bytes:
    """Single slide using a blank layout (no title placeholder).

    Used for: FR-03 AC5 (title absent -> "").
    """
    prs = PptxPresentation()
    prs.slides.add_slide(prs.slide_layouts[6])  # Blank — no title placeholder
    return _save_to_bytes(prs)


@pytest.fixture
def pptx_zero_slides() -> bytes:
    """Presentation with 0 slides.

    Used for: FR-03 AC7.
    """
    prs = PptxPresentation()
    # Do not add any slides — python-pptx starts with 0 slides
    return _save_to_bytes(prs)


@pytest.fixture
def pptx_with_table() -> bytes:
    """Single slide with a 2x3 table.

    Used for: FR-04 AC2/AC7.
    """
    prs = PptxPresentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank

    tbl = slide.shapes.add_table(2, 3, Inches(1), Inches(1), Inches(6), Inches(2)).table
    tbl.cell(0, 0).text = "A"
    tbl.cell(0, 1).text = "B"
    tbl.cell(0, 2).text = "C"
    tbl.cell(1, 0).text = "D"
    tbl.cell(1, 1).text = "E"
    tbl.cell(1, 2).text = "F"

    return _save_to_bytes(prs)


@pytest.fixture
def pptx_with_empty_table() -> bytes:
    """Single slide with a table where all cells are empty.

    Used for: FR-04 AC7 (empty cells -> "").
    """
    prs = PptxPresentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(2))
    # Cells remain empty (default)

    return _save_to_bytes(prs)


@pytest.fixture
def pptx_with_image(tmp_path: Path) -> bytes:
    """Single slide with a PNG image shape.

    Used for: FR-04 AC3.
    """
    # Write the minimal PNG to a temp file so python-pptx can add it
    png_file = tmp_path / "test.png"
    png_file.write_bytes(_PNG_BYTES)

    prs = PptxPresentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(str(png_file), Inches(1), Inches(1), Inches(3), Inches(2))

    return _save_to_bytes(prs)


@pytest.fixture
def pptx_with_group() -> bytes:
    """Single slide with a group shape containing one text child.

    Uses XML direct assembly because python-pptx 1.0.2 does not expose
    group_shapes().  The resulting PPTX is valid and python-pptx can
    read it back with shape_type == MSO_SHAPE_TYPE.GROUP.

    Used for: FR-04 AC4.
    """
    from lxml import etree  # noqa: PLC0415 — local import keeps top-level clean

    prs = PptxPresentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    sp_tree = slide.shapes._spTree

    # Manually build a grpSp element with one text child.
    grp_xml = (
        "<p:grpSp"
        ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<p:nvGrpSpPr>"
        '<p:cNvPr id="10" name="Group 1"/>'
        "<p:cNvGrpSpPr/>"
        "<p:nvPr/>"
        "</p:nvGrpSpPr>"
        "<p:grpSpPr>"
        "<a:xfrm>"
        '<a:off x="0" y="0"/><a:ext cx="6858000" cy="4572000"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="6858000" cy="4572000"/>'
        "</a:xfrm>"
        "</p:grpSpPr>"
        "<p:sp>"
        "<p:nvSpPr>"
        '<p:cNvPr id="11" name="TextInGroup"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        "<p:nvPr/>"
        "</p:nvSpPr>"
        "<p:spPr>"
        "<a:xfrm>"
        '<a:off x="914400" y="914400"/>'
        '<a:ext cx="2743200" cy="914400"/>'
        "</a:xfrm>"
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        "</p:spPr>"
        "<p:txBody>"
        "<a:bodyPr/><a:lstStyle/>"
        "<a:p><a:r><a:t>child text</a:t></a:r></a:p>"
        "</p:txBody>"
        "</p:sp>"
        "</p:grpSp>"
    )
    grp_elem = etree.fromstring(grp_xml)
    sp_tree.append(grp_elem)

    return _save_to_bytes(prs)


@pytest.fixture
def pptx_mixed_shapes(tmp_path: Path) -> bytes:
    """Single slide with text + table + image (mixed types, z-order test).

    Used for: FR-04 AC5.
    """
    png_file = tmp_path / "test.png"
    png_file.write_bytes(_PNG_BYTES)

    prs = PptxPresentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Add in z-order: text, table, image
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(2), Inches(0.8))
    tb.text_frame.text = "Text shape"

    slide.shapes.add_table(2, 2, Inches(0.5), Inches(1.5), Inches(3), Inches(1.5))

    slide.shapes.add_picture(
        str(png_file), Inches(4), Inches(1), Inches(2), Inches(1.5)
    )

    return _save_to_bytes(prs)


@pytest.fixture
def other_shape_pptx_path() -> Path:
    """Path to tests/fixtures/other_shape.pptx (chart = OTHER shape).

    Used for: FR-04 AC6.
    Created by the fixture generation script; checked into repo (ADR-207).
    """
    return Path(__file__).parent / "fixtures" / "other_shape.pptx"
