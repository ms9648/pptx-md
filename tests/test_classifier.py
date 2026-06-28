"""Tests for src/pptx_md/classifier.py and src/pptx_md/image_pipeline.py.

FR-06 AC1~AC9 + image_pipeline enrich_images integration.

Test naming convention: test_ac<N>_<description>
Each AC has at least one covering test that asserts the Then clause directly.

Synthetic fixtures are generated deterministically by Pillow (no binary files
checked in — ADR-207 philosophy).  Every fixture input is fixed-content, so
tests are fully reproducible (AC9 determinism).
"""

from __future__ import annotations

import io
import pathlib
import random

import pytest
from PIL import Image

from pptx_md.classifier import classify_image
from pptx_md.ir import (
    GroupShapeIR,
    ImageClass,
    ImageShapeIR,
    PresentationIR,
    ShapeKind,
    SlideIR,
)

# ---------------------------------------------------------------------------
# Synthetic image builders — deterministic (no random seed; pixel formulae)
# ---------------------------------------------------------------------------


def _to_png_bytes(img: Image.Image) -> bytes:
    """Encode a Pillow image to PNG bytes."""
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _make_photo_like_png(width: int = 200, height: int = 200) -> bytes:
    """Return PNG bytes resembling a high-colour photograph.

    Uses a deterministic gradient + salt pattern with fixed seed (no os.urandom).
    Result: n_unique_colors >> PHOTO_COLORS_MIN, color_fraction high.
    """
    rng = random.Random(42)  # fixed seed -> deterministic
    img = Image.new("RGB", (width, height))
    pixels = []
    for y in range(height):
        for x in range(width):
            # Gradient base (ensures wide colour spread)
            r = int(255 * x / width)
            g = int(255 * y / height)
            b = int(255 * (x + y) / (width + height))
            # Salt: add small jitter to maximise unique colour count
            r = min(255, r + rng.randint(0, 20))
            g = min(255, g + rng.randint(0, 20))
            b = min(255, b + rng.randint(0, 20))
            pixels.append((r, g, b))
    img.putdata(pixels)  # type: ignore[arg-type]
    return _to_png_bytes(img)


def _make_logo_like_png(width: int = 64, height: int = 64) -> bytes:
    """Return PNG bytes resembling a simple logo with transparency.

    Small canvas + alpha channel + very few colours -> LOGO.
    """
    img = Image.new("RGBA", (width, height), (255, 255, 255, 0))  # transparent bg
    # Draw a solid filled rectangle in a single colour
    for y in range(height // 4, 3 * height // 4):
        for x in range(width // 4, 3 * width // 4):
            img.putpixel((x, y), (50, 120, 200, 255))
    return _to_png_bytes(img)


def _make_diagram_like_png(width: int = 256, height: int = 256) -> bytes:
    """Return PNG bytes resembling a diagram/chart (grid lines on white bg).

    White background + thin dark grid lines -> few colours, high bg fraction,
    moderate edge density -> DIAGRAM.
    """
    img = Image.new("RGB", (width, height), (255, 255, 255))  # white bg
    pixels = list(img.getdata())
    line_color = (30, 30, 30)  # near-black grid
    step = 20  # grid spacing

    # Draw horizontal and vertical grid lines
    for y in range(0, height, step):
        for x in range(width):
            pixels[y * width + x] = line_color
    for x in range(0, width, step):
        for y in range(height):
            pixels[y * width + x] = line_color

    # Add a few flat-colour filled rectangles (chart bars)
    bar_color = (70, 130, 180)  # steel blue
    for y in range(height // 2, height - 20):
        for x in range(30, 80):
            pixels[y * width + x] = bar_color
    for y in range(height // 3, height - 20):
        for x in range(100, 150):
            pixels[y * width + x] = bar_color

    img.putdata(pixels)  # type: ignore[arg-type]
    return _to_png_bytes(img)


def _make_text_like_png(width: int = 256, height: int = 256) -> bytes:
    """Return PNG bytes resembling a text screenshot.

    White background + many thin dark horizontal stripes (simulated text lines)
    -> high edge density, very few unique colours, high bg fraction -> TEXT.
    """
    img = Image.new("RGB", (width, height), (255, 255, 255))  # white bg
    pixels = list(img.getdata())
    text_color = (10, 10, 10)  # near-black "text"

    # Simulate ~12 lines of text across the image
    line_height = height // 14
    for line_idx in range(12):
        y_start = line_idx * line_height + (line_height // 4)
        y_end = y_start + max(1, line_height // 2)
        # Each "text line" is a band of dark pixels with gaps (word spacing)
        for y in range(y_start, min(y_end, height)):
            for x in range(10, width - 10):
                # Leave small gaps every ~40px to simulate word spacing
                if (x % 42) not in range(38, 42):
                    pixels[y * width + x] = text_color

    img.putdata(pixels)  # type: ignore[arg-type]
    return _to_png_bytes(img)


# ---------------------------------------------------------------------------
# Helpers to build minimal IR objects for integration tests
# ---------------------------------------------------------------------------


def _make_image_ir(
    image_bytes: bytes,
    image_ext: str = "png",
    shape_id: int = 1,
) -> ImageShapeIR:
    return ImageShapeIR(
        shape_id=shape_id,
        name=f"Picture {shape_id}",
        kind=ShapeKind.IMAGE,
        image_bytes=image_bytes,
        image_format=image_ext,
        image_ext=image_ext,
        alt_text="",
    )


def _make_presentation_with_shapes(
    *shapes: ImageShapeIR,
) -> PresentationIR:
    slide = SlideIR(index=0, title="", notes="", shapes=list(shapes))
    return PresentationIR(source_path="test.pptx", slides=[slide])


# ---------------------------------------------------------------------------
# Module-level synthetic fixtures (built once, reused as bytes constants)
# ---------------------------------------------------------------------------

_PHOTO_PNG: bytes = _make_photo_like_png()
_LOGO_PNG: bytes = _make_logo_like_png()
_DIAGRAM_PNG: bytes = _make_diagram_like_png()
_TEXT_PNG: bytes = _make_text_like_png()


# ===========================================================================
# AC1 — Given decodable ImageShapeIR When classifier called
#        Then classification slot set to one of 4 values (None -> non-None)
# ===========================================================================


def test_ac1_분류_후_non_none() -> None:
    """AC1: classify_image on a decodable image returns a non-None ImageClass."""
    result = classify_image(_PHOTO_PNG)
    assert result is not None
    assert result in list(ImageClass)


def test_ac1_분류_전_classification_none() -> None:
    """AC1: ImageShapeIR.classification starts as None before classification."""
    shape = _make_image_ir(_PHOTO_PNG)
    assert shape.classification is None


# ===========================================================================
# AC2 — Given high-colour photo fixture When classifier called Then PHOTO
# ===========================================================================


def test_ac2_photo_분류() -> None:
    """AC2: High-colour gradient image classifies as PHOTO."""
    result = classify_image(_PHOTO_PNG)
    assert result == ImageClass.PHOTO


# ===========================================================================
# AC3 — Given small+transparent+few-colour logo fixture Then LOGO
# ===========================================================================


def test_ac3_logo_분류() -> None:
    """AC3: Small RGBA image with few colours classifies as LOGO."""
    result = classify_image(_LOGO_PNG)
    assert result == ImageClass.LOGO


def test_ac3_logo_알파없이_단순_작은_이미지() -> None:
    """AC3: Very small opaque image with few unique colours also classifies as LOGO."""
    img = Image.new("RGB", (32, 32), (200, 50, 50))  # solid red, tiny
    png = _to_png_bytes(img)
    result = classify_image(png)
    assert result == ImageClass.LOGO


# ===========================================================================
# AC4 — Given grid/chart diagram fixture Then DIAGRAM
# ===========================================================================


def test_ac4_diagram_분류() -> None:
    """AC4: Grid-line diagram image classifies as DIAGRAM."""
    result = classify_image(_DIAGRAM_PNG)
    assert result == ImageClass.DIAGRAM


# ===========================================================================
# AC5 — Given text-screenshot fixture Then TEXT
# ===========================================================================


def test_ac5_text_분류() -> None:
    """AC5: Text-stripe image (high-contrast lines on white) classifies as TEXT."""
    result = classify_image(_TEXT_PNG)
    assert result == ImageClass.TEXT


# ===========================================================================
# AC6 — in-place mutation via enrich_images: only classification changes
# ===========================================================================


def test_ac6_enrich_images_in_place_갱신() -> None:
    """AC6: enrich_images mutates classification in-place, other fields preserved."""
    from pptx_md.image_pipeline import enrich_images

    shape = _make_image_ir(_PHOTO_PNG, image_ext="png", shape_id=10)
    original_bytes = shape.image_bytes
    original_format = shape.image_format
    original_name = shape.name
    original_id = shape.shape_id
    original_alt = shape.alt_text
    original_kind = shape.kind

    pres = _make_presentation_with_shapes(shape)
    assert shape.classification is None  # pre-condition

    enrich_images(pres)

    # classification slot must be filled (non-None)
    assert shape.classification is not None

    # All other fields must be unchanged
    assert shape.image_bytes is original_bytes
    assert shape.image_format == original_format
    assert shape.name == original_name
    assert shape.shape_id == original_id
    assert shape.alt_text == original_alt
    assert shape.kind == original_kind


def test_ac6_enrich_images_반환_none() -> None:
    """AC6: enrich_images returns None (in-place semantics, ADR-210)."""
    from pptx_md.image_pipeline import enrich_images

    pres = _make_presentation_with_shapes(_make_image_ir(_PHOTO_PNG))
    result = enrich_images(pres)
    assert result is None


def test_ac6_ir_스키마_변경_없음() -> None:
    """AC6: IR schema unchanged — no new fields added by enrich_images."""
    import dataclasses

    from pptx_md.image_pipeline import enrich_images

    shape = _make_image_ir(_PHOTO_PNG)
    field_names_before = {f.name for f in dataclasses.fields(shape)}

    pres = _make_presentation_with_shapes(shape)
    enrich_images(pres)

    field_names_after = {f.name for f in dataclasses.fields(shape)}
    assert field_names_before == field_names_after


# ===========================================================================
# AC7 — Given empty/corrupt bytes Then no exception, classification=None
# ===========================================================================


def test_ac7_디코딩_불가_bytes_none_반환() -> None:
    """AC7: Non-decodable bytes return None without raising."""
    result = classify_image(b"not_an_image")
    assert result is None


def test_ac7_빈_bytes_none_반환() -> None:
    """AC7: Empty bytes return None without raising."""
    result = classify_image(b"")
    assert result is None


def test_ac7_truncated_png_none_반환() -> None:
    """AC7: Truncated PNG header returns None without raising."""
    result = classify_image(b"\x89PNG\r\n\x1a\n\x00\x00")
    assert result is None


def test_ac7_예외_전파_없음() -> None:
    """AC7: classify_image never raises — even on garbage input."""
    try:
        classify_image(b"\x00" * 100)
        classify_image(b"RIFF\x00\x00\x00\x00AVI ")  # fake AVI
        classify_image(b"\xff\xd8\xff\xe0")  # truncated JPEG header
    except Exception as exc:
        pytest.fail(f"classify_image raised unexpectedly: {exc}")


# ===========================================================================
# AC8 — VLM/LibreOffice import 0 건
# ===========================================================================


def test_ac8_vlm_import_없음_classifier() -> None:
    """AC8: classifier.py source contains no anthropic/openai imports."""
    source_file = (
        pathlib.Path(__file__).parent.parent / "src" / "pptx_md" / "classifier.py"
    )
    source = source_file.read_text(encoding="utf-8")
    assert "anthropic" not in source, "classifier.py must not import anthropic"
    assert "openai" not in source, "classifier.py must not import openai"


def test_ac8_vlm_import_없음_image_pipeline() -> None:
    """AC8: image_pipeline.py source contains no anthropic/openai imports."""
    source_file = (
        pathlib.Path(__file__).parent.parent / "src" / "pptx_md" / "image_pipeline.py"
    )
    source = source_file.read_text(encoding="utf-8")
    assert "anthropic" not in source, "image_pipeline.py must not import anthropic"
    assert "openai" not in source, "image_pipeline.py must not import openai"


def test_ac8_vlm_import_없음_vector() -> None:
    """AC8: vector.py source contains no anthropic/openai imports."""
    source_file = pathlib.Path(__file__).parent.parent / "src" / "pptx_md" / "vector.py"
    source = source_file.read_text(encoding="utf-8")
    assert "anthropic" not in source, "vector.py must not import anthropic"
    assert "openai" not in source, "vector.py must not import openai"


def test_ac8_import_no_importerror() -> None:
    """AC8: Modules import without ImportError in any environment."""
    import importlib

    importlib.import_module("pptx_md.classifier")
    importlib.import_module("pptx_md.image_pipeline")
    importlib.import_module("pptx_md.vector")


# ===========================================================================
# AC9 — Determinism (same input -> same output, 2 calls)
# ===========================================================================


def test_ac9_결정성_photo() -> None:
    """AC9: Same photo bytes classified identically on two calls."""
    result1 = classify_image(_PHOTO_PNG)
    result2 = classify_image(_PHOTO_PNG)
    assert result1 == result2


def test_ac9_결정성_logo() -> None:
    """AC9: Same logo bytes classified identically on two calls."""
    assert classify_image(_LOGO_PNG) == classify_image(_LOGO_PNG)


def test_ac9_결정성_diagram() -> None:
    """AC9: Same diagram bytes classified identically on two calls."""
    assert classify_image(_DIAGRAM_PNG) == classify_image(_DIAGRAM_PNG)


def test_ac9_결정성_text() -> None:
    """AC9: Same text-like bytes classified identically on two calls."""
    assert classify_image(_TEXT_PNG) == classify_image(_TEXT_PNG)


def test_ac9_결정성_none_입력() -> None:
    """AC9: Corrupt bytes consistently return None on two calls."""
    corrupt = b"deadbeef" * 10
    assert classify_image(corrupt) == classify_image(corrupt)


# ===========================================================================
# Integration: enrich_images with multiple shapes including grouped
# ===========================================================================


def test_integration_enrich_multiple_shapes() -> None:
    """enrich_images classifies all image shapes in a multi-shape slide."""
    from pptx_md.image_pipeline import enrich_images

    shape1 = _make_image_ir(_PHOTO_PNG, shape_id=1)
    shape2 = _make_image_ir(_LOGO_PNG, shape_id=2)
    shape3 = _make_image_ir(_DIAGRAM_PNG, shape_id=3)

    pres = _make_presentation_with_shapes(shape1, shape2, shape3)
    enrich_images(pres)

    assert shape1.classification is not None
    assert shape2.classification is not None
    assert shape3.classification is not None


def test_integration_enrich_grouped_images() -> None:
    """enrich_images reaches ImageShapeIR nested inside a GroupShapeIR."""
    from pptx_md.image_pipeline import enrich_images

    image_shape = _make_image_ir(_PHOTO_PNG, shape_id=99)
    group = GroupShapeIR(
        shape_id=50,
        name="Group 1",
        kind=ShapeKind.GROUP,
        children=[image_shape],
    )
    slide = SlideIR(index=0, shapes=[group])
    pres = PresentationIR(source_path="test.pptx", slides=[slide])

    enrich_images(pres)

    assert image_shape.classification is not None


def test_integration_enrich_corrupt_image_isolation() -> None:
    """One corrupt image does not prevent other images from being classified."""
    from pptx_md.image_pipeline import enrich_images

    corrupt_shape = _make_image_ir(b"not_an_image", shape_id=1)
    good_shape = _make_image_ir(_PHOTO_PNG, shape_id=2)

    pres = _make_presentation_with_shapes(corrupt_shape, good_shape)
    enrich_images(pres)

    # corrupt image: classification stays None (graceful)
    assert corrupt_shape.classification is None
    # good image: classified normally
    assert good_shape.classification == ImageClass.PHOTO


def test_integration_enrich_emf_no_libreoffice(monkeypatch: pytest.MonkeyPatch) -> None:
    """EMF shapes with no LibreOffice available keep classification=None."""
    import pptx_md.vector as vec
    from pptx_md.image_pipeline import enrich_images

    # Force libreoffice_available() -> False
    monkeypatch.setattr(vec, "libreoffice_available", lambda: False)
    # Also patch shutil.which inside vector module to return None
    import shutil as _shutil

    monkeypatch.setattr(_shutil, "which", lambda _name: None)

    emf_shape = ImageShapeIR(
        shape_id=5,
        name="EMF Picture",
        kind=ShapeKind.IMAGE,
        image_bytes=b"EMF_STUB_DATA",
        image_format="emf",
        image_ext="emf",
        alt_text="",
    )
    pres = _make_presentation_with_shapes(emf_shape)
    enrich_images(pres)

    # No LibreOffice -> classification stays None, no exception
    assert emf_shape.classification is None


def test_integration_enrich_vector_classify_after_conversion(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """When vector conversion succeeds, the PNG result is classified."""
    import pptx_md.vector as vec
    from pptx_md.image_pipeline import enrich_images

    # Monkeypatch convert_vector_to_png to return a real photo PNG
    monkeypatch.setattr(vec, "convert_vector_to_png", lambda _b, _e: _PHOTO_PNG)

    emf_shape = ImageShapeIR(
        shape_id=6,
        name="EMF Picture",
        kind=ShapeKind.IMAGE,
        image_bytes=b"EMF_STUB",
        image_format="emf",
        image_ext="emf",
        alt_text="",
    )
    pres = _make_presentation_with_shapes(emf_shape)
    enrich_images(pres)

    assert emf_shape.classification == ImageClass.PHOTO


# vector.py 전용 테스트는 tests/test_vector.py 에서 관리 (AC1~AC6 전부)
# 이 파일에서는 vector 를 오케스트레이터(image_pipeline) 통합 테스트에서만 사용
