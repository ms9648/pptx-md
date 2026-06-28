"""Tests for src/pptx_md/description_pipeline.py (FR-10, issue #28).

AC1~AC7 전 항목 커버. 모든 describe 호출은 FakeDescriber (mock 기반, 실제 API 호출 0).

Test naming convention: test_ac<N>_<description>
Each AC has at least one covering test that asserts the Then clause directly.
"""

from __future__ import annotations

import logging
import pathlib
import subprocess
import sys
from typing import Any
from unittest.mock import patch

import pytest

from pptx_md.ir import (
    GroupShapeIR,
    ImageClass,
    ImageShapeIR,
    PresentationIR,
    ShapeKind,
    SlideIR,
    TextShapeIR,
)

# ---------------------------------------------------------------------------
# Fixtures / helpers
# ---------------------------------------------------------------------------

_PROJECT_ROOT = pathlib.Path(__file__).parent.parent


class FakeDescriber:
    """Minimal ImageDescriber-compatible test double (no SDK, no inheritance).

    Stores each call's arguments for inspection.
    Records whether it was invoked, and what shape_hint values it received.
    Raises if ``raise_on_call`` is True (simulates provider failure).
    """

    def __init__(
        self,
        return_text: str = "fake description",
        raise_on_call: bool = False,
        raise_exc: Exception | None = None,
    ) -> None:
        self.return_text = return_text
        self.raise_on_call = raise_on_call
        self.raise_exc = raise_exc or RuntimeError("fake describe error")
        self.calls: list[dict[str, Any]] = []

    def describe(
        self,
        image_bytes: bytes,
        image_ext: str,
        shape_hint: str | None,
    ) -> str:
        self.calls.append(
            {
                "image_bytes": image_bytes,
                "image_ext": image_ext,
                "shape_hint": shape_hint,
            }
        )
        if self.raise_on_call:
            raise self.raise_exc
        return self.return_text


def _make_presentation(
    slides: list[SlideIR] | None = None,
) -> PresentationIR:
    """Build a minimal PresentationIR for testing."""
    return PresentationIR(source_path="test.pptx", slides=slides or [])


def _make_image_shape(
    shape_id: int = 1,
    image_ext: str = "png",
    image_bytes: bytes = b"\x89PNG",
    classification: ImageClass | None = None,
    alt_text: str = "",
) -> ImageShapeIR:
    return ImageShapeIR(
        shape_id=shape_id,
        name=f"image_{shape_id}",
        kind=ShapeKind.IMAGE,
        image_bytes=image_bytes,
        image_ext=image_ext,
        image_format=image_ext,
        alt_text=alt_text,
        classification=classification,
        description=None,
    )


def _make_slide(shapes: list[Any]) -> SlideIR:  # type: ignore[misc]
    return SlideIR(index=0, shapes=shapes)


# ---------------------------------------------------------------------------
# AC1 — classification 채워진 ImageShapeIR + FakeDescriber ->
#         enrich_descriptions 후 description 슬롯 채워짐, 반환값 None (ADR-210)
# ---------------------------------------------------------------------------


def test_ac1_description_in_place_채워짐() -> None:
    """AC1: enrich_descriptions fills description slot in-place."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(classification=ImageClass.DIAGRAM)
    slide = _make_slide([shape])
    pres = _make_presentation([slide])
    describer = FakeDescriber(return_text="a nice diagram")

    result = enrich_descriptions(pres, describer)

    assert result is None, "enrich_descriptions must return None (ADR-210)"
    assert shape.description == "a nice diagram"


def test_ac1_반환값_None() -> None:
    """AC1: Return value of enrich_descriptions is None (ADR-210)."""
    from pptx_md.description_pipeline import enrich_descriptions

    pres = _make_presentation()
    ret = enrich_descriptions(pres, FakeDescriber())
    assert ret is None


def test_ac1_여러_이미지_모두_채워짐() -> None:
    """AC1: All ImageShapeIR descriptions are filled when multiple images exist."""
    from pptx_md.description_pipeline import enrich_descriptions

    shapes = [
        _make_image_shape(shape_id=1, classification=ImageClass.TEXT),
        _make_image_shape(shape_id=2, classification=ImageClass.PHOTO),
    ]
    pres = _make_presentation([_make_slide(shapes)])
    describer = FakeDescriber(return_text="desc")

    enrich_descriptions(pres, describer)

    assert shapes[0].description == "desc"
    assert shapes[1].description == "desc"
    assert len(describer.calls) == 2


# ---------------------------------------------------------------------------
# AC2 — classification=DIAGRAM -> shape_hint="diagram...", None -> shape_hint=None
# ---------------------------------------------------------------------------


def test_ac2_diagram_shape_hint_전달() -> None:
    """AC2: ImageClass.DIAGRAM -> shape_hint with 'diagram' passed to describer."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(classification=ImageClass.DIAGRAM)
    pres = _make_presentation([_make_slide([shape])])
    describer = FakeDescriber()

    enrich_descriptions(pres, describer)

    assert len(describer.calls) == 1
    hint = describer.calls[0]["shape_hint"]
    assert hint is not None
    assert "diagram" in hint.lower()


def test_ac2_classification_none_shape_hint_none() -> None:
    """AC2: classification=None -> shape_hint=None passed to describer."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(classification=None)
    pres = _make_presentation([_make_slide([shape])])
    describer = FakeDescriber()

    enrich_descriptions(pres, describer)

    assert len(describer.calls) == 1
    assert describer.calls[0]["shape_hint"] is None


def test_ac2_text_class_shape_hint() -> None:
    """AC2: ImageClass.TEXT -> shape_hint containing 'text'."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(classification=ImageClass.TEXT)
    pres = _make_presentation([_make_slide([shape])])
    describer = FakeDescriber()

    enrich_descriptions(pres, describer)

    hint = describer.calls[0]["shape_hint"]
    assert hint is not None
    assert "text" in hint.lower()


def test_ac2_photo_class_shape_hint() -> None:
    """AC2: ImageClass.PHOTO -> shape_hint containing 'photograph'."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(classification=ImageClass.PHOTO)
    pres = _make_presentation([_make_slide([shape])])
    describer = FakeDescriber()

    enrich_descriptions(pres, describer)

    hint = describer.calls[0]["shape_hint"]
    assert hint is not None
    assert "photograph" in hint.lower()


def test_ac2_logo_class_shape_hint() -> None:
    """AC2: ImageClass.LOGO -> shape_hint containing 'logo'."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(classification=ImageClass.LOGO)
    pres = _make_presentation([_make_slide([shape])])
    describer = FakeDescriber()

    enrich_descriptions(pres, describer)

    hint = describer.calls[0]["shape_hint"]
    assert hint is not None
    assert "logo" in hint.lower()


# ---------------------------------------------------------------------------
# _shape_hint unit tests (5 deterministic cases)
# ---------------------------------------------------------------------------


def test_shape_hint_text() -> None:
    """_shape_hint(TEXT) returns non-empty string with 'text'."""
    from pptx_md.description_pipeline import _shape_hint

    result = _shape_hint(ImageClass.TEXT)
    assert result is not None
    assert "text" in result.lower()


def test_shape_hint_diagram() -> None:
    """_shape_hint(DIAGRAM) returns non-empty string with 'diagram'."""
    from pptx_md.description_pipeline import _shape_hint

    result = _shape_hint(ImageClass.DIAGRAM)
    assert result is not None
    assert "diagram" in result.lower()


def test_shape_hint_photo() -> None:
    """_shape_hint(PHOTO) returns non-empty string with 'photograph'."""
    from pptx_md.description_pipeline import _shape_hint

    result = _shape_hint(ImageClass.PHOTO)
    assert result is not None
    assert "photograph" in result.lower()


def test_shape_hint_logo() -> None:
    """_shape_hint(LOGO) returns non-empty string with 'logo'."""
    from pptx_md.description_pipeline import _shape_hint

    result = _shape_hint(ImageClass.LOGO)
    assert result is not None
    assert "logo" in result.lower()


def test_shape_hint_none_returns_none() -> None:
    """_shape_hint(None) returns None."""
    from pptx_md.description_pipeline import _shape_hint

    result = _shape_hint(None)
    assert result is None


# ---------------------------------------------------------------------------
# AC3 (경계/NFR-08) — describer=None -> 모든 description None 유지, 예외 0건
# ---------------------------------------------------------------------------


def test_ac3_describer_none_모든_description_none_유지() -> None:
    """AC3: describer=None -> all descriptions stay None, no exceptions."""
    from pptx_md.description_pipeline import enrich_descriptions

    shapes = [
        _make_image_shape(shape_id=1, classification=ImageClass.DIAGRAM),
        _make_image_shape(shape_id=2, classification=ImageClass.TEXT),
    ]
    pres = _make_presentation([_make_slide(shapes)])

    enrich_descriptions(pres, None)

    for shape in shapes:
        assert (
            shape.description is None
        ), "description must remain None when describer=None"


def test_ac3_describer_none_예외_0건() -> None:
    """AC3: describer=None -> no exceptions raised."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape()
    pres = _make_presentation([_make_slide([shape])])

    try:
        enrich_descriptions(pres, None)
    except Exception as exc:  # noqa: BLE001
        pytest.fail(f"enrich_descriptions(describer=None) raised unexpectedly: {exc}")


def test_ac3_describer_none_빈_프레젠테이션() -> None:
    """AC3: describer=None with empty presentation -> returns None, no error."""
    from pptx_md.description_pipeline import enrich_descriptions

    pres = _make_presentation()
    result = enrich_descriptions(pres, None)
    assert result is None


# ---------------------------------------------------------------------------
# AC4 (예외/격리) — 첫 이미지 describe() 예외 -> 해당 None + WARNING, 후속 정상
# ---------------------------------------------------------------------------


def test_ac4_첫_이미지_예외_후속_정상_채워짐() -> None:
    """AC4: First image describe() raises -> that description=None, next filled."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape1 = _make_image_shape(shape_id=1, classification=ImageClass.TEXT)
    shape2 = _make_image_shape(shape_id=2, classification=ImageClass.PHOTO)
    pres = _make_presentation([_make_slide([shape1, shape2])])

    call_count = [0]

    class _FailFirstDescriber:
        def describe(
            self,
            image_bytes: bytes,
            image_ext: str,
            shape_hint: str | None,
        ) -> str:
            call_count[0] += 1
            if call_count[0] == 1:
                raise RuntimeError("simulated first failure")
            return "second ok"

    enrich_descriptions(pres, _FailFirstDescriber())

    assert shape1.description is None, "failed shape must have description=None"
    assert shape2.description == "second ok", "subsequent shape must be filled"


def test_ac4_예외_전파_안됨() -> None:
    """AC4: describe() exception is isolated — never propagated to caller."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape()
    pres = _make_presentation([_make_slide([shape])])

    from pptx_md.errors import DescribeError

    describer = FakeDescriber(raise_on_call=True, raise_exc=DescribeError("boom"))

    try:
        enrich_descriptions(pres, describer)
    except Exception as exc:  # noqa: BLE001
        pytest.fail(f"Exception must not propagate from enrich_descriptions: {exc}")

    assert shape.description is None


def test_ac4_예외시_warning_로그(caplog: pytest.LogCaptureFixture) -> None:
    """AC4: describe() failure emits a WARNING log with meta (not PII)."""
    from pptx_md.description_pipeline import enrich_descriptions
    from pptx_md.errors import DescribeError

    shape = _make_image_shape(shape_id=42, classification=ImageClass.DIAGRAM)
    pres = _make_presentation([_make_slide([shape])])
    describer = FakeDescriber(raise_on_call=True, raise_exc=DescribeError("api fail"))

    with caplog.at_level(logging.WARNING, logger="pptx_md.description_pipeline"):
        enrich_descriptions(pres, describer)

    assert any(
        "42" in r.message for r in caplog.records
    ), "WARNING must include shape_id=42"
    assert any(r.levelno == logging.WARNING for r in caplog.records)


def test_ac4_여러_이미지_중간_실패_나머지_정상() -> None:
    """AC4: Middle image fails -> only that one has None, others filled."""
    from pptx_md.description_pipeline import enrich_descriptions

    shapes = [
        _make_image_shape(shape_id=i, classification=ImageClass.TEXT)
        for i in range(1, 5)
    ]
    pres = _make_presentation([_make_slide(shapes)])

    call_count = [0]

    class _FailMiddle:
        def describe(
            self,
            image_bytes: bytes,
            image_ext: str,
            shape_hint: str | None,
        ) -> str:
            call_count[0] += 1
            if call_count[0] == 2:
                raise RuntimeError("middle fail")
            return f"ok_{call_count[0]}"

    enrich_descriptions(pres, _FailMiddle())

    assert shapes[0].description is not None
    assert shapes[1].description is None
    assert shapes[2].description is not None
    assert shapes[3].description is not None


# ---------------------------------------------------------------------------
# AC5 — GroupShapeIR 중첩 안 ImageShapeIR -> iter_shapes 재귀로 도달해 채워짐
# ---------------------------------------------------------------------------


def test_ac5_group_중첩_이미지_도달() -> None:
    """AC5: ImageShapeIR nested in GroupShapeIR is reached via iter_shapes."""
    from pptx_md.description_pipeline import enrich_descriptions

    inner_image = _make_image_shape(shape_id=10, classification=ImageClass.LOGO)
    group = GroupShapeIR(
        shape_id=5,
        name="group_5",
        kind=ShapeKind.GROUP,
        children=[inner_image],
    )
    text_shape = TextShapeIR(
        shape_id=1,
        name="title",
        kind=ShapeKind.TEXT,
    )
    slide = _make_slide([text_shape, group])
    pres = _make_presentation([slide])
    describer = FakeDescriber(return_text="logo desc")

    enrich_descriptions(pres, describer)

    assert inner_image.description == "logo desc"
    assert len(describer.calls) == 1


def test_ac5_깊은_중첩_그룹_이미지_도달() -> None:
    """AC5: ImageShapeIR nested two levels deep in GroupShapeIR is reached."""
    from pptx_md.description_pipeline import enrich_descriptions

    deep_image = _make_image_shape(shape_id=99, classification=ImageClass.PHOTO)
    inner_group = GroupShapeIR(
        shape_id=20,
        name="inner_group",
        kind=ShapeKind.GROUP,
        children=[deep_image],
    )
    outer_group = GroupShapeIR(
        shape_id=10,
        name="outer_group",
        kind=ShapeKind.GROUP,
        children=[inner_group],
    )
    pres = _make_presentation([_make_slide([outer_group])])
    describer = FakeDescriber(return_text="deep photo")

    enrich_descriptions(pres, describer)

    assert deep_image.description == "deep photo"


# ---------------------------------------------------------------------------
# AC6 — 로그에 image_bytes 원본·describe 반환 텍스트 전문 미출력 (NFR-06)
# ---------------------------------------------------------------------------


def test_ac6_로그에_image_bytes_미출력(caplog: pytest.LogCaptureFixture) -> None:
    """AC6: image_bytes raw content must not appear in any log output (NFR-06)."""
    from pptx_md.description_pipeline import enrich_descriptions

    sensitive_bytes = b"SENSITIVE_IMAGE_CONTENT_NFR06_TEST"
    shape = _make_image_shape(
        classification=ImageClass.TEXT,
        image_bytes=sensitive_bytes,
    )
    pres = _make_presentation([_make_slide([shape])])
    describer = FakeDescriber(return_text="some description")

    with caplog.at_level(logging.DEBUG, logger="pptx_md.description_pipeline"):
        enrich_descriptions(pres, describer)

    all_logs = caplog.text
    assert "SENSITIVE_IMAGE_CONTENT_NFR06_TEST" not in all_logs


def test_ac6_로그에_description_전문_미출력(caplog: pytest.LogCaptureFixture) -> None:
    """AC6: Full description text must not appear in log output (NFR-06)."""
    from pptx_md.description_pipeline import enrich_descriptions

    sensitive_desc = "HIGHLY_SENSITIVE_DESCRIPTION_TEXT_PII_CONTENT"
    shape = _make_image_shape(classification=ImageClass.DIAGRAM)
    pres = _make_presentation([_make_slide([shape])])
    describer = FakeDescriber(return_text=sensitive_desc)

    with caplog.at_level(logging.DEBUG, logger="pptx_md.description_pipeline"):
        enrich_descriptions(pres, describer)

    all_logs = caplog.text
    assert sensitive_desc not in all_logs


def test_ac6_로그에_메타만_출력(caplog: pytest.LogCaptureFixture) -> None:
    """AC6: Log contains only meta (shape_id, hint presence) not image data."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(shape_id=77, classification=ImageClass.DIAGRAM)
    pres = _make_presentation([_make_slide([shape])])
    describer = FakeDescriber()

    with caplog.at_level(logging.DEBUG, logger="pptx_md.description_pipeline"):
        enrich_descriptions(pres, describer)

    # At least one debug log should mention shape_id
    assert any("77" in r.message for r in caplog.records)


# ---------------------------------------------------------------------------
# AC7 — anthropic/openai grep 0건 + mypy exit 0 + ruff/black exit 0
# ---------------------------------------------------------------------------


def test_ac7_description_pipeline_sdk_import_0건() -> None:
    """AC7: description_pipeline.py contains no 'anthropic' or 'openai' imports."""
    source_file = _PROJECT_ROOT / "src" / "pptx_md" / "description_pipeline.py"
    source = source_file.read_text(encoding="utf-8")

    assert (
        "anthropic" not in source
    ), "description_pipeline.py must not reference 'anthropic'"
    assert "openai" not in source, "description_pipeline.py must not reference 'openai'"


def test_ac7_mypy_strict_exit0() -> None:
    """AC7: mypy src/ exits with code 0 (description_pipeline included)."""
    result = subprocess.run(
        [sys.executable, "-m", "mypy", "src/"],
        capture_output=True,
        text=True,
        cwd=str(_PROJECT_ROOT),
    )
    assert result.returncode == 0, (
        f"mypy src/ failed (exit {result.returncode}).\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )


def test_ac7_ruff_exit0() -> None:
    """AC7: ruff check . exits with code 0."""
    result = subprocess.run(
        [sys.executable, "-m", "ruff", "check", "."],
        capture_output=True,
        text=True,
        cwd=str(_PROJECT_ROOT),
    )
    assert result.returncode == 0, (
        f"ruff check failed (exit {result.returncode}).\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )


def test_ac7_black_check_exit0() -> None:
    """AC7: black --check . exits with code 0."""
    result = subprocess.run(
        [sys.executable, "-m", "black", "--check", "."],
        capture_output=True,
        text=True,
        cwd=str(_PROJECT_ROOT),
    )
    assert result.returncode == 0, (
        f"black --check failed (exit {result.returncode}).\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )


# ---------------------------------------------------------------------------
# EMF/WMF 변환 분기 테스트 (ADR-216)
# ---------------------------------------------------------------------------


def test_emf_변환_성공시_describe_호출() -> None:
    """EMF/WMF: successful vector->PNG conversion -> describe called with PNG bytes."""
    from pptx_md.description_pipeline import enrich_descriptions

    emf_shape = _make_image_shape(
        shape_id=3,
        image_ext="emf",
        image_bytes=b"EMF_BYTES",
        classification=ImageClass.DIAGRAM,
    )
    pres = _make_presentation([_make_slide([emf_shape])])
    describer = FakeDescriber(return_text="vector diagram")

    fake_png = b"\x89PNG"
    with patch(
        "pptx_md.description_pipeline.vector.convert_vector_to_png"
    ) as mock_conv:
        mock_conv.return_value = fake_png
        enrich_descriptions(pres, describer)

    assert emf_shape.description == "vector diagram"
    assert describer.calls[0]["image_ext"] == "png"
    assert describer.calls[0]["image_bytes"] == fake_png


def test_emf_변환_실패시_description_none() -> None:
    """EMF/WMF: conversion failure -> description=None, describe NOT called."""
    from pptx_md.description_pipeline import enrich_descriptions

    emf_shape = _make_image_shape(
        shape_id=4,
        image_ext="emf",
        image_bytes=b"EMF_BYTES",
        classification=ImageClass.DIAGRAM,
    )
    pres = _make_presentation([_make_slide([emf_shape])])
    describer = FakeDescriber()

    with patch(
        "pptx_md.description_pipeline.vector.convert_vector_to_png"
    ) as mock_conv:
        mock_conv.return_value = None
        enrich_descriptions(pres, describer)

    assert emf_shape.description is None
    assert (
        len(describer.calls) == 0
    ), "describe must NOT be called when conversion fails"


def test_wmf_변환_성공시_describe_호출() -> None:
    """WMF: successful conversion -> describe called."""
    from pptx_md.description_pipeline import enrich_descriptions

    wmf_shape = _make_image_shape(
        shape_id=5,
        image_ext="wmf",
        image_bytes=b"WMF_BYTES",
        classification=None,
    )
    pres = _make_presentation([_make_slide([wmf_shape])])
    describer = FakeDescriber(return_text="wmf result")

    with patch(
        "pptx_md.description_pipeline.vector.convert_vector_to_png"
    ) as mock_conv:
        mock_conv.return_value = b"\x89PNG_WMF"
        enrich_descriptions(pres, describer)

    assert wmf_shape.description == "wmf result"
    assert describer.calls[0]["shape_hint"] is None  # classification=None -> None


# ---------------------------------------------------------------------------
# import 가능성 테스트 (NFR-08)
# ---------------------------------------------------------------------------


def test_import_description_pipeline_성공() -> None:
    """NFR-08: import pptx_md.description_pipeline succeeds without any VLM SDK."""
    import importlib

    mod = importlib.import_module("pptx_md.description_pipeline")
    assert mod is not None


def test_description_pipeline_sdk_없이_enrich_호출() -> None:
    """NFR-08: enrich_descriptions works with FakeDescriber (no SDK needed)."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(classification=ImageClass.TEXT)
    pres = _make_presentation([_make_slide([shape])])

    enrich_descriptions(pres, FakeDescriber(return_text="no sdk needed"))
    assert shape.description == "no sdk needed"
