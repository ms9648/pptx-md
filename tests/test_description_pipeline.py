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
import threading
import time
from concurrent.futures import ThreadPoolExecutor as _RealThreadPoolExecutor
from typing import Any
from unittest.mock import patch

import pytest

from pptx_md import mermaid
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
        _make_image_shape(
            shape_id=1, image_bytes=b"\x89PNG_ONE", classification=ImageClass.TEXT
        ),
        _make_image_shape(
            shape_id=2, image_bytes=b"\x89PNG_TWO", classification=ImageClass.PHOTO
        ),
    ]
    pres = _make_presentation([_make_slide(shapes)])
    describer = FakeDescriber(return_text="desc")

    enrich_descriptions(pres, describer)

    assert shapes[0].description == "desc"
    assert shapes[1].description == "desc"
    # distinct image bytes -> 2 unique hashes -> 2 describe() calls (AC6, issue #68)
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

    shape1 = _make_image_shape(
        shape_id=1, image_bytes=b"\x89PNG_FIRST", classification=ImageClass.TEXT
    )
    shape2 = _make_image_shape(
        shape_id=2, image_bytes=b"\x89PNG_SECOND", classification=ImageClass.PHOTO
    )
    pres = _make_presentation([_make_slide([shape1, shape2])])

    class _FailFirstDescriber:
        def describe(
            self,
            image_bytes: bytes,
            image_ext: str,
            shape_hint: str | None,
        ) -> str:
            if image_bytes == b"\x89PNG_FIRST":
                raise RuntimeError("simulated first failure")
            return "second ok"

    # max_workers=1: distinct bytes -> distinct hash groups; identity-based
    # failure (not call order) keeps this test valid under concurrency (#68).
    enrich_descriptions(pres, _FailFirstDescriber(), max_workers=1)

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
        _make_image_shape(
            shape_id=i,
            image_bytes=b"\x89PNG_" + str(i).encode("ascii"),
            classification=ImageClass.TEXT,
        )
        for i in range(1, 5)
    ]
    pres = _make_presentation([_make_slide(shapes)])

    fail_bytes = b"\x89PNG_2"

    class _FailMiddle:
        def describe(
            self,
            image_bytes: bytes,
            image_ext: str,
            shape_hint: str | None,
        ) -> str:
            if image_bytes == fail_bytes:
                raise RuntimeError("middle fail")
            return f"ok_{image_bytes!r}"

    # identity-based failure (not call-order-based) is valid under
    # concurrency (#68); distinct bytes per shape -> distinct hash groups.
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


# ===========================================================================
# FR-27 (issue #68, ARCH-M12) — hash caching / concurrency / Mermaid diagram
#
# Test naming: test_ac<N>_fr27_<desc> where <N> matches issue #68's AC
# numbering (NOT the FR-10 AC numbering used above in this same file).
# ===========================================================================


class HashAwareDescriber:
    """FakeDescriber whose response is a deterministic function of the input
    bytes (so distinct hashes are distinguishable in assertions)."""

    def __init__(self) -> None:
        self.calls: list[bytes] = []

    def describe(
        self, image_bytes: bytes, image_ext: str, shape_hint: str | None
    ) -> str:
        self.calls.append(image_bytes)
        return f"desc-for-{image_bytes!r}"


class SelectiveFailDescriber:
    """Raises only when the input bytes match ``fail_bytes``."""

    def __init__(self, fail_bytes: bytes, ok_text: str = "ok") -> None:
        self.fail_bytes = fail_bytes
        self.ok_text = ok_text
        self.calls: list[bytes] = []

    def describe(
        self, image_bytes: bytes, image_ext: str, shape_hint: str | None
    ) -> str:
        self.calls.append(image_bytes)
        if image_bytes == self.fail_bytes:
            raise RuntimeError("simulated failure for this image only")
        return self.ok_text


class ConcurrencyCountingDescriber:
    """Counts the maximum number of concurrently-active describe() calls."""

    def __init__(self, delay: float = 0.03) -> None:
        self.delay = delay
        self._lock = threading.Lock()
        self._active = 0
        self.max_observed = 0
        self.calls = 0

    def describe(
        self, image_bytes: bytes, image_ext: str, shape_hint: str | None
    ) -> str:
        with self._lock:
            self._active += 1
            self.max_observed = max(self.max_observed, self._active)
            self.calls += 1
        time.sleep(self.delay)
        with self._lock:
            self._active -= 1
        return "counted"


class MermaidFenceDescriber:
    """Returns a fixed VLM response text (mermaid-fenced or plain)."""

    def __init__(self, text: str) -> None:
        self.text = text
        self.calls: list[str | None] = []

    def describe(
        self, image_bytes: bytes, image_ext: str, shape_hint: str | None
    ) -> str:
        self.calls.append(shape_hint)
        return self.text


# ---------------------------------------------------------------------------
# issue AC2 (경계, NFR-08 회귀) — describer=None -> pool/cache 미생성, 예외 0
# ---------------------------------------------------------------------------


def test_ac2_fr27_describer_none_pool_미생성() -> None:
    """issue AC2: describer=None -> ThreadPoolExecutor is never constructed."""
    from pptx_md.description_pipeline import enrich_descriptions

    shapes = [
        _make_image_shape(shape_id=1, image_bytes=b"AAA"),
        _make_image_shape(shape_id=2, image_bytes=b"BBB"),
    ]
    pres = _make_presentation([_make_slide(shapes)])

    with patch("pptx_md.description_pipeline.ThreadPoolExecutor") as mock_pool_cls:
        enrich_descriptions(pres, None)

    mock_pool_cls.assert_not_called()
    for shape in shapes:
        assert shape.description is None


# ---------------------------------------------------------------------------
# issue AC3 (예외, ADR-214 회귀) — 특정 이미지만 예외 -> 해당만 None, 나머지 정상
# ---------------------------------------------------------------------------


def test_ac3_fr27_특정_이미지만_예외_나머지_정상() -> None:
    """issue AC3: describer raises for one specific image only -> isolated failure."""
    from pptx_md.description_pipeline import enrich_descriptions

    fail_bytes = b"WILL_FAIL_BYTES"
    shapes = [
        _make_image_shape(shape_id=1, image_bytes=b"OK_ONE"),
        _make_image_shape(shape_id=2, image_bytes=fail_bytes),
        _make_image_shape(shape_id=3, image_bytes=b"OK_TWO"),
    ]
    pres = _make_presentation([_make_slide(shapes)])
    describer = SelectiveFailDescriber(fail_bytes=fail_bytes, ok_text="fine")

    try:
        enrich_descriptions(pres, describer)
    except Exception as exc:  # noqa: BLE001
        pytest.fail(f"enrich_descriptions must not raise: {exc}")

    assert shapes[0].description == "fine"
    assert shapes[1].description is None
    assert shapes[2].description == "fine"


# ---------------------------------------------------------------------------
# issue AC4 (Mermaid 옵션) — DIAGRAM + diagram_mermaid=True -> flowchart 펜스
# ---------------------------------------------------------------------------


def test_ac4_fr27_diagram_mermaid_활성_펜스_산출() -> None:
    """issue AC4: DIAGRAM + diagram_mermaid=True + valid fence response -> rendered."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(
        shape_id=1, image_bytes=b"DIAG1", classification=ImageClass.DIAGRAM
    )
    pres = _make_presentation([_make_slide([shape])])
    vlm_response = "Here is the flow:\n```mermaid\nflowchart TD\nA-->B\n```\nend."
    describer = MermaidFenceDescriber(vlm_response)

    enrich_descriptions(pres, describer, diagram_mermaid=True)

    assert shape.description is not None
    assert shape.description.startswith("```mermaid")
    assert "flowchart TD" in shape.description
    # hint must have been augmented with the mermaid request suffix (ADR-611)
    assert describer.calls[0] is not None
    assert mermaid.DIAGRAM_HINT_SUFFIX in describer.calls[0]


def test_ac4_fr27_옵션_비활성_기본값_평문_설명() -> None:
    """issue AC4: diagram_mermaid=False (default) -> plain text, no mermaid hint."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(
        shape_id=1, image_bytes=b"DIAG2", classification=ImageClass.DIAGRAM
    )
    pres = _make_presentation([_make_slide([shape])])
    vlm_response = "```mermaid\nflowchart TD\nA-->B\n```"
    describer = MermaidFenceDescriber(vlm_response)

    enrich_descriptions(pres, describer)  # diagram_mermaid defaults to False

    assert shape.description == vlm_response
    assert describer.calls[0] is not None
    assert mermaid.DIAGRAM_HINT_SUFFIX not in describer.calls[0]


# ---------------------------------------------------------------------------
# issue AC5 (Mermaid fallback) — 구조화 불가 응답 -> 일반 텍스트, 본문 손실 0
# ---------------------------------------------------------------------------


def test_ac5_fr27_구조화_불가_응답_평문_fallback() -> None:
    """issue AC5: unstructured VLM response -> falls back to plain text, no loss."""
    from pptx_md.description_pipeline import enrich_descriptions

    shape = _make_image_shape(
        shape_id=1, image_bytes=b"DIAG3", classification=ImageClass.DIAGRAM
    )
    pres = _make_presentation([_make_slide([shape])])
    plain_response = "This is just a plain description with no mermaid fence."
    describer = MermaidFenceDescriber(plain_response)

    try:
        enrich_descriptions(pres, describer, diagram_mermaid=True)
    except Exception as exc:  # noqa: BLE001
        pytest.fail(f"enrich_descriptions must not raise on fallback: {exc}")

    assert shape.description == plain_response  # zero content loss


# ---------------------------------------------------------------------------
# issue AC6 (해시 캐싱) — 동일 바이트 2회 이상 -> describe 호출 수 == 고유 해시 수
# ---------------------------------------------------------------------------


def test_ac6_fr27_동일_바이트_중복_이미지_describe_1회() -> None:
    """issue AC6: identical image bytes appearing 3x -> describe called once
    for that hash (call count == unique hash count)."""
    from pptx_md.description_pipeline import enrich_descriptions

    logo_bytes = b"LOGO_REPEATED_BYTES"
    shapes = [
        _make_image_shape(shape_id=1, image_bytes=logo_bytes),
        _make_image_shape(shape_id=2, image_bytes=logo_bytes),
        _make_image_shape(shape_id=3, image_bytes=logo_bytes),
        _make_image_shape(shape_id=4, image_bytes=b"DIFFERENT_BYTES"),
    ]
    pres = _make_presentation([_make_slide(shapes)])
    describer = HashAwareDescriber()

    enrich_descriptions(pres, describer)

    # 2 unique hashes -> exactly 2 describe() calls (not 4)
    assert len(describer.calls) == 2
    assert shapes[0].description == shapes[1].description == shapes[2].description
    assert shapes[3].description != shapes[0].description


def test_ac6_fr27_원본_바이트_미저장_로그_0(caplog: pytest.LogCaptureFixture) -> None:
    """issue AC6: original image bytes never appear in log output (NFR-06)."""
    from pptx_md.description_pipeline import enrich_descriptions

    sensitive = b"NFR06_SENSITIVE_HASH_INPUT_BYTES"
    shapes = [
        _make_image_shape(shape_id=1, image_bytes=sensitive),
        _make_image_shape(shape_id=2, image_bytes=sensitive),
    ]
    pres = _make_presentation([_make_slide(shapes)])
    describer = HashAwareDescriber()

    with caplog.at_level(logging.DEBUG, logger="pptx_md.description_pipeline"):
        enrich_descriptions(pres, describer)

    assert b"NFR06_SENSITIVE_HASH_INPUT_BYTES".decode() not in caplog.text


# ---------------------------------------------------------------------------
# issue AC7 (동시성 제한) — 동시 describe 호출 수 <= N, 옵션 조정 가능
# ---------------------------------------------------------------------------


def test_ac7_fr27_동시성_상한_준수_기본값4() -> None:
    """issue AC7: at no point do concurrent describe() calls exceed max_workers."""
    from pptx_md.description_pipeline import enrich_descriptions

    shapes = [
        _make_image_shape(shape_id=i, image_bytes=f"IMG_{i}".encode())
        for i in range(10)
    ]
    pres = _make_presentation([_make_slide(shapes)])
    describer = ConcurrencyCountingDescriber(delay=0.03)

    enrich_descriptions(pres, describer, max_workers=3)

    assert describer.calls == 10
    assert describer.max_observed <= 3


def test_ac7_fr27_동시성_옵션_조정_max_workers_1_순차() -> None:
    """issue AC7: max_workers=1 -> effectively sequential (max_observed == 1)."""
    from pptx_md.description_pipeline import enrich_descriptions

    shapes = [
        _make_image_shape(shape_id=i, image_bytes=f"SEQ_{i}".encode()) for i in range(5)
    ]
    pres = _make_presentation([_make_slide(shapes)])
    describer = ConcurrencyCountingDescriber(delay=0.02)

    enrich_descriptions(pres, describer, max_workers=1)

    assert describer.max_observed == 1


def test_ac7_fr27_max_workers_클램프_고유수_상한() -> None:
    """issue AC7: max_workers larger than unique image count is clamped down."""
    from pptx_md.description_pipeline import enrich_descriptions

    shapes = [
        _make_image_shape(shape_id=i, image_bytes=f"CLAMP_{i}".encode())
        for i in range(3)
    ]
    pres = _make_presentation([_make_slide(shapes)])

    captured_max_workers: list[int | None] = []

    class RecordingExecutor(_RealThreadPoolExecutor):
        def __init__(self, max_workers: int | None = None, *a: Any, **kw: Any) -> None:
            captured_max_workers.append(max_workers)
            super().__init__(max_workers=max_workers, *a, **kw)

    describer = FakeDescriber(return_text="clamped")
    with patch("pptx_md.description_pipeline.ThreadPoolExecutor", RecordingExecutor):
        enrich_descriptions(pres, describer, max_workers=100)

    assert captured_max_workers == [3]


def test_ac7_fr27_max_workers_0이하_방어_최소1() -> None:
    """issue AC7: max_workers<=0 is clamped to 1 (defensive floor)."""
    from pptx_md.description_pipeline import enrich_descriptions

    shapes = [_make_image_shape(shape_id=1, image_bytes=b"ZERO_WORKERS")]
    pres = _make_presentation([_make_slide(shapes)])

    captured_max_workers: list[int | None] = []

    class RecordingExecutor(_RealThreadPoolExecutor):
        def __init__(self, max_workers: int | None = None, *a: Any, **kw: Any) -> None:
            captured_max_workers.append(max_workers)
            super().__init__(max_workers=max_workers, *a, **kw)

    describer = FakeDescriber(return_text="floor")
    with patch("pptx_md.description_pipeline.ThreadPoolExecutor", RecordingExecutor):
        enrich_descriptions(pres, describer, max_workers=0)

    assert captured_max_workers == [1]


# ---------------------------------------------------------------------------
# issue AC8 (결정성 골든 회귀) — max_workers=1 vs N -> 최종 매핑 동일
# ---------------------------------------------------------------------------


def test_ac8_fr27_동시성_수준과_무관하게_결과_동일() -> None:
    """issue AC8: same IR described with max_workers=1 vs 8 -> identical mapping."""
    from pptx_md.description_pipeline import enrich_descriptions

    def _build() -> list[ImageShapeIR]:
        return [
            _make_image_shape(shape_id=i, image_bytes=f"DET_{i % 4}".encode())
            for i in range(9)
        ]

    shapes_seq = _build()
    pres_seq = _make_presentation([_make_slide(shapes_seq)])
    enrich_descriptions(pres_seq, HashAwareDescriber(), max_workers=1)

    shapes_par = _build()
    pres_par = _make_presentation([_make_slide(shapes_par)])
    enrich_descriptions(pres_par, HashAwareDescriber(), max_workers=8)

    seq_result = [s.description for s in shapes_seq]
    par_result = [s.description for s in shapes_par]
    assert seq_result == par_result


# ---------------------------------------------------------------------------
# issue AC9 (커버리지 지표) — 7장 골든 픽스처 + describer 주입 -> 커버리지 100%
# ---------------------------------------------------------------------------


def test_ac9_fr27_7장_이미지_커버리지_100퍼센트() -> None:
    """issue AC9: 7 distinct images + working describer -> 100% description coverage."""
    from pptx_md.description_pipeline import enrich_descriptions

    shapes = [
        _make_image_shape(shape_id=i, image_bytes=f"GOLDEN_{i}".encode())
        for i in range(1, 8)
    ]
    pres = _make_presentation([_make_slide(shapes)])
    describer = HashAwareDescriber()

    enrich_descriptions(pres, describer)

    filled = sum(1 for s in shapes if s.description is not None)
    assert filled == 7
    assert filled / len(shapes) == 1.0


def test_ac9_fr27_실패_이미지_FR22_마커로_대체() -> None:
    """issue AC9: one of 7 images fails -> assembled Markdown falls back to the
    standard FR-22 positional marker for that image; others render description."""
    from pptx_md.assembler import assemble_document
    from pptx_md.description_pipeline import enrich_descriptions

    fail_bytes = b"GOLDEN_FAIL_5"
    shapes = [
        _make_image_shape(shape_id=i, image_bytes=f"GOLDEN_{i}".encode())
        for i in range(1, 8)
        if i != 5
    ]
    shapes.insert(4, _make_image_shape(shape_id=5, image_bytes=fail_bytes))
    slide = _make_slide(shapes)
    pres = PresentationIR(source_path="golden.pptx", slides=[slide])

    describer = SelectiveFailDescriber(fail_bytes=fail_bytes, ok_text="golden desc")
    enrich_descriptions(pres, describer)

    md = assemble_document(pres)

    assert md.count("golden desc") == 6
    assert "슬라이드 1 이미지" in md  # FR-22 standard positional marker present


# ---------------------------------------------------------------------------
# issue AC10 (조건부 COM) — convert_via_com 부재, COM 모듈 신설 금지 (N/A 부채)
# ---------------------------------------------------------------------------


def test_ac10_fr27_convert_via_com_부재_확인() -> None:
    """issue AC10: convert_via_com does not exist anywhere in pptx_md (N/A debt,
    ARCH-M12 §9.3/§10). No COM module was introduced by this change."""
    import pptx_md

    assert not hasattr(pptx_md, "convert_via_com")

    pkg_dir = _PROJECT_ROOT / "src" / "pptx_md"
    for py_file in pkg_dir.glob("*.py"):
        source = py_file.read_text(encoding="utf-8")
        assert "convert_via_com" not in source
        assert "win32com" not in source


# ---------------------------------------------------------------------------
# NFR-08 — SDK import 0 grep on description_pipeline.py (M12 addition)
# ---------------------------------------------------------------------------


def test_fr27_sdk_import_0건_hashlib_concurrent_futures_only() -> None:
    """NFR-08: description_pipeline.py's new imports are stdlib only."""
    source_file = _PROJECT_ROOT / "src" / "pptx_md" / "description_pipeline.py"
    source = source_file.read_text(encoding="utf-8")

    assert "anthropic" not in source
    assert "openai" not in source
    assert "import hashlib" in source
    assert "from concurrent.futures import" in source
