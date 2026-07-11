"""Tests for src/pptx_md/slide_render.py — FR-30 슬라이드 렌더 (이슈 #91 AC1~AC8).

Test naming convention: test_ac<N>_<description> (AC 번호는 이슈 #91 본문 기준).

pypdfium2/soffice 미가용 환경 대비:
    - soffice 는 monkeypatch(``shutil.which``/``subprocess.run``)로 항상 시뮬레이션
      가능하므로 대부분의 테스트는 실제 LibreOffice 설치 없이 실행된다.
    - pypdfium2 로 만든 실제(합성) PDF 바이트를 fake soffice 출력물로 사용해 진짜
      pypdfium2 rasterisation 경로까지 검증한다(``_requires_pypdfium2`` 로 게이팅).
    - pypdfium2 가 설치되지 않은 환경에서는 해당 테스트가 skip 되고, AC7(지연 import)
      과 soffice-부재/타임아웃 등 경로는 pypdfium2 유무와 무관하게 항상 실행된다.
"""

from __future__ import annotations

import ast
import io
import logging
import shutil as _shutil
import subprocess
import sys
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock

import pytest

from pptx_md.slide_render import (
    RENDER_DPI_DEFAULT,
    RENDER_TIMEOUT_S,
    render_deck,
    render_slide,
    slide_render_available,
)

try:
    import pypdfium2 as _pdfium

    _PYPDFIUM2_AVAILABLE = True
except Exception:  # pragma: no cover - depends on optional [render] extra
    _PYPDFIUM2_AVAILABLE = False

_requires_pypdfium2 = pytest.mark.skipif(
    not _PYPDFIUM2_AVAILABLE, reason="pypdfium2 not installed ([render] extra)"
)


def _make_fake_pdf_bytes(
    num_pages: int = 1, width: float = 720.0, height: float = 405.0
) -> bytes:
    """실제 pypdfium2 로 최소 유효 PDF 생성 (fake soffice 출력 대체용, 테스트 전용).

    width/height 는 PDF point 단위(72pt = 1in). 기본값은 16:9 비율(720x405).
    """
    doc = _pdfium.PdfDocument.new()
    for _ in range(num_pages):
        doc.new_page(width, height)
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()


def _fake_soffice_run_factory(pdf_bytes: bytes, call_counter: list[int] | None = None):
    """subprocess.run 을 대체해 ``<outdir>/deck.pdf`` 에 pdf_bytes 를 기록하는 fake."""

    def _fake_run(cmd: list[str], **kwargs: Any) -> MagicMock:
        if call_counter is not None:
            call_counter.append(1)
        out_dir = Path(cmd[cmd.index("--outdir") + 1])
        out_file = out_dir / "deck.pdf"
        out_file.write_bytes(pdf_bytes)
        mock = MagicMock()
        mock.returncode = 0
        return mock

    return _fake_run


# ===========================================================================
# AC1 — render_deck/render_slide 성공 경로 + soffice 1회 호출 + 편의 시그니처
# ===========================================================================


@_requires_pypdfium2
def test_ac1_render_slide_returns_png_bytes(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """AC1: soffice+pypdfium2 가용 시 render_slide() 가 PNG bytes 를 반환한다."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    pdf_bytes = _make_fake_pdf_bytes(num_pages=3)
    monkeypatch.setattr(subprocess, "run", _fake_soffice_run_factory(pdf_bytes))

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake-pptx-bytes")

    result = render_slide(dummy_pptx, 0)
    assert result is not None
    assert result[:4] == b"\x89PNG"


@_requires_pypdfium2
def test_ac1_render_slide_delegates_to_render_deck(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """AC1: render_slide(path, n) == render_deck(path, [n])[n] (편의 시그니처)."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    pdf_bytes = _make_fake_pdf_bytes(num_pages=2)
    monkeypatch.setattr(subprocess, "run", _fake_soffice_run_factory(pdf_bytes))

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake-pptx-bytes")

    via_slide = render_slide(dummy_pptx, 1)
    via_deck = render_deck(dummy_pptx, [1])[1]
    assert via_slide is not None
    assert via_deck is not None
    assert via_slide[:4] == via_deck[:4] == b"\x89PNG"


@_requires_pypdfium2
def test_ac1_render_deck_calls_soffice_exactly_once(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """AC1: 여러 인덱스를 요청해도 soffice(--convert-to pdf) 는 정확히 1회 호출."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    pdf_bytes = _make_fake_pdf_bytes(num_pages=5)
    call_counter: list[int] = []
    monkeypatch.setattr(
        subprocess, "run", _fake_soffice_run_factory(pdf_bytes, call_counter)
    )

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake-pptx-bytes")

    result = render_deck(dummy_pptx, [0, 2, 4])
    assert (
        len(call_counter) == 1
    ), "soffice subprocess 는 deck 당 정확히 1회 호출되어야 함"
    assert set(result.keys()) == {0, 2, 4}
    assert all(v is not None and v[:4] == b"\x89PNG" for v in result.values())


def test_ac1_render_slide_none_when_soffice_missing(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """AC1: soffice 미가용 시 예외 없이 None 반환 (graceful skip)."""
    monkeypatch.setattr(_shutil, "which", lambda _name: None)
    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake")

    result = render_slide(dummy_pptx, 0)
    assert result is None


def test_ac1_render_page_to_png_returns_none_on_render_exception() -> None:
    """AC1 내부: 단일 페이지 렌더 중 예외 -> 해당 페이지만 None (예외 미전파)."""
    import pptx_md.slide_render as sr

    class _RaisingPdfDoc:
        def __getitem__(self, index: int) -> Any:
            raise RuntimeError("simulated page decode failure")

    result = sr._render_page_to_png(_RaisingPdfDoc(), 0, scale=1.0)
    assert result is None


def test_ac1_render_page_to_png_returns_none_on_empty_bytes() -> None:
    """AC1 내부: 렌더 결과 PNG bytes 가 비어있으면 None."""
    import pptx_md.slide_render as sr

    class _FakePilImage:
        def save(self, buf: Any, format: str) -> None:
            pass  # buf 에 아무것도 쓰지 않음 -> getvalue() == b""

    class _FakeBitmap:
        def to_pil(self) -> _FakePilImage:
            return _FakePilImage()

    class _FakePage:
        def render(self, scale: float) -> _FakeBitmap:
            return _FakeBitmap()

    class _FakePdfDoc:
        def __getitem__(self, index: int) -> _FakePage:
            return _FakePage()

    result = sr._render_page_to_png(_FakePdfDoc(), 0, scale=1.0)
    assert result is None


# ===========================================================================
# AC2 — soffice/pypdfium2 미가용·손상 PPTX·범위초과 인덱스 → None (예외 미전파)
# ===========================================================================


def test_ac2_soffice_missing_all_indices_none(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """AC2: soffice 미가용 -> 요청한 전 인덱스가 None, 경고 로그 1건 이상."""
    monkeypatch.setattr(_shutil, "which", lambda _name: None)
    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake")

    with caplog.at_level(logging.WARNING, logger="pptx_md.slide_render"):
        result = render_deck(dummy_pptx, [0, 1, 2])

    assert result == {0: None, 1: None, 2: None}
    assert any(r.levelno >= logging.WARNING for r in caplog.records)


def test_ac2_pypdfium2_missing_all_indices_none(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """AC2: soffice 는 있으나 pypdfium2 import 실패 -> 전 인덱스 None, 예외 미전파."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    # pypdfium2 가 실제로 설치돼 있어도 import 실패를 강제 시뮬레이션(sys.modules 오염).
    monkeypatch.setitem(sys.modules, "pypdfium2", None)

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake")

    with caplog.at_level(logging.WARNING, logger="pptx_md.slide_render"):
        result = render_deck(dummy_pptx, [0, 1])

    assert result == {0: None, 1: None}
    assert any(r.levelno >= logging.WARNING for r in caplog.records)


def test_ac2_corrupt_pptx_soffice_nonzero_exit_returns_none(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """AC2: 손상 PPTX 로 soffice 변환 실패(returncode!=0) -> None, 예외 미전파."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")

    def _fake_run_fail(cmd: list[str], **kwargs: Any) -> MagicMock:
        mock = MagicMock()
        mock.returncode = 1
        return mock

    monkeypatch.setattr(subprocess, "run", _fake_run_fail)

    corrupt_pptx = tmp_path / "corrupt.pptx"
    corrupt_pptx.write_bytes(b"not a real pptx file")

    with caplog.at_level(logging.WARNING, logger="pptx_md.slide_render"):
        result = render_deck(corrupt_pptx, [0])

    assert result == {0: None}
    assert any(r.levelno >= logging.WARNING for r in caplog.records)


@_requires_pypdfium2
def test_ac2_out_of_range_index_returns_none_others_succeed(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """AC2 (경계): 범위초과/음수 인덱스는 None, 유효 인덱스는 정상 렌더."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    pdf_bytes = _make_fake_pdf_bytes(num_pages=3)  # valid indices: 0,1,2
    monkeypatch.setattr(subprocess, "run", _fake_soffice_run_factory(pdf_bytes))

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake-pptx-bytes")

    result = render_deck(dummy_pptx, [0, 5, -1])
    assert result[5] is None
    assert result[-1] is None
    assert result[0] is not None
    assert result[0][:4] == b"\x89PNG"


def test_ac2_missing_input_file_returns_none(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """AC2: 존재하지 않는 입력 경로 -> None, 예외 미전파."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    missing_path = tmp_path / "does-not-exist.pptx"

    result = render_deck(missing_path, [0])
    assert result == {0: None}


def test_ac2_empty_indices_returns_empty_dict_no_soffice_call(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """AC2 (경계): 빈 indices -> 빈 dict, soffice 미호출(무의미한 변환 회피)."""
    which_called = False

    def _spy_which(name: str) -> str | None:
        nonlocal which_called
        which_called = True
        return "/usr/bin/soffice"

    monkeypatch.setattr(_shutil, "which", _spy_which)
    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake")

    result = render_deck(dummy_pptx, [])
    assert result == {}
    assert not which_called, "빈 indices 에서는 soffice 탐색조차 불필요"


def test_ac2_soffice_ok_but_pdf_output_missing(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """AC2: soffice returncode=0 이나 예상 PDF 산출물이 없음 -> None, WARNING 로그."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")

    def _fake_run_no_output(cmd: list[str], **kwargs: Any) -> MagicMock:
        mock = MagicMock()
        mock.returncode = 0
        return mock  # deck.pdf 를 실제로 만들지 않음

    monkeypatch.setattr(subprocess, "run", _fake_run_no_output)

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake-pptx-bytes")

    with caplog.at_level(logging.WARNING, logger="pptx_md.slide_render"):
        result = render_deck(dummy_pptx, [0])

    assert result == {0: None}
    assert any(r.levelno >= logging.WARNING for r in caplog.records)


@_requires_pypdfium2
def test_ac2_soffice_ok_but_pdf_output_corrupt(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """AC2: PDF 산출물이 손상돼 pypdfium2 가 열지 못함 -> None, WARNING 로그."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    monkeypatch.setattr(
        subprocess, "run", _fake_soffice_run_factory(b"not-a-real-pdf-file")
    )

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake-pptx-bytes")

    with caplog.at_level(logging.WARNING, logger="pptx_md.slide_render"):
        result = render_deck(dummy_pptx, [0])

    assert result == {0: None}
    assert any(r.levelno >= logging.WARNING for r in caplog.records)


def test_ac2_soffice_raises_generic_exception_all_none(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """AC2: subprocess.run 이 예상 밖 Exception 을 던지면 전 인덱스 None."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")

    def _raise_os_error(*args: Any, **kwargs: Any) -> None:
        raise OSError("디스크 공간 부족 시뮬레이션")

    monkeypatch.setattr(subprocess, "run", _raise_os_error)

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake-pptx-bytes")

    with caplog.at_level(logging.WARNING, logger="pptx_md.slide_render"):
        try:
            result = render_deck(dummy_pptx, [0, 1])
        except Exception as exc:
            pytest.fail(f"예상 밖 예외 발생 시 예외가 전파됨: {exc}")

    assert result == {0: None, 1: None}
    assert any(r.levelno >= logging.WARNING for r in caplog.records)


# ===========================================================================
# AC3 — 종횡비 ±1% 이내, dpi 파라미터 반영(scale=dpi/72)
# ===========================================================================


@_requires_pypdfium2
def test_ac3_aspect_ratio_within_1_percent(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """AC3: 16:9 슬라이드(720x405pt) 렌더 결과 PNG 종횡비가 원본과 ±1% 이내."""
    from PIL import Image

    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    pdf_bytes = _make_fake_pdf_bytes(num_pages=1, width=720.0, height=405.0)
    monkeypatch.setattr(subprocess, "run", _fake_soffice_run_factory(pdf_bytes))

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake-pptx-bytes")

    png_bytes = render_slide(dummy_pptx, 0)
    assert png_bytes is not None
    img = Image.open(io.BytesIO(png_bytes))

    expected_ratio = 720.0 / 405.0
    actual_ratio = img.width / img.height
    assert abs(actual_ratio - expected_ratio) / expected_ratio <= 0.01


@_requires_pypdfium2
def test_ac3_dpi_parameter_scales_output_size(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """AC3: dpi 파라미터가 scale=dpi/72 로 반영되어 출력 픽셀 크기가 비례한다."""
    from PIL import Image

    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    pdf_bytes = _make_fake_pdf_bytes(num_pages=1, width=720.0, height=405.0)
    monkeypatch.setattr(subprocess, "run", _fake_soffice_run_factory(pdf_bytes))

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake-pptx-bytes")

    png_72 = render_slide(dummy_pptx, 0, dpi=72)  # scale=1.0 -> ~720x405
    png_150 = render_slide(dummy_pptx, 0, dpi=150)  # scale~2.083 -> ~1500x844

    assert png_72 is not None and png_150 is not None
    img_72 = Image.open(io.BytesIO(png_72))
    img_150 = Image.open(io.BytesIO(png_150))

    assert img_72.width == pytest.approx(720, abs=2)
    assert img_72.height == pytest.approx(405, abs=2)
    # dpi 150 이 dpi 72 대비 약 150/72 배 커야 함 (±5% 허용)
    ratio = img_150.width / img_72.width
    assert abs(ratio - (150 / 72)) / (150 / 72) <= 0.05


def test_ac3_render_dpi_default_is_150() -> None:
    """AC3 전제: RENDER_DPI_DEFAULT == 150 (ARCH-v020 §3.1 명시값)."""
    assert RENDER_DPI_DEFAULT == 150


# ===========================================================================
# AC4 — 타임아웃 -> None (vector.py 규약 계승)
# ===========================================================================


def test_ac4_soffice_timeout_returns_none(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """AC4: TimeoutExpired -> 전 인덱스 None, 예외 전파 없음, WARNING 로그."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")

    def _raise_timeout(*args: Any, **kwargs: Any) -> None:
        raise subprocess.TimeoutExpired(cmd=["soffice"], timeout=RENDER_TIMEOUT_S)

    monkeypatch.setattr(subprocess, "run", _raise_timeout)

    dummy_pptx = tmp_path / "deck.pptx"
    dummy_pptx.write_bytes(b"fake-pptx-bytes")

    with caplog.at_level(logging.WARNING, logger="pptx_md.slide_render"):
        try:
            result = render_deck(dummy_pptx, [0, 1])
        except Exception as exc:
            pytest.fail(f"TimeoutExpired 시 예외가 전파됨: {exc}")

    assert result == {0: None, 1: None}
    assert any(r.levelno >= logging.WARNING for r in caplog.records)


def test_ac4_render_timeout_s_constant_defined() -> None:
    """AC4 전제: RENDER_TIMEOUT_S 는 vector.VECTOR_TIMEOUT_S 규약을 계승한 int 상수."""
    assert isinstance(RENDER_TIMEOUT_S, int)
    assert RENDER_TIMEOUT_S > 0


# ===========================================================================
# AC5 — slide_render_available() 게이트: soffice AND pypdfium2 둘 다 확인
# ===========================================================================


def test_ac5_available_both_present(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC5: soffice+pypdfium2 둘 다 가용 -> True."""
    import pptx_md.slide_render as sr

    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    monkeypatch.setattr(sr, "_pypdfium2_importable", lambda: True)
    assert sr.slide_render_available() is True


def test_ac5_available_soffice_missing(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC5: soffice 미가용 -> False (pypdfium2 가용 여부 무관)."""
    import pptx_md.slide_render as sr

    monkeypatch.setattr(_shutil, "which", lambda _name: None)
    monkeypatch.setattr(sr, "_pypdfium2_importable", lambda: True)
    assert sr.slide_render_available() is False


def test_ac5_available_pypdfium2_missing(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC5: pypdfium2 미가용 -> False (soffice 가용 여부 무관)."""
    import pptx_md.slide_render as sr

    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    monkeypatch.setattr(sr, "_pypdfium2_importable", lambda: False)
    assert sr.slide_render_available() is False


def test_ac5_available_both_missing(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC5: 둘 다 미가용 -> False."""
    import pptx_md.slide_render as sr

    monkeypatch.setattr(_shutil, "which", lambda _name: None)
    monkeypatch.setattr(sr, "_pypdfium2_importable", lambda: False)
    assert sr.slide_render_available() is False


def test_ac5_available_returns_bool_never_raises() -> None:
    """AC5: slide_render_available() 은 항상 bool 을 반환하고 예외를 던지지 않는다."""
    result = slide_render_available()
    assert isinstance(result, bool)


@_requires_pypdfium2
def test_ac5_pypdfium2_importable_true_when_installed() -> None:
    """AC5 내부: pypdfium2 설치 환경에서 _pypdfium2_importable() -> True."""
    import pptx_md.slide_render as sr

    assert sr._pypdfium2_importable() is True


def test_ac5_pypdfium2_importable_false_when_blocked(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """AC5 내부: pypdfium2 import 실패 시 _pypdfium2_importable() -> False."""
    import pptx_md.slide_render as sr

    monkeypatch.setitem(sys.modules, "pypdfium2", None)
    assert sr._pypdfium2_importable() is False


# ===========================================================================
# AC6 — PII 미로그 (파일 경로·원문 텍스트 로그 출력 금지)
# ===========================================================================


def test_ac6_original_path_not_logged(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """AC6: 원본 파일 경로(민감할 수 있는 실제 파일명)가 로그에 노출되지 않는다."""
    monkeypatch.setattr(_shutil, "which", lambda _name: None)  # graceful skip 경로

    sensitive_dir = tmp_path / "고객사_민감정보_2026"
    sensitive_dir.mkdir()
    sensitive_pptx = sensitive_dir / "고객제안서_기밀.pptx"
    sensitive_pptx.write_bytes(b"fake")

    with caplog.at_level(logging.DEBUG, logger="pptx_md.slide_render"):
        render_deck(sensitive_pptx, [0])

    log_text = "\n".join(r.getMessage() for r in caplog.records)
    assert "고객사_민감정보_2026" not in log_text
    assert "고객제안서_기밀" not in log_text
    assert str(sensitive_pptx) not in log_text


def test_ac6_no_vlm_or_pii_related_imports_in_source() -> None:
    """AC6/PII 인접 확인: slide_render.py 소스에 anthropic/openai import 없음."""
    src = Path(__file__).parent.parent / "src" / "pptx_md" / "slide_render.py"
    text = src.read_text(encoding="utf-8")
    assert "anthropic" not in text
    assert "openai" not in text


# ===========================================================================
# AC7 — core-only import 성공: pypdfium2 지연 import(함수 내부)
# ===========================================================================


def test_ac7_import_always_succeeds() -> None:
    """AC7: pypdfium2 설치 여부와 무관하게 import pptx_md.slide_render 는 항상 성공."""
    import importlib

    mod = importlib.import_module("pptx_md.slide_render")
    assert mod is not None


def test_ac7_no_pypdfium2_top_level_import() -> None:
    """AC7: 모듈 최상단(함수/클래스 밖) 에 pypdfium2 import 가 없음(AST 검사)."""
    src_path = Path(__file__).parent.parent / "src" / "pptx_md" / "slide_render.py"
    tree = ast.parse(src_path.read_text(encoding="utf-8"))

    top_level_imports: list[str] = []
    for node in tree.body:  # 모듈 최상단 statement 만 (함수 내부 제외)
        if isinstance(node, ast.Import):
            for alias in node.names:
                top_level_imports.append(alias.name.split(".")[0])
        elif isinstance(node, ast.ImportFrom) and node.module:
            top_level_imports.append(node.module.split(".")[0])

    assert (
        "pypdfium2" not in top_level_imports
    ), "slide_render.py 최상단에 pypdfium2 import 가 있음 (지연 import 위반)"


def test_ac7_module_stdlib_only_top_level() -> None:
    """AC7: 최상단 import 는 stdlib(+typing) 만 (pypdfium2/anthropic/openai 등 0건)."""
    src_path = Path(__file__).parent.parent / "src" / "pptx_md" / "slide_render.py"
    tree = ast.parse(src_path.read_text(encoding="utf-8"))

    top_level_imports: list[str] = []
    for node in tree.body:
        if isinstance(node, ast.Import):
            for alias in node.names:
                top_level_imports.append(alias.name.split(".")[0])
        elif isinstance(node, ast.ImportFrom) and node.module:
            top_level_imports.append(node.module.split(".")[0])

    allowed_stdlib = {
        "__future__",
        "io",
        "logging",
        "os",
        "shutil",
        "subprocess",
        "tempfile",
        "collections",
        "pathlib",
        "typing",
    }
    disallowed = [m for m in top_level_imports if m not in allowed_stdlib]
    assert (
        disallowed == []
    ), f"slide_render.py 에 허용되지 않은 최상단 import: {disallowed}"


def test_ac7_import_succeeds_even_when_pypdfium2_blocked(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """AC7: sys.modules 에서 pypdfium2 를 차단해도 slide_render 재로드가 성공한다."""
    import importlib

    monkeypatch.setitem(sys.modules, "pypdfium2", None)
    reloaded = importlib.reload(sys.modules["pptx_md.slide_render"])
    assert reloaded is not None
    # 원상복구를 위해 다시 reload (monkeypatch teardown 후 다음 테스트 안전)
    monkeypatch.undo()
    importlib.reload(sys.modules["pptx_md.slide_render"])


# ===========================================================================
# AC8 — mypy/ruff/black exit 0 (프로파일 §4 명령)
# ===========================================================================


def test_ac8_mypy_exit_0() -> None:
    """AC8: mypy src/ 가 exit 0 (strict 통과, slide_render.py 포함)."""
    result = subprocess.run(
        ["python", "-m", "mypy", "src/"],
        capture_output=True,
        text=True,
        cwd=str(Path(__file__).parent.parent),
    )
    assert (
        result.returncode == 0
    ), f"mypy exit {result.returncode}:\n{result.stdout}\n{result.stderr}"


def test_ac8_ruff_check_exit_0() -> None:
    """AC8: ruff check 가 slide_render.py 에 대해 exit 0."""
    result = subprocess.run(
        ["python", "-m", "ruff", "check", "src/pptx_md/slide_render.py"],
        capture_output=True,
        text=True,
        cwd=str(Path(__file__).parent.parent),
    )
    assert (
        result.returncode == 0
    ), f"ruff exit {result.returncode}:\n{result.stdout}\n{result.stderr}"


def test_ac8_black_check_exit_0() -> None:
    """AC8: black --check 가 slide_render.py 에 대해 exit 0."""
    result = subprocess.run(
        ["python", "-m", "black", "--check", "src/pptx_md/slide_render.py"],
        capture_output=True,
        text=True,
        cwd=str(Path(__file__).parent.parent),
    )
    assert (
        result.returncode == 0
    ), f"black exit {result.returncode}:\n{result.stdout}\n{result.stderr}"
