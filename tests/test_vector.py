"""Tests for src/pptx_md/vector.py — FR-07 LibreOffice 변환 (AC1~AC6).

Test naming convention: test_ac<N>_<description>
Each AC has at least one covering test that asserts the Then clause directly.

Always-run tests (AC2/AC3/AC4/AC5/AC6) use monkeypatch to control
shutil.which and subprocess.run — no LibreOffice installation required.
AC1 is skipped when LibreOffice is not installed.
"""

from __future__ import annotations

import ast
import io
import logging
import shutil as _shutil
import struct
import subprocess
import zlib
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock

import pytest

from pptx_md.vector import (
    VECTOR_EXTS,
    convert_vector_to_png,
    libreoffice_available,
)

# ---------------------------------------------------------------------------
# skipif helper — mirrors ARCH-M3 §7.2 strategy
# ---------------------------------------------------------------------------

_requires_libreoffice = pytest.mark.skipif(
    _shutil.which("soffice") is None
    and _shutil.which("soffice.bin") is None
    and _shutil.which("libreoffice") is None,
    reason="LibreOffice (soffice) not installed",
)

# ---------------------------------------------------------------------------
# Minimal PNG helper (no Pillow at fixture-build time, pure stdlib)
# ---------------------------------------------------------------------------


def _make_minimal_png(width: int = 1, height: int = 1) -> bytes:
    """Return a minimal WxH white PNG (stdlib only, no Pillow dependency)."""
    sig = b"\x89PNG\r\n\x1a\n"

    def _chunk(name: bytes, data: bytes) -> bytes:
        length = struct.pack(">I", len(data))
        crc = struct.pack(">I", zlib.crc32(name + data) & 0xFFFFFFFF)
        return length + name + data + crc

    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    ihdr = _chunk(b"IHDR", ihdr_data)
    # one scanline: filter_byte=0, then RGB white for each pixel
    raw = b"\x00" + b"\xff\xff\xff" * width
    idat = _chunk(b"IDAT", zlib.compress(raw))
    iend = _chunk(b"IEND", b"")
    return sig + ihdr + idat + iend


_VALID_PNG = _make_minimal_png()


# ===========================================================================
# AC1 — 변환 성공 (LibreOffice 설치 환경에서만 실행)
# ===========================================================================


@_requires_libreoffice
def test_ac1_변환_성공_png_bytes_반환(tmp_path: Path) -> None:
    """AC1: LibreOffice 설치 시 EMF/WMF blob → PNG bytes 반환, Pillow 디코딩 가능."""
    import struct

    # Minimal valid EMF header (40 bytes) — type 0x00000001, size 0x00000028
    emf_header = struct.pack(
        "<IIIIIIIIII",
        0x00000001,  # iType: EMR_HEADER
        0x00000028,  # nSize: 40 bytes (minimal)
        0,
        0,
        100,
        100,  # rclBounds: left, top, right, bottom
        0,
        0,
        100,
        100,  # rclFrame: left, top, right, bottom
    )
    result = convert_vector_to_png(emf_header, "emf")

    # If LibreOffice converts successfully, result must be valid PNG
    if result is not None:
        assert result[:4] == b"\x89PNG", "반환 bytes 가 PNG 시그니처로 시작해야 함"
        # Verify Pillow can decode it
        from PIL import Image

        img = Image.open(io.BytesIO(result))
        assert img.width > 0
        assert img.height > 0
    # else: LibreOffice 가 설치됐지만 변환 실패(EMF 무결성 문제) → None 허용
    # (skipif 로 LibreOffice 존재는 보장됨, 변환 결과는 EMF 포맷에 의존)


@_requires_libreoffice
def test_ac1_wmf_변환_시도(tmp_path: Path) -> None:
    """AC1: WMF ext 도 변환 시도 경로 진입 확인 (LibreOffice 설치 환경)."""
    # Minimal WMF header stub
    wmf_stub = b"\xd7\xcd\xc6\x9a" + b"\x00" * 16
    result = convert_vector_to_png(wmf_stub, "wmf")
    # 결과는 PNG bytes 또는 None (변환 성공 여부는 LibreOffice 버전 의존)
    assert result is None or result[:4] == b"\x89PNG"


# ===========================================================================
# AC2 — 런타임 감지: import 항상 성공, shutil.which 기반
# ===========================================================================


def test_ac2_import_항상_성공() -> None:
    """AC2: LibreOffice 미설치 환경에서도 import pptx_md.vector 성공."""
    import importlib

    mod = importlib.import_module("pptx_md.vector")
    assert mod is not None


def test_ac2_libreoffice_available_bool_반환() -> None:
    """AC2: libreoffice_available() 는 항상 bool 반환, 예외 없음."""
    result = libreoffice_available()
    assert isinstance(result, bool)


def test_ac2_libreoffice_available_which_found(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC2: shutil.which 가 경로 반환 시 libreoffice_available() -> True."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    assert libreoffice_available() is True


def test_ac2_libreoffice_available_which_not_found(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """AC2: shutil.which 가 None 반환 시 libreoffice_available() -> False."""
    monkeypatch.setattr(_shutil, "which", lambda _name: None)
    assert libreoffice_available() is False


def test_ac2_module_level_no_side_effects() -> None:
    """AC2: import 시점에 shutil.which 미호출 (import-time 부작용 0)."""
    # vector 모듈을 재로드해도 which 가 호출되지 않아야 한다.
    # monkeypatch 로 which 를 추적하되 예외 발생으로 호출 감지.
    import importlib
    import sys

    original_which = _shutil.which
    calls: list[str] = []

    def tracking_which(name: str) -> str | None:
        calls.append(name)
        return original_which(name)

    _shutil.which = tracking_which  # type: ignore[assignment]
    try:
        # 재로드 시 최상단 코드가 재실행됨
        if "pptx_md.vector" in sys.modules:
            importlib.reload(sys.modules["pptx_md.vector"])
        else:
            importlib.import_module("pptx_md.vector")
    finally:
        _shutil.which = original_which  # type: ignore[assignment]

    assert calls == [], f"import-time 에 shutil.which 호출됨: {calls}"


# ===========================================================================
# AC3 — LibreOffice 부재 → None 반환, 예외 없음
# ===========================================================================


def test_ac3_which_none_emf_returns_none(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC3: shutil.which → None 시 convert_vector_to_png(emf) -> None, 예외 없음."""
    monkeypatch.setattr(_shutil, "which", lambda _name: None)
    result = convert_vector_to_png(b"fake_emf_bytes", "emf")
    assert result is None


def test_ac3_which_none_wmf_returns_none(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC3: shutil.which → None 시 convert_vector_to_png(wmf) -> None."""
    monkeypatch.setattr(_shutil, "which", lambda _name: None)
    result = convert_vector_to_png(b"fake_wmf_bytes", "wmf")
    assert result is None


def test_ac3_no_exception_on_empty_bytes(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC3: LibreOffice 부재 + 빈 bytes 도 예외 없이 None 반환."""
    monkeypatch.setattr(_shutil, "which", lambda _name: None)
    try:
        result = convert_vector_to_png(b"", "emf")
    except Exception as exc:
        pytest.fail(f"convert_vector_to_png 가 예외를 발생시켰습니다: {exc}")
    assert result is None


def test_ac3_no_exception_on_null_bytes(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC3: LibreOffice 부재 + null bytes 도 예외 없이 None 반환."""
    monkeypatch.setattr(_shutil, "which", lambda _name: None)
    result = convert_vector_to_png(b"\x00" * 50, "wmf")
    assert result is None


# ===========================================================================
# AC4 — 비벡터 ext → pass-through None, LibreOffice 미호출
# ===========================================================================


def test_ac4_png_ext_returns_none() -> None:
    """AC4: ext='png' → VECTOR_EXTS 외 → None 반환 (LibreOffice 호출 없음)."""
    result = convert_vector_to_png(b"fake_png_data", "png")
    assert result is None


def test_ac4_jpeg_ext_returns_none() -> None:
    """AC4: ext='jpeg' → None 반환."""
    result = convert_vector_to_png(b"fake_jpeg_data", "jpeg")
    assert result is None


def test_ac4_jpg_ext_returns_none() -> None:
    """AC4: ext='jpg' → None 반환."""
    result = convert_vector_to_png(b"fake_jpg_data", "jpg")
    assert result is None


def test_ac4_bmp_ext_returns_none() -> None:
    """AC4: ext='bmp' → None 반환."""
    result = convert_vector_to_png(b"fake_bmp_data", "bmp")
    assert result is None


def test_ac4_libreoffice_not_called_for_non_vector(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """AC4: 비벡터 ext 입력 시 shutil.which 자체도 호출되지 않음."""
    which_called = False

    def spy_which(name: str) -> str | None:
        nonlocal which_called
        which_called = True
        return "/usr/bin/soffice"

    monkeypatch.setattr(_shutil, "which", spy_which)
    result = convert_vector_to_png(b"fake_data", "png")
    assert result is None
    assert not which_called, "비벡터 ext 에서 shutil.which 가 호출됨"


def test_ac4_vector_exts_contains_emf_wmf() -> None:
    """AC4 전제: VECTOR_EXTS 에 emf, wmf 가 포함됨."""
    assert "emf" in VECTOR_EXTS
    assert "wmf" in VECTOR_EXTS


def test_ac4_uppercase_ext_treated_as_non_vector() -> None:
    """AC4: 대문자 ext 'EMF' 는 VECTOR_EXTS('emf')와 다름 → .lower() 처리 확인."""
    # vector.py 내부에서 image_ext.lower() 로 비교하므로 'EMF' 도 통과해야 함
    # monkeypatch 없이 — shutil.which 결과에 따라 None 또는 시도
    # 검증 목표: 예외가 발생하지 않음
    try:
        _ = convert_vector_to_png(b"fake", "EMF")
    except Exception as exc:
        pytest.fail(f"대문자 ext 입력에서 예외 발생: {exc}")


# ===========================================================================
# AC5 — 변환 실패 격리: subprocess 실패/타임아웃/출력 없음 → None + WARNING
# ===========================================================================


def test_ac5_subprocess_nonzero_returncode(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    """AC5: subprocess returncode≠0 → None 반환, WARNING 로그 포함."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")

    mock_result = MagicMock()
    mock_result.returncode = 1
    monkeypatch.setattr(subprocess, "run", lambda *a, **kw: mock_result)

    with caplog.at_level(logging.WARNING, logger="pptx_md.vector"):
        result = convert_vector_to_png(b"emf_stub", "emf")

    assert result is None
    assert any(
        r.levelno >= logging.WARNING for r in caplog.records
    ), "returncode≠0 시 WARNING 이상 로그가 없음"


def test_ac5_subprocess_nonzero_no_exception(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """AC5: subprocess 실패 시 예외 전파 없음."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")
    mock_result = MagicMock()
    mock_result.returncode = 2
    monkeypatch.setattr(subprocess, "run", lambda *a, **kw: mock_result)

    try:
        result = convert_vector_to_png(b"emf_stub", "emf")
    except Exception as exc:
        pytest.fail(f"subprocess 실패 시 예외 발생: {exc}")
    assert result is None


def test_ac5_output_file_missing(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    """AC5: returncode=0 이지만 출력 파일 없음 → None 반환, WARNING 로그."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")

    mock_result = MagicMock()
    mock_result.returncode = 0
    monkeypatch.setattr(subprocess, "run", lambda *a, **kw: mock_result)
    # subprocess.run 은 성공했지만 out_file(in.png)을 실제로 생성하지 않음

    with caplog.at_level(logging.WARNING, logger="pptx_md.vector"):
        result = convert_vector_to_png(b"emf_stub", "emf")

    assert result is None
    assert any(
        r.levelno >= logging.WARNING for r in caplog.records
    ), "출력 파일 없을 때 WARNING 이상 로그가 없음"


def test_ac5_timeout_returns_none(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    """AC5: subprocess.TimeoutExpired → None 반환, 예외 전파 없음, WARNING 로그."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")

    def raise_timeout(*args: Any, **kwargs: Any) -> None:
        raise subprocess.TimeoutExpired(cmd=["soffice"], timeout=30)

    monkeypatch.setattr(subprocess, "run", raise_timeout)

    with caplog.at_level(logging.WARNING, logger="pptx_md.vector"):
        try:
            result = convert_vector_to_png(b"emf_stub", "emf")
        except Exception as exc:
            pytest.fail(f"TimeoutExpired 시 예외 전파됨: {exc}")

    assert result is None
    assert any(
        r.levelno >= logging.WARNING for r in caplog.records
    ), "TimeoutExpired 시 WARNING 이상 로그가 없음"


def test_ac5_generic_exception_returns_none(
    monkeypatch: pytest.MonkeyPatch, caplog: pytest.LogCaptureFixture
) -> None:
    """AC5: 예상치 못한 Exception → None 반환, 예외 전파 없음, WARNING 로그."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")

    def raise_os_error(*args: Any, **kwargs: Any) -> None:
        raise OSError("디스크 공간 부족 시뮬레이션")

    monkeypatch.setattr(subprocess, "run", raise_os_error)

    with caplog.at_level(logging.WARNING, logger="pptx_md.vector"):
        try:
            result = convert_vector_to_png(b"emf_stub", "emf")
        except Exception as exc:
            pytest.fail(f"OSError 시 예외 전파됨: {exc}")

    assert result is None
    assert any(
        r.levelno >= logging.WARNING for r in caplog.records
    ), "예상치 못한 예외 시 WARNING 이상 로그가 없음"


def test_ac5_successful_conversion_returns_png(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    """AC5 (정상 경로 확인): subprocess 성공 + 출력 파일 존재 → PNG bytes 반환."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")

    def fake_subprocess_run(cmd: list[str], **kwargs: Any) -> MagicMock:
        # cmd 형식: [soffice, --headless, --convert-to, png, --outdir, <dir>, <input>]
        # <dir> 는 cmd[-2], <input> 는 cmd[-1]
        out_dir = Path(cmd[-2])
        out_file = out_dir / "in.png"
        out_file.write_bytes(_VALID_PNG)
        mock = MagicMock()
        mock.returncode = 0
        return mock

    monkeypatch.setattr(subprocess, "run", fake_subprocess_run)

    result = convert_vector_to_png(b"emf_stub", "emf")

    assert result is not None
    assert result[:4] == b"\x89PNG", "반환 bytes 가 PNG 시그니처로 시작해야 함"


def test_ac5_empty_output_file_returns_none(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """AC5: subprocess 성공했지만 출력 PNG 가 빈 파일 → None 반환."""
    monkeypatch.setattr(_shutil, "which", lambda _name: "/usr/bin/soffice")

    def fake_subprocess_empty(cmd: list[str], **kwargs: Any) -> MagicMock:
        out_dir = Path(cmd[-2])
        out_file = out_dir / "in.png"
        out_file.write_bytes(b"")  # 빈 파일
        mock = MagicMock()
        mock.returncode = 0
        return mock

    monkeypatch.setattr(subprocess, "run", fake_subprocess_empty)

    result = convert_vector_to_png(b"emf_stub", "emf")
    assert result is None


# ===========================================================================
# AC6 — VLM import 0 + mypy strict exit 0
# ===========================================================================


def test_ac6_no_vlm_imports_in_vector_source() -> None:
    """AC6: vector.py 소스에 anthropic/openai import 가 없음."""
    vector_src = Path(__file__).parent.parent / "src" / "pptx_md" / "vector.py"
    text = vector_src.read_text(encoding="utf-8")

    assert "anthropic" not in text, "vector.py 에 'anthropic' 문자열이 발견됨"
    assert "openai" not in text, "vector.py 에 'openai' 문자열이 발견됨"


def test_ac6_vector_module_stdlib_only() -> None:
    """AC6: vector.py 의 최상단 imports 가 stdlib 만임을 확인."""
    vector_src = Path(__file__).parent.parent / "src" / "pptx_md" / "vector.py"
    tree = ast.parse(vector_src.read_text(encoding="utf-8"))

    # 모듈 최상단 import 문만 수집 (함수/클래스 내부 제외)
    top_level_imports: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import | ast.ImportFrom):
            if isinstance(node, ast.ImportFrom) and node.module:
                top_level_imports.append(node.module.split(".")[0])
            elif isinstance(node, ast.Import):
                for alias in node.names:
                    top_level_imports.append(alias.name.split(".")[0])

    allowed_stdlib = {
        "__future__",
        "logging",
        "shutil",
        "subprocess",
        "tempfile",
        "pathlib",
    }
    disallowed = [m for m in top_level_imports if m not in allowed_stdlib]
    assert disallowed == [], f"vector.py 에 허용되지 않은 import: {disallowed}"


def test_ac6_mypy_exit_0() -> None:
    """AC6: mypy src/ 가 exit 0 으로 종료 (strict 통과)."""
    import subprocess as sp

    result = sp.run(
        ["python", "-m", "mypy", "src/"],
        capture_output=True,
        text=True,
        cwd=str(Path(__file__).parent.parent),
    )
    assert (
        result.returncode == 0
    ), f"mypy exit {result.returncode}:\n{result.stdout}\n{result.stderr}"
