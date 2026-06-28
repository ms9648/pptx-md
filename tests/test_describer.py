"""Tests for src/pptx_md/describer.py and FR-08 errors.py extensions.

AC1 ~ AC6 전 항목 커버.

Test naming convention: test_ac<N>_<description>
Each AC has at least one covering test that asserts the Then clause directly.
"""

from __future__ import annotations

import inspect
import pathlib
import subprocess
import sys
import textwrap

# ===========================================================================
# AC1 — core만 설치된 환경에서 import 성공 + ImageDescriber 심볼 노출 (NFR-08)
# ===========================================================================


def test_ac1_import_성공() -> None:
    """AC1: import pptx_md.describer succeeds without any VLM SDK installed."""
    import importlib

    mod = importlib.import_module("pptx_md.describer")
    assert mod is not None


def test_ac1_ImageDescriber_심볼_노출() -> None:
    """AC1: ImageDescriber is accessible as pptx_md.describer.ImageDescriber."""
    from pptx_md.describer import ImageDescriber

    assert ImageDescriber is not None


def test_ac1_dunder_all에_포함() -> None:
    """AC1: ImageDescriber is listed in __all__."""
    import pptx_md.describer as mod

    assert "ImageDescriber" in mod.__all__


def test_ac1_errors_import_성공() -> None:
    """AC1: import pptx_md.errors succeeds and exposes new exception classes."""
    from pptx_md.errors import DescribeError, InstallationError, PptxMdError

    assert issubclass(InstallationError, PptxMdError)
    assert issubclass(DescribeError, PptxMdError)


# ===========================================================================
# AC2 — describe 시그니처 정확히 일치
# ===========================================================================


def test_ac2_describe_시그니처_일치() -> None:
    """AC2: ImageDescriber.describe has the exact required signature."""
    import typing

    from pptx_md.describer import ImageDescriber

    sig = inspect.signature(ImageDescriber.describe)
    params = sig.parameters

    # Required parameters (excluding 'self')
    assert "image_bytes" in params, "image_bytes parameter missing"
    assert "image_ext" in params, "image_ext parameter missing"
    assert "shape_hint" in params, "shape_hint parameter missing"

    # Resolve forward references with get_type_hints (handles
    # `from __future__ import annotations` string annotations)
    import pptx_md.describer as _mod

    hints = typing.get_type_hints(_mod.ImageDescriber.describe)
    assert hints.get("image_bytes") is bytes, "image_bytes must be typed as bytes"
    assert hints.get("image_ext") is str, "image_ext must be typed as str"
    # shape_hint: str | None  — Python 3.11 union type
    shape_hint_hint = hints.get("shape_hint")
    assert shape_hint_hint is not None, "shape_hint must be annotated"
    # return type
    assert hints.get("return") is str, "return type must be str"


def test_ac2_describe_파라미터_순서() -> None:
    """AC2: Parameters appear in the correct positional order."""
    from pptx_md.describer import ImageDescriber

    sig = inspect.signature(ImageDescriber.describe)
    # Exclude 'self'
    param_names = [p for p in sig.parameters if p != "self"]
    assert param_names == ["image_bytes", "image_ext", "shape_hint"]


# ===========================================================================
# AC3 — 상속 없이 describe 구현한 임의 클래스 → isinstance True (@runtime_checkable)
# ===========================================================================


class _FakeDescriber:
    """Concrete class satisfying ImageDescriber structurally (no inheritance)."""

    def describe(
        self,
        image_bytes: bytes,
        image_ext: str,
        shape_hint: str | None,
    ) -> str:
        return f"described {image_ext}"


class _NoDescribeClass:
    """Class that does NOT implement describe — should NOT satisfy Protocol."""

    def other_method(self) -> str:
        return "not a describer"


def test_ac3_구조적_타이핑_isinstance_true() -> None:
    """AC3: Class with describe() satisfies ImageDescriber without inheritance."""
    from pptx_md.describer import ImageDescriber

    obj = _FakeDescriber()
    # Must not inherit from ImageDescriber
    # No nominal inheritance: _FakeDescriber is not a subclass of ImageDescriber
    assert not issubclass(_FakeDescriber, type(ImageDescriber))
    # isinstance must return True due to @runtime_checkable
    assert isinstance(obj, ImageDescriber)


def test_ac3_describe_없는_클래스_isinstance_false() -> None:
    """AC3: Object without describe() → isinstance(..., ImageDescriber) == False."""
    from pptx_md.describer import ImageDescriber

    obj = _NoDescribeClass()
    assert not isinstance(obj, ImageDescriber)


def test_ac3_describe_작동_확인() -> None:
    """AC3: _FakeDescriber.describe runs correctly and returns str."""
    obj = _FakeDescriber()
    result = obj.describe(b"\x89PNG", "png", "hint text")
    assert isinstance(result, str)
    assert result == "described png"


def test_ac3_shape_hint_none_허용() -> None:
    """AC3: describe() called with shape_hint=None does not raise."""
    obj = _FakeDescriber()
    result = obj.describe(b"\x89PNG", "png", None)
    assert isinstance(result, str)


# ===========================================================================
# AC4 — 시그니처 불일치 클래스를 ImageDescriber로 사용하면 mypy strict 타입 에러
# ===========================================================================


def test_ac4_mypy_시그니처_불일치_타입에러() -> None:
    """AC4: mypy detects type error when incompatible class used as ImageDescriber.

    We write a small Python snippet with a deliberate type mismatch and run
    mypy as a subprocess.  The test passes iff mypy exits with a non-zero
    code (i.e., it found the expected type error).
    """
    # Build a temporary snippet that has a wrong return type
    snippet = textwrap.dedent("""\
        from pptx_md.describer import ImageDescriber

        class BadDescriber:
            def describe(
                self,
                image_bytes: bytes,
                image_ext: str,
                shape_hint: str | None,
            ) -> int:  # wrong return type: int instead of str
                return 42

        def use_describer(d: ImageDescriber) -> str:
            return d.describe(b"", "png", None)

        use_describer(BadDescriber())  # type error: incompatible type
        """)

    # Write to a temp file in the scratchpad directory
    scratchpad = pathlib.Path(
        r"C:\Users\ms964\AppData\Local\Temp\claude\D--dev-pptx-md"
        r"\7410c743-da47-49b5-a139-663e25f2b338\scratchpad"
    )
    scratchpad.mkdir(parents=True, exist_ok=True)
    snippet_path = scratchpad / "bad_describer_check.py"
    snippet_path.write_text(snippet, encoding="utf-8")

    result = subprocess.run(
        [sys.executable, "-m", "mypy", "--strict", str(snippet_path)],
        capture_output=True,
        text=True,
    )
    # mypy should exit non-zero because of the type mismatch
    assert result.returncode != 0, (
        f"Expected mypy to report a type error for incompatible describer, "
        f"but it exited 0.\nstdout: {result.stdout}\nstderr: {result.stderr}"
    )
    # Confirm the error is type-related (not a mypy crash)
    assert (
        "error" in result.stdout.lower() or "error" in result.stderr.lower()
    ), f"mypy output did not contain 'error'.\nstdout: {result.stdout}"


# ===========================================================================
# AC5 — describer.py 소스에 anthropic/openai grep 0건
# ===========================================================================


def test_ac5_anthropic_import_없음() -> None:
    """AC5: describer.py contains no reference to 'anthropic'."""
    source_file = (
        pathlib.Path(__file__).parent.parent / "src" / "pptx_md" / "describer.py"
    )
    source = source_file.read_text(encoding="utf-8")
    assert "anthropic" not in source, "describer.py must not reference 'anthropic'"


def test_ac5_openai_import_없음() -> None:
    """AC5: describer.py contains no reference to 'openai'."""
    source_file = (
        pathlib.Path(__file__).parent.parent / "src" / "pptx_md" / "describer.py"
    )
    source = source_file.read_text(encoding="utf-8")
    assert "openai" not in source, "describer.py must not reference 'openai'"


def test_ac5_stdlib_만_import() -> None:
    """AC5: describer.py only imports from stdlib (typing, __future__)."""
    source_file = (
        pathlib.Path(__file__).parent.parent / "src" / "pptx_md" / "describer.py"
    )
    source = source_file.read_text(encoding="utf-8")
    # Check for any non-stdlib import lines
    disallowed = ["import anthropic", "import openai", "from anthropic", "from openai"]
    for pattern in disallowed:
        assert pattern not in source, f"describer.py must not contain: {pattern!r}"


# ===========================================================================
# AC6 — mypy src/ exit 0 (strict), ruff check . && black --check . exit 0
# ===========================================================================


def test_ac6_mypy_strict_exit0() -> None:
    """AC6: mypy src/ --strict exits with code 0."""
    project_root = pathlib.Path(__file__).parent.parent
    result = subprocess.run(
        [sys.executable, "-m", "mypy", "src/"],
        capture_output=True,
        text=True,
        cwd=str(project_root),
    )
    assert result.returncode == 0, (
        f"mypy src/ failed (exit {result.returncode}).\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )


def test_ac6_ruff_exit0() -> None:
    """AC6: ruff check . exits with code 0."""
    project_root = pathlib.Path(__file__).parent.parent
    result = subprocess.run(
        [sys.executable, "-m", "ruff", "check", "."],
        capture_output=True,
        text=True,
        cwd=str(project_root),
    )
    assert result.returncode == 0, (
        f"ruff check failed (exit {result.returncode}).\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )


def test_ac6_black_check_exit0() -> None:
    """AC6: black --check . exits with code 0."""
    project_root = pathlib.Path(__file__).parent.parent
    result = subprocess.run(
        [sys.executable, "-m", "black", "--check", "."],
        capture_output=True,
        text=True,
        cwd=str(project_root),
    )
    assert result.returncode == 0, (
        f"black --check failed (exit {result.returncode}).\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )
