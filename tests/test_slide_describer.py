"""Tests for src/pptx_md/slide_describer.py (FR-32, ADR-625/626, DEC-8).

AC1 ~ AC5 전 항목 커버 (이슈 #93 수용 기준):
  AC1: SlideReconstructor Protocol + SlideContext 데이터클래스
       (제목·텍스트 아웃라인 등 컨텍스트 필드)
  AC2: reconstruct(image_bytes, ext, context)->str 시그니처
  AC3: @runtime_checkable
  AC4: SDK import 0 (순수 계약 모듈, INV-5 격리
       — anthropic/openai/pptx/Pillow import 금지)
  AC5: mypy exit 0, 단위 테스트(계약·타입·runtime_checkable 동작)

Test naming convention: test_ac<N>_<description> (describer.py 전례 계승).
"""

from __future__ import annotations

import inspect
import pathlib
import subprocess
import sys
import textwrap

# ===========================================================================
# AC1 — SlideReconstructor Protocol + SlideContext 데이터클래스 (컨텍스트 필드)
# ===========================================================================


def test_ac1_import_성공() -> None:
    """AC1: import pptx_md.slide_describer succeeds without any VLM SDK installed."""
    import importlib

    mod = importlib.import_module("pptx_md.slide_describer")
    assert mod is not None


def test_ac1_SlideReconstructor_심볼_노출() -> None:
    """AC1: SlideReconstructor is accessible as pptx_md.slide_describer.<name>."""
    from pptx_md.slide_describer import SlideReconstructor

    assert SlideReconstructor is not None


def test_ac1_SlideContext_심볼_노출() -> None:
    """AC1: SlideContext is accessible as pptx_md.slide_describer.SlideContext."""
    from pptx_md.slide_describer import SlideContext

    assert SlideContext is not None


def test_ac1_dunder_all에_포함() -> None:
    """AC1: SlideContext and SlideReconstructor are listed in __all__."""
    import pptx_md.slide_describer as mod

    assert "SlideContext" in mod.__all__
    assert "SlideReconstructor" in mod.__all__


def test_ac1_SlideContext_필드_및_값() -> None:
    """AC1: SlideContext carries slide_index/title/text_outline context fields."""
    from pptx_md.slide_describer import SlideContext

    ctx = SlideContext(slide_index=0, title="제목", text_outline="- 불릿 1")
    assert ctx.slide_index == 0
    assert ctx.title == "제목"
    assert ctx.text_outline == "- 불릿 1"


def test_ac1_SlideContext_빈문자열_허용() -> None:
    """AC1: title/text_outline accept empty string (no-title / no-text slides)."""
    from pptx_md.slide_describer import SlideContext

    ctx = SlideContext(slide_index=3, title="", text_outline="")
    assert ctx.title == ""
    assert ctx.text_outline == ""


def test_ac1_SlideContext_frozen_불변() -> None:
    """AC1: SlideContext is a frozen dataclass — attribute mutation raises."""
    import dataclasses

    from pptx_md.slide_describer import SlideContext

    assert dataclasses.is_dataclass(SlideContext)
    ctx = SlideContext(slide_index=0, title="t", text_outline="o")
    try:
        ctx.title = "changed"  # type: ignore[misc]
        raised = False
    except dataclasses.FrozenInstanceError:
        raised = True
    assert raised, "SlideContext must be frozen (immutable)"


# ===========================================================================
# AC2 — reconstruct(image_bytes, ext, context) -> str 시그니처
# ===========================================================================


def test_ac2_reconstruct_시그니처_일치() -> None:
    """AC2: SlideReconstructor.reconstruct has the exact required signature."""
    import typing

    from pptx_md.slide_describer import SlideContext, SlideReconstructor

    sig = inspect.signature(SlideReconstructor.reconstruct)
    params = sig.parameters

    assert "image_bytes" in params, "image_bytes parameter missing"
    assert "image_ext" in params, "image_ext parameter missing"
    assert "context" in params, "context parameter missing"

    import pptx_md.slide_describer as _mod

    hints = typing.get_type_hints(_mod.SlideReconstructor.reconstruct)
    assert hints.get("image_bytes") is bytes, "image_bytes must be typed as bytes"
    assert hints.get("image_ext") is str, "image_ext must be typed as str"
    assert hints.get("context") is SlideContext, "context must be typed as SlideContext"
    assert hints.get("return") is str, "return type must be str"


def test_ac2_reconstruct_파라미터_순서() -> None:
    """AC2: Parameters appear in the correct positional order."""
    from pptx_md.slide_describer import SlideReconstructor

    sig = inspect.signature(SlideReconstructor.reconstruct)
    param_names = [p for p in sig.parameters if p != "self"]
    assert param_names == ["image_bytes", "image_ext", "context"]


# ===========================================================================
# AC3 — @runtime_checkable: 상속 없이 reconstruct 구현한 임의 클래스 → isinstance True
# ===========================================================================


class _FakeReconstructor:
    """Concrete class satisfying SlideReconstructor structurally (no inheritance)."""

    def reconstruct(
        self,
        image_bytes: bytes,
        image_ext: str,
        context: object,
    ) -> str:
        return f"reconstructed slide {getattr(context, 'slide_index', '?')}"


class _NoReconstructClass:
    """Class that does NOT implement reconstruct — should NOT satisfy Protocol."""

    def other_method(self) -> str:
        return "not a reconstructor"


def test_ac3_구조적_타이핑_isinstance_true() -> None:
    """AC3: Class with reconstruct() satisfies SlideReconstructor w/o inheritance."""
    from pptx_md.slide_describer import SlideReconstructor

    obj = _FakeReconstructor()
    assert not issubclass(_FakeReconstructor, type(SlideReconstructor))
    assert isinstance(obj, SlideReconstructor)


def test_ac3_reconstruct_없는_클래스_isinstance_false() -> None:
    """AC3: Object w/o reconstruct() → isinstance(..., SlideReconstructor) == False."""
    from pptx_md.slide_describer import SlideReconstructor

    obj = _NoReconstructClass()
    assert not isinstance(obj, SlideReconstructor)


def test_ac3_reconstruct_작동_확인() -> None:
    """AC3: _FakeReconstructor.reconstruct runs correctly and returns str."""
    from pptx_md.slide_describer import SlideContext

    obj = _FakeReconstructor()
    ctx = SlideContext(slide_index=2, title="t", text_outline="o")
    result = obj.reconstruct(b"\x89PNG", "png", ctx)
    assert isinstance(result, str)
    assert result == "reconstructed slide 2"


# ===========================================================================
# AC4 — SDK import 0 (순수 계약 모듈, INV-5 격리)
# ===========================================================================


def test_ac4_anthropic_import_없음() -> None:
    """AC4: slide_describer.py contains no reference to 'anthropic'."""
    source = _read_source()
    assert (
        "anthropic" not in source
    ), "slide_describer.py must not reference 'anthropic'"


def test_ac4_openai_import_없음() -> None:
    """AC4: slide_describer.py contains no reference to 'openai'."""
    source = _read_source()
    assert "openai" not in source, "slide_describer.py must not reference 'openai'"


def test_ac4_pptx_import_없음() -> None:
    """AC4: slide_describer.py has no python-pptx import (excl. 'pptx_md' self-refs)."""
    import re

    source = _read_source()
    assert not re.search(
        r"\bimport pptx\b(?!_md)", source
    ), "slide_describer.py must not import python-pptx"
    assert not re.search(
        r"\bfrom pptx\b(?!_md)", source
    ), "slide_describer.py must not import python-pptx"


def test_ac4_pillow_import_없음() -> None:
    """AC4: slide_describer.py contains no reference to Pillow (PIL)."""
    source = _read_source()
    assert "import PIL" not in source, "slide_describer.py must not import Pillow"
    assert "from PIL" not in source, "slide_describer.py must not import Pillow"


def test_ac4_stdlib_만_import() -> None:
    """AC4: slide_describer.py only imports stdlib (dataclasses, typing, __future__)."""
    import re

    source = _read_source()
    disallowed_literal = [
        "import anthropic",
        "import openai",
        "from anthropic",
        "from openai",
        "import PIL",
        "from PIL",
    ]
    for pattern in disallowed_literal:
        assert (
            pattern not in source
        ), f"slide_describer.py must not contain: {pattern!r}"
    # 'pptx' excludes 'pptx_md' self-references (module docstring mentions the
    # package name and python-pptx by name without importing it).
    assert not re.search(r"\bimport pptx\b(?!_md)", source)
    assert not re.search(r"\bfrom pptx\b(?!_md)", source)


def test_ac4_core_only_import_실제성공() -> None:
    """AC4: importing the module in-process succeeds (core-only environment proxy)."""
    import importlib
    import sys

    sys.modules.pop("pptx_md.slide_describer", None)
    mod = importlib.import_module("pptx_md.slide_describer")
    assert hasattr(mod, "SlideReconstructor")
    assert hasattr(mod, "SlideContext")


def _read_source() -> str:
    source_file = (
        pathlib.Path(__file__).parent.parent / "src" / "pptx_md" / "slide_describer.py"
    )
    return source_file.read_text(encoding="utf-8")


# ===========================================================================
# AC5 — mypy exit 0 / ruff / black exit 0 (계약·타입 검증)
# ===========================================================================


def test_ac5_mypy_strict_exit0() -> None:
    """AC5: mypy src/ exits with code 0 (includes slide_describer.py)."""
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


def test_ac5_mypy_올바른_시그니처_에러없음() -> None:
    """AC5 GOOD: mypy --strict reports no error for correctly-typed reconstructor."""
    snippet = textwrap.dedent("""\
        from pptx_md.slide_describer import SlideContext, SlideReconstructor

        class GoodReconstructor:
            def reconstruct(
                self,
                image_bytes: bytes,
                image_ext: str,
                context: SlideContext,
            ) -> str:
                return "ok"

        def use_reconstructor(r: SlideReconstructor) -> str:
            ctx = SlideContext(slide_index=0, title="t", text_outline="o")
            return r.reconstruct(b"", "png", ctx)

        use_reconstructor(GoodReconstructor())
        """)

    result = _run_mypy_with_snippet(snippet)
    assert result.returncode == 0, (
        "Expected mypy exit 0 for correct SlideReconstructor signature, "
        f"but got exit {result.returncode}.\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )


def test_ac5_mypy_시그니처_불일치_타입에러() -> None:
    """AC5 BAD: incompatible reconstructor class triggers mypy strict error."""
    snippet = textwrap.dedent("""\
        from pptx_md.slide_describer import SlideContext, SlideReconstructor

        class BadReconstructor:
            def reconstruct(
                self,
                image_bytes: bytes,
                image_ext: str,
                context: SlideContext,
            ) -> int:  # wrong return type: int instead of str
                return 42

        def use_reconstructor(r: SlideReconstructor) -> str:
            ctx = SlideContext(slide_index=0, title="t", text_outline="o")
            return r.reconstruct(b"", "png", ctx)

        use_reconstructor(BadReconstructor())  # type error: incompatible type
        """)

    result = _run_mypy_with_snippet(snippet)
    assert result.returncode != 0, (
        "Expected mypy to report a type error for incompatible reconstructor, "
        f"but it exited 0.\nstdout: {result.stdout}\nstderr: {result.stderr}"
    )
    assert (
        "error" in result.stdout.lower() or "error" in result.stderr.lower()
    ), f"mypy output did not contain 'error'.\nstdout: {result.stdout}"


def test_ac5_ruff_exit0() -> None:
    """AC5: ruff check . exits with code 0."""
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


def test_ac5_black_check_exit0() -> None:
    """AC5: black --check . exits with code 0."""
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


def _run_mypy_with_snippet(snippet: str) -> subprocess.CompletedProcess[str]:
    """Write *snippet* to a temp file and run ``mypy --strict src/ <file>``."""
    project_root = pathlib.Path(__file__).parent.parent
    tmp_dir = project_root / "tests" / "_ac5_tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    snippet_path = tmp_dir / "snippet.py"
    snippet_path.write_text(snippet, encoding="utf-8")
    try:
        return subprocess.run(
            [sys.executable, "-m", "mypy", "--strict", "src/", str(snippet_path)],
            capture_output=True,
            text=True,
            cwd=str(project_root),
        )
    finally:
        snippet_path.unlink(missing_ok=True)
        try:
            tmp_dir.rmdir()
        except OSError:
            pass
