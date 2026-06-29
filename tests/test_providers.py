"""Tests for src/pptx_md/providers/ package (FR-09, issue #27).

AC1 ~ AC7 전 항목 커버.  모든 SDK 호출은 mock 기반 (실제 API 호출 0건).

Test naming convention: test_ac<N>_<description>
Each AC has at least one covering test that asserts the Then clause directly.
"""

from __future__ import annotations

import importlib
import logging
import pathlib
import subprocess
import sys
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_PROJECT_ROOT = pathlib.Path(__file__).parent.parent


def _remove_sdk_modules(*names: str) -> dict[str, Any]:
    """Remove SDK module entries from sys.modules and return them for restore."""
    removed: dict[str, Any] = {}
    keys_to_remove = [
        k
        for k in list(sys.modules.keys())
        if any(k == n or k.startswith(n + ".") for n in names)
    ]
    for key in keys_to_remove:
        removed[key] = sys.modules.pop(key)
    return removed


def _restore_modules(saved: dict[str, Any]) -> None:
    """Restore previously removed modules back into sys.modules."""
    sys.modules.update(saved)


# ---------------------------------------------------------------------------
# AC1 — core만 설치된 환경에서 providers 패키지 import 성공 (NFR-08)
# ---------------------------------------------------------------------------


def test_ac1_import_providers_패키지_성공() -> None:
    """AC1: import pptx_md.providers succeeds without VLM SDKs installed."""
    # Remove providers cache to force reimport
    saved = _remove_sdk_modules("pptx_md.providers")
    try:
        # Simulate SDK absence by temporarily hiding SDK modules
        saved_anthropic = sys.modules.pop("anthropic", None)
        saved_openai = sys.modules.pop("openai", None)
        # Temporarily inject ImportError-raising fakes at top-level to prevent
        # any accidentally installed SDK from interfering
        sys.modules["anthropic"] = None  # type: ignore[assignment]
        sys.modules["openai"] = None  # type: ignore[assignment]
        try:
            mod = importlib.import_module("pptx_md.providers")
            assert mod is not None
        finally:
            # Restore
            if saved_anthropic is not None:
                sys.modules["anthropic"] = saved_anthropic
            else:
                sys.modules.pop("anthropic", None)
            if saved_openai is not None:
                sys.modules["openai"] = saved_openai
            else:
                sys.modules.pop("openai", None)
    finally:
        _restore_modules(saved)


def test_ac1_import_providers_anthropic_모듈_성공() -> None:
    """AC1: import pptx_md.providers.anthropic succeeds without SDK."""
    saved = _remove_sdk_modules("pptx_md.providers.anthropic")
    saved_sdk = sys.modules.pop("anthropic", None)
    sys.modules["anthropic"] = None  # type: ignore[assignment]
    try:
        mod = importlib.import_module("pptx_md.providers.anthropic")
        assert mod is not None
    finally:
        if saved_sdk is not None:
            sys.modules["anthropic"] = saved_sdk
        else:
            sys.modules.pop("anthropic", None)
        _restore_modules(saved)


def test_ac1_import_providers_openai_모듈_성공() -> None:
    """AC1: import pptx_md.providers.openai succeeds without SDK."""
    saved = _remove_sdk_modules("pptx_md.providers.openai")
    saved_sdk = sys.modules.pop("openai", None)
    sys.modules["openai"] = None  # type: ignore[assignment]
    try:
        mod = importlib.import_module("pptx_md.providers.openai")
        assert mod is not None
    finally:
        if saved_sdk is not None:
            sys.modules["openai"] = saved_sdk
        else:
            sys.modules.pop("openai", None)
        _restore_modules(saved)


# ---------------------------------------------------------------------------
# AC2 — SDK 미설치 환경에서 인스턴스화 → InstallationError (pip install 안내 포함)
# ---------------------------------------------------------------------------


def test_ac2_anthropic_sdk_미설치_InstallationError() -> None:
    """AC2: AnthropicDescriber(...) raises InstallationError when SDK absent."""
    from pptx_md.errors import InstallationError

    # Force ImportError when 'anthropic' is imported inside constructor
    with patch.dict(sys.modules, {"anthropic": None}):  # type: ignore[dict-item]
        # Re-import the module to clear any cached _anthropic_sdk reference
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            with pytest.raises(InstallationError) as exc_info:
                AnthropicDescriber()
            assert "pip install pptx-md[vlm]" in str(exc_info.value)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac2_openai_sdk_미설치_InstallationError() -> None:
    """AC2: OpenAIDescriber(...) raises InstallationError when SDK absent."""
    from pptx_md.errors import InstallationError

    with patch.dict(sys.modules, {"openai": None}):  # type: ignore[dict-item]
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import OpenAIDescriber

            with pytest.raises(InstallationError) as exc_info:
                OpenAIDescriber()
            assert "pip install pptx-md[vlm]" in str(exc_info.value)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


def test_ac2_ModuleNotFoundError_미노출_anthropic() -> None:
    """AC2: Raw ModuleNotFoundError is NOT propagated (wrapped in InstallationError)."""
    from pptx_md.errors import InstallationError

    with patch.dict(sys.modules, {"anthropic": None}):  # type: ignore[dict-item]
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            with pytest.raises(InstallationError):
                AnthropicDescriber()
            # If we reach here, it was InstallationError not ModuleNotFoundError
        except ModuleNotFoundError:
            pytest.fail("Raw ModuleNotFoundError must not be propagated (AC2)")
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac2_get_describer_unknown_ValueError() -> None:
    """AC2-extra: get_describer('unknown') raises ValueError."""
    from pptx_md.providers import get_describer

    with pytest.raises(ValueError, match="unknown"):
        get_describer("unknown")


# ---------------------------------------------------------------------------
# AC3 — mock SDK → describe 반환 텍스트 + shape_hint 프롬프트 포함
# ---------------------------------------------------------------------------


def _make_anthropic_mock_response(text: str) -> MagicMock:
    """Create a mock Anthropic messages.create response with *text*."""
    block = MagicMock()
    block.text = text
    response = MagicMock()
    response.content = [block]
    return response


def _make_openai_mock_response(text: str) -> MagicMock:
    """Create a mock OpenAI chat.completions.create response with *text*."""
    msg = MagicMock()
    msg.content = text
    choice = MagicMock()
    choice.message = msg
    response = MagicMock()
    response.choices = [choice]
    return response


def _build_mock_anthropic_sdk(response_text: str = "mocked description") -> MagicMock:
    """Return a mock anthropic module with Anthropic client."""
    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.messages.create.return_value = _make_anthropic_mock_response(
        response_text
    )
    mock_sdk.Anthropic.return_value = mock_client
    return mock_sdk


def _build_mock_openai_sdk(response_text: str = "mocked description") -> MagicMock:
    """Return a mock openai module with OpenAI client."""
    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.chat.completions.create.return_value = _make_openai_mock_response(
        response_text
    )
    mock_sdk.OpenAI.return_value = mock_client
    return mock_sdk


def test_ac3_anthropic_describe_텍스트_반환() -> None:
    """AC3: AnthropicDescriber.describe returns text from mock response."""
    mock_sdk = _build_mock_anthropic_sdk("a diagram showing flow")

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        # Re-import to use patched SDK
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            describer = AnthropicDescriber(api_key="test-key")
            result = describer.describe(b"\x89PNG", "png", "diagram")
            assert result == "a diagram showing flow"
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac3_anthropic_shape_hint_프롬프트_포함() -> None:
    """AC3: shape_hint text is included in the prompt sent to Anthropic API."""
    mock_sdk = _build_mock_anthropic_sdk("some description")

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            mock_client = mock_sdk.Anthropic.return_value
            describer = AnthropicDescriber(api_key="test-key")
            describer.describe(b"\x89PNG", "png", "diagram")

            # Verify messages.create was called with shape_hint in text content
            call_args = mock_client.messages.create.call_args
            assert call_args is not None
            messages = call_args.kwargs.get("messages") or call_args.args[0]
            # Find the text content block in messages
            found_hint = False
            for msg in messages:
                for content_block in msg.get("content", []):
                    if content_block.get("type") == "text":
                        text = content_block.get("text", "")
                        if "diagram" in text:
                            found_hint = True
            assert found_hint, "shape_hint 'diagram' must appear in prompt text"
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac3_openai_describe_텍스트_반환() -> None:
    """AC3: OpenAIDescriber.describe returns text from mock response."""
    mock_sdk = _build_mock_openai_sdk("a photo of a mountain")

    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import OpenAIDescriber

            describer = OpenAIDescriber(api_key="test-key")
            result = describer.describe(b"\x89PNG", "png", "photo")
            assert result == "a photo of a mountain"
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


def test_ac3_openai_shape_hint_프롬프트_포함() -> None:
    """AC3: shape_hint text is included in the prompt sent to OpenAI API."""
    mock_sdk = _build_mock_openai_sdk("some description")

    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import OpenAIDescriber

            mock_client = mock_sdk.OpenAI.return_value
            describer = OpenAIDescriber(api_key="test-key")
            describer.describe(b"\x89PNG", "png", "diagram")

            call_args = mock_client.chat.completions.create.call_args
            assert call_args is not None
            messages = call_args.kwargs.get("messages") or call_args.args[0]
            found_hint = False
            for msg in messages:
                for content_block in msg.get("content", []):
                    if content_block.get("type") == "text":
                        text = content_block.get("text", "")
                        if "diagram" in text:
                            found_hint = True
            assert found_hint, "shape_hint 'diagram' must appear in prompt text"
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


# ---------------------------------------------------------------------------
# AC4 — isinstance(obj, ImageDescriber) == True (Protocol 충족)
# ---------------------------------------------------------------------------


def test_ac4_anthropic_isinstance_ImageDescriber() -> None:
    """AC4: AnthropicDescriber instance satisfies ImageDescriber Protocol."""
    from pptx_md.describer import ImageDescriber

    mock_sdk = _build_mock_anthropic_sdk()
    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            obj = AnthropicDescriber(api_key="test-key")
            assert isinstance(obj, ImageDescriber)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac4_openai_isinstance_ImageDescriber() -> None:
    """AC4: OpenAIDescriber instance satisfies ImageDescriber Protocol."""
    from pptx_md.describer import ImageDescriber

    mock_sdk = _build_mock_openai_sdk()
    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import OpenAIDescriber

            obj = OpenAIDescriber(api_key="test-key")
            assert isinstance(obj, ImageDescriber)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


def test_ac4_no_describe_isinstance_false() -> None:
    """AC4 negative: object without describe() is not an ImageDescriber."""
    from pptx_md.describer import ImageDescriber

    class NotADescriber:
        pass

    assert not isinstance(NotADescriber(), ImageDescriber)


# ---------------------------------------------------------------------------
# AC5 — mock SDK 예외 → DescribeError 전파, key/bytes 미노출
# ---------------------------------------------------------------------------


def test_ac5_anthropic_sdk_예외_DescribeError_전파() -> None:
    """AC5: Anthropic SDK exception is wrapped and propagated as DescribeError."""
    from pptx_md.errors import DescribeError

    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.messages.create.side_effect = RuntimeError("network failure")
    mock_sdk.Anthropic.return_value = mock_client

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            describer = AnthropicDescriber(api_key="secret-key-123")
            with pytest.raises(DescribeError):
                describer.describe(b"sensitive-bytes", "png", None)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac5_anthropic_예외_메시지_api_key_미노출() -> None:
    """AC5: DescribeError message must not contain the API key."""
    from pptx_md.errors import DescribeError

    secret_key = "sk-ant-very-secret-key"
    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.messages.create.side_effect = RuntimeError("api error")
    mock_sdk.Anthropic.return_value = mock_client

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            describer = AnthropicDescriber(api_key=secret_key)
            with pytest.raises(DescribeError) as exc_info:
                describer.describe(b"\x89PNG", "png", None)
            # API key must not appear in the exception message
            assert secret_key not in str(exc_info.value)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac5_anthropic_예외_메시지_image_bytes_미노출() -> None:
    """AC5: DescribeError message must not contain raw image bytes repr."""
    from pptx_md.errors import DescribeError

    sensitive_bytes = b"SENSITIVE_IMAGE_CONTENT_XYZ"
    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.messages.create.side_effect = RuntimeError("api error")
    mock_sdk.Anthropic.return_value = mock_client

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            describer = AnthropicDescriber(api_key="test-key")
            with pytest.raises(DescribeError) as exc_info:
                describer.describe(sensitive_bytes, "png", None)
            assert "SENSITIVE_IMAGE_CONTENT_XYZ" not in str(exc_info.value)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac5_openai_sdk_예외_DescribeError_전파() -> None:
    """AC5: OpenAI SDK exception is wrapped and propagated as DescribeError."""
    from pptx_md.errors import DescribeError

    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.chat.completions.create.side_effect = RuntimeError("network failure")
    mock_sdk.OpenAI.return_value = mock_client

    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import OpenAIDescriber

            describer = OpenAIDescriber(api_key="secret-key")
            with pytest.raises(DescribeError):
                describer.describe(b"sensitive", "png", None)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


def test_ac5_openai_예외_메시지_api_key_미노출() -> None:
    """AC5: DescribeError message must not contain the OpenAI API key."""
    from pptx_md.errors import DescribeError

    secret_key = "sk-openai-super-secret"
    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.chat.completions.create.side_effect = RuntimeError("api error")
    mock_sdk.OpenAI.return_value = mock_client

    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import OpenAIDescriber

            describer = OpenAIDescriber(api_key=secret_key)
            with pytest.raises(DescribeError) as exc_info:
                describer.describe(b"\x89PNG", "png", None)
            assert secret_key not in str(exc_info.value)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


# ---------------------------------------------------------------------------
# AC6 — describe()가 API key 또는 image_bytes를 로그에 출력하지 않음 (NFR-05/06)
# ---------------------------------------------------------------------------


def test_ac6_anthropic_로그_api_key_미출력(caplog: pytest.LogCaptureFixture) -> None:
    """AC6: API key must not appear in any log output during describe()."""
    secret_key = "sk-ant-log-test-secret"
    mock_sdk = _build_mock_anthropic_sdk("ok")

    with caplog.at_level(logging.DEBUG, logger="pptx_md.providers.anthropic"):
        with patch.dict(sys.modules, {"anthropic": mock_sdk}):
            saved = sys.modules.pop("pptx_md.providers.anthropic", None)
            try:
                from pptx_md.providers.anthropic import AnthropicDescriber

                describer = AnthropicDescriber(api_key=secret_key)
                describer.describe(b"\x89PNG", "png", None)
            finally:
                if saved is not None:
                    sys.modules["pptx_md.providers.anthropic"] = saved

    all_logs = caplog.text
    assert secret_key not in all_logs, "API key must not appear in logs"


def test_ac6_anthropic_로그_image_bytes_미출력(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """AC6: image_bytes content must not appear in any log output."""
    sensitive_content = b"SUPER_SENSITIVE_IMAGE_DATA_FOR_LOG_TEST"
    mock_sdk = _build_mock_anthropic_sdk("ok")

    with caplog.at_level(logging.DEBUG, logger="pptx_md.providers.anthropic"):
        with patch.dict(sys.modules, {"anthropic": mock_sdk}):
            saved = sys.modules.pop("pptx_md.providers.anthropic", None)
            try:
                from pptx_md.providers.anthropic import AnthropicDescriber

                describer = AnthropicDescriber(api_key="test-key")
                describer.describe(sensitive_content, "png", None)
            finally:
                if saved is not None:
                    sys.modules["pptx_md.providers.anthropic"] = saved

    all_logs = caplog.text
    assert "SUPER_SENSITIVE_IMAGE_DATA_FOR_LOG_TEST" not in all_logs


def test_ac6_openai_로그_api_key_미출력(caplog: pytest.LogCaptureFixture) -> None:
    """AC6: OpenAI API key must not appear in any log output during describe()."""
    secret_key = "sk-openai-log-test-secret"
    mock_sdk = _build_mock_openai_sdk("ok")

    with caplog.at_level(logging.DEBUG, logger="pptx_md.providers.openai"):
        with patch.dict(sys.modules, {"openai": mock_sdk}):
            saved = sys.modules.pop("pptx_md.providers.openai", None)
            try:
                from pptx_md.providers.openai import OpenAIDescriber

                describer = OpenAIDescriber(api_key=secret_key)
                describer.describe(b"\x89PNG", "png", None)
            finally:
                if saved is not None:
                    sys.modules["pptx_md.providers.openai"] = saved

    all_logs = caplog.text
    assert secret_key not in all_logs, "API key must not appear in logs"


def test_ac6_openai_로그_image_bytes_미출력(caplog: pytest.LogCaptureFixture) -> None:
    """AC6: image_bytes content must not appear in any log output for OpenAI."""
    sensitive_content = b"OPENAI_SENSITIVE_IMAGE_DATA_FOR_LOG_TEST"
    mock_sdk = _build_mock_openai_sdk("ok")

    with caplog.at_level(logging.DEBUG, logger="pptx_md.providers.openai"):
        with patch.dict(sys.modules, {"openai": mock_sdk}):
            saved = sys.modules.pop("pptx_md.providers.openai", None)
            try:
                from pptx_md.providers.openai import OpenAIDescriber

                describer = OpenAIDescriber(api_key="test-key")
                describer.describe(sensitive_content, "png", None)
            finally:
                if saved is not None:
                    sys.modules["pptx_md.providers.openai"] = saved

    all_logs = caplog.text
    assert "OPENAI_SENSITIVE_IMAGE_DATA_FOR_LOG_TEST" not in all_logs


# ---------------------------------------------------------------------------
# AC7 — mypy src/ exit 0 + ruff/black exit 0 (static quality gate)
# ---------------------------------------------------------------------------


def test_ac7_mypy_strict_exit0() -> None:
    """AC7: mypy src/ exits with code 0 (providers included)."""
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
# AC5 extra: providers/__init__.py — get_describer and __getattr__ coverage
# (issue #44: providers/__init__.py was 45% — these tests cover missing lines)
# ---------------------------------------------------------------------------


def test_ac5_extra_get_describer_anthropic_경로_커버() -> None:
    """AC5/coverage: get_describer('anthropic') must instantiate AnthropicDescriber.

    Covers providers/__init__.py lines 57-60 (the 'anthropic' branch of get_describer).
    """
    mock_sdk = _build_mock_anthropic_sdk("result text")

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers import get_describer

            describer = get_describer("anthropic", api_key="test-key")
            assert describer is not None
            # The returned object must have a describe method (ImageDescriber protocol)
            assert callable(getattr(describer, "describe", None))
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac5_extra_get_describer_openai_경로_커버() -> None:
    """AC5/coverage: get_describer('openai') must instantiate OpenAIDescriber.

    Covers providers/__init__.py lines 61-64 (the 'openai' branch of get_describer).
    """
    mock_sdk = _build_mock_openai_sdk("result text")

    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers import get_describer

            describer = get_describer("openai", api_key="test-key")
            assert describer is not None
            assert callable(getattr(describer, "describe", None))
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


def test_ac5_extra_getattr_anthropic_describer_반환() -> None:
    """AC5/coverage: providers.__getattr__('AnthropicDescriber') must return the class.

    Covers providers/__init__.py lines 73-76 (__getattr__ AnthropicDescriber branch).
    """
    mock_sdk = _build_mock_anthropic_sdk()

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        # Clear cached __getattr__ result by removing from sys.modules
        saved_pkg = sys.modules.pop("pptx_md.providers", None)
        saved_cls = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            import pptx_md.providers as providers_mod

            # Access via attribute lookup (triggers __getattr__)
            klass = providers_mod.AnthropicDescriber
            assert klass is not None
            # Instantiate to verify it's the real class
            obj = klass(api_key="test-key")
            assert callable(getattr(obj, "describe", None))
        finally:
            if saved_pkg is not None:
                sys.modules["pptx_md.providers"] = saved_pkg
            if saved_cls is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved_cls


def test_ac5_extra_getattr_openai_describer_반환() -> None:
    """AC5/coverage: providers.__getattr__('OpenAIDescriber') must return the class.

    Covers providers/__init__.py lines 77-80 (__getattr__ OpenAIDescriber branch).
    """
    mock_sdk = _build_mock_openai_sdk()

    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved_pkg = sys.modules.pop("pptx_md.providers", None)
        saved_cls = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            import pptx_md.providers as providers_mod

            klass = providers_mod.OpenAIDescriber
            assert klass is not None
            obj = klass(api_key="test-key")
            assert callable(getattr(obj, "describe", None))
        finally:
            if saved_pkg is not None:
                sys.modules["pptx_md.providers"] = saved_pkg
            if saved_cls is not None:
                sys.modules["pptx_md.providers.openai"] = saved_cls


def test_ac5_extra_getattr_unknown_AttributeError() -> None:
    """AC5/coverage: providers.__getattr__('NonExistent') raises AttributeError.

    Covers providers/__init__.py line 81 (AttributeError raise path).
    """
    import pptx_md.providers as providers_mod

    with pytest.raises(AttributeError, match="no attribute"):
        _ = providers_mod.NonExistentClass  # type: ignore[attr-defined]
