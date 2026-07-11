"""Tests for provider `reconstruct()` — FR-32 (issue #94, ADR-625/626).

Covers `SlideReconstructor.reconstruct(image_bytes, image_ext, context) -> str`
as implemented by `AnthropicDescriber`/`OpenAIDescriber`. All SDK calls are
mocked (no real VLM API calls).

Test naming convention: test_ac<N>_<description>, N = FR-32 AC number
(REQ-v020.md §4). AC3/AC4 map to protocol-shape and prompt-strategy checks
respectively (no AC3 orchestration-level test here — that belongs to #93/#95).
"""

from __future__ import annotations

import logging
import sys
from unittest.mock import MagicMock, patch

import pytest

from pptx_md.slide_describer import SlideContext, SlideReconstructor

# ---------------------------------------------------------------------------
# Shared mock-response helpers (mirrors tests/test_providers.py conventions)
# ---------------------------------------------------------------------------


def _make_anthropic_mock_response(text: str) -> MagicMock:
    block = MagicMock()
    block.text = text
    response = MagicMock()
    response.content = [block]
    return response


def _make_openai_mock_response(text: str) -> MagicMock:
    msg = MagicMock()
    msg.content = text
    choice = MagicMock()
    choice.message = msg
    response = MagicMock()
    response.choices = [choice]
    return response


def _build_mock_anthropic_sdk(response_text: str = "## mocked\n\nbody") -> MagicMock:
    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.messages.create.return_value = _make_anthropic_mock_response(
        response_text
    )
    mock_sdk.Anthropic.return_value = mock_client
    return mock_sdk


def _build_mock_openai_sdk(response_text: str = "## mocked\n\nbody") -> MagicMock:
    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.chat.completions.create.return_value = _make_openai_mock_response(
        response_text
    )
    mock_sdk.OpenAI.return_value = mock_client
    return mock_sdk


_CTX = SlideContext(slide_index=3, title="S5 4-column mapping", text_outline="")

_TABLE_MD = (
    "| Region | Q1 | Q2 | Q3 |\n"
    "| --- | --- | --- | --- |\n"
    "| APAC | 10 | 20 | 30 |\n"
)


# ---------------------------------------------------------------------------
# AC1 — 유효 Markdown 조각 반환
# ---------------------------------------------------------------------------


def test_ac1_anthropic_reconstruct_유효_md_반환() -> None:
    """AC1: AnthropicDescriber.reconstruct returns the provider's MD fragment."""
    mock_sdk = _build_mock_anthropic_sdk(_TABLE_MD)

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            provider = AnthropicDescriber(api_key="test-key")
            result = provider.reconstruct(b"\x89PNG", "png", _CTX)
            assert result == _TABLE_MD
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac1_openai_reconstruct_유효_md_반환() -> None:
    """AC1: OpenAIDescriber.reconstruct returns the provider's MD fragment."""
    mock_sdk = _build_mock_openai_sdk(_TABLE_MD)

    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import OpenAIDescriber

            provider = OpenAIDescriber(api_key="test-key")
            result = provider.reconstruct(b"\x89PNG", "png", _CTX)
            assert result == _TABLE_MD
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


# ---------------------------------------------------------------------------
# AC2 — pipe table 정합(헤더+구분행+데이터행) 요청이 프롬프트에 포함되고
#        provider 는 응답을 있는 그대로 통과시켜 표 구조를 보존한다.
# ---------------------------------------------------------------------------


def test_ac2_anthropic_reconstruct_표_정합_요청_프롬프트_포함() -> None:
    """AC2: prompt requests header+separator+data-row pipe tables."""
    mock_sdk = _build_mock_anthropic_sdk(_TABLE_MD)

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            mock_client = mock_sdk.Anthropic.return_value
            provider = AnthropicDescriber(api_key="test-key")
            result = provider.reconstruct(b"\x89PNG", "png", _CTX)

            call_args = mock_client.messages.create.call_args
            messages = call_args.kwargs.get("messages") or call_args.args[0]
            prompt_text = ""
            for msg in messages:
                for block in msg.get("content", []):
                    if block.get("type") == "text":
                        prompt_text += block.get("text", "")

            assert "header row" in prompt_text
            assert "separator row" in prompt_text
            assert "data row" in prompt_text
            # Table structure is preserved verbatim in the returned fragment.
            assert result.count("|") == _TABLE_MD.count("|")
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


# ---------------------------------------------------------------------------
# AC3 — 프로토콜 관계: reconstruct() 구현체가 SlideReconstructor 구조 만족
# ---------------------------------------------------------------------------


def test_ac3_anthropic_isinstance_SlideReconstructor() -> None:
    """AC3: AnthropicDescriber instance satisfies SlideReconstructor Protocol."""
    mock_sdk = _build_mock_anthropic_sdk()
    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            obj = AnthropicDescriber(api_key="test-key")
            assert isinstance(obj, SlideReconstructor)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac3_openai_isinstance_SlideReconstructor() -> None:
    """AC3: OpenAIDescriber instance satisfies SlideReconstructor Protocol."""
    mock_sdk = _build_mock_openai_sdk()
    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import OpenAIDescriber

            obj = OpenAIDescriber(api_key="test-key")
            assert isinstance(obj, SlideReconstructor)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


def test_ac3_anthropic_describe_와_reconstruct_병행_구현() -> None:
    """AC3: a single provider instance implements both describe() and
    reconstruct() (protocols are separate but co-implementable, ARCH-v020 §3.3)."""
    mock_sdk = _build_mock_anthropic_sdk("desc")
    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.describer import ImageDescriber
            from pptx_md.providers.anthropic import AnthropicDescriber

            obj = AnthropicDescriber(api_key="test-key")
            assert isinstance(obj, ImageDescriber)
            assert isinstance(obj, SlideReconstructor)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


# ---------------------------------------------------------------------------
# AC4 — 프롬프트 전략: 표/중첩리스트/흐름/차트 우선 구조화, 불확실시 원문 보존,
#        순수 Markdown만 출력(코드펜스 금지) 이 모듈 상수 + 실제 전송 프롬프트에 반영
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("module_name", ["anthropic", "openai"])
def test_ac4_prompt_strategy_모듈_상수_추적성(module_name: str) -> None:
    """AC4: RECONSTRUCT_PROMPT_STRATEGY module constant encodes the required
    prompt policy (tables/nested lists/flow/chart priority, uncertain->verbatim,
    pure Markdown only / no code fences)."""
    import importlib

    mod = importlib.import_module(f"pptx_md.providers.{module_name}")
    strategy = mod.RECONSTRUCT_PROMPT_STRATEGY
    assert "table" in strategy.lower()
    assert "nested list" in strategy.lower()
    assert "flow" in strategy.lower()
    assert "chart" in strategy.lower()
    assert "verbatim" in strategy.lower()
    assert "no code fences" in strategy.lower() or "code fences" in strategy.lower()
    assert "no explanatory prose" in strategy.lower() or "prose" in strategy.lower()


def test_ac4_anthropic_context_title_and_outline_프롬프트_포함() -> None:
    """AC4: SlideContext.title / text_outline are woven into the sent prompt."""
    mock_sdk = _build_mock_anthropic_sdk("ok")
    ctx = SlideContext(
        slide_index=0, title="AS-IS -> TO-BE", text_outline="line one\nline two"
    )

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            mock_client = mock_sdk.Anthropic.return_value
            provider = AnthropicDescriber(api_key="test-key")
            provider.reconstruct(b"\x89PNG", "png", ctx)

            call_args = mock_client.messages.create.call_args
            messages = call_args.kwargs.get("messages") or call_args.args[0]
            prompt_text = ""
            for msg in messages:
                for block in msg.get("content", []):
                    if block.get("type") == "text":
                        prompt_text += block.get("text", "")

            assert "AS-IS -> TO-BE" in prompt_text
            assert "line one" in prompt_text
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac4_anthropic_reconstruct_uses_separate_model_from_describe() -> None:
    """AC4/설계: reconstruct() 는 describe() 와 별도의 고성능 모델 상수를 쓴다."""
    mock_sdk = _build_mock_anthropic_sdk("ok")
    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import (
                _DEFAULT_MODEL,
                _RECONSTRUCT_MODEL,
                AnthropicDescriber,
            )

            mock_client = mock_sdk.Anthropic.return_value
            provider = AnthropicDescriber(api_key="test-key")
            provider.reconstruct(b"\x89PNG", "png", _CTX)

            call_args = mock_client.messages.create.call_args
            assert call_args.kwargs["model"] == _RECONSTRUCT_MODEL
            assert _RECONSTRUCT_MODEL != _DEFAULT_MODEL
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac4_openai_reconstruct_uses_separate_model_from_describe() -> None:
    """AC4/설계: reconstruct() 는 describe() 와 별도의 고성능 모델 상수를 쓴다."""
    mock_sdk = _build_mock_openai_sdk("ok")
    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import (
                _DEFAULT_MODEL,
                _RECONSTRUCT_MODEL,
                OpenAIDescriber,
            )

            mock_client = mock_sdk.OpenAI.return_value
            provider = OpenAIDescriber(api_key="test-key")
            provider.reconstruct(b"\x89PNG", "png", _CTX)

            call_args = mock_client.chat.completions.create.call_args
            assert call_args.kwargs["model"] == _RECONSTRUCT_MODEL
            assert _RECONSTRUCT_MODEL != _DEFAULT_MODEL
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


# ---------------------------------------------------------------------------
# AC5 — 실패/빈 응답 → DescribeError (기존 예외 표준 재사용)
# ---------------------------------------------------------------------------


def test_ac5_anthropic_sdk_예외_DescribeError_전파() -> None:
    """AC5: Anthropic SDK exception during reconstruct() -> DescribeError."""
    from pptx_md.errors import DescribeError

    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.messages.create.side_effect = RuntimeError("network failure")
    mock_sdk.Anthropic.return_value = mock_client

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            provider = AnthropicDescriber(api_key="secret-key")
            with pytest.raises(DescribeError):
                provider.reconstruct(b"sensitive-bytes", "png", _CTX)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac5_anthropic_빈_응답_DescribeError() -> None:
    """AC5: an empty reconstruct() response is treated as failure -> DescribeError."""
    from pptx_md.errors import DescribeError

    mock_sdk = _build_mock_anthropic_sdk("")
    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            provider = AnthropicDescriber(api_key="test-key")
            with pytest.raises(DescribeError):
                provider.reconstruct(b"\x89PNG", "png", _CTX)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_ac5_openai_sdk_예외_DescribeError_전파() -> None:
    """AC5: OpenAI SDK exception during reconstruct() -> DescribeError."""
    from pptx_md.errors import DescribeError

    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.chat.completions.create.side_effect = RuntimeError("network failure")
    mock_sdk.OpenAI.return_value = mock_client

    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import OpenAIDescriber

            provider = OpenAIDescriber(api_key="secret-key")
            with pytest.raises(DescribeError):
                provider.reconstruct(b"sensitive-bytes", "png", _CTX)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


def test_ac5_openai_빈_응답_DescribeError() -> None:
    """AC5: an empty reconstruct() response is treated as failure -> DescribeError."""
    from pptx_md.errors import DescribeError

    mock_sdk = _build_mock_openai_sdk("")
    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import OpenAIDescriber

            provider = OpenAIDescriber(api_key="test-key")
            with pytest.raises(DescribeError):
                provider.reconstruct(b"\x89PNG", "png", _CTX)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved


def test_ac5_anthropic_예외_메시지_api_key_bytes_미노출() -> None:
    """AC5/AC7: DescribeError message must not leak the API key or image bytes."""
    from pptx_md.errors import DescribeError

    secret_key = "sk-ant-reconstruct-secret"
    sensitive_bytes = b"SENSITIVE_SLIDE_PNG_CONTENT"
    mock_sdk = MagicMock()
    mock_client = MagicMock()
    mock_client.messages.create.side_effect = RuntimeError("api error")
    mock_sdk.Anthropic.return_value = mock_client

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            provider = AnthropicDescriber(api_key=secret_key)
            with pytest.raises(DescribeError) as exc_info:
                provider.reconstruct(sensitive_bytes, "png", _CTX)
            assert secret_key not in str(exc_info.value)
            assert "SENSITIVE_SLIDE_PNG_CONTENT" not in str(exc_info.value)
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


# ---------------------------------------------------------------------------
# AC6 — 구조화 실패(산문) 시에도 원문 그대로 통과(콘텐츠 손실 0)
# ---------------------------------------------------------------------------


def test_ac6_anthropic_산문_응답_원문_그대로_보존() -> None:
    """AC6: an unstructured prose response is returned verbatim (no mangling,
    no content loss) rather than raising or truncating."""
    prose = (
        "This slide could not be confidently structured; "
        "original text follows: foo bar baz."
    )
    mock_sdk = _build_mock_anthropic_sdk(prose)

    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import AnthropicDescriber

            provider = AnthropicDescriber(api_key="test-key")
            result = provider.reconstruct(b"\x89PNG", "png", _CTX)
            assert result == prose
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


# ---------------------------------------------------------------------------
# AC7 — PII: PNG/프롬프트/응답/API key 로깅 시 원문 미출력
# ---------------------------------------------------------------------------


def test_ac7_anthropic_로그_api_key_미출력(caplog: pytest.LogCaptureFixture) -> None:
    """AC7: API key must not appear in any log output during reconstruct()."""
    secret_key = "sk-ant-reconstruct-log-secret"
    mock_sdk = _build_mock_anthropic_sdk("ok")

    with caplog.at_level(logging.DEBUG, logger="pptx_md.providers.anthropic"):
        with patch.dict(sys.modules, {"anthropic": mock_sdk}):
            saved = sys.modules.pop("pptx_md.providers.anthropic", None)
            try:
                from pptx_md.providers.anthropic import AnthropicDescriber

                provider = AnthropicDescriber(api_key=secret_key)
                provider.reconstruct(b"\x89PNG", "png", _CTX)
            finally:
                if saved is not None:
                    sys.modules["pptx_md.providers.anthropic"] = saved

    assert secret_key not in caplog.text


def test_ac7_anthropic_로그_슬라이드_제목_텍스트_미출력(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """AC7/PII: SlideContext.title / text_outline (potential PII) never logged."""
    mock_sdk = _build_mock_anthropic_sdk("ok")
    ctx = SlideContext(
        slide_index=0,
        title="CONFIDENTIAL_TITLE_PII",
        text_outline="CONFIDENTIAL_BODY_PII_TEXT",
    )

    with caplog.at_level(logging.DEBUG, logger="pptx_md.providers.anthropic"):
        with patch.dict(sys.modules, {"anthropic": mock_sdk}):
            saved = sys.modules.pop("pptx_md.providers.anthropic", None)
            try:
                from pptx_md.providers.anthropic import AnthropicDescriber

                provider = AnthropicDescriber(api_key="test-key")
                provider.reconstruct(b"\x89PNG", "png", ctx)
            finally:
                if saved is not None:
                    sys.modules["pptx_md.providers.anthropic"] = saved

    assert "CONFIDENTIAL_TITLE_PII" not in caplog.text
    assert "CONFIDENTIAL_BODY_PII_TEXT" not in caplog.text


def test_ac7_anthropic_로그_이미지_바이트_미출력(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """AC7: rendered PNG bytes content must not appear in any log output."""
    sensitive = b"SUPER_SENSITIVE_SLIDE_PNG_BYTES_FOR_LOG_TEST"
    mock_sdk = _build_mock_anthropic_sdk("ok")

    with caplog.at_level(logging.DEBUG, logger="pptx_md.providers.anthropic"):
        with patch.dict(sys.modules, {"anthropic": mock_sdk}):
            saved = sys.modules.pop("pptx_md.providers.anthropic", None)
            try:
                from pptx_md.providers.anthropic import AnthropicDescriber

                provider = AnthropicDescriber(api_key="test-key")
                provider.reconstruct(sensitive, "png", _CTX)
            finally:
                if saved is not None:
                    sys.modules["pptx_md.providers.anthropic"] = saved

    assert "SUPER_SENSITIVE_SLIDE_PNG_BYTES_FOR_LOG_TEST" not in caplog.text


def test_ac7_openai_로그_api_key_미출력(caplog: pytest.LogCaptureFixture) -> None:
    """AC7: OpenAI API key must not appear in any log output during reconstruct()."""
    secret_key = "sk-openai-reconstruct-log-secret"
    mock_sdk = _build_mock_openai_sdk("ok")

    with caplog.at_level(logging.DEBUG, logger="pptx_md.providers.openai"):
        with patch.dict(sys.modules, {"openai": mock_sdk}):
            saved = sys.modules.pop("pptx_md.providers.openai", None)
            try:
                from pptx_md.providers.openai import OpenAIDescriber

                provider = OpenAIDescriber(api_key=secret_key)
                provider.reconstruct(b"\x89PNG", "png", _CTX)
            finally:
                if saved is not None:
                    sys.modules["pptx_md.providers.openai"] = saved

    assert secret_key not in caplog.text


def test_ac7_openai_로그_슬라이드_제목_텍스트_미출력(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """AC7/PII: SlideContext.title / text_outline never logged (OpenAI provider)."""
    mock_sdk = _build_mock_openai_sdk("ok")
    ctx = SlideContext(
        slide_index=0,
        title="CONFIDENTIAL_TITLE_PII_OAI",
        text_outline="CONFIDENTIAL_BODY_PII_TEXT_OAI",
    )

    with caplog.at_level(logging.DEBUG, logger="pptx_md.providers.openai"):
        with patch.dict(sys.modules, {"openai": mock_sdk}):
            saved = sys.modules.pop("pptx_md.providers.openai", None)
            try:
                from pptx_md.providers.openai import OpenAIDescriber

                provider = OpenAIDescriber(api_key="test-key")
                provider.reconstruct(b"\x89PNG", "png", ctx)
            finally:
                if saved is not None:
                    sys.modules["pptx_md.providers.openai"] = saved

    assert "CONFIDENTIAL_TITLE_PII_OAI" not in caplog.text
    assert "CONFIDENTIAL_BODY_PII_TEXT_OAI" not in caplog.text


def test_ac7_openai_로그_이미지_바이트_미출력(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """AC7: rendered PNG bytes content must not appear in any log output (OpenAI)."""
    sensitive = b"OPENAI_SENSITIVE_SLIDE_PNG_BYTES_FOR_LOG_TEST"
    mock_sdk = _build_mock_openai_sdk("ok")

    with caplog.at_level(logging.DEBUG, logger="pptx_md.providers.openai"):
        with patch.dict(sys.modules, {"openai": mock_sdk}):
            saved = sys.modules.pop("pptx_md.providers.openai", None)
            try:
                from pptx_md.providers.openai import OpenAIDescriber

                provider = OpenAIDescriber(api_key="test-key")
                provider.reconstruct(sensitive, "png", _CTX)
            finally:
                if saved is not None:
                    sys.modules["pptx_md.providers.openai"] = saved

    assert "OPENAI_SENSITIVE_SLIDE_PNG_BYTES_FOR_LOG_TEST" not in caplog.text


# ---------------------------------------------------------------------------
# 기존 describe() 무변경(회귀 0) — 신규 reconstruct() 도입 후에도 describe()가
# 여전히 별도의 저비용 모델을 사용함을 재확인 (anthropic._DEFAULT_MODEL 등).
# ---------------------------------------------------------------------------


def test_regression_anthropic_describe_여전히_기본_모델_사용() -> None:
    """Regression: describe() must still use self._model (unaffected by
    reconstruct()'s separate self._reconstruct_model)."""
    mock_sdk = _build_mock_anthropic_sdk("a description")
    with patch.dict(sys.modules, {"anthropic": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.anthropic", None)
        try:
            from pptx_md.providers.anthropic import _DEFAULT_MODEL, AnthropicDescriber

            mock_client = mock_sdk.Anthropic.return_value
            provider = AnthropicDescriber(api_key="test-key")
            provider.describe(b"\x89PNG", "png", "hint")

            call_args = mock_client.messages.create.call_args
            assert call_args.kwargs["model"] == _DEFAULT_MODEL
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.anthropic"] = saved


def test_regression_openai_describe_여전히_기본_모델_사용() -> None:
    """Regression: describe() must still use self._model (unaffected by
    reconstruct()'s separate self._reconstruct_model)."""
    mock_sdk = _build_mock_openai_sdk("a description")
    with patch.dict(sys.modules, {"openai": mock_sdk}):
        saved = sys.modules.pop("pptx_md.providers.openai", None)
        try:
            from pptx_md.providers.openai import _DEFAULT_MODEL, OpenAIDescriber

            mock_client = mock_sdk.OpenAI.return_value
            provider = OpenAIDescriber(api_key="test-key")
            provider.describe(b"\x89PNG", "png", "hint")

            call_args = mock_client.chat.completions.create.call_args
            assert call_args.kwargs["model"] == _DEFAULT_MODEL
        finally:
            if saved is not None:
                sys.modules["pptx_md.providers.openai"] = saved
