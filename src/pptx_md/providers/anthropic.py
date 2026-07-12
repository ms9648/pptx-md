"""AnthropicDescriber — VLM provider backed by the Anthropic SDK (FR-09).

SDK import is deferred to the constructor (ADR-213 fail-fast):
    ``import pptx_md.providers.anthropic`` succeeds without the SDK (NFR-08).
    ``AnthropicDescriber(...)`` raises :class:`~pptx_md.errors.InstallationError`
    if ``anthropic`` is not installed.

Logging: only metadata (shape_id, media_type, hint class-name, exception type).
API key, image bytes, prompt body, and VLM response text are NEVER logged
(NFR-05/06).
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING, Any, Literal

from pptx_md.errors import DescribeError, InstallationError
from pptx_md.slide_describer import SlideContext

if TYPE_CHECKING:  # pragma: no cover — type-only, never imported at runtime
    from anthropic.types import ImageBlockParam, MessageParam, TextBlockParam

__all__ = ["AnthropicDescriber"]

_logger = logging.getLogger("pptx_md.providers.anthropic")

_DEFAULT_MODEL = "claude-haiku-4-5-20251001"

#: FR-32/ADR-626 (DEC-3/DEC-6, 사람 승인 대상): 슬라이드 재구성(reconstruct)은
#: describe()의 저비용 haiku 모델과 **분리된** 별도의 고성능(품질 우선) 모델을
#: 기본값으로 쓴다. 실값(모델명/토큰 상한)은 사람 승인 전 미고정 — provider
#: 상수로 노출해 추적성을 확보한다(ARCH-v020 §3.3 "모델·비용 수치 제안" 표).
_RECONSTRUCT_MODEL = "claude-opus-4-8"

#: 재구성 산출(표/중첩리스트 다량)은 describe(1024)보다 분량이 크므로 상향
#: (ADR-626 제안).
_RECONSTRUCT_MAX_TOKENS = 4096

_MediaType = Literal["image/png", "image/jpeg", "image/gif", "image/webp"]

#: FR-32 AC4 프롬프트 전략 — "표/중첩리스트/흐름/차트 데이터 우선 구조화, 불확실
#: 하면 원문 텍스트 보존. 순수 Markdown만 출력(설명 산문·코드펜스 금지)."
#: 모듈 상수로 두어 추적성을 확보한다(mermaid.DIAGRAM_HINT_SUFFIX 전례,
#: ARCH-v020 §3.3).
RECONSTRUCT_PROMPT_STRATEGY: str = (
    "You are reconstructing a single presentation slide (given as a whole-slide "
    "image) into a structured Markdown fragment. Prioritize, in this order: "
    "(1) tables -- emit a well-formed Markdown pipe table with a header row, a "
    "separator row, and at least one data row; "
    "(2) nested lists -- represent hierarchy using indentation; "
    "(3) flows/processes (arrows, connectors, sequential steps) -- represent as "
    "an ordered list or an explicit left-to-right/top-to-bottom sequence; "
    "(4) chart data (bar/line/pie/donut) -- extract the underlying category/value "
    "data as a table instead of describing the visual. "
    "If the structure of any element is uncertain, preserve the original visible "
    "text verbatim rather than emit a broken or guessed structure -- never drop "
    "content. "
    "Output ONLY the reconstructed Markdown body: no explanatory prose, no "
    "preamble/epilogue, and no code fences (do not wrap the answer in ``` marks)."
)


def _build_reconstruct_prompt(context: SlideContext) -> str:
    """Assemble the user-turn text for the slide-reconstruction call.

    Neither ``context.title`` nor ``context.text_outline`` content is logged
    (NFR-06) -- only debug metadata (lengths/booleans) is logged by the
    caller.
    """
    parts: list[str] = [RECONSTRUCT_PROMPT_STRATEGY]
    if context.title:
        parts.append(f"Slide title: {context.title}")
    if context.text_outline:
        parts.append(
            "Deterministic text-path extraction for this slide (grounding "
            "reference; do not just copy it back unless structure is "
            f"uncertain):\n{context.text_outline}"
        )
    return "\n\n".join(parts)


def _build_prompt(shape_hint: str | None, alt_text: str) -> str:
    """Assemble the user-turn text for the messages API.

    Neither shape_hint content nor alt_text content is logged (NFR-06).
    """
    parts: list[str] = []
    if shape_hint:
        parts.append(shape_hint)
    if alt_text:
        parts.append(f"Alt text: {alt_text}")
    parts.append("Describe this image in detail.")
    return " ".join(parts)


def _extract_text(response: Any) -> str:
    """Extract the text string from an Anthropic messages response."""
    content = response.content
    if isinstance(content, list) and content:
        block = content[0]
        if hasattr(block, "text") and isinstance(block.text, str):
            return block.text
    return ""


class AnthropicDescriber:
    """Anthropic Claude-based image describer (FR-09, ADR-212/213).

    Structurally satisfies :class:`~pptx_md.describer.ImageDescriber` Protocol
    without nominal inheritance (ADR-217).

    SDK (``anthropic``) is imported inside ``__init__`` only — fail-fast on
    missing installation (ADR-213).  ``import pptx_md.providers.anthropic``
    always succeeds (NFR-08).

    Args:
        api_key: Anthropic API key.  ``None`` (default) → the SDK reads the
            ``ANTHROPIC_API_KEY`` environment variable (NFR-05).
        model: Claude model identifier used by ``describe()``.  Defaults to
            ``claude-haiku-4-5-20251001`` (low-cost).
        reconstruct_model: Claude model identifier used by ``reconstruct()``
            (FR-32).  Defaults to a **separate**, higher-tier model
            (``_RECONSTRUCT_MODEL``) per the quality-first policy (DEC-3/5,
            ADR-626) — independent from ``model`` above.

    Raises:
        InstallationError: If the ``anthropic`` package is not installed.
    """

    def __init__(
        self,
        api_key: str | None = None,
        model: str = _DEFAULT_MODEL,
        reconstruct_model: str = _RECONSTRUCT_MODEL,
    ) -> None:
        # ADR-213: conditional import in constructor — fail-fast
        try:
            import anthropic as _anthropic_sdk  # noqa: PLC0415
        except ImportError as exc:
            raise InstallationError(
                "Anthropic provider requires the VLM extras: "
                "pip install pptx-md[vlm]"
            ) from exc

        # api_key=None → SDK reads ANTHROPIC_API_KEY env var (NFR-05)
        self._client = _anthropic_sdk.Anthropic(api_key=api_key)
        self._model = model
        self._reconstruct_model = reconstruct_model

    def describe(
        self,
        image_bytes: bytes,
        image_ext: str,
        shape_hint: str | None,
    ) -> str:
        """Return a natural-language description of the image.

        Calls the Anthropic messages API with the image as a base64 payload.
        API key, image bytes, and prompt body are never logged (NFR-05/06).

        Args:
            image_bytes: Raw raster image bytes.
            image_ext: Lowercase extension without leading dot (e.g. ``"png"``).
            shape_hint: Optional classification-derived hint (FR-10).
                ``None`` → generic prompt.

        Returns:
            Description string.

        Raises:
            DescribeError: On any API/network/quota failure (ADR-214).
        """
        import base64  # stdlib

        media_type_map: dict[str, _MediaType] = {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "webp": "image/webp",
        }
        media_type: _MediaType = media_type_map.get(image_ext.lower(), "image/png")

        prompt_text = _build_prompt(shape_hint, "")
        image_data = base64.standard_b64encode(image_bytes).decode("ascii")

        _logger.debug(
            "anthropic describe: media_type=%r, hint_set=%s, bytes_len=%d",
            media_type,
            bool(shape_hint),
            len(image_bytes),
        )

        image_block: ImageBlockParam = {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": media_type,
                "data": image_data,
            },
        }
        text_block: TextBlockParam = {
            "type": "text",
            "text": prompt_text,
        }
        message: MessageParam = {
            "role": "user",
            "content": [image_block, text_block],
        }

        try:
            response = self._client.messages.create(
                model=self._model,
                max_tokens=1024,
                messages=[message],
            )
        except Exception as exc:
            # Wrap all SDK/network failures — never expose key or image_bytes
            # (NFR-05/06). Log only exception type name.
            _logger.warning(
                "anthropic describe failed: %s",
                type(exc).__name__,
            )
            raise DescribeError(
                f"Anthropic describe failed: {type(exc).__name__}"
            ) from exc

        result = _extract_text(response)
        _logger.debug(
            "anthropic describe: success, result_empty=%s",
            not result,
        )
        return result

    def reconstruct(
        self,
        image_bytes: bytes,
        image_ext: str,
        context: SlideContext,
    ) -> str:
        """Return a structured Markdown fragment reconstructed from a slide image.

        Structurally satisfies
        :class:`~pptx_md.slide_describer.SlideReconstructor` (FR-32, ADR-625).
        Calls the Anthropic messages API with the whole-slide render as a
        base64 payload plus the FR-32 AC4 prompt strategy
        (:data:`RECONSTRUCT_PROMPT_STRATEGY`). Uses a **separate**,
        higher-tier model (``self._reconstruct_model``) from ``describe()``
        (quality-first, DEC-3/5, ADR-626). API key, image bytes, prompt
        body, and the raw response text are never logged (NFR-05/06,
        FR-32 AC7).

        Args:
            image_bytes: Raw rendered whole-slide PNG bytes.
            image_ext: Lowercase extension without leading dot (e.g. ``"png"``).
            context: Deterministic slide context (title, text outline,
                slide index) grounding the reconstruction.

        Returns:
            A structured Markdown fragment (FR-32 AC1/AC2).

        Raises:
            DescribeError: On any API/network/quota failure, or an empty
                response (FR-32 AC5; orchestrator isolates and falls back
                to the text path, INV-4).
        """
        import base64  # stdlib

        media_type_map: dict[str, _MediaType] = {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "webp": "image/webp",
        }
        media_type: _MediaType = media_type_map.get(image_ext.lower(), "image/png")

        prompt_text = _build_reconstruct_prompt(context)
        image_data = base64.standard_b64encode(image_bytes).decode("ascii")

        _logger.debug(
            "anthropic reconstruct: slide_index=%d, media_type=%r, bytes_len=%d",
            context.slide_index,
            media_type,
            len(image_bytes),
        )

        image_block: ImageBlockParam = {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": media_type,
                "data": image_data,
            },
        }
        text_block: TextBlockParam = {
            "type": "text",
            "text": prompt_text,
        }
        message: MessageParam = {
            "role": "user",
            "content": [image_block, text_block],
        }

        try:
            response = self._client.messages.create(
                model=self._reconstruct_model,
                max_tokens=_RECONSTRUCT_MAX_TOKENS,
                messages=[message],
            )
        except Exception as exc:
            # Wrap all SDK/network failures — never expose key or image_bytes
            # (NFR-05/06). Log only exception type name.
            _logger.warning(
                "anthropic reconstruct failed: slide_index=%d, error=%s",
                context.slide_index,
                type(exc).__name__,
            )
            raise DescribeError(
                f"Anthropic reconstruct failed: {type(exc).__name__}"
            ) from exc

        result = _extract_text(response)
        if not result:
            _logger.warning(
                "anthropic reconstruct: empty response, slide_index=%d",
                context.slide_index,
            )
            raise DescribeError("Anthropic reconstruct returned an empty response")

        _logger.debug(
            "anthropic reconstruct: success, slide_index=%d",
            context.slide_index,
        )
        return result
