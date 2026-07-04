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

if TYPE_CHECKING:  # pragma: no cover — type-only, never imported at runtime
    from anthropic.types import ImageBlockParam, MessageParam, TextBlockParam

__all__ = ["AnthropicDescriber"]

_logger = logging.getLogger("pptx_md.providers.anthropic")

_DEFAULT_MODEL = "claude-haiku-4-5-20251001"

_MediaType = Literal["image/png", "image/jpeg", "image/gif", "image/webp"]


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
        model: Claude model identifier.  Defaults to ``claude-haiku-4-5-20251001``.

    Raises:
        InstallationError: If the ``anthropic`` package is not installed.
    """

    def __init__(
        self,
        api_key: str | None = None,
        model: str = _DEFAULT_MODEL,
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
