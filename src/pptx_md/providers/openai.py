"""OpenAIDescriber — VLM provider backed by the OpenAI SDK (FR-09).

SDK import is deferred to the constructor (ADR-213 fail-fast):
    ``import pptx_md.providers.openai`` succeeds without the SDK (NFR-08).
    ``OpenAIDescriber(...)`` raises :class:`~pptx_md.errors.InstallationError`
    if ``openai`` is not installed.

Logging: only metadata (media_type, hint presence, bytes length, exception type).
API key, image bytes, prompt body, and VLM response text are NEVER logged
(NFR-05/06).
"""

from __future__ import annotations

import logging
from typing import Any

from pptx_md.errors import DescribeError, InstallationError

__all__ = ["OpenAIDescriber"]

_logger = logging.getLogger("pptx_md.providers.openai")

_DEFAULT_MODEL = "gpt-4o-mini"


def _build_prompt(shape_hint: str | None, alt_text: str) -> str:
    """Assemble the user-turn text for the chat completions API.

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
    """Extract the text string from an OpenAI chat completion response."""
    choices = response.choices
    if isinstance(choices, list) and choices:
        choice = choices[0]
        msg = getattr(choice, "message", None)
        if msg is not None:
            content = getattr(msg, "content", None)
            if isinstance(content, str):
                return content
    return ""


class OpenAIDescriber:
    """OpenAI GPT-4-vision-based image describer (FR-09, ADR-212/213).

    Structurally satisfies :class:`~pptx_md.describer.ImageDescriber` Protocol
    without nominal inheritance (ADR-217).

    SDK (``openai``) is imported inside ``__init__`` only — fail-fast on
    missing installation (ADR-213).  ``import pptx_md.providers.openai``
    always succeeds (NFR-08).

    Args:
        api_key: OpenAI API key.  ``None`` (default) → the SDK reads the
            ``OPENAI_API_KEY`` environment variable (NFR-05).
        model: OpenAI model identifier.  Defaults to ``gpt-4o-mini``.

    Raises:
        InstallationError: If the ``openai`` package is not installed.
    """

    def __init__(
        self,
        api_key: str | None = None,
        model: str = _DEFAULT_MODEL,
    ) -> None:
        # ADR-213: conditional import in constructor — fail-fast
        try:
            import openai as _openai_sdk  # noqa: PLC0415
        except ImportError as exc:
            raise InstallationError(
                "OpenAI provider requires the VLM extras: " "pip install pptx-md[vlm]"
            ) from exc

        # api_key=None → SDK reads OPENAI_API_KEY env var (NFR-05)
        self._client = _openai_sdk.OpenAI(api_key=api_key)
        self._model = model

    def describe(
        self,
        image_bytes: bytes,
        image_ext: str,
        shape_hint: str | None,
    ) -> str:
        """Return a natural-language description of the image.

        Calls the OpenAI chat completions API with the image as a base64 URL.
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

        media_type_map: dict[str, str] = {
            "png": "image/png",
            "jpg": "image/jpeg",
            "jpeg": "image/jpeg",
            "gif": "image/gif",
            "webp": "image/webp",
        }
        media_type = media_type_map.get(image_ext.lower(), "image/png")

        prompt_text = _build_prompt(shape_hint, "")
        image_data = base64.standard_b64encode(image_bytes).decode("ascii")
        image_url = f"data:{media_type};base64,{image_data}"

        _logger.debug(
            "openai describe: media_type=%r, hint_set=%s, bytes_len=%d",
            media_type,
            bool(shape_hint),
            len(image_bytes),
        )

        try:
            response = self._client.chat.completions.create(
                model=self._model,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image_url",
                                "image_url": {"url": image_url},
                            },
                            {
                                "type": "text",
                                "text": prompt_text,
                            },
                        ],
                    }
                ],
                max_tokens=1024,
            )
        except Exception as exc:
            # Wrap all SDK/network failures — never expose key or image_bytes
            # (NFR-05/06). Log only exception type name.
            _logger.warning(
                "openai describe failed: %s",
                type(exc).__name__,
            )
            raise DescribeError(
                f"OpenAI describe failed: {type(exc).__name__}"
            ) from exc

        result = _extract_text(response)
        _logger.debug(
            "openai describe: success, result_empty=%s",
            not result,
        )
        return result
