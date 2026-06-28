"""VLM provider package (FR-09, ADR-212).

Provides :class:`AnthropicDescriber` and :class:`OpenAIDescriber` — both
structurally satisfy the :class:`~pptx_md.describer.ImageDescriber` Protocol
(ADR-217).

**NFR-08 guarantee**: ``import pptx_md.providers`` always succeeds regardless
of whether the VLM SDKs are installed.  SDK imports happen inside each
provider's constructor (ADR-213).

Usage::

    from pptx_md.providers import AnthropicDescriber, OpenAIDescriber

    # raises InstallationError if SDK absent
    describer = AnthropicDescriber(api_key="...")
    text = describer.describe(image_bytes, "png", "diagram")

Factory::

    from pptx_md.providers import get_describer

    describer = get_describer("anthropic", api_key="...")
"""

from __future__ import annotations

from pptx_md.describer import ImageDescriber
from pptx_md.errors import InstallationError as InstallationError  # re-export

__all__ = [
    "AnthropicDescriber",
    "OpenAIDescriber",
    "get_describer",
]


def get_describer(name: str, **config: object) -> ImageDescriber:
    """Return a reference describer instance by provider *name*.

    Lazy-imports the concrete provider module so that unused providers'
    SDK absence does not affect users of other providers (ADR-212 partial
    availability).

    Args:
        name: Provider name — ``"anthropic"`` or ``"openai"``.
        **config: Keyword arguments forwarded to the describer constructor
            (e.g. ``api_key``, ``model``).

    Returns:
        An :class:`~pptx_md.describer.ImageDescriber`-compatible instance.

    Raises:
        ValueError: For an unknown provider name.
        InstallationError: If the required SDK is not installed.
    """
    if name == "anthropic":
        from pptx_md.providers.anthropic import AnthropicDescriber  # noqa: PLC0415

        return AnthropicDescriber(**config)  # type: ignore[arg-type]
    if name == "openai":
        from pptx_md.providers.openai import OpenAIDescriber  # noqa: PLC0415

        return OpenAIDescriber(**config)  # type: ignore[arg-type]
    raise ValueError(
        f"Unknown VLM provider {name!r}. Supported: 'anthropic', 'openai'."
    )


# Re-export concrete classes for convenience (lazy import to preserve NFR-08:
# top-level import of this package must not trigger SDK imports).
def __getattr__(name: str) -> object:
    if name == "AnthropicDescriber":
        from pptx_md.providers.anthropic import AnthropicDescriber  # noqa: PLC0415

        return AnthropicDescriber
    if name == "OpenAIDescriber":
        from pptx_md.providers.openai import OpenAIDescriber  # noqa: PLC0415

        return OpenAIDescriber
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
