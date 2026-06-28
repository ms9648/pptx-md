"""ImageDescriber protocol — structural interface for VLM image-description providers.

FR-08: Defines the contract that any VLM provider must satisfy.  Uses
``typing.Protocol`` + ``@runtime_checkable`` so providers can be plug-in
without nominal inheritance (ADR-217, RFI-1 Q1).

SDK-agnostic: this module imports NO VLM SDK, so
``import pptx_md.describer`` always succeeds in core-only environments
(NFR-08).

Internal module (skill §1): NOT re-exported from package root until M6.
"""

from __future__ import annotations

from typing import Protocol, runtime_checkable

__all__ = ["ImageDescriber"]


@runtime_checkable
class ImageDescriber(Protocol):
    """Structural interface for VLM image-description providers (FR-08).

    Any object implementing ``describe(...)`` with this exact signature
    satisfies the protocol through structural typing — no nominal
    inheritance required.  This allows users to plug in external
    providers (e.g. Gemini) without importing or subclassing this class
    (RFI-1 Q1).

    SDK-agnostic: this module imports NO VLM SDK, so
    ``import pptx_md.describer`` always succeeds in core-only
    environments (NFR-08).

    Example::

        class MyDescriber:
            def describe(
                self,
                image_bytes: bytes,
                image_ext: str,
                shape_hint: str | None,
            ) -> str:
                return "an image"

        from pptx_md.describer import ImageDescriber
        obj = MyDescriber()
        assert isinstance(obj, ImageDescriber)  # True — no inheritance needed
    """

    def describe(
        self,
        image_bytes: bytes,
        image_ext: str,
        shape_hint: str | None,
    ) -> str:
        """Return a natural-language description of the image.

        Args:
            image_bytes: Raw image bytes (PNG, JPEG, …).  EMF/WMF images
                are converted to PNG upstream before reaching a describer.
            image_ext: File-extension hint for the image format, e.g.
                ``"png"``, ``"jpeg"``.  Used by providers to determine the
                MIME type for the API payload.
            shape_hint: Optional classification-derived hint string
                (FR-10); ``None`` means no hint — the provider falls back
                to a generic prompt.

        Returns:
            A non-empty natural-language description string on success.
            Returning an empty string is allowed but indicates the provider
            could not produce a useful description.

        Raises:
            DescribeError: On provider/API failure.  The provider MUST NOT
                swallow failures silently; the orchestrator (FR-10) is
                responsible for per-image isolation (ADR-214).
        """
        ...
