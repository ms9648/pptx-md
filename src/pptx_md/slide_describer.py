"""SlideReconstructor protocol — structural interface for VLM slide reconstructors.

FR-32 (ADR-625/626, DEC-8): Defines the contract that any VLM slide-reconstruction
provider must satisfy. This is a **new** protocol distinct from
``pptx_md.describer.ImageDescriber`` — the input/output units differ:

- ``ImageDescriber.describe(image_bytes, image_ext, shape_hint) -> str``
  takes a single embedded image + a hint string and returns a natural-language
  description.
- ``SlideReconstructor.reconstruct(image_bytes, image_ext, context) -> str``
  takes a whole rendered slide PNG + slide context (title/text outline) and
  returns a structured Markdown fragment (tables/nested lists/flow).

A single provider class MAY implement both protocols (shared HTTP client),
but the interfaces themselves stay separate (ARCH-v020 §3.3, §5.3 ADR-625).

Uses ``typing.Protocol`` + ``@runtime_checkable`` so providers can be
plug-in without nominal inheritance (describer.py / ADR-217 precedent).

SDK-agnostic: this module imports NO VLM SDK and NO ``python-pptx``/``Pillow``,
so ``import pptx_md.slide_describer`` always succeeds in core-only
environments (INV-5, NFR-08 precedent).

Internal module (skill §1): NOT re-exported from package root until M6.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol, runtime_checkable

__all__ = ["SlideContext", "SlideReconstructor"]


@dataclass(frozen=True)
class SlideContext:
    """Immutable per-slide context passed to a slide-reconstruction provider.

    Carries just enough deterministic, IR-derived information for the
    provider to ground its reconstruction — no VLM/rendering state.

    Attributes:
        slide_index: 0-based slide index (traceability; matches
            ``SlideIR`` ordering, ADR-220).
        title: Slide title text. Empty string (``""``) is allowed when the
            slide has no title shape.
        text_outline: The deterministic text-path rendered body for this
            slide (i.e. what the existing text extraction would produce).
            Used by providers to ground/verify the reconstruction without
            depending on OCR. Empty string (``""``) is allowed.
    """

    slide_index: int
    title: str
    text_outline: str


@runtime_checkable
class SlideReconstructor(Protocol):
    """Structural interface for VLM slide-reconstruction providers (FR-32).

    Any object implementing ``reconstruct(...)`` with this exact signature
    satisfies the protocol through structural typing — no nominal
    inheritance required (describer.py / RFI-1 Q1 precedent). This allows
    users to plug in external providers without importing or subclassing
    this class.

    Distinct from ``ImageDescriber`` (FR-08): input is a whole rendered
    slide image + ``SlideContext`` (not a single embedded image + hint
    string), and output is a structured Markdown fragment (not a
    natural-language description) (ARCH-v020 §3.3, DEC-8).

    SDK-agnostic: this module imports NO VLM SDK, so
    ``import pptx_md.slide_describer`` always succeeds in core-only
    environments (INV-5).

    Example::

        class MyReconstructor:
            def reconstruct(
                self,
                image_bytes: bytes,
                image_ext: str,
                context: SlideContext,
            ) -> str:
                return f"## {context.title}\\n\\n..."

        from pptx_md.slide_describer import SlideReconstructor
        obj = MyReconstructor()
        assert isinstance(obj, SlideReconstructor)  # True — no inheritance needed
    """

    def reconstruct(
        self,
        image_bytes: bytes,
        image_ext: str,
        context: SlideContext,
    ) -> str:
        """Return a structured Markdown fragment reconstructed from a slide image.

        Args:
            image_bytes: Raw rendered slide PNG bytes (whole-slide render,
                not a single embedded shape).
            image_ext: File-extension hint for the image format, e.g.
                ``"png"``. Used by providers to determine the MIME type
                for the API payload.
            context: Deterministic slide context (title, text outline,
                slide index) grounding the reconstruction.

        Returns:
            A Markdown fragment: pure Markdown body (no code fences, no
            explanatory prose) with pipe tables (header + separator + at
            least one data row), indentation-based nested lists, and
            ordered flow where applicable (FR-32 AC1/AC2/AC4). If
            structuring is uncertain, providers preserve the original text
            verbatim rather than emit a broken structure (FR-32 AC6).

        Raises:
            DescribeError: On provider/API failure, timeout, or empty
                response. The provider MUST NOT swallow failures silently;
                the orchestrator (FR-33) is responsible for per-slide
                isolation and text-path fallback (INV-4, FR-32 AC5).
        """
        ...
