"""Library-wide exception hierarchy (introduced in M2; reused by M4/M5)."""


class PptxMdError(Exception):
    """Base class for all pptx-md errors."""


class ParseError(PptxMdError):
    """Raised when a PPTX file cannot be opened/parsed (file-level failure)."""


class InstallationError(PptxMdError):
    """Raised when an optional VLM SDK is required but not installed (FR-09).

    Message guides the user to: pip install pptx-md[vlm] (스킬 §4).
    """


class DescribeError(PptxMdError):
    """Raised when a VLM provider fails to describe an image (FR-09).

    Wraps the underlying SDK/HTTP error; the orchestrator isolates it
    (ADR-214).
    """
