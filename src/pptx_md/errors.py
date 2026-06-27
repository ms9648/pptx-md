"""Library-wide exception hierarchy (introduced in M2; reused by M4/M5)."""


class PptxMdError(Exception):
    """Base class for all pptx-md errors."""


class ParseError(PptxMdError):
    """Raised when a PPTX file cannot be opened/parsed (file-level failure)."""
