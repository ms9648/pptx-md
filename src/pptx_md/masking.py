"""Personal information masking — FR-15 (#37).

Public interface:
    MASK_TOKEN: str                  -- replacement token ("[REDACTED]")
    MaskingOptions                   -- dataclass controlling masking behaviour
    mask_text(text, options) -> str  -- apply masking to a single string

Design constraints:
    - NFR-06: original (pre-mask) text MUST NOT appear in any log output.
              Only match-count metadata is logged.
    - Opt-in: masking is disabled by default (MaskingOptions.enabled=False).
    - Deterministic: same input -> same output (pure function, no side effects).
    - Extensible: callers supply additional re.Pattern objects via
      MaskingOptions.patterns.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

__all__ = ["MASK_TOKEN", "MaskingOptions", "mask_text"]

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Built-in default patterns
# ---------------------------------------------------------------------------

_DEFAULT_PATTERNS: list[re.Pattern[str]] = [
    # Email address (RFC 5321-ish)
    re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}"),
    # Korean phone number: 010-1234-5678 / 02.1234.5678 / 031 1234 5678 etc.
    re.compile(r"0\d{1,2}[-.\s]\d{3,4}[-.\s]\d{4}"),
]

MASK_TOKEN: str = "[REDACTED]"


# ---------------------------------------------------------------------------
# MaskingOptions dataclass
# ---------------------------------------------------------------------------


@dataclass
class MaskingOptions:
    """Options controlling personal-information masking (FR-15).

    Attributes:
        enabled:  When False (default), mask_text() is a no-op.
        patterns: List of compiled regex patterns to redact.
                  Defaults to the built-in email + phone patterns.
    """

    enabled: bool = False
    patterns: list[re.Pattern[str]] = field(
        default_factory=lambda: list(_DEFAULT_PATTERNS)
    )


# ---------------------------------------------------------------------------
# Public function
# ---------------------------------------------------------------------------


def mask_text(text: str, options: MaskingOptions) -> str:
    """Replace every PII match in *text* with MASK_TOKEN.

    Args:
        text:    Input string (may contain PII).
        options: MaskingOptions controlling whether masking is active and
                 which patterns to apply.

    Returns:
        The masked string if options.enabled is True; *text* unchanged otherwise.

    NFR-06 guarantee:
        No log record emitted by this function contains the original text or
        any matched PII fragment — only the total match count is logged at DEBUG.
    """
    if not options.enabled:
        return text

    result = text
    total_matches = 0

    for pattern in options.patterns:
        matches = pattern.findall(result)
        total_matches += len(matches)
        result = pattern.sub(MASK_TOKEN, result)

    # NFR-06: log metadata only — never the original text
    logger.debug("masking: %d match(es) replaced", total_matches)

    return result
