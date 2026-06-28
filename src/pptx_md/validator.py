"""Markdown structural validator — validate_markdown (FR-14, issue #36).

Public interface:
    ValidationResult  -- dataclass(valid: bool, warnings: list[str])
    validate_markdown(md: str) -> ValidationResult

Design constraints:
    - Non-destructive, non-raising: always returns ValidationResult (AC7).
    - Standard library only — no external Markdown parser (AC8).
    - Code-block interior lines are excluded from heading analysis (AC6).
    - valid=False triggers: empty document (AC2), unclosed code fence (AC3).
    - Heading anomalies produce warnings but do NOT affect valid (AC4, AC5).
"""

from __future__ import annotations

from dataclasses import dataclass, field

__all__ = ["ValidationResult", "validate_markdown"]

# ---------------------------------------------------------------------------
# Named constants
# ---------------------------------------------------------------------------
_FENCE_MARKER: str = "```"
_WARN_EMPTY: str = "빈 문서: 내용이 없습니다."
_WARN_UNCLOSED_FENCE: str = "닫히지 않은 코드블록: ``` 펜스 개수가 홀수입니다."
_WARN_HEADING_START: str = (
    "헤딩 시작 레벨 이상: 첫 헤딩이 h1(#) 또는 h2(##) 가 아닌"
    " 더 깊은 레벨로 시작합니다."
)


def _warn_heading_jump(prev_level: int, cur_level: int) -> str:
    return f"헤딩 레벨 점프: h{prev_level}에서 h{cur_level}로 2단계 이상 깊어졌습니다."


# ---------------------------------------------------------------------------
# Public dataclass
# ---------------------------------------------------------------------------


@dataclass
class ValidationResult:
    """Result of validate_markdown.

    Attributes:
        valid:    False when a fatal structural defect is detected
                  (empty document or unclosed code fence).
        warnings: Human-readable diagnostic messages.  May be non-empty
                  even when valid=True (heading-order anomalies).
    """

    valid: bool
    warnings: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Public function
# ---------------------------------------------------------------------------


def validate_markdown(md: str) -> ValidationResult:
    """Validate structural integrity of a Markdown string (FR-14).

    Checks performed (in order):
    1. Empty document: md.strip() == "" -> valid=False.
    2. Unclosed code fence: odd number of ``` occurrences -> valid=False.
    3. Heading order anomalies (code-block exterior only):
       - First heading is not h1/h2 -> warning.
       - Heading level skips 2+ levels deeper in a single step -> warning.
       valid is NOT affected by heading warnings.

    Args:
        md: Markdown string to validate.

    Returns:
        ValidationResult — never raises.
    """
    try:
        return _validate(md)
    except Exception:  # noqa: BLE001
        # AC7: unconditionally return ValidationResult regardless of input
        return ValidationResult(
            valid=False, warnings=["내부 검증 오류가 발생했습니다."]
        )


# ---------------------------------------------------------------------------
# Private implementation
# ---------------------------------------------------------------------------


def _validate(md: str) -> ValidationResult:
    warnings: list[str] = []
    valid = True

    # --- Check 1: empty document ---
    if md.strip() == "":
        warnings.append(_WARN_EMPTY)
        return ValidationResult(valid=False, warnings=warnings)

    # --- Check 2: unclosed code fence ---
    fence_count = md.count(_FENCE_MARKER)
    if fence_count % 2 != 0:
        warnings.append(_WARN_UNCLOSED_FENCE)
        valid = False

    # --- Check 3: heading order (code-block exterior only) ---
    heading_warnings = _check_heading_order(md)
    warnings.extend(heading_warnings)

    return ValidationResult(valid=valid, warnings=warnings)


def _check_heading_order(md: str) -> list[str]:
    """Extract ATX headings outside code fences and check order rules.

    Returns a list of warning strings (may be empty).
    """
    warnings: list[str] = []
    in_code_block = False
    prev_level: int | None = None
    first_heading_seen = False

    for line in md.splitlines():
        # Toggle code-block state on fence lines
        if line.startswith(_FENCE_MARKER):
            in_code_block = not in_code_block
            continue

        # Skip lines inside a code block (AC6)
        if in_code_block:
            continue

        # Detect ATX heading: line starts with 1–6 '#' followed by space or end
        heading_level = _atx_heading_level(line)
        if heading_level is None:
            continue

        if not first_heading_seen:
            first_heading_seen = True
            # AC4: first heading must be h1 or h2
            if heading_level > 2:
                warnings.append(_WARN_HEADING_START)
            prev_level = heading_level
        else:
            # AC5: level must not increase (deepen) by more than 1 step at once
            assert prev_level is not None
            if heading_level > prev_level + 1:
                warnings.append(_warn_heading_jump(prev_level, heading_level))
            prev_level = heading_level

    return warnings


def _atx_heading_level(line: str) -> int | None:
    """Return ATX heading depth (1-6) or None if the line is not a heading."""
    if not line.startswith("#"):
        return None
    # Count leading '#' characters
    level = 0
    for ch in line:
        if ch == "#":
            level += 1
        else:
            break
    if level > 6:
        return None
    # Must be followed by a space or be the entire line (bare '####')
    rest = line[level:]
    if rest == "" or rest.startswith(" "):
        return level
    return None
