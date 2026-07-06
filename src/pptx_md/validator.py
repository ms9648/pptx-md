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

FR-29 (issue #80, ARCH-Wave3 §3.7, ADR-619) extends the above with four
additional structural rules. These are executed unconditionally (D-5: no
new opt-in flag) and only ever append warnings — ``valid`` is never
affected by them:
    - _check_pipe_tables         -- broken / column-mismatched / all-blank
                                     GFM pipe tables (FR-29 AC1/AC2/AC5).
    - _check_duplicate_headings  -- adjacent identical headings (AC3).
    - _check_control_chars       -- residual U+000B (``\\v``) (AC4).
    - _check_empty_slides        -- slide blocks with no body content
                                     beyond heading/comment lines (AC6).
Both a leading YAML frontmatter block and code-fence interiors are
excluded from analysis to avoid false positives (§3.7 common
preprocessing).
"""

from __future__ import annotations

import re
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

# --- FR-29 신규 경고 상수 (ARCH-Wave3 §3.7) ---
_WARN_TABLE_NO_DATA: str = "깨진 표: 데이터 행이 없습니다."
_WARN_TABLE_COL_MISMATCH: str = "표 열 수 불일치: 데이터 행의 열 수가 헤더와 다릅니다."
_WARN_TABLE_ALL_BLANK: str = "전공백 표: 모든 셀이 비어 있습니다."
_WARN_DUP_HEADING: str = "중복 인접 헤딩: 동일 텍스트 헤딩이 연속됩니다."
_WARN_EMPTY_SLIDE: str = "빈 슬라이드: 헤딩 외 본문이 없습니다."

# Matches a standalone slide separator line (assembler _SLIDE_SEPARATOR body).
_SEPARATOR_RE = re.compile(r"^-{3,}$")
# Matches an unescaped pipe character (used to detect table-row candidates
# and to split a row into cells while ignoring `\|` escapes).
_UNESCAPED_PIPE_RE = re.compile(r"(?<!\\)\|")
# Matches a single GFM table separator cell, e.g. "---", ":---", "---:", ":-:".
_SEPARATOR_CELL_RE = re.compile(r"^:?-{1,}:?$")


def _warn_heading_jump(prev_level: int, cur_level: int) -> str:
    return f"헤딩 레벨 점프: h{prev_level}에서 h{cur_level}로 2단계 이상 깊어졌습니다."


def _warn_vertical_tab(count: int) -> str:
    return f"제어문자 \\v 잔존: {count}개."


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

    # --- FR-29: structural quality rules (always on, warning-only, D-5) ---
    warnings.extend(_check_control_chars(md))
    warnings.extend(_check_pipe_tables(md))
    warnings.extend(_check_duplicate_headings(md))
    warnings.extend(_check_empty_slides(md))

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


# ---------------------------------------------------------------------------
# FR-29: structural quality rules (issue #80, ARCH-Wave3 §3.7, ADR-619)
# ---------------------------------------------------------------------------


def _strip_frontmatter_lines(lines: list[str]) -> list[str]:
    """Drop a single leading YAML frontmatter block, if present (§3.3/§3.7).

    A frontmatter block is a standalone ``---`` line at (or near, modulo
    leading blank lines) the very start of the document, followed later by
    another standalone ``---`` line. Everything up to and including the
    closing ``---`` is skipped so downstream rules do not mistake it for a
    slide separator or a pipe-table row.
    """
    idx = 0
    while idx < len(lines) and lines[idx].strip() == "":
        idx += 1
    if idx >= len(lines) or lines[idx].strip() != "---":
        return lines
    for closing in range(idx + 1, len(lines)):
        if lines[closing].strip() == "---":
            return lines[closing + 1 :]
    return lines


def _check_control_chars(md: str) -> list[str]:
    """FR-29 AC4: warn when U+000B (``\\v``) remains in the document."""
    count = md.count("\v")
    if count > 0:
        return [_warn_vertical_tab(count)]
    return []


def _split_table_row(line: str) -> list[str]:
    """Split a pipe-table row into cells, ignoring escaped pipes (`\\|`).

    Optional leading/trailing pipe delimiters are dropped (GFM convention).
    """
    cells = _UNESCAPED_PIPE_RE.split(line.strip())
    if cells and cells[0] == "":
        cells = cells[1:]
    if cells and cells[-1] == "":
        cells = cells[:-1]
    return [c.replace("\\|", "|") for c in cells]


def _is_separator_row(line: str) -> bool:
    """True if `line` is a GFM table separator row (e.g. `|---|:--:|`)."""
    cells = _split_table_row(line)
    if not cells:
        return False
    return all(_SEPARATOR_CELL_RE.match(c.strip()) for c in cells)


def _check_pipe_tables(md: str) -> list[str]:
    """FR-29 AC1/AC2/AC5: broken / column-mismatched / all-blank pipe tables.

    Only pipe-row runs outside code fences and outside a leading
    frontmatter block are considered. A run is confirmed as a GFM table
    only when its second line is a valid separator row (header + spec);
    plain `|` occurrences elsewhere (including inside fenced code) never
    trigger a false positive (AC1).
    """
    warnings: list[str] = []
    lines = _strip_frontmatter_lines(md.splitlines())
    in_code_block = False
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        if line.startswith(_FENCE_MARKER):
            in_code_block = not in_code_block
            i += 1
            continue
        if in_code_block or not _UNESCAPED_PIPE_RE.search(line):
            i += 1
            continue

        # Collect a run of consecutive pipe-containing lines (table candidate).
        run: list[str] = []
        while (
            i < n
            and not lines[i].startswith(_FENCE_MARKER)
            and _UNESCAPED_PIPE_RE.search(lines[i])
        ):
            run.append(lines[i])
            i += 1

        if len(run) < 2 or not _is_separator_row(run[1]):
            continue  # not a GFM table (no valid separator row)

        header_cells = _split_table_row(run[0])
        data_rows = run[2:]
        if not data_rows:
            warnings.append(_WARN_TABLE_NO_DATA)
            continue

        col_mismatch = False
        all_blank = True
        for row in data_rows:
            row_cells = _split_table_row(row)
            if len(row_cells) != len(header_cells):
                col_mismatch = True
            if any(c.strip() for c in row_cells):
                all_blank = False
        if col_mismatch:
            warnings.append(_WARN_TABLE_COL_MISMATCH)
        if all_blank:
            warnings.append(_WARN_TABLE_ALL_BLANK)

    return warnings


def _check_duplicate_headings(md: str) -> list[str]:
    """FR-29 AC3: warn on adjacent identical-text ATX headings.

    "Adjacent" (D-4/§3.7) means no non-blank, non-heading line (body text,
    a `---` separator, an HTML comment, etc.) occurs between the two
    headings. Blank lines alone do not break adjacency.
    """
    warnings: list[str] = []
    lines = _strip_frontmatter_lines(md.splitlines())
    in_code_block = False
    prev_heading_text: str | None = None
    intervening_content = False

    for line in lines:
        if line.startswith(_FENCE_MARKER):
            in_code_block = not in_code_block
            continue
        if in_code_block:
            continue

        level = _atx_heading_level(line)
        if level is not None:
            text = line[level:].strip()
            if (
                prev_heading_text is not None
                and not intervening_content
                and text == prev_heading_text
            ):
                warnings.append(_WARN_DUP_HEADING)
            prev_heading_text = text
            intervening_content = False
        elif line.strip() != "":
            intervening_content = True

    return warnings


def _check_empty_slides(md: str) -> list[str]:
    """FR-29 AC6: warn on slide blocks with no body beyond heading/comment.

    Blocks are delimited by standalone ``---`` separator lines outside
    code fences (D-4 판정 기준, FR-28 AC4 와 공통). A block is "empty" when,
    after removing ATX heading lines, HTML comment lines (slide metadata),
    and blank lines, no non-blank content remains.
    """
    lines = _strip_frontmatter_lines(md.splitlines())

    blocks: list[list[str]] = [[]]
    in_code_block = False
    for line in lines:
        if not in_code_block and _SEPARATOR_RE.match(line.strip()):
            blocks.append([])
            continue
        if line.startswith(_FENCE_MARKER):
            in_code_block = not in_code_block
        blocks[-1].append(line)

    warnings: list[str] = []
    for block in blocks:
        remaining_non_blank = False
        block_in_code = False
        for line in block:
            if line.startswith(_FENCE_MARKER):
                block_in_code = not block_in_code
                continue
            if block_in_code:
                if line.strip() != "":
                    remaining_non_blank = True
                continue
            stripped = line.strip()
            if stripped == "":
                continue
            if _atx_heading_level(line) is not None:
                continue
            if stripped.startswith("<!--") and stripped.endswith("-->"):
                continue
            remaining_non_blank = True
        if not remaining_non_blank:
            warnings.append(_WARN_EMPTY_SLIDE)

    return warnings
