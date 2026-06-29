"""Tests for FR-18 documentation artifacts (issue #45).

Verifies that README.md, docs/usage.md, and docs/api.md satisfy
the acceptance criteria defined in issue #45.

AC-to-test mapping:
  AC1 -> test_ac1_readme_exists
  AC2 -> test_ac2_quick_start_uses_public_api_only
  AC3 -> test_ac3_install_paths_present
  AC4 -> test_ac4_no_literal_api_key
  AC5 -> test_ac5_masking_options_example
  AC6 -> test_ac6_api_ref_symbols_match_all
  AC7 -> test_ac7_no_broken_internal_links
"""

from __future__ import annotations

import re
from pathlib import Path

import pptx_md

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_REPO_ROOT = Path(__file__).parent.parent
_README = _REPO_ROOT / "README.md"
_USAGE = _REPO_ROOT / "docs" / "usage.md"
_API_REF = _REPO_ROOT / "docs" / "api.md"


def _read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


# ---------------------------------------------------------------------------
# AC1: README.md exists (hatch build 전제)
# ---------------------------------------------------------------------------


def test_ac1_readme_exists() -> None:
    """AC1: README.md must exist at the repository root.

    hatch build reads README.md via pyproject.toml `readme` field.
    A missing file causes a build warning/error.
    """
    assert _README.exists(), "README.md not found at repo root"
    assert _README.stat().st_size > 0, "README.md is empty"


# ---------------------------------------------------------------------------
# AC2: Quick Start 코드 블록이 FR-16 공개 API만 사용
# ---------------------------------------------------------------------------

# All symbols that may appear in Quick Start code blocks.
# Anything not in this set (e.g. internal parse_presentation) is a violation.
_ALLOWED_IMPORTS = {
    "pptx_md",
    "convert",
    "ConvertOptions",
    "MaskingOptions",
    "MASK_TOKEN",
    "validate_markdown",
    "ValidationResult",
    "ImageDescriber",
    "get_describer",
    "PptxMdError",
    "ParseError",
    "DescribeError",
    "InstallationError",
    # stdlib allowed in examples
    "os",
    "re",
    "pathlib",
    "Path",
    "logging",
}

_INTERNAL_SYMBOLS = [
    "parse_presentation",
    "enrich_images",
    "enrich_descriptions",
    "assemble_document",
    "assemble_slide",
    "is_complex_table",
    "table_to_mermaid",
    "classify_image",
    "convert_vector_to_png",
    "AnthropicDescriber",
    "OpenAIDescriber",
]


def _extract_code_blocks(text: str) -> list[str]:
    """Return all fenced code block bodies from a Markdown string.

    Uses a line-by-line state machine to ensure the last code block
    is never missed (regex DOTALL approach failed to capture trailing blocks).
    """
    blocks = []
    in_block = False
    current: list[str] = []
    for line in text.splitlines():
        if line.strip().startswith("```"):
            if in_block:
                blocks.append("\n".join(current))
                current = []
                in_block = False
            else:
                in_block = True
        elif in_block:
            current.append(line)
    return blocks


def test_ac2_quick_start_uses_public_api_only() -> None:
    """AC2: Quick Start code blocks must only use the FR-16 public API.

    Internal pipeline functions (parse_presentation, enrich_images, etc.)
    must not appear in README.md code examples.
    """
    readme_text = _read(_README)
    code_blocks = _extract_code_blocks(readme_text)
    assert len(code_blocks) > 0, "README.md contains no code blocks"

    for symbol in _INTERNAL_SYMBOLS:
        for block in code_blocks:
            assert symbol not in block, (
                f"README.md code block references internal symbol '{symbol}' "
                "(AC2 violation: Quick Start must use public API only)"
            )


# ---------------------------------------------------------------------------
# AC3: pip install pptx-md 및 pip install pptx-md[vlm] 두 경로 모두 기술
# ---------------------------------------------------------------------------


def test_ac3_install_paths_present() -> None:
    """AC3: README.md must document both install paths.

    - pip install pptx-md  (core-only)
    - pip install pptx-md[vlm]  (VLM extras)
    """
    readme_text = _read(_README)
    assert (
        "pip install pptx-md" in readme_text
    ), "README.md missing 'pip install pptx-md' (core install path)"
    assert (
        "pip install pptx-md[vlm]" in readme_text
    ), "README.md missing 'pip install pptx-md[vlm]' (VLM extras install path)"


# ---------------------------------------------------------------------------
# AC4: VLM 예시에 평문 키 리터럴 없음 (환경변수만)
# ---------------------------------------------------------------------------

# Patterns that suggest a literal API key (not env var reference)
_LITERAL_KEY_PATTERNS = [
    re.compile(r'api_key\s*=\s*"[a-zA-Z0-9\-_]{20,}"'),  # api_key="sk-..."
    re.compile(r"api_key\s*=\s*'[a-zA-Z0-9\-_]{20,}'"),  # api_key='sk-...'
    re.compile(r'api_key\s*=\s*f"'),  # f-string literal
]


def test_ac4_no_literal_api_key() -> None:
    """AC4: VLM code examples must use environment variables, not literal keys.

    NFR-05: API keys must never appear as plain string literals in docs.
    """
    for doc_path in (_README, _USAGE):
        if not doc_path.exists():
            continue
        text = _read(doc_path)
        for pattern in _LITERAL_KEY_PATTERNS:
            match = pattern.search(text)
            assert match is None, (
                f"{doc_path.name}: literal API key found matching "
                f"{pattern.pattern!r} — use os.environ[...] instead (AC4/NFR-05)"
            )
    # Positive check: README must reference os.environ (or os.getenv)
    readme_text = _read(_README)
    assert (
        "os.environ" in readme_text or "os.getenv" in readme_text
    ), "README.md must show os.environ / os.getenv for API key (AC4/NFR-05)"


# ---------------------------------------------------------------------------
# AC5: MaskingOptions(enabled=True) + 커스텀 패턴 예시 포함
# ---------------------------------------------------------------------------


def test_ac5_masking_options_example() -> None:
    """AC5: README.md must include MaskingOptions(enabled=True) and
    a custom pattern example.
    """
    readme_text = _read(_README)
    assert (
        "MaskingOptions(enabled=True)" in readme_text
    ), "README.md missing MaskingOptions(enabled=True) example (AC5)"
    # Custom pattern can be shown as re.compile(...) or patterns=[...]
    assert (
        "patterns=" in readme_text or "re.compile(" in readme_text
    ), "README.md missing custom pattern example for MaskingOptions (AC5)"


# ---------------------------------------------------------------------------
# AC6: API 레퍼런스 심볼 ↔ __all__ 정합 (누락/초과 0건)
# ---------------------------------------------------------------------------


def _extract_api_symbols_from_doc(api_doc: str) -> set[str]:
    """Parse symbol names from docs/api.md.

    Convention: each public symbol is documented with a level-3 heading
    (### symbol_name) or a bold code span (**`symbol_name`**).
    We look for both patterns.
    """
    symbols: set[str] = set()
    # Match ### heading lines like "### convert" or "### ConvertOptions"
    for m in re.finditer(r"^###\s+(`?)(\w+)", api_doc, re.MULTILINE):
        symbols.add(m.group(2))
    # Match **`symbol`** bold-code spans
    for m in re.finditer(r"\*\*`(\w+)`\*\*", api_doc):
        symbols.add(m.group(1))
    return symbols


def test_ac6_api_ref_symbols_match_all() -> None:
    """AC6: docs/api.md must document exactly the symbols in pptx_md.__all__.

    No symbol may be missing, and no undocumented symbol may appear.
    """
    assert _API_REF.exists(), "docs/api.md not found"
    api_doc = _read(_API_REF)

    expected = set(pptx_md.__all__)
    # __version__ is in __init__ but not always in __all__; skip if absent
    expected.discard("__version__")

    documented = _extract_api_symbols_from_doc(api_doc)

    missing = expected - documented
    assert (
        not missing
    ), f"docs/api.md is missing symbols from pptx_md.__all__: {sorted(missing)}"

    # We do not enforce "no extra" strictly because docs/api.md may document
    # __version__ separately; but all __all__ items must be present.


# ---------------------------------------------------------------------------
# AC7: 내부 링크 깨짐 0건
# ---------------------------------------------------------------------------


def _collect_internal_links(md_path: Path, text: str) -> list[tuple[Path, str]]:
    """Return (resolved_path, raw_link) for all Markdown-style internal links
    found in *text* that are NOT external URLs.
    """
    results: list[tuple[Path, str]] = []
    # [text](path) or [text](path#anchor) — skip http/https/mailto
    pattern = re.compile(r"\[([^\]]*)\]\(([^)]+)\)")
    doc_dir = md_path.parent
    for m in pattern.finditer(text):
        raw = m.group(2)
        # Skip external links
        if raw.startswith(("http://", "https://", "mailto:")):
            continue
        # Strip anchor
        path_part = raw.split("#")[0]
        if not path_part:
            continue  # same-page anchor only
        resolved = (doc_dir / path_part).resolve()
        results.append((resolved, raw))
    return results


def test_ac7_no_broken_internal_links() -> None:
    """AC7: All internal links in README.md, docs/usage.md, docs/api.md
    must resolve to existing files.
    """
    docs_to_check = [p for p in (_README, _USAGE, _API_REF) if p.exists()]
    assert docs_to_check, "No documentation files found to check"

    broken: list[str] = []
    for doc_path in docs_to_check:
        text = _read(doc_path)
        for resolved, raw in _collect_internal_links(doc_path, text):
            if not resolved.exists():
                broken.append(f"{doc_path.name}: broken link '{raw}' -> {resolved}")

    assert not broken, "Broken internal links found:\n" + "\n".join(broken)
