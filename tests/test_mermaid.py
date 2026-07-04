"""Tests for src/pptx_md/mermaid.py diagram flowchart extraction (FR-27, issue #68).

Covers ``DIAGRAM_HINT_SUFFIX`` and ``render_diagram_mermaid`` (ADR-611/612,
ARCH-M12 §3.4, WBS W1). Existing table-serialisation symbols
(``is_complex_table`` / ``table_to_mermaid``) are exercised in
``tests/test_mermaid_fallback.py``; this file only adds coverage for the M12
diagram concern and confirms zero regression on those symbols.

Test naming: test_ac<N>_fr27_<desc> where <N> matches issue #68's AC numbering.
"""

from __future__ import annotations

import pathlib
import subprocess
import sys

from pptx_md.mermaid import DIAGRAM_HINT_SUFFIX, render_diagram_mermaid

_PROJECT_ROOT = pathlib.Path(__file__).parent.parent


# ---------------------------------------------------------------------------
# issue AC4 — valid mermaid flowchart fence -> normalised string returned
# ---------------------------------------------------------------------------


def test_ac4_fr27_유효한_flowchart_펜스_정규화_반환() -> None:
    """issue AC4: a well-formed ```mermaid flowchart fence is extracted."""
    text = "Sure, here it is:\n```mermaid\nflowchart TD\n  A --> B\n```\nThanks."

    result = render_diagram_mermaid(text)

    assert result is not None
    assert result.startswith("```mermaid")
    assert result.endswith("```")
    assert "flowchart TD" in result
    assert "A --> B" in result


def test_ac4_fr27_graph_키워드도_허용() -> None:
    """issue AC4: 'graph' keyword (Mermaid alias for flowchart) is also accepted."""
    text = "```mermaid\ngraph LR\n  X --> Y\n```"

    result = render_diagram_mermaid(text)

    assert result is not None
    assert "graph LR" in result


def test_ac4_fr27_앞뒤_공백_정규화() -> None:
    """issue AC4: surrounding blank lines inside the fence are trimmed."""
    text = "```mermaid\n\n\nflowchart TD\nA-->B\n\n\n```"

    result = render_diagram_mermaid(text)

    assert result is not None
    assert result == "```mermaid\nflowchart TD\nA-->B\n```"


# ---------------------------------------------------------------------------
# issue AC5 — 구조화 불가 응답 -> None (호출부가 원본 텍스트로 fallback)
# ---------------------------------------------------------------------------


def test_ac5_fr27_펜스_없음_none() -> None:
    """issue AC5: no ```mermaid fence present -> None."""
    text = "This is just a plain description with no code block at all."

    assert render_diagram_mermaid(text) is None


def test_ac5_fr27_키워드_불일치_none() -> None:
    """issue AC5: fence present but first line is not a flowchart keyword -> None."""
    text = "```mermaid\nsequenceDiagram\n  Alice->>Bob: Hello\n```"

    assert render_diagram_mermaid(text) is None


def test_ac5_fr27_빈_본문_none() -> None:
    """issue AC5: fence with only whitespace body -> None."""
    text = "```mermaid\n   \n\n```"

    assert render_diagram_mermaid(text) is None


def test_ac5_fr27_무예외_임의_입력() -> None:
    """issue AC5: render_diagram_mermaid never raises on arbitrary input."""
    candidates = [
        "",
        "```mermaid",
        "```mermaid\n```",
        "flowchart TD\nA-->B",
        "```",
    ]
    for candidate in candidates:
        try:
            render_diagram_mermaid(candidate)
        except Exception as exc:  # noqa: BLE001
            raise AssertionError(
                f"must never raise; input={candidate!r} exc={exc}"
            ) from exc


def test_fr27_결정적_동일_입력_동일_출력() -> None:
    """ADR-611: identical input -> identical output (determinism)."""
    text = "```mermaid\nflowchart TD\nA-->B\n```"
    assert render_diagram_mermaid(text) == render_diagram_mermaid(text)


# ---------------------------------------------------------------------------
# DIAGRAM_HINT_SUFFIX
# ---------------------------------------------------------------------------


def test_fr27_diagram_hint_suffix_mermaid_언급() -> None:
    """DIAGRAM_HINT_SUFFIX must be a non-empty str mentioning mermaid/flowchart."""
    assert isinstance(DIAGRAM_HINT_SUFFIX, str)
    assert DIAGRAM_HINT_SUFFIX
    assert "mermaid" in DIAGRAM_HINT_SUFFIX.lower()
    assert "flowchart" in DIAGRAM_HINT_SUFFIX.lower()


# ---------------------------------------------------------------------------
# Regression — existing table symbols unaffected (ADR-612)
# ---------------------------------------------------------------------------


def test_fr27_table_함수_회귀_없음() -> None:
    """ADR-612: table_to_mermaid / is_complex_table remain importable & unchanged."""
    from pptx_md.ir import ShapeKind, TableShapeIR
    from pptx_md.mermaid import is_complex_table, table_to_mermaid

    table = TableShapeIR(
        shape_id=1,
        name="t",
        kind=ShapeKind.TABLE,
        rows=[["a", "b"], ["1", "2"]],
        n_rows=2,
        n_cols=2,
    )
    assert is_complex_table(table) is False
    rendered = table_to_mermaid(table)
    assert rendered.startswith("```mermaid")
    assert "headers: a | b" in rendered


# ---------------------------------------------------------------------------
# NFR-08 — no VLM SDK import; only stdlib `re` added
# ---------------------------------------------------------------------------


def test_fr27_sdk_import_0건() -> None:
    """NFR-08: mermaid.py contains no 'anthropic' or 'openai' references."""
    source_file = _PROJECT_ROOT / "src" / "pptx_md" / "mermaid.py"
    source = source_file.read_text(encoding="utf-8")

    assert "anthropic" not in source
    assert "openai" not in source
    assert "import re" in source


def test_fr27_mypy_strict_exit0() -> None:
    """NFR-03: mypy src/ exits 0 (mermaid.py included)."""
    result = subprocess.run(
        [sys.executable, "-m", "mypy", "src/"],
        capture_output=True,
        text=True,
        cwd=str(_PROJECT_ROOT),
    )
    assert result.returncode == 0, (
        f"mypy src/ failed (exit {result.returncode}).\n"
        f"stdout:\n{result.stdout}\nstderr:\n{result.stderr}"
    )
