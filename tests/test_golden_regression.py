"""Golden regression gate for FR-28 structured output (FR-29 AC8-AC10, issue #81, W-E).

Covers the dispatch AC list for issue #81 (mapped 1:1 onto REQ-wave3 §3
FR-29 AC8-AC10, ARCH-Wave3 §3.8, ADR-620):
    AC8  -> 2 golden fixtures baked, path + ConvertOptions combo documented.
    AC9  -> byte-for-byte regression comparison (hard-fail + unified diff on
            mismatch) + determinism (INV-1: same options -> same bytes twice).
    AC10 -> DoD warning categories (duplicate heading / \\v / all-blank
            table / broken table) == 0 on both goldens via validate_markdown.
    AC5 (issue-level, not REQ) -> image-description-coverage 100% is
            VLM/FR-27 scope; this no-VLM golden gate explicitly marks it N/A.
    AC6  -> PPTXMD_UPDATE_GOLDEN=1 env-guard: only rewrites when set, so an
            accidental `pytest` run never silently drifts the goldens.

Golden source (ADR-620): a hand-assembled, deterministic `PresentationIR`
built by `_build_presentation()` below -- NOT a real .pptx file / parser
run. Classification heuristics and python-pptx's own non-determinism are
out of scope for this gate (they are covered by the classifier/parser unit
tests); FR-28/FR-29 are assembler/validator concerns, so this test drives
`assemble_document` directly. The fixture intentionally includes, in one
deterministic IR (AC1 of the issue dispatch):
    - a section-cover slide with a real title and an EMPTY body
      (`slides[0]`, "표지") -- exercises the emit_toc "kept, not dropped"
      path (ARCH-Wave3 §3.2) and the FR-29 AC6 empty-slide validator rule.
    - a slide with a '>'-decomposable hierarchical title "A > B > C"
      (`slides[1]`) containing a text paragraph, a well-formed 2x2 GFM
      table (no column-mismatch / all-blank rows, so it never trips the
      FR-29 AC1/AC2/AC5 table warnings), an `ImageShapeIR` classified as
      `ImageClass.DIAGRAM` (exercises `derive_has_diagram`, FR-28 AC6),
      and non-empty `SlideIR.notes` (exercises `include_notes`).
    - a closing slide with an ordinary title and body (`slides[2]`,
      "마무리") for variety / determinism coverage.

Golden fixture paths + exact `ConvertOptions`/`assemble_document` keyword
combination (AC8, D-7):
    tests/golden/wave3_off.md
        = assemble_document(pres)   # all 4 FR-28 options at their False
                                     # default -- pre-Wave-3 equivalent,
                                     # backward-compatibility regression
                                     # target (see `_OFF_OPTIONS` below).
    tests/golden/wave3_on.md
        = assemble_document(
              pres,
              heading_hierarchy=True,
              emit_toc=True,
              emit_frontmatter=True,
              include_notes=True,
          )                          # all 4 options True -- structured-
                                      # quality-metrics target (see
                                      # `_ON_OPTIONS` below).

Regression judgment (AC9, D-6): both goldens are re-rendered from the same
`_build_presentation()` IR on every `pytest` run (existing CI `pytest`
step, ci.yml:37-38 -- no new workflow file) and compared **byte-for-byte**
against the baked `.md` files. Any diff raises `AssertionError` with a
`difflib.unified_diff` body, hard-failing the existing pytest step and
blocking the merge (branch protection, profile §5). No diff -> pass.
Determinism (INV-1) is asserted separately: rendering the same options
twice from two independently-built `PresentationIR` instances must yield
byte-identical output.

DoD warning-category gate (AC10): `validate_markdown` is run on both
rendered goldens; warnings are filtered against `_DOD_WARNING_SUBSTRINGS`
(substring match against the *exact* message text emitted by
`pptx_md.validator`, verified unique per category -- see the comment
above that dict) and the filtered count must be 0 for both goldens. The
"빈 슬라이드" (empty-slide) warning category is **excluded** from this DoD
set per REQ-wave3 §3 AC10 / ARCH-Wave3 §3.2: the retained section-cover
block ("표지") is intentionally still `body_empty` under `emit_toc`, so it
legitimately trips the FR-29 AC6 empty-slide rule -- that is expected, not
a regression (see `TestAc10DodWarningCategoriesZero
.test_ac10_empty_slide_warning_present_but_excluded_from_dod_set`).

Image-description-coverage 100% (issue AC5 / REQ-wave3 §6 Out of Scope):
this gate is a **no-VLM** golden (the DIAGRAM image above carries no
`description`, only a `classification`) -- description-coverage DoD is
VLM/FR-27's own gate, **N/A** here by design (ADR-620 / §3.8).
(이미지 설명 커버리지 100% DoD 는 VLM/FR-27 소관이며, 본 no-VLM 골든
게이트에서는 **N/A** 이다.)

Golden update procedure (AC6, D-6): the two `.md` files under
`tests/golden/` are only rewritten when the `PPTXMD_UPDATE_GOLDEN`
environment variable is set to the literal string ``"1"`` (see
`_compare_or_update` below); an ordinary `pytest` invocation only
compares and never mutates them (accidental-drift prevention). To
intentionally update the goldens after a reviewed, intended
assembler/validator change:

    PowerShell:
        $env:PPTXMD_UPDATE_GOLDEN=1; pytest tests/test_golden_regression.py
        Remove-Item Env:/PPTXMD_UPDATE_GOLDEN

    POSIX shell:
        PPTXMD_UPDATE_GOLDEN=1 pytest tests/test_golden_regression.py

Regenerating the goldens is **not** itself sufficient sign-off -- the
change must go through a dedicated golden-update PR with explicit human
approval (D-6); this test file and `ci.yml` are not modified by that
procedure.

AC name format: ac<N>_<short_description>.
"""

from __future__ import annotations

import difflib
import os
from pathlib import Path

import pytest

from pptx_md.assembler import assemble_document
from pptx_md.ir import (
    ImageClass,
    ImageShapeIR,
    ParagraphIR,
    PresentationIR,
    ShapeKind,
    SlideIR,
    TableShapeIR,
    TextShapeIR,
)
from pptx_md.validator import validate_markdown

# ---------------------------------------------------------------------------
# Golden fixture paths (AC8)
# ---------------------------------------------------------------------------
_GOLDEN_DIR: Path = Path(__file__).parent / "golden"
_OFF_GOLDEN_PATH: Path = _GOLDEN_DIR / "wave3_off.md"
_ON_GOLDEN_PATH: Path = _GOLDEN_DIR / "wave3_on.md"

# Exact ConvertOptions/assemble_document keyword combinations (AC8, D-7).
_OFF_OPTIONS: dict[str, bool] = {}
_ON_OPTIONS: dict[str, bool] = {
    "heading_hierarchy": True,
    "emit_toc": True,
    "emit_frontmatter": True,
    "include_notes": True,
}

# DoD warning categories (REQ-wave3 §3 AC10, ARCH-Wave3 §3.8) matched by a
# substring that is verified unique to that category's exact validator
# message (pptx_md/validator.py):
#   _WARN_DUP_HEADING       = "중복 인접 헤딩: 동일 텍스트 헤딩이 연속됩니다."
#   _warn_vertical_tab(...) = "제어문자 \\v 잔존: {count}개."
#   _WARN_TABLE_ALL_BLANK   = "전공백 표: 모든 셀이 비어 있습니다."
#   _WARN_TABLE_NO_DATA     = "깨진 표: 데이터 행이 없습니다."
# "빈 슬라이드" (_WARN_EMPTY_SLIDE) is deliberately absent from this dict
# (AC10 명시 제외).
_DOD_WARNING_SUBSTRINGS: dict[str, str] = {
    "중복_헤딩": "중복",
    "수직탭": "\\v",
    "전공백_표": "전공백 표",
    "깨진_표": "깨진 표",
}


def _dod_warnings(warnings: list[str]) -> list[str]:
    """Filter *warnings* down to the AC10 DoD category set (빈 슬라이드 제외)."""
    substrings = tuple(_DOD_WARNING_SUBSTRINGS.values())
    return [w for w in warnings if any(s in w for s in substrings)]


# ---------------------------------------------------------------------------
# Deterministic hand-built PresentationIR (ADR-620, issue AC1)
# ---------------------------------------------------------------------------


def _build_presentation() -> PresentationIR:
    """Build a fresh, deterministic golden `PresentationIR` (ADR-620).

    Called anew by every test that needs one so that determinism checks
    (AC9/INV-1) exercise two independently-constructed object graphs
    rather than a single shared/mutated instance.

    Returns:
        A 3-slide `PresentationIR`: a section-cover slide with an empty
        body, a slide with a hierarchical '>' title / normal table /
        DIAGRAM image / notes, and an ordinary closing slide.
    """
    cover = SlideIR(
        index=0,
        title="표지",
        shapes=[],
    )
    body_slide = SlideIR(
        index=1,
        title="A > B > C",
        notes="발표자 노트: 핵심 논지를 강조할 것",
        shapes=[
            TextShapeIR(
                shape_id=101,
                name="본문 텍스트",
                kind=ShapeKind.TEXT,
                paragraphs=[ParagraphIR(text="핵심 내용 설명입니다.", level=0)],
            ),
            TableShapeIR(
                shape_id=102,
                name="표",
                kind=ShapeKind.TABLE,
                rows=[["항목", "값"], ["속도", "빠름"]],
                n_rows=2,
                n_cols=2,
            ),
            ImageShapeIR(
                shape_id=103,
                name="다이어그램",
                kind=ShapeKind.IMAGE,
                classification=ImageClass.DIAGRAM,
            ),
        ],
    )
    closing = SlideIR(
        index=2,
        title="마무리",
        shapes=[
            TextShapeIR(
                shape_id=201,
                name="마무리 텍스트",
                kind=ShapeKind.TEXT,
                paragraphs=[ParagraphIR(text="감사합니다.", level=0)],
            ),
        ],
    )
    return PresentationIR(
        source_path="golden-fixture.pptx",
        slides=[cover, body_slide, closing],
    )


# ---------------------------------------------------------------------------
# AC6: env-guarded compare-or-update helper (D-6 — accidental-drift guard)
# ---------------------------------------------------------------------------


def _compare_or_update(path: Path, content: str) -> None:
    """Byte-compare *content* against the golden file at *path*, or rewrite it.

    AC6/D-6: rewriting only happens when the ``PPTXMD_UPDATE_GOLDEN``
    environment variable is set to the literal string ``"1"`` — an
    ordinary `pytest` run only compares (accidental-regeneration guard).
    On mismatch, raises `AssertionError` with a `difflib.unified_diff`
    body so the existing CI `pytest` step hard-fails with an actionable
    message (AC9).

    Args:
        path: Golden `.md` file path.
        content: Freshly-rendered Markdown to compare/write.

    Raises:
        AssertionError: When not in update mode and *content* differs
            from the file at *path* (byte-for-byte).
    """
    content_bytes = content.encode("utf-8")
    if os.environ.get("PPTXMD_UPDATE_GOLDEN") == "1":
        path.write_bytes(content_bytes)
        return

    expected_bytes = path.read_bytes()
    if expected_bytes != content_bytes:
        expected_text = expected_bytes.decode("utf-8")
        diff = "\n".join(
            difflib.unified_diff(
                expected_text.splitlines(),
                content.splitlines(),
                fromfile=str(path),
                tofile="<rendered>",
                lineterm="",
            )
        )
        raise AssertionError(
            f"골든 회귀 실패(AC9): {path} 와 재산출 결과가 다릅니다(hard-fail).\n"
            f"의도된 변경이면 PPTXMD_UPDATE_GOLDEN=1 로 재기록 후 별도 골든 갱신 "
            f"PR 로 사람 승인을 받으세요(D-6).\n{diff}"
        )


# ---------------------------------------------------------------------------
# AC8: golden fixtures — paths + ConvertOptions combo documented
# ---------------------------------------------------------------------------


class TestAc8GoldenFixturesPathsAndOptions:
    """골든 2종 박제 + 경로·옵션 조합 명시 (AC8, D-7)."""

    def test_ac8_off_golden_path_exists(self) -> None:
        """ac8_off_경로존재: tests/golden/wave3_off.md 가 박제되어 있다."""
        assert _OFF_GOLDEN_PATH.exists()

    def test_ac8_on_golden_path_exists(self) -> None:
        """ac8_on_경로존재: tests/golden/wave3_on.md 가 박제되어 있다."""
        assert _ON_GOLDEN_PATH.exists()

    def test_ac8_off_options_are_all_default_false(self) -> None:
        """ac8_off_옵션조합: off 골든은 FR-28 4옵션 전부 default(False)로 산출된다."""
        assert _OFF_OPTIONS == {}
        # assemble_document(pres) with no keyword overrides == all 4
        # heading_hierarchy/emit_toc/emit_frontmatter/include_notes at
        # their False default (현행 동등, 하위호환 회귀용).

    def test_ac8_on_options_are_all_true(self) -> None:
        """ac8_on_옵션조합: on 골든은 FR-28 4옵션 전부 True 로 산출된다."""
        assert _ON_OPTIONS == {
            "heading_hierarchy": True,
            "emit_toc": True,
            "emit_frontmatter": True,
            "include_notes": True,
        }


# ---------------------------------------------------------------------------
# AC9: byte-for-byte regression + determinism (INV-1)
# ---------------------------------------------------------------------------


class TestAc9ByteRegressionAndDeterminism:
    """골든 대비 바이트 회귀(hard-fail) + 결정성 (AC9, D-6, INV-1)."""

    def test_ac9_off_golden_byte_identical(self) -> None:
        """ac9_off_바이트동일: 재산출 결과가 off 골든과 바이트 동일하다."""
        rendered = assemble_document(_build_presentation(), **_OFF_OPTIONS)
        _compare_or_update(_OFF_GOLDEN_PATH, rendered)

    def test_ac9_on_golden_byte_identical(self) -> None:
        """ac9_on_바이트동일: 재산출 결과가 on 골든과 바이트 동일하다."""
        rendered = assemble_document(_build_presentation(), **_ON_OPTIONS)
        _compare_or_update(_ON_GOLDEN_PATH, rendered)

    def test_ac9_determinism_off_two_renders_identical(self) -> None:
        """ac9_off_결정성: 동일 off 옵션 2회 산출이 상호 바이트 동일하다(INV-1)."""
        first = assemble_document(_build_presentation(), **_OFF_OPTIONS)
        second = assemble_document(_build_presentation(), **_OFF_OPTIONS)
        assert first == second

    def test_ac9_determinism_on_two_renders_identical(self) -> None:
        """ac9_on_결정성: 동일 on 옵션 2회 산출이 상호 바이트 동일하다(INV-1)."""
        first = assemble_document(_build_presentation(), **_ON_OPTIONS)
        second = assemble_document(_build_presentation(), **_ON_OPTIONS)
        assert first == second


# ---------------------------------------------------------------------------
# AC10: DoD warning categories == 0 on both goldens
# ---------------------------------------------------------------------------


class TestAc10DodWarningCategoriesZero:
    """DoD 지표(중복 헤딩0/`\\v`0/전공백·깨진 표0) 골든 2종 경고 0 검증 (AC10)."""

    def test_ac10_off_golden_dod_warnings_zero(self) -> None:
        """ac10_off_DoD경고0: off 골든에서 4개 DoD 카테고리 경고가 0건이다."""
        rendered = assemble_document(_build_presentation(), **_OFF_OPTIONS)
        result = validate_markdown(rendered)
        assert _dod_warnings(result.warnings) == []

    def test_ac10_on_golden_dod_warnings_zero(self) -> None:
        """ac10_on_DoD경고0: on 골든에서 4개 DoD 카테고리 경고가 0건이다."""
        rendered = assemble_document(_build_presentation(), **_ON_OPTIONS)
        result = validate_markdown(rendered)
        assert _dod_warnings(result.warnings) == []

    def test_ac10_empty_slide_warning_present_but_excluded_from_dod_set(self) -> None:
        """ac10_빈슬라이드_집합제외: 유지된 섹션표지는 빈 슬라이드 경고 대상이나
        AC10 DoD 집합에는 포함되지 않는다(REQ AC10 명시, §3.2/§3.8)."""
        rendered = assemble_document(_build_presentation(), **_ON_OPTIONS)
        result = validate_markdown(rendered)
        assert any("빈 슬라이드" in w for w in result.warnings)
        assert _dod_warnings(result.warnings) == []


# ---------------------------------------------------------------------------
# AC5 (issue-level): image-description-coverage 100% is N/A for this gate
# ---------------------------------------------------------------------------


class TestAc5ImageDescriptionCoverageMarkedNA:
    """이미지 설명 커버리지 100% DoD 는 VLM/FR-27 소관 -> no-VLM 골든 N/A (AC5)."""

    def test_ac5_module_docstring_marks_image_coverage_na(self) -> None:
        """ac5_문서화: 모듈 docstring 이 이미지 설명 커버리지 N/A 를 명시한다."""
        doc = __doc__ or ""
        assert "N/A" in doc
        assert "이미지" in doc and "커버리지" in doc

    def test_ac5_golden_diagram_image_has_no_description(self) -> None:
        """ac5_no_VLM_고정: 골든의 DIAGRAM 이미지는 description 이 없다(no-VLM)."""
        pres = _build_presentation()
        diagram_images = [
            shape
            for slide in pres.slides
            for shape in slide.shapes
            if isinstance(shape, ImageShapeIR)
            and shape.classification is ImageClass.DIAGRAM
        ]
        assert len(diagram_images) == 1
        assert diagram_images[0].description is None


# ---------------------------------------------------------------------------
# AC6: PPTXMD_UPDATE_GOLDEN env-guard — isolated tmp_path, real goldens untouched
# ---------------------------------------------------------------------------


class TestAc6UpdateGoldenEnvGuard:
    """PPTXMD_UPDATE_GOLDEN 환경변수 가드 — 우발 갱신 방지 (AC6, D-6).

    Uses an isolated `tmp_path` fixture file, never the real
    `tests/golden/*.md` fixtures, so this test suite cannot itself cause
    golden drift.
    """

    def test_ac6_without_env_var_mismatch_raises_with_diff(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """ac6_미설정_비교만: env var 미설정 시 불일치는 diff 포함 실패만 하고
        파일은 변경되지 않는다(우발 갱신 방지)."""
        monkeypatch.delenv("PPTXMD_UPDATE_GOLDEN", raising=False)
        golden = tmp_path / "sample.md"
        golden.write_text("old content", encoding="utf-8")

        with pytest.raises(AssertionError, match="골든 회귀 실패"):
            _compare_or_update(golden, "new content")

        assert golden.read_text(encoding="utf-8") == "old content"

    def test_ac6_with_env_var_rewrites_file(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """ac6_설정시_재기록: PPTXMD_UPDATE_GOLDEN=1 이면 골든이 재기록된다."""
        monkeypatch.setenv("PPTXMD_UPDATE_GOLDEN", "1")
        golden = tmp_path / "sample.md"
        golden.write_text("old content", encoding="utf-8")

        _compare_or_update(golden, "new content")

        assert golden.read_bytes() == b"new content"

    def test_ac6_matching_content_passes_without_env_var(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """ac6_일치시_통과: env var 미설정 상태에서 내용이 같으면 예외 없이 통과한다."""
        monkeypatch.delenv("PPTXMD_UPDATE_GOLDEN", raising=False)
        golden = tmp_path / "sample.md"
        golden.write_text("same content", encoding="utf-8")

        _compare_or_update(golden, "same content")  # no raise
