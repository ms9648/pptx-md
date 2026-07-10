"""Tests for FR-28 notes.py: speaker-notes blockquote rendering (issue #79, W-C).

Covers AC1-AC5 of issue #79 (include_notes).
AC name format: ac<N>_<short_description>
"""

from __future__ import annotations

import inspect

from pptx_md import notes as notes_module
from pptx_md.notes import render_notes_block

# ---------------------------------------------------------------------------
# AC1: notes != "" -> "> notes" header + each note line "> " prefixed
# ---------------------------------------------------------------------------


class TestAc1RenderNotesBlock:
    """render_notes_block — notes != '' 이면 '> notes' blockquote 반환 (AC1/AC8)."""

    def test_ac1_single_line_notes_renders_header_and_line(self) -> None:
        """ac1_단일줄_notes: 헤더 '> notes' + 노트 줄이 '> ' 프리픽스로 렌더된다."""
        result = render_notes_block("발표자 메모입니다")
        assert result == "> notes\n> 발표자 메모입니다"

    def test_ac1_multi_line_notes_prefixes_each_line(self) -> None:
        """ac1_여러줄_notes: 각 줄이 개별적으로 '> ' 프리픽스를 받는다."""
        result = render_notes_block("줄1\n줄2\n줄3")
        assert result == "> notes\n> 줄1\n> 줄2\n> 줄3"

    def test_ac1_notes_with_leading_trailing_whitespace_content(self) -> None:
        """ac1_공백_포함_notes: 비공백 내용이 있으면(주변 공백 포함) 블록이 생성된다."""
        result = render_notes_block("  실제 내용  ")
        assert result == "> notes\n>   실제 내용  "


# ---------------------------------------------------------------------------
# AC2: notes == "" (or whitespace-only) -> "" returned, no empty blockquote
# ---------------------------------------------------------------------------


class TestAc2EmptyNotes:
    """render_notes_block — 빈/공백 notes는 빈 문자열을 반환한다 (AC2/AC9)."""

    def test_ac2_empty_string_returns_empty_string(self) -> None:
        """ac2_빈문자열_notes: notes == '' 이면 '' 반환(블록 미생성)."""
        assert render_notes_block("") == ""

    def test_ac2_whitespace_only_returns_empty_string(self) -> None:
        """ac2_공백만_notes: 공백/개행만 있는 notes도 '' 반환."""
        assert render_notes_block("   \n\t  ") == ""

    def test_ac2_result_has_no_bare_blockquote_header(self) -> None:
        """ac2_빈_blockquote_미생성: 빈 notes 결과에 '> notes' 헤더가 없다."""
        result = render_notes_block("")
        assert "> notes" not in result


# ---------------------------------------------------------------------------
# AC3: \v (U+000B) -> \n normalization before line split (FR-24 parity)
# ---------------------------------------------------------------------------


class TestAc3VerticalTabNormalization:
    """render_notes_block — '\\v' 를 '\\n' 으로 정규화 후 줄 분해한다 (AC3)."""

    def test_ac3_single_vt_normalized_to_newline(self) -> None:
        """ac3_vt_단일_정규화: '\\v' 1개가 줄바꿈으로 처리되어 별도 줄이 된다."""
        result = render_notes_block("첫째줄\v둘째줄")
        assert result == "> notes\n> 첫째줄\n> 둘째줄"

    def test_ac3_multiple_vt_each_becomes_newline(self) -> None:
        """ac3_vt_여러개_정규화: 여러 '\\v' 가 각각 줄바꿈으로 정규화된다."""
        result = render_notes_block("A\vB\vC")
        assert result == "> notes\n> A\n> B\n> C"

    def test_ac3_vt_mixed_with_existing_newline(self) -> None:
        """ac3_vt_기존개행과_혼재: '\\v' 와 '\\n' 이 섞여도 모두 개행으로 처리된다."""
        result = render_notes_block("A\vB\nC")
        assert result == "> notes\n> A\n> B\n> C"


# ---------------------------------------------------------------------------
# AC4: deterministic, no exceptions, stdlib-only (no VLM SDK/python-pptx/Pillow)
# ---------------------------------------------------------------------------


class TestAc4DeterministicAndIsolated:
    """render_notes_block — 결정적·무예외·의존 격리 (AC4/INV-1/INV-5)."""

    def test_ac4_deterministic_repeated_calls_identical(self) -> None:
        """ac4_결정적: 동일 입력 2회 호출 결과가 바이트 동일하다."""
        notes = "반복 호출 검증\v두번째 줄"
        assert render_notes_block(notes) == render_notes_block(notes)

    def test_ac4_no_exception_on_edge_case_inputs(self) -> None:
        """ac4_무예외: 다양한 경계 입력에 대해 예외가 발생하지 않는다."""
        edge_cases = ["", " ", "\v", "\n", "\v\v\v", "a" * 10_000, "특수문자!@#$%^&*()"]
        for case in edge_cases:
            render_notes_block(case)  # no exception raised

    def test_ac4_no_forbidden_imports_in_module_source(self) -> None:
        """ac4_의존_격리: notes.py 소스에 SDK/pptx/Pillow import가 없다(INV-5)."""
        source = inspect.getsource(notes_module)
        forbidden = ["anthropic", "openai", "pptx", "PIL"]
        for name in forbidden:
            assert (
                f"import {name}" not in source
            ), f"notes.py must not import {name} (INV-5)"
