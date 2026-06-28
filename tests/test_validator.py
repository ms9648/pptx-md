"""Tests for FR-14: validate_markdown (issue #36).

AC1 — 정상 MD → valid=True, warnings=[]
AC2 — 빈 문자열/공백 → valid=False, warnings 1건 이상
AC3 — 펜스 홀수 → valid=False, warnings에 미닫힘 사유
AC4 — ###으로 시작(첫 헤딩이 h1/h2 아님) → warnings에 시작 레벨 이상
AC5 — # → #### 점프(h1→h4, 2단계 초과) → warnings에 레벨 점프 사유
AC6 — 코드블록 내 # 텍스트 → 헤딩 오인 없음
AC7 — 어떤 입력이든 예외 없이 ValidationResult 반환
AC8 — 외부 Markdown 파서 의존 없음 (import 체크)
"""

from __future__ import annotations

import pytest

from pptx_md.validator import ValidationResult, validate_markdown

# ---------------------------------------------------------------------------
# AC1 — 정상 MD
# ---------------------------------------------------------------------------


class TestAc1NormalMarkdown:
    def test_ac1_simple_heading_and_paragraph(self) -> None:
        """AC1: 정상 헤딩·단락 → valid=True, warnings=[]."""
        md = "# Title\n\nSome paragraph."
        result = validate_markdown(md)
        assert result.valid is True
        assert result.warnings == []

    def test_ac1_with_closed_code_block(self) -> None:
        """AC1: 닫힌 코드블록 포함 정상 MD → valid=True, warnings=[]."""
        md = "## Heading\n\n```python\nprint('hello')\n```\n\nText."
        result = validate_markdown(md)
        assert result.valid is True
        assert result.warnings == []

    def test_ac1_sequential_headings(self) -> None:
        """AC1: h1→h2→h3 순차 헤딩 → valid=True, warnings=[]."""
        md = "# Top\n\n## Sub\n\n### Sub-sub\n\nContent."
        result = validate_markdown(md)
        assert result.valid is True
        assert result.warnings == []

    def test_ac1_returns_validation_result_type(self) -> None:
        """AC1: 반환 타입은 ValidationResult."""
        result = validate_markdown("# Hello")
        assert isinstance(result, ValidationResult)


# ---------------------------------------------------------------------------
# AC2 — 빈 문서
# ---------------------------------------------------------------------------


class TestAc2EmptyDocument:
    def test_ac2_empty_string(self) -> None:
        """AC2: 빈 문자열 → valid=False, warnings 1건 이상."""
        result = validate_markdown("")
        assert result.valid is False
        assert len(result.warnings) >= 1

    def test_ac2_whitespace_only(self) -> None:
        """AC2: 공백만 있는 문자열 → valid=False."""
        result = validate_markdown("   ")
        assert result.valid is False
        assert len(result.warnings) >= 1

    def test_ac2_newlines_only(self) -> None:
        """AC2: 개행만 있는 문자열 → valid=False."""
        result = validate_markdown("\n\n\n")
        assert result.valid is False
        assert len(result.warnings) >= 1

    def test_ac2_warning_message_contains_hint(self) -> None:
        """AC2: warnings 메시지에 빈 문서 사유 포함."""
        result = validate_markdown("")
        assert any("빈 문서" in w or "empty" in w.lower() for w in result.warnings)


# ---------------------------------------------------------------------------
# AC3 — 홀수 펜스
# ---------------------------------------------------------------------------


class TestAc3UnclosedCodeFence:
    def test_ac3_single_opening_fence(self) -> None:
        """AC3: 여는 펜스만 1개 → valid=False, warnings에 미닫힘 사유."""
        md = "# Title\n\n```python\ncode here"
        result = validate_markdown(md)
        assert result.valid is False
        assert len(result.warnings) >= 1

    def test_ac3_warning_describes_unclosed_fence(self) -> None:
        """AC3: warnings 메시지에 코드블록 미닫힘 사유 포함."""
        md = "# Title\n\n```\nsome code"
        result = validate_markdown(md)
        assert any(
            "코드블록" in w or "펜스" in w or "fence" in w.lower()
            for w in result.warnings
        )

    def test_ac3_three_fences(self) -> None:
        """AC3: 펜스 3개(홀수) → valid=False."""
        md = "# Title\n\n```\ncode\n```\n\ntext\n\n```\nunclosed"
        result = validate_markdown(md)
        assert result.valid is False

    def test_ac3_two_fences_closed(self) -> None:
        """AC3(역): 펜스 2개(짝수, 정상 닫힘) → valid=True."""
        md = "# Title\n\n```\ncode\n```\n\nAfter."
        result = validate_markdown(md)
        assert result.valid is True


# ---------------------------------------------------------------------------
# AC4 — 첫 헤딩이 h3+ (h1/h2 아님)
# ---------------------------------------------------------------------------


class TestAc4HeadingStartLevel:
    def test_ac4_starts_with_h3(self) -> None:
        """AC4: 문서가 ### 으로 시작 → warnings에 시작 레벨 이상 사유."""
        md = "### Deep Start\n\nContent."
        result = validate_markdown(md)
        assert any(
            "시작" in w or "start" in w.lower() or "레벨" in w
            for w in result.warnings
        )

    def test_ac4_starts_with_h4(self) -> None:
        """AC4: 첫 헤딩이 h4 → warnings 포함."""
        md = "#### Very Deep\n\nText."
        result = validate_markdown(md)
        assert len(result.warnings) >= 1

    def test_ac4_does_not_set_valid_false(self) -> None:
        """AC4: 헤딩 시작 이상은 valid에 영향 없음."""
        md = "### Just Warning\n\nContent."
        result = validate_markdown(md)
        assert result.valid is True

    def test_ac4_h2_start_is_ok(self) -> None:
        """AC4(경계): ## 시작은 경고 없음."""
        md = "## Allowed Start\n\nContent."
        result = validate_markdown(md)
        assert result.warnings == []


# ---------------------------------------------------------------------------
# AC5 — 헤딩 레벨 2단계 이상 점프
# ---------------------------------------------------------------------------


class TestAc5HeadingLevelJump:
    def test_ac5_h1_to_h4_jump(self) -> None:
        """AC5: # 다음 #### (h1→h4, 3단계 점프) → warnings에 레벨 점프 사유."""
        md = "# Top\n\n#### Deep\n\nContent."
        result = validate_markdown(md)
        assert any(
            "점프" in w or "jump" in w.lower() or "레벨" in w
            for w in result.warnings
        )

    def test_ac5_h1_to_h3_jump(self) -> None:
        """AC5: # 다음 ### (h1→h3, 2단계 점프) → warning 포함."""
        md = "# Top\n\n### Skip\n\nContent."
        result = validate_markdown(md)
        assert len(result.warnings) >= 1

    def test_ac5_does_not_set_valid_false(self) -> None:
        """AC5: 레벨 점프는 valid에 영향 없음."""
        md = "# Top\n\n#### Deep"
        result = validate_markdown(md)
        assert result.valid is True

    def test_ac5_h1_to_h2_sequential_ok(self) -> None:
        """AC5(역): h1→h2 는 1단계 점프로 경고 없음."""
        md = "# Top\n\n## Sub\n\nText."
        result = validate_markdown(md)
        assert result.warnings == []

    def test_ac5_h2_to_h4_jump(self) -> None:
        """AC5: ## 다음 #### (h2→h4, 2단계 점프) → warning."""
        md = "## Section\n\n#### Deep Jump\n\nContent."
        result = validate_markdown(md)
        assert len(result.warnings) >= 1


# ---------------------------------------------------------------------------
# AC6 — 코드블록 내부 # 헤딩 오인 방지
# ---------------------------------------------------------------------------


class TestAc6CodeBlockHeadingExclusion:
    def test_ac6_hash_inside_closed_code_block_not_treated_as_heading(
        self,
    ) -> None:
        """AC6: 닫힌 코드블록 내 # 텍스트 → 헤딩으로 처리하지 않음."""
        md = "## Normal Heading\n\n```\n# not a heading\n## also not\n```\n\nText."
        result = validate_markdown(md)
        assert result.valid is True
        assert result.warnings == []

    def test_ac6_deep_hash_in_code_block_no_jump_warning(self) -> None:
        """AC6: 코드블록 내 #### 텍스트 → 레벨 점프 경고 없음."""
        md = "# Title\n\n```\n#### code comment\n```\n\n## After"
        result = validate_markdown(md)
        assert result.valid is True
        assert result.warnings == []

    def test_ac6_hash_outside_code_block_is_detected(self) -> None:
        """AC6(대조): 코드블록 외부 #### 는 경고 대상."""
        md = "# Title\n\n#### Outside\n\nContent."
        result = validate_markdown(md)
        assert len(result.warnings) >= 1

    def test_ac6_multiple_code_blocks_no_false_positives(self) -> None:
        """AC6: 여러 코드블록 내 # 텍스트 모두 무시."""
        md = (
            "# Heading\n\n"
            "```\n# in block 1\n```\n\n"
            "## Sub\n\n"
            "```\n### in block 2\n```\n\n"
            "Text."
        )
        result = validate_markdown(md)
        assert result.valid is True
        assert result.warnings == []


# ---------------------------------------------------------------------------
# AC7 — 비예외 보장
# ---------------------------------------------------------------------------


class TestAc7NoException:
    def test_ac7_empty_string(self) -> None:
        """AC7: 빈 문자열 → 예외 없이 ValidationResult 반환."""
        result = validate_markdown("")
        assert isinstance(result, ValidationResult)

    def test_ac7_very_long_string(self) -> None:
        """AC7: 매우 긴 문자열 → 예외 없이 반환."""
        md = "# Title\n\n" + ("word " * 100_000)
        result = validate_markdown(md)
        assert isinstance(result, ValidationResult)

    def test_ac7_binary_like_text(self) -> None:
        """AC7: 이진 유사 텍스트(null bytes 등) → 예외 없이 반환."""
        md = "# Title\n\x00\xff\xfe binary-ish"
        result = validate_markdown(md)
        assert isinstance(result, ValidationResult)

    def test_ac7_unicode_content(self) -> None:
        """AC7: 유니코드 문자 포함 → 예외 없이 반환."""
        md = "# 한글 제목\n\n내용입니다. 🎉"
        result = validate_markdown(md)
        assert isinstance(result, ValidationResult)

    @pytest.mark.parametrize(
        "md",
        [
            "###",
            "#",
            "```",
            "   #   ",
            "\t# Tab heading",
            "###### h6",
            "####### not a heading",
        ],
    )
    def test_ac7_edge_case_inputs(self, md: str) -> None:
        """AC7: 경계 입력들 → 예외 없이 ValidationResult 반환."""
        result = validate_markdown(md)
        assert isinstance(result, ValidationResult)


# ---------------------------------------------------------------------------
# AC8 — 외부 Markdown 파서 의존 없음
# ---------------------------------------------------------------------------


class TestAc8NoDependency:
    def test_ac8_import_validator_without_markdown_parsers(self) -> None:
        """AC8: markdown/mistune/commonmark 없이 validator import 성공."""
        # validator가 표준 라이브러리만 사용하므로 외부 파서 없이도 동작해야 한다.
        import importlib.util

        spec = importlib.util.find_spec("pptx_md.validator")
        assert spec is not None, "pptx_md.validator 모듈을 찾을 수 없습니다"

    def test_ac8_no_markdown_parser_in_module_imports(self) -> None:
        """AC8: validator 모듈의 의존 모듈에 외부 Markdown 파서가 없음."""
        import pptx_md.validator as validator_module

        source_file = validator_module.__file__
        assert source_file is not None
        with open(source_file, encoding="utf-8") as f:
            source = f.read()
        # 외부 Markdown 파서 라이브러리 이름이 소스에 포함되지 않아야 함
        forbidden = ["import markdown", "import mistune", "import commonmark",
                     "from markdown", "from mistune", "from commonmark"]
        for forbidden_import in forbidden:
            assert forbidden_import not in source, (
                f"외부 Markdown 파서 의존 발견: {forbidden_import}"
            )
