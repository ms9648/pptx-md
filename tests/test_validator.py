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
            "시작" in w or "start" in w.lower() or "레벨" in w for w in result.warnings
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
            "점프" in w or "jump" in w.lower() or "레벨" in w for w in result.warnings
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
        forbidden = [
            "import markdown",
            "import mistune",
            "import commonmark",
            "from markdown",
            "from mistune",
            "from commonmark",
        ]
        for forbidden_import in forbidden:
            assert (
                forbidden_import not in source
            ), f"외부 Markdown 파서 의존 발견: {forbidden_import}"


# ---------------------------------------------------------------------------
# FR-29 (issue #80): validator 신규 4규칙 — ARCH-Wave3 §3.7, ADR-619
#
# AC1 — 데이터 행 0인 pipe table → 깨진 표 경고 1건 (펜스 내 `|` 미오탐)
# AC2 — 헤더와 열 수 불일치 데이터 행 → 열 수 불일치 경고 1건
# AC3 — 인접 동일 텍스트 헤딩 2줄 → 중복 인접 헤딩 경고 1건
# AC4 — `\v` 잔존 → 잔존 개수 포함 경고 1건
# AC5 — 전공백 pipe table → 전공백 표 경고 1건
# AC6 — 본문 없는 슬라이드 블록(D-4) → 빈 슬라이드 경고 1건
# AC7 — 정상 문서 → 신규 경고 0, FR-14 회귀 0, 무예외, valid 불변
# AC8 — ruff/black/mypy/pytest 게이트 (별도 커맨드 실행으로 증명, 본 파일은 통과로 기여)
# ---------------------------------------------------------------------------


class TestFr29Ac1BrokenPipeTable:
    def test_ac1_header_and_separator_only_warns_no_data(self) -> None:
        """AC1: 헤더+구분행만 있고 데이터 행 0 → 깨진 표 경고 1건."""
        md = "## Title\n\n| A | B |\n| --- | --- |\n\nText."
        result = validate_markdown(md)
        assert result.valid is True
        assert sum("깨진 표" in w for w in result.warnings) == 1

    def test_ac1_fenced_pipe_lines_not_misdetected(self) -> None:
        """AC1: 코드펜스 내부 `|` 라인은 표로 오인하지 않음(경고 0)."""
        md = "## Title\n\n```\n| A | B |\n| --- | --- |\n```\n\nText."
        result = validate_markdown(md)
        assert not any("표" in w for w in result.warnings)

    def test_ac1_normal_table_with_data_no_warning(self) -> None:
        """AC1(대조): 데이터 행이 있는 정상 표 → 깨진 표 경고 없음."""
        md = "## Title\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\nText."
        result = validate_markdown(md)
        assert not any("깨진 표" in w for w in result.warnings)

    def test_ac1_second_line_not_a_valid_separator_is_not_a_table(self) -> None:
        """AC1(대조): 2번째 행이 구분행이 아니면 표로 확정하지 않음(경고 없음)."""
        md = "## Title\n\n| A |\n|\n\nText."
        result = validate_markdown(md)
        assert not any("표" in w for w in result.warnings)


class TestFr29Ac2ColumnMismatch:
    def test_ac2_data_row_fewer_columns_warns(self) -> None:
        """AC2: 데이터 행의 열 수가 헤더보다 적음 → 열 수 불일치 경고 1건."""
        md = "## Title\n\n| A | B | C |\n| --- | --- | --- |\n| 1 | 2 |\n\nText."
        result = validate_markdown(md)
        assert result.valid is True
        assert sum("열 수 불일치" in w for w in result.warnings) == 1

    def test_ac2_data_row_more_columns_warns(self) -> None:
        """AC2: 데이터 행의 열 수가 헤더보다 많음 → 열 수 불일치 경고 1건."""
        md = "## Title\n\n| A | B |\n| --- | --- |\n| 1 | 2 | 3 |\n\nText."
        result = validate_markdown(md)
        assert sum("열 수 불일치" in w for w in result.warnings) == 1

    def test_ac2_matching_columns_no_warning(self) -> None:
        """AC2(대조): 열 수가 일치하면 경고 없음."""
        md = "## Title\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n\nText."
        result = validate_markdown(md)
        assert not any("열 수 불일치" in w for w in result.warnings)


class TestFr29Ac3DuplicateAdjacentHeadings:
    def test_ac3_adjacent_identical_headings_warns(self) -> None:
        """AC3: 인접(사이 본문 없음) 동일 텍스트 헤딩 2줄 → 중복 인접 헤딩 경고 1건."""
        md = "## Same\n\n## Same\n\nContent."
        result = validate_markdown(md)
        assert result.valid is True
        assert sum("중복 인접 헤딩" in w for w in result.warnings) == 1

    def test_ac3_headings_separated_by_body_no_warning(self) -> None:
        """AC3(대조): 사이에 본문이 있으면 인접 아님 → 경고 없음."""
        md = "## Same\n\nSome content.\n\n## Same\n\nMore."
        result = validate_markdown(md)
        assert not any("중복 인접 헤딩" in w for w in result.warnings)

    def test_ac3_headings_separated_by_slide_separator_no_warning(self) -> None:
        """AC3(대조): `---` 슬라이드 구분자가 개입하면 인접 아님 → 경고 없음."""
        md = "## Same\n\nBody 1.\n\n---\n\n## Same\n\nBody 2."
        result = validate_markdown(md)
        assert not any("중복 인접 헤딩" in w for w in result.warnings)

    def test_ac3_different_text_headings_no_warning(self) -> None:
        """AC3(대조): 텍스트가 다르면 인접해도 경고 없음."""
        md = "## First\n\n## Second\n\nContent."
        result = validate_markdown(md)
        assert not any("중복 인접 헤딩" in w for w in result.warnings)


class TestFr29Ac4ResidualVerticalTab:
    def test_ac4_single_vertical_tab_warns_with_count(self) -> None:
        """AC4: `\\v` 1개 잔존 → 잔존 개수(1개) 포함 경고 1건."""
        md = "## Title\n\nLine one\vLine two."
        result = validate_markdown(md)
        assert result.valid is True
        assert any("\\v" in w and "1개" in w for w in result.warnings)

    def test_ac4_multiple_vertical_tabs_count_correct(self) -> None:
        """AC4: `\\v` 3개 잔존 → 경고 메시지에 개수(3개) 포함."""
        md = "## Title\n\nA\vB\vC\vD."
        result = validate_markdown(md)
        assert any("3개" in w for w in result.warnings)

    def test_ac4_no_vertical_tab_no_warning(self) -> None:
        """AC4(대조): `\\v` 없으면 경고 없음."""
        md = "## Title\n\nNormal text."
        result = validate_markdown(md)
        assert not any("\\v" in w for w in result.warnings)


class TestFr29Ac5AllBlankPipeTable:
    def test_ac5_all_cells_blank_warns(self) -> None:
        """AC5: 모든 셀이 공백인 pipe table → 전공백 표 경고 1건."""
        md = "## Title\n\n| A | B |\n| --- | --- |\n|  |  |\n\nText."
        result = validate_markdown(md)
        assert result.valid is True
        assert sum("전공백 표" in w for w in result.warnings) == 1

    def test_ac5_partial_content_no_warning(self) -> None:
        """AC5(대조): 일부 셀에 내용이 있으면 전공백 경고 없음."""
        md = "## Title\n\n| A | B |\n| --- | --- |\n| 1 |  |\n\nText."
        result = validate_markdown(md)
        assert not any("전공백 표" in w for w in result.warnings)


class TestFr29Ac6EmptySlideBlock:
    def test_ac6_slide_block_with_only_heading_warns(self) -> None:
        """AC6: 헤딩만 있고 본문 없는 슬라이드 블록 → 빈 슬라이드 경고 1건."""
        md = (
            "## Slide 1\n\nBody.\n\n---\n\n"
            "## Slide 2\n\n---\n\n"
            "## Slide 3\n\nMore body."
        )
        result = validate_markdown(md)
        assert result.valid is True
        assert sum("빈 슬라이드" in w for w in result.warnings) == 1

    def test_ac6_slide_block_with_heading_and_comment_only_warns(self) -> None:
        """AC6: 헤딩+HTML 주석만 있고 본문 없는 블록도 빈 슬라이드로 판정."""
        md = (
            "## Slide 1\n\nBody.\n\n---\n\n"
            "<!-- slide_index: 2 | section: Slide 2 | has_diagram: false -->\n"
            "## Slide 2\n\n---\n\n"
            "## Slide 3\n\nMore body."
        )
        result = validate_markdown(md)
        assert sum("빈 슬라이드" in w for w in result.warnings) == 1

    def test_ac6_all_blocks_with_body_no_warning(self) -> None:
        """AC6(대조): 모든 블록에 본문이 있으면 빈 슬라이드 경고 없음."""
        md = "## Slide 1\n\nBody 1.\n\n---\n\n## Slide 2\n\nBody 2."
        result = validate_markdown(md)
        assert not any("빈 슬라이드" in w for w in result.warnings)

    def test_ac6_zero_line_block_between_adjacent_separators_warns(self) -> None:
        """AC6: 연속된 `---` 사이 완전 빈 블록(0줄)도 빈 슬라이드로 판정."""
        md = "## Slide 1\n\nBody.\n\n---\n---\n\n## Slide 2\n\nBody 2."
        result = validate_markdown(md)
        assert sum("빈 슬라이드" in w for w in result.warnings) == 1

    def test_ac6_frontmatter_leading_blank_lines_still_stripped(self) -> None:
        """AC6: frontmatter 앞에 공백 줄이 있어도 정상 스킵되어 오탐 없음."""
        md = (
            "\n\n---\ngenerator: pptx-md\n---\n\n"
            "## Slide 1\n\nBody 1.\n\n---\n\n## Slide 2\n\nBody 2."
        )
        result = validate_markdown(md)
        assert not any("빈 슬라이드" in w for w in result.warnings)


class TestFr29Ac7NormalDocumentNoNewWarnings:
    def test_ac7_normal_document_zero_new_warnings(self) -> None:
        """AC7: 결함 없는 정상 문서 → FR-29 신규 경고 0건, valid=True."""
        md = (
            "## 표지\n\n본문 텍스트입니다.\n\n"
            "| A | B |\n| --- | --- |\n| 1 | 2 |\n\n"
            "---\n\n"
            "## Ⅱ. 사업의 이해\n\n본문 2.\n\n"
            "---\n\n"
            "## 결론\n\n마지막 본문."
        )
        result = validate_markdown(md)
        assert result.valid is True
        assert result.warnings == []

    def test_ac7_existing_fr14_warnings_unaffected(self) -> None:
        """AC7: FR-14 헤딩 시작/점프 경고는 회귀 없이 유지, 신규 경고와 안 섞임."""
        md = "### Deep Start\n\nContent."
        result = validate_markdown(md)
        assert any("시작" in w for w in result.warnings)
        new_rule_hits = [
            w
            for w in result.warnings
            if "깨진 표" in w
            or "열 수 불일치" in w
            or "전공백 표" in w
            or "중복 인접 헤딩" in w
            or "\\v" in w
            or "빈 슬라이드" in w
        ]
        assert new_rule_hits == []

    def test_ac7_no_exception_on_malformed_pipe_content(self) -> None:
        """AC7: 기형적인 pipe 콘텐츠에도 예외 없이 ValidationResult 반환."""
        md = "## Title\n\n|||||\n|-|\n\nContent."
        result = validate_markdown(md)
        assert isinstance(result, ValidationResult)

    def test_ac7_frontmatter_not_misdetected(self) -> None:
        """AC7: frontmatter 블록은 표/빈 슬라이드 분석에서 제외되어 오탐 없음."""
        md = (
            "---\n"
            "generator: pptx-md\n"
            "slide_count: 2\n"
            "---\n\n"
            "## Slide 1\n\nBody 1.\n\n---\n\n## Slide 2\n\nBody 2."
        )
        result = validate_markdown(md)
        assert not any("빈 슬라이드" in w for w in result.warnings)
        assert not any("표" in w for w in result.warnings)
