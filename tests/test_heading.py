"""Tests for FR-28 heading.py: heading-path decomposition (issue #77, W-A).

Covers AC1-AC3 of issue #77 (heading_hierarchy) plus the shared D-4
is_body_empty helper (used by both FR-28 AC4 and FR-29 AC6).
AC name format: ac<N>_<short_description>
"""

from __future__ import annotations

from pptx_md.heading import is_body_empty, render_heading_lines, split_heading_path

# ---------------------------------------------------------------------------
# split_heading_path
# ---------------------------------------------------------------------------


class TestSplitHeadingPath:
    """split_heading_path — '>' 경로 분해 + trim + 빈 세그먼트 제거"""

    def test_ac1_splits_and_trims_segments(self) -> None:
        """ac1_경로_분해_trim: '>' 로 분해하고 각 세그먼트를 trim한다."""
        result = split_heading_path("Ⅱ. 사업의 이해 > ⑦ 현황·문제점 > 세부")
        assert result == ["Ⅱ. 사업의 이해", "⑦ 현황·문제점", "세부"]

    def test_ac1_drops_empty_segments(self) -> None:
        """ac1_빈_세그먼트_제거: 선행/후행/중복 구분자로 생긴 빈 세그먼트는 제거된다."""
        result = split_heading_path(">A>>B>")
        assert result == ["A", "B"]

    def test_ac2_no_separator_returns_single_segment(self) -> None:
        """ac2_구분자_없음: '>' 이 없으면 단일 세그먼트 리스트를 반환한다."""
        result = split_heading_path("단순 제목")
        assert result == ["단순 제목"]

    def test_ac2_empty_title_returns_empty_list(self) -> None:
        """ac2_빈_제목: 빈 문자열은 빈 리스트를 반환한다."""
        assert split_heading_path("") == []


# ---------------------------------------------------------------------------
# render_heading_lines
# ---------------------------------------------------------------------------


class TestRenderHeadingLinesHierarchyOn:
    """render_heading_lines(hierarchy=True) — 계층 분해 (AC1)"""

    def test_ac1_two_segments_h2_h3(self) -> None:
        """ac1_2세그먼트_h2_h3: 첫 세그먼트 ##, 두번째 ###."""
        result = render_heading_lines("Ⅱ. 사업의 이해 > ⑦ 현황·문제점", hierarchy=True)
        assert result == ["## Ⅱ. 사업의 이해", "### ⑦ 현황·문제점"]

    def test_ac1_three_segments_h2_h3_h4(self) -> None:
        """ac1_3세그먼트_h2_h3_h4: 세 세그먼트가 각각 ##/###/#### 로 분해된다."""
        result = render_heading_lines(
            "Ⅱ. 사업의 이해 > ⑦ 현황·문제점 > 세부", hierarchy=True
        )
        assert result == ["## Ⅱ. 사업의 이해", "### ⑦ 현황·문제점", "#### 세부"]

    def test_ac1_segments_are_trimmed(self) -> None:
        """ac1_세그먼트_trim: 각 세그먼트 앞뒤 공백이 제거된다."""
        result = render_heading_lines("  A  >  B  ", hierarchy=True)
        assert result == ["## A", "### B"]

    def test_ac1_h6_clamp_beyond_max_level(self) -> None:
        """ac1_h6_클램프: 세그먼트가 5개 이상이면 6단계(h6)에서 클램프된다."""
        title = "A > B > C > D > E > F > G"  # 7 segments
        result = render_heading_lines(title, hierarchy=True)
        assert result == [
            "## A",
            "### B",
            "#### C",
            "##### D",
            "###### E",
            "###### F",  # clamped at h6
            "###### G",  # clamped at h6
        ]

    def test_ac2_no_separator_single_heading(self) -> None:
        """ac2_구분자_없음_단일헤딩: hierarchy=True 라도 구분자 없으면 단일 ##."""
        result = render_heading_lines("단순 제목", hierarchy=True)
        assert result == ["## 단순 제목"]

    def test_ac2_single_segment_after_trim_is_single_heading(self) -> None:
        """ac2_trim후_단일세그먼트: 선행/후행 구분자만 있어 세그먼트가 1개면 단일 ##."""
        result = render_heading_lines(">단일 세그먼트>", hierarchy=True)
        assert result == ["## >단일 세그먼트>"]


class TestRenderHeadingLinesHierarchyOff:
    """render_heading_lines(hierarchy=False) — 바이트 동일 보존 (AC3)"""

    def test_ac3_off_always_single_heading_verbatim(self) -> None:
        """ac3_off_원본_그대로: '>' 포함 제목도 단일 ##, 원본 title 무가공."""
        title = "  Ⅱ. 사업의 이해 > ⑦ 현황·문제점  "
        result = render_heading_lines(title, hierarchy=False)
        assert result == [f"## {title}"]

    def test_ac3_off_simple_title_unchanged(self) -> None:
        """ac3_off_단순제목: 구분자 없는 제목도 기존과 동일하게 렌더된다."""
        result = render_heading_lines("단순 제목", hierarchy=False)
        assert result == ["## 단순 제목"]


# ---------------------------------------------------------------------------
# is_body_empty (D-4, shared with FR-29 AC6)
# ---------------------------------------------------------------------------


class TestIsBodyEmpty:
    """is_body_empty — D-4 빈/구분 슬라이드 판정"""

    def test_ac4_empty_list_is_empty(self) -> None:
        """ac4_빈리스트: body_parts가 비어 있으면 True."""
        assert is_body_empty([]) is True

    def test_ac4_whitespace_only_parts_is_empty(self) -> None:
        """ac4_공백만: 공백만 있는 부분들은 비어있는 것으로 판정된다."""
        assert is_body_empty(["   ", "\n\t"]) is True

    def test_ac4_non_empty_part_is_not_empty(self) -> None:
        """ac4_비공백_존재: 비공백 텍스트가 있으면 False."""
        assert is_body_empty(["", "본문 내용", ""]) is False
