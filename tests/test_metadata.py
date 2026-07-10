"""Tests for FR-28 metadata.py: frontmatter / slide metadata (issue #78, W-B).

Covers the dispatch AC list for issue #78 (mapped 1:1 onto REQ-wave3 §2
FR-28 AC5-AC7):
    AC1 -> REQ AC5 (frontmatter block, stdlib, no source_path/PII)
    AC2 -> REQ AC5 (slide HTML-comment metadata, deterministic, '\\n'/'|'
           sanitisation)
    AC3 -> REQ AC6 (has_diagram derivation via iter_shapes DFS, incl. groups)
    AC4 -> REQ AC5 (section derivation = top-level '##' heading text)
    AC5 -> REQ AC7 ('---'/FR-14 non-conflict — HTML comment only)
    AC6 -> INV-2/INV-5 (IR unchanged, no forbidden imports)
    AC7 -> gate (mypy/ruff/black/pytest/coverage — verified out-of-band)

AC name format: ac<N>_<short_description>
"""

from __future__ import annotations

import pptx_md.metadata as metadata_module
from pptx_md.assembler import assemble_document
from pptx_md.ir import (
    GroupShapeIR,
    ImageClass,
    ImageShapeIR,
    ParagraphIR,
    PresentationIR,
    ShapeKind,
    SlideIR,
    TextShapeIR,
)
from pptx_md.metadata import (
    build_frontmatter,
    build_slide_comment,
    derive_has_diagram,
    derive_section,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _image_shape(
    classification: ImageClass | None = None,
    shape_id: int = 30,
) -> ImageShapeIR:
    return ImageShapeIR(
        shape_id=shape_id,
        name="image",
        kind=ShapeKind.IMAGE,
        image_bytes=b"",
        image_format="",
        image_ext="",
        classification=classification,
    )


def _group_shape(children: list, shape_id: int = 50) -> GroupShapeIR:
    return GroupShapeIR(
        shape_id=shape_id,
        name="group",
        kind=ShapeKind.GROUP,
        children=children,
    )


def _text_shape(text: str, shape_id: int = 10, is_title: bool = False) -> TextShapeIR:
    return TextShapeIR(
        shape_id=shape_id,
        name="text",
        kind=ShapeKind.TEXT,
        paragraphs=[ParagraphIR(text=text, level=0)],
        is_title=is_title,
    )


# ---------------------------------------------------------------------------
# AC1: build_frontmatter — YAML block, stdlib, no source_path/PII
# ---------------------------------------------------------------------------


class TestAc1BuildFrontmatter:
    """build_frontmatter — 최상단 YAML frontmatter 1회, stdlib 조립"""

    def test_ac1_frontmatter_has_delimiters_and_slide_count(self) -> None:
        """ac1_frontmatter_구조: '---' 로 열고 닫으며 slide_count 를 담는다."""
        pres = PresentationIR(
            source_path="/tmp/secret/deck.pptx",
            slides=[SlideIR(index=0), SlideIR(index=1), SlideIR(index=2)],
        )
        fm = build_frontmatter(pres)
        lines = fm.split("\n")
        assert lines[0] == "---"
        assert lines[-1] == "---"
        assert "slide_count: 3" in fm

    def test_ac1_frontmatter_excludes_source_path(self) -> None:
        """ac1_PII_회피: source_path(로컬 경로)는 frontmatter 에 포함되지 않는다."""
        pres = PresentationIR(
            source_path=r"C:\Users\alice\Documents\confidential.pptx",
            slides=[SlideIR(index=0)],
        )
        fm = build_frontmatter(pres)
        assert "confidential" not in fm
        assert "alice" not in fm
        assert "source_path" not in fm

    def test_ac1_frontmatter_is_deterministic(self) -> None:
        """ac1_결정성: 동일 IR -> 동일 frontmatter 문자열(2회 호출 바이트 동일)."""
        pres = PresentationIR(source_path="x.pptx", slides=[SlideIR(index=0)])
        assert build_frontmatter(pres) == build_frontmatter(pres)

    def test_ac1_frontmatter_empty_presentation(self) -> None:
        """ac1_빈_프레젠테이션: 슬라이드가 없어도 예외 없이 slide_count: 0."""
        pres = PresentationIR(source_path="x.pptx", slides=[])
        fm = build_frontmatter(pres)
        assert "slide_count: 0" in fm


# ---------------------------------------------------------------------------
# AC2: build_slide_comment — HTML comment, deterministic, '\n'/'|' sanitised
# ---------------------------------------------------------------------------


class TestAc2BuildSlideComment:
    """build_slide_comment — 슬라이드별 HTML 주석 메타"""

    def test_ac2_comment_contains_all_three_keys(self) -> None:
        """ac2_3키_포함: slide_index/section/has_diagram 이 모두 포함된다."""
        comment = build_slide_comment(1, "표지", False)
        assert comment.startswith("<!--")
        assert comment.endswith("-->")
        assert "slide_index: 1" in comment
        assert "section: 표지" in comment
        assert "has_diagram: false" in comment

    def test_ac2_comment_has_diagram_true(self) -> None:
        """ac2_has_diagram_true: has_diagram=True -> 'has_diagram: true'."""
        comment = build_slide_comment(2, "본문", True)
        assert "has_diagram: true" in comment

    def test_ac2_comment_sanitises_newline_in_section(self) -> None:
        """ac2_개행_치환: section 내 '\\n' 은 공백으로 치환되어 단일 라인을 유지한다."""
        comment = build_slide_comment(1, "줄1\n줄2", False)
        assert "\n" not in comment
        assert "줄1 줄2" in comment

    def test_ac2_comment_sanitises_pipe_in_section(self) -> None:
        """ac2_파이프_치환: section 내 '|' 는 '/' 로 치환되어 구분자 파손을 방지한다."""
        comment = build_slide_comment(1, "A|B", False)
        assert "A/B" in comment
        # Only the two structural pipes (key separators) remain.
        assert comment.count("|") == 2

    def test_ac2_comment_is_deterministic(self) -> None:
        """ac2_결정성: 동일 인자 -> 동일 문자열."""
        assert build_slide_comment(5, "s", True) == build_slide_comment(5, "s", True)


# ---------------------------------------------------------------------------
# AC3: derive_has_diagram — iter_shapes DFS, incl. groups
# ---------------------------------------------------------------------------


class TestAc3DeriveHasDiagram:
    """derive_has_diagram — DIAGRAM 분류 이미지 존재 여부"""

    def test_ac3_diagram_present_top_level(self) -> None:
        """ac3_최상위_DIAGRAM: 최상위 도형에 DIAGRAM 이미지가 있으면 True."""
        slide = SlideIR(
            index=0, shapes=[_image_shape(classification=ImageClass.DIAGRAM)]
        )
        assert derive_has_diagram(slide) is True

    def test_ac3_no_diagram_returns_false(self) -> None:
        """ac3_DIAGRAM_없음: DIAGRAM 이미지가 없으면 False."""
        slide = SlideIR(index=0, shapes=[_image_shape(classification=ImageClass.PHOTO)])
        assert derive_has_diagram(slide) is False

    def test_ac3_no_images_at_all_returns_false(self) -> None:
        """ac3_이미지_없음: 이미지 도형 자체가 없으면 False."""
        slide = SlideIR(index=0, shapes=[_text_shape("hello")])
        assert derive_has_diagram(slide) is False

    def test_ac3_diagram_nested_in_group_returns_true(self) -> None:
        """ac3_그룹_내부_DIAGRAM: 그룹(iter_shapes DFS) 내부의 DIAGRAM 도 감지된다."""
        nested = _image_shape(classification=ImageClass.DIAGRAM, shape_id=31)
        group = _group_shape([_text_shape("caption", shape_id=11), nested])
        slide = SlideIR(index=0, shapes=[group])
        assert derive_has_diagram(slide) is True

    def test_ac3_classification_none_returns_false(self) -> None:
        """ac3_미분류: classification=None 인 이미지는 DIAGRAM 이 아니다."""
        slide = SlideIR(index=0, shapes=[_image_shape(classification=None)])
        assert derive_has_diagram(slide) is False


# ---------------------------------------------------------------------------
# AC4: derive_section — top-level '##' heading text, "" if absent
# ---------------------------------------------------------------------------


class TestAc4DeriveSection:
    """derive_section — 최상위 '##' 헤딩 텍스트 파생"""

    def test_ac4_returns_h2_text(self) -> None:
        """ac4_h2_텍스트: 최상위 '##' 라인의 텍스트를 반환한다."""
        assert derive_section(["## 표지"]) == "표지"

    def test_ac4_picks_first_h2_ignoring_h3_plus(self) -> None:
        """ac4_h3이하_무시: '###' 이후 세그먼트는 무시하고 첫 '##' 만 사용한다."""
        result = derive_section(["## 사업의 이해", "### 현황", "#### 세부"])
        assert result == "사업의 이해"

    def test_ac4_no_h2_returns_empty_string(self) -> None:
        """ac4_h2_없음: '##' 라인이 없으면 빈 문자열을 반환한다."""
        assert derive_section(["### 서브만"]) == ""

    def test_ac4_empty_heading_lines_returns_empty_string(self) -> None:
        """ac4_빈_리스트: heading_lines 가 비어 있으면 빈 문자열을 반환한다."""
        assert derive_section([]) == ""

    def test_ac4_strips_surrounding_whitespace(self) -> None:
        """ac4_공백_제거: 헤딩 텍스트 앞뒤 공백은 제거된다."""
        assert derive_section(["##   여백 제목   "]) == "여백 제목"


# ---------------------------------------------------------------------------
# AC5: '---'/FR-14 non-conflict — HTML comment only, integrated via assembler
# ---------------------------------------------------------------------------


class TestAc5NoConflictWithSeparatorsOrHeadings:
    """메타는 HTML 주석 전용 -> '---'/FR-14 헤딩 규칙과 무충돌"""

    def test_ac5_slide_comment_is_not_a_heading_or_fence(self) -> None:
        """ac5_주석_형식: 슬라이드 주석은 '#'/'```'/'|' 표 라인으로 시작하지 않는다."""
        comment = build_slide_comment(1, "표지", False)
        assert not comment.startswith("#")
        assert not comment.startswith("```")
        assert not comment.lstrip().startswith("|")

    def test_ac5_frontmatter_yaml_delimiter_appears_once_at_document_start(
        self,
    ) -> None:
        """ac5_YAML_1회: 문서 최상단 YAML '---' 블록은 1회로 한정된다."""
        pres = PresentationIR(
            source_path="x.pptx",
            slides=[
                SlideIR(index=0, title="표지", shapes=[]),
                SlideIR(index=1, title="본문", shapes=[_text_shape("내용")]),
            ],
        )
        doc = assemble_document(pres, emit_frontmatter=True)
        # The YAML frontmatter block occupies exactly the first 4 lines
        # (open '---', 2 keys, close '---'), followed by a blank separator
        # line and then the first slide's HTML-comment metadata — never
        # another standalone '---' immediately after the opening block.
        expected_prefix = "---\ngenerator: pptx-md\nslide_count: 2\n---\n\n"
        assert doc.startswith(expected_prefix)
        remainder = doc[len(expected_prefix) :]
        assert remainder.startswith("<!-- slide_index: 1")
        # Later slide-separator '---' lines (ADR-220 block join) are a
        # distinct, later occurrence — not part of the YAML block itself.
        assert "\n\n---\n\n" in doc

    def test_ac5_html_comment_attached_per_slide_in_document(self) -> None:
        """ac5_슬라이드별_주석부착: on 시 슬라이드에 HTML 주석 메타가 부착된다."""
        pres = PresentationIR(
            source_path="x.pptx",
            slides=[SlideIR(index=0, title="표지", shapes=[])],
        )
        doc = assemble_document(pres, emit_frontmatter=True)
        assert "<!-- slide_index: 1 | section: 표지 | has_diagram: false -->" in doc


# ---------------------------------------------------------------------------
# AC6: IR unchanged (INV-2) / no forbidden imports (INV-5)
# ---------------------------------------------------------------------------


class TestAc6IrUnchangedAndImportIsolation:
    """IR 무변경(INV-2) / VLM SDK·python-pptx·Pillow import 0 (INV-5)"""

    def test_ac6_module_imports_only_stdlib_and_ir(self) -> None:
        """ac6_import_격리: metadata 모듈은 'pptx_md.ir' 만 내부 의존으로 갖는다."""
        import ast

        forbidden = {"pptx", "PIL", "anthropic", "openai"}
        source = metadata_module.__file__
        assert source is not None
        with open(source, encoding="utf-8") as f:
            tree = ast.parse(f.read(), filename=source)

        imported_roots: set[str] = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imported_roots.update(alias.name.split(".")[0] for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                imported_roots.add(node.module.split(".")[0])

        assert imported_roots & forbidden == set()
        # Only the internal `pptx_md` package (for `pptx_md.ir`) is imported.
        assert imported_roots <= {"pptx_md", "__future__"}

    def test_ac6_derive_has_diagram_does_not_mutate_slide(self) -> None:
        """ac6_IR_무변경: derive_has_diagram 호출 후 shapes 리스트는 그대로 유지된다."""
        img = _image_shape(classification=ImageClass.DIAGRAM)
        slide = SlideIR(index=0, shapes=[img])
        before = list(slide.shapes)
        derive_has_diagram(slide)
        assert slide.shapes == before
        assert slide.shapes[0] is img
