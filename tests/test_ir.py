"""Unit tests for src/pptx_md/ir.py — FR-05 중간 IR 정의.

Test naming convention: test_ac<N>_<description>
Each AC (AC1~AC8) has at least one test.  python-pptx is NOT mocked here;
IR is pure stdlib dataclass — direct instantiation suffices (ADR-206/207).
"""

from __future__ import annotations

import dataclasses
import importlib
import importlib.util
import pathlib

from pptx_md.ir import (
    GroupShapeIR,
    ImageClass,
    ImageShapeIR,
    OtherShapeIR,
    ParagraphIR,
    PresentationIR,
    ShapeIR,
    ShapeKind,
    SlideIR,
    TableShapeIR,
    TextShapeIR,
    iter_shapes,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_text_shape(shape_id: int = 1, name: str = "Title 1") -> TextShapeIR:
    return TextShapeIR(
        shape_id=shape_id,
        name=name,
        kind=ShapeKind.TEXT,
        paragraphs=[ParagraphIR(text="Hello", level=0)],
        is_title=True,
    )


def _make_table_shape(shape_id: int = 2) -> TableShapeIR:
    return TableShapeIR(
        shape_id=shape_id,
        name="Table 1",
        kind=ShapeKind.TABLE,
        rows=[["A", "B"], ["C", "D"]],
        n_rows=2,
        n_cols=2,
    )


def _make_image_shape(shape_id: int = 3) -> ImageShapeIR:
    return ImageShapeIR(
        shape_id=shape_id,
        name="Picture 1",
        kind=ShapeKind.IMAGE,
        image_bytes=b"\x89PNG\r\n",
        image_format="png",
        image_ext="png",
        alt_text="",
    )


def _make_group_shape(
    shape_id: int = 4, children: list[ShapeIR] | None = None
) -> GroupShapeIR:
    return GroupShapeIR(
        shape_id=shape_id,
        name="Group 1",
        kind=ShapeKind.GROUP,
        children=children or [],
    )


def _make_slide(index: int = 0, shapes: list[ShapeIR] | None = None) -> SlideIR:
    return SlideIR(
        index=index, title="Slide Title", notes="Notes here", shapes=shapes or []
    )


def _make_presentation(slides: list[SlideIR] | None = None) -> PresentationIR:
    return PresentationIR(
        source_path="/tmp/test.pptx",
        slide_width_emu=9144000,
        slide_height_emu=5143500,
        slides=slides or [],
    )


# ---------------------------------------------------------------------------
# AC1 — Presentation → Slide → Shape 3단 계층
# ---------------------------------------------------------------------------


def test_ac1_계층_구조_presentation_slide_shape() -> None:
    """AC1: PresentationIR → SlideIR → ShapeIR 3단 계층이 표현된다."""
    text_shape = _make_text_shape()
    slide = _make_slide(shapes=[text_shape])
    pres = _make_presentation(slides=[slide])

    # Presentation holds slides
    assert len(pres.slides) == 1
    assert isinstance(pres.slides[0], SlideIR)

    # Slide holds shapes
    assert len(pres.slides[0].shapes) == 1
    assert isinstance(pres.slides[0].shapes[0], ShapeIR)

    # Shape is the text shape we created
    shape = pres.slides[0].shapes[0]
    assert isinstance(shape, TextShapeIR)
    assert shape.kind == ShapeKind.TEXT


def test_ac1_계층_구조_다중_슬라이드() -> None:
    """AC1: Presentation 은 여러 SlideIR 를 순서대로 보유한다."""
    slides = [_make_slide(index=i) for i in range(3)]
    pres = _make_presentation(slides=slides)

    assert len(pres.slides) == 3
    for i, sld in enumerate(pres.slides):
        assert sld.index == i


# ---------------------------------------------------------------------------
# AC2 — 슬라이드 메타 필드 (index, title, notes, 빈 값 규약 "")
# ---------------------------------------------------------------------------


def test_ac2_슬라이드_메타_필드_존재() -> None:
    """AC2: SlideIR 는 index, title, notes 필드를 보유한다."""
    slide = SlideIR(index=0, title="My Title", notes="My Notes")

    assert slide.index == 0
    assert slide.title == "My Title"
    assert slide.notes == "My Notes"


def test_ac2_빈_title_notes_는_빈_문자열() -> None:
    """AC2: 빈 title/notes 는 None 이 아니라 빈 문자열 '' 로 표현된다 (ADR-202)."""
    slide = SlideIR(index=0)

    assert slide.title == ""
    assert slide.notes == ""
    # 타입이 str 이어야 함 (None 아님)
    assert isinstance(slide.title, str)
    assert isinstance(slide.notes, str)


# ---------------------------------------------------------------------------
# AC3 — 도형 타입 판별 (ShapeKind enum, 5종 서브클래스)
# ---------------------------------------------------------------------------


def test_ac3_도형_타입_판별_text() -> None:
    """AC3: TextShapeIR 는 kind=TEXT 로 판별 가능하다."""
    shape = _make_text_shape()
    assert shape.kind == ShapeKind.TEXT
    assert isinstance(shape, TextShapeIR)


def test_ac3_도형_타입_판별_table() -> None:
    """AC3: TableShapeIR 는 kind=TABLE 로 판별 가능하다."""
    shape = _make_table_shape()
    assert shape.kind == ShapeKind.TABLE
    assert isinstance(shape, TableShapeIR)
    assert shape.rows == [["A", "B"], ["C", "D"]]
    assert shape.n_rows == 2
    assert shape.n_cols == 2


def test_ac3_도형_타입_판별_image() -> None:
    """AC3: ImageShapeIR 는 kind=IMAGE 로 판별 가능하다."""
    shape = _make_image_shape()
    assert shape.kind == ShapeKind.IMAGE
    assert isinstance(shape, ImageShapeIR)
    assert shape.image_bytes == b"\x89PNG\r\n"
    assert shape.image_format == "png"


def test_ac3_도형_타입_판별_group() -> None:
    """AC3: GroupShapeIR 는 kind=GROUP 로 판별 가능하다."""
    shape = _make_group_shape()
    assert shape.kind == ShapeKind.GROUP
    assert isinstance(shape, GroupShapeIR)


def test_ac3_도형_타입_판별_other() -> None:
    """AC3: OtherShapeIR 는 kind=OTHER 로 판별 가능하다."""
    shape = OtherShapeIR(
        shape_id=5,
        name="Chart 1",
        kind=ShapeKind.OTHER,
        mso_shape_type="CHART",
        fallback_text="",
    )
    assert shape.kind == ShapeKind.OTHER
    assert isinstance(shape, OtherShapeIR)
    assert shape.mso_shape_type == "CHART"


def test_ac3_shape_kind_enum_모든_값() -> None:
    """AC3: ShapeKind enum 은 TEXT/TABLE/IMAGE/GROUP/OTHER 5종을 포함한다."""
    expected = {"text", "table", "image", "group", "other"}
    actual = {k.value for k in ShapeKind}
    assert actual == expected


# ---------------------------------------------------------------------------
# AC4 — 그룹 재귀 표현 (GroupShapeIR.children 중첩)
# ---------------------------------------------------------------------------


def test_ac4_그룹_재귀_표현_1depth() -> None:
    """AC4: GroupShapeIR 는 자식 ShapeIR 목록을 보유한다."""
    text = _make_text_shape(shape_id=10)
    group = _make_group_shape(shape_id=20, children=[text])

    assert len(group.children) == 1
    assert isinstance(group.children[0], TextShapeIR)


def test_ac4_그룹_재귀_표현_중첩() -> None:
    """AC4: 그룹 안 그룹 임의 깊이를 GroupShapeIR 중첩으로 표현한다."""
    inner_text = _make_text_shape(shape_id=1)
    inner_group = _make_group_shape(shape_id=2, children=[inner_text])
    outer_group = _make_group_shape(shape_id=3, children=[inner_group])

    assert isinstance(outer_group, GroupShapeIR)
    assert isinstance(outer_group.children[0], GroupShapeIR)
    assert isinstance(outer_group.children[0].children[0], TextShapeIR)  # type: ignore[union-attr]


def test_ac4_iter_shapes_그룹_재귀_평탄화() -> None:
    """AC4: iter_shapes 는 그룹을 재귀적으로 펼쳐 모든 도형을 깊이우선으로 반환한다."""
    text1 = _make_text_shape(shape_id=1)
    text2 = _make_text_shape(shape_id=2)
    inner_group = _make_group_shape(shape_id=10, children=[text2])
    outer_group = _make_group_shape(shape_id=11, children=[text1, inner_group])
    slide = _make_slide(shapes=[outer_group])

    flat = list(iter_shapes(slide))
    # depth-first pre-order: outer_group, text1, inner_group, text2
    shape_ids = [s.shape_id for s in flat]
    assert shape_ids == [11, 1, 10, 2]


# ---------------------------------------------------------------------------
# AC5 — 이미지 확장 슬롯 (classification, description 기본값 None)
# ---------------------------------------------------------------------------


def test_ac5_이미지_확장_슬롯_기본값_none() -> None:
    """AC5: v1 ImageShapeIR 의 classification, description 은 기본값 None 이다."""
    shape = _make_image_shape()
    assert shape.classification is None
    assert shape.description is None


def test_ac5_이미지_확장_슬롯_설정_가능() -> None:
    """AC5: 확장 슬롯은 M3/M4 가 값을 채울 수 있다."""
    shape = _make_image_shape()
    shape.classification = ImageClass.PHOTO
    shape.description = "A photograph of a mountain"

    assert shape.classification == ImageClass.PHOTO
    assert shape.description == "A photograph of a mountain"


def test_ac5_이미지_확장_슬롯이_비이미지_사용처를_깨지_않음() -> None:
    """AC5: ImageShapeIR 확장 슬롯은 TextShapeIR 등 다른 IR 에 영향 없다."""
    text_shape = _make_text_shape()
    table_shape = _make_table_shape()

    # 다른 shape IR 은 classification/description 필드가 없다
    assert not hasattr(text_shape, "classification")
    assert not hasattr(table_shape, "classification")


def test_ac5_image_class_enum_값() -> None:
    """AC5: ImageClass enum 은 TEXT/DIAGRAM/PHOTO/LOGO 4종을 포함한다."""
    expected = {"text", "diagram", "photo", "logo"}
    actual = {c.value for c in ImageClass}
    assert actual == expected


# ---------------------------------------------------------------------------
# AC6 — mypy strict (별도 CLI 검증; 테스트에서는 import 성공으로 대리 확인)
# ---------------------------------------------------------------------------


def test_ac6_ir_모듈_정상_import() -> None:
    """AC6: ir.py 가 에러 없이 import 되고 핵심 심볼이 노출된다.

    mypy strict exit 0 은 로컬 빌드 체크(ruff/black/mypy 명령)로 별도 확인한다.
    이 테스트는 런타임 import 가능성을 보장한다.
    """
    from pptx_md.ir import (  # noqa: F401
        GroupShapeIR,
        ImageClass,
        ImageShapeIR,
        OtherShapeIR,
        ParagraphIR,
        PresentationIR,
        ShapeIR,
        ShapeKind,
        SlideIR,
        TableShapeIR,
        TextShapeIR,
        iter_shapes,
    )


def test_ac6_ir_모듈에_pptx_import_없음() -> None:
    """AC6/ADR-206: ir.py 는 python-pptx 를 import 하지 않는다 (순수 stdlib)."""
    # ir 모듈을 fresh load (이미 sys.modules 에 있어도 spec 만 확인)
    spec = importlib.util.find_spec("pptx_md.ir")
    assert spec is not None
    assert spec.origin is not None

    # ir.py 소스에 pptx import 가 없음을 소스 파일로 확인
    source = pathlib.Path(spec.origin).read_text(encoding="utf-8")
    assert "from pptx" not in source, "ir.py must not import python-pptx (ADR-206)"
    assert "import pptx" not in source, "ir.py must not import python-pptx (ADR-206)"


# ---------------------------------------------------------------------------
# AC7 — 경계 (빈 컬렉션)
# ---------------------------------------------------------------------------


def test_ac7_빈_슬라이드_목록_presentation() -> None:
    """AC7: slides=[] 인 PresentationIR 는 정상 인스턴스화된다."""
    pres = PresentationIR(source_path="empty.pptx")
    assert pres.slides == []
    assert isinstance(pres.slides, list)


def test_ac7_빈_도형_목록_slide() -> None:
    """AC7: shapes=[] 인 SlideIR 는 정상 인스턴스화된다."""
    slide = SlideIR(index=0)
    assert slide.shapes == []
    assert isinstance(slide.shapes, list)


def test_ac7_빈_컬렉션_iter_shapes_는_빈_반환() -> None:
    """AC7: 도형 0개인 SlideIR 의 iter_shapes 는 아무것도 yield 하지 않는다."""
    slide = SlideIR(index=0)
    result = list(iter_shapes(slide))
    assert result == []


def test_ac7_빈_그룹_children_정상() -> None:
    """AC7: children=[] 인 GroupShapeIR 는 정상 인스턴스화된다."""
    group = GroupShapeIR(shape_id=1, name="EmptyGroup", kind=ShapeKind.GROUP)
    assert group.children == []


# ---------------------------------------------------------------------------
# AC8 — 직렬화 가능성 (dataclasses.asdict, image_bytes 제외 비교)
# ---------------------------------------------------------------------------


def test_ac8_직렬화_dataclasses_asdict() -> None:
    """AC8: dataclasses.asdict() 로 dict 변환 가능하다.

    image_bytes 를 제외한 구조를 동등 비교할 수 있다.
    """
    image_shape = _make_image_shape(shape_id=3)
    slide = SlideIR(
        index=0,
        title="Test Slide",
        notes="",
        shapes=[image_shape],
    )
    pres = PresentationIR(source_path="test.pptx", slides=[slide])

    result = dataclasses.asdict(pres)

    # image_bytes 를 제외하고 비교
    shapes_in_result = result["slides"][0]["shapes"]
    assert len(shapes_in_result) == 1

    shape_dict = shapes_in_result[0].copy()
    shape_dict.pop("image_bytes")  # 바이너리 제외

    assert shape_dict == {
        "shape_id": 3,
        "name": "Picture 1",
        "kind": "image",
        "left": 0,  # FR-23: coordinate fields default 0
        "top": 0,
        "width": 0,
        "height": 0,
        "image_format": "png",
        "image_ext": "png",
        "alt_text": "",
        "classification": None,
        "description": None,
    }


def test_ac8_직렬화_텍스트_도형_계층() -> None:
    """AC8: TextShapeIR 과 ParagraphIR 를 asdict 로 직렬화하면 결정적으로 표현된다."""
    para = ParagraphIR(text="Hello World", level=0)
    shape = TextShapeIR(
        shape_id=1,
        name="Title 1",
        kind=ShapeKind.TEXT,
        paragraphs=[para],
        is_title=True,
    )

    d = dataclasses.asdict(shape)

    assert d == {
        "shape_id": 1,
        "name": "Title 1",
        "kind": "text",
        "left": 0,  # FR-23: coordinate fields default 0
        "top": 0,
        "width": 0,
        "height": 0,
        "paragraphs": [{"text": "Hello World", "level": 0}],
        "is_title": True,
        "is_footer": False,  # FR-21: new field with default False
    }


def test_ac8_직렬화_표_도형() -> None:
    """AC8: TableShapeIR 을 asdict 하면 rows 2차원 배열이 결정적으로 나온다."""
    shape = _make_table_shape()
    d = dataclasses.asdict(shape)

    assert d["rows"] == [["A", "B"], ["C", "D"]]
    assert d["n_rows"] == 2
    assert d["n_cols"] == 2


def test_ac8_직렬화_슬라이드_계층() -> None:
    """AC8: SlideIR 전체를 asdict 하면 index/title/notes 및 shapes 계층이 나온다."""
    slide = SlideIR(
        index=2,
        title="Summary",
        notes="These are notes",
        shapes=[
            OtherShapeIR(
                shape_id=99,
                name="Chart",
                kind=ShapeKind.OTHER,
                mso_shape_type="CHART",
                fallback_text="",
            )
        ],
    )

    d = dataclasses.asdict(slide)

    assert d["index"] == 2
    assert d["title"] == "Summary"
    assert d["notes"] == "These are notes"
    assert d["shapes"][0]["mso_shape_type"] == "CHART"


# ---------------------------------------------------------------------------
# FR-23 (#58) — IR 좌표 확장 (AC1)
# ---------------------------------------------------------------------------


class TestFR23CoordinateFields:
    """FR-23 (#58): ShapeIR 좌표 필드 추가 — AC1"""

    def test_ac1_좌표필드_기본값_zero(self) -> None:
        """ac1_좌표필드_기본값_zero: ShapeIR 서브클래스를 좌표 인자 없이 생성하면
        left/top/width/height 모두 0이다."""
        text = TextShapeIR(shape_id=1, name="T", kind=ShapeKind.TEXT)
        assert text.left == 0
        assert text.top == 0
        assert text.width == 0
        assert text.height == 0

    def test_ac1_모든_서브클래스_no_TypeError(self) -> None:
        """ac1_모든_서브클래스_no_TypeError: 5종 서브클래스 모두 좌표 인자 없이
        TypeError 없이 생성된다."""
        TextShapeIR(shape_id=1, name="T", kind=ShapeKind.TEXT)
        TableShapeIR(shape_id=2, name="Tbl", kind=ShapeKind.TABLE)
        ImageShapeIR(shape_id=3, name="Img", kind=ShapeKind.IMAGE)
        GroupShapeIR(shape_id=4, name="Grp", kind=ShapeKind.GROUP)
        OtherShapeIR(shape_id=5, name="Oth", kind=ShapeKind.OTHER)

    def test_ac1_좌표_값_설정_가능(self) -> None:
        """ac1_좌표_값_설정: 좌표 인자를 명시하면 해당 값이 저장된다."""
        shape = TextShapeIR(
            shape_id=1,
            name="T",
            kind=ShapeKind.TEXT,
            left=914400,
            top=457200,
            width=4572000,
            height=1371600,
        )
        assert shape.left == 914400
        assert shape.top == 457200
        assert shape.width == 4572000
        assert shape.height == 1371600

    def test_ac1_좌표필드_int_타입(self) -> None:
        """ac1_좌표필드_타입: 필드 타입은 int 이다."""
        shape = TextShapeIR(shape_id=1, name="T", kind=ShapeKind.TEXT)
        assert isinstance(shape.left, int)
        assert isinstance(shape.top, int)
        assert isinstance(shape.width, int)
        assert isinstance(shape.height, int)

    def test_ac1_좌표_asdict_포함(self) -> None:
        """ac1_직렬화_포함: dataclasses.asdict 결과에 4개 좌표 필드가 포함된다."""
        shape = OtherShapeIR(
            shape_id=9,
            name="O",
            kind=ShapeKind.OTHER,
            left=100,
            top=200,
            width=300,
            height=400,
        )
        d = dataclasses.asdict(shape)
        assert d["left"] == 100
        assert d["top"] == 200
        assert d["width"] == 300
        assert d["height"] == 400
