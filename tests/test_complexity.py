"""Unit tests for src/pptx_md/complexity.py — FR-31 시각 복잡도 탐지 (issue #92).

Test naming convention: test_ac<N>_<description>
Each AC (AC1~AC6, +oberride AC7 design note) has at least one covering test
that asserts the Then clause directly. All fixtures are synthetic SlideIR
instances built in-process (no PPTX/VLM/render dependency, ADR-624).
"""

from __future__ import annotations

from pptx_md.complexity import (
    CHART_SMARTART_MSO_SHAPE_TYPES,
    COMPLEXITY_THRESHOLD,
    CONNECTOR_COUNT_THRESHOLD,
    CONNECTOR_MSO_SHAPE_TYPES,
    complexity_score,
    is_visually_complex,
)
from pptx_md.ir import (
    GroupShapeIR,
    OtherShapeIR,
    ParagraphIR,
    ShapeIR,
    ShapeKind,
    SlideIR,
    TableShapeIR,
    TextShapeIR,
)

# ---------------------------------------------------------------------------
# Helpers — synthetic shape/slide builders (EMU coordinates)
# ---------------------------------------------------------------------------

_EMU_PER_INCH = 914_400


def _text(
    shape_id: int,
    *,
    left: int = 0,
    top: int = 0,
    text: str = "body",
    is_title: bool = False,
) -> TextShapeIR:
    return TextShapeIR(
        shape_id=shape_id,
        name=f"TextBox {shape_id}",
        kind=ShapeKind.TEXT,
        left=left,
        top=top,
        width=_EMU_PER_INCH,
        height=_EMU_PER_INCH,
        paragraphs=[ParagraphIR(text=text, level=0)],
        is_title=is_title,
    )


def _connector(shape_id: int) -> OtherShapeIR:
    return OtherShapeIR(
        shape_id=shape_id,
        name=f"Connector {shape_id}",
        kind=ShapeKind.OTHER,
        mso_shape_type="LINE",
        fallback_text="",
    )


def _chart(shape_id: int = 900) -> OtherShapeIR:
    return OtherShapeIR(
        shape_id=shape_id,
        name="Chart 1",
        kind=ShapeKind.OTHER,
        mso_shape_type="CHART",
        fallback_text="",
    )


def _smartart(shape_id: int = 901) -> OtherShapeIR:
    return OtherShapeIR(
        shape_id=shape_id,
        name="SmartArt 1",
        kind=ShapeKind.OTHER,
        mso_shape_type="DIAGRAM",
        fallback_text="",
    )


def _slide(index: int, shapes: list[ShapeIR]) -> SlideIR:
    return SlideIR(index=index, title="Slide", shapes=shapes)


# ---------------------------------------------------------------------------
# AC1 — 순수·결정적, IR 파생 특징만, VLM/렌더 미호출 (모듈 최상단 import 로 대리 확인)
# ---------------------------------------------------------------------------


def test_ac1_is_visually_complex_는_bool_반환() -> None:
    """AC1: is_visually_complex(slide) -> bool."""
    slide = _slide(0, [_text(1, is_title=True)])
    result = is_visually_complex(slide)
    assert isinstance(result, bool)


def test_ac1_complexity_score_는_int_반환() -> None:
    """AC1: complexity_score(slide) -> int."""
    slide = _slide(0, [_text(1, is_title=True)])
    result = complexity_score(slide)
    assert isinstance(result, int)


def test_ac1_모듈에_vlm_sdk_import_없음() -> None:
    """AC1/AC6/INV-5: complexity.py 소스에 VLM SDK/렌더 import 가 없다."""
    import pathlib

    import pptx_md.complexity as complexity_module

    source = pathlib.Path(complexity_module.__file__).read_text(encoding="utf-8")
    for forbidden in (
        "import anthropic",
        "import openai",
        "import pypdfium2",
        "import PIL",
        "from PIL",
        "import pptx",
        "from pptx import",
        "from pptx.",
    ):
        assert forbidden not in source, f"complexity.py must not import {forbidden!r}"


# ---------------------------------------------------------------------------
# AC2 — 멀티컬럼(동일 top 밴드 3+ 텍스트, 서로 다른 left) -> True
# ---------------------------------------------------------------------------


def test_ac2_멀티컬럼_4열_매핑_슬라이드는_True() -> None:
    """AC2: 동일 top 밴드에 4개 열(다른 left)로 정렬된 텍스트 도형 -> True (S5 재현)."""
    band_top = 2 * _EMU_PER_INCH
    shapes: list[ShapeIR] = [_text(1, is_title=True, top=0)]
    for i, left_in in enumerate([0, 2, 4, 6]):
        shapes.append(_text(10 + i, left=left_in * _EMU_PER_INCH, top=band_top))

    slide = _slide(4, shapes)
    assert is_visually_complex(slide) is True


def test_ac2_동일_left에_쌓인_텍스트는_멀티컬럼_아님() -> None:
    """AC2 경계: 같은 top 밴드라도 left 가 동일(세로 스택)하면 멀티컬럼 신호 없음."""
    band_top = 1 * _EMU_PER_INCH
    shapes: list[ShapeIR] = [_text(1, is_title=True, top=0)]
    for i in range(4):
        shapes.append(_text(20 + i, left=0, top=band_top))

    slide = _slide(5, shapes)
    # 멀티컬럼 신호는 없으므로 다른 신호(커넥터/차트/표/그룹)가 없는 한 False
    assert is_visually_complex(slide) is False


# ---------------------------------------------------------------------------
# AC3 — 커넥터/흐름·차트/SmartArt -> True
# ---------------------------------------------------------------------------


def test_ac3_커넥터_임계_이상_슬라이드는_True() -> None:
    """AC3: CONNECTOR_COUNT_THRESHOLD 개 이상 커넥터(LINE) -> True (S7 흐름 재현)."""
    shapes: list[ShapeIR] = [_text(1, is_title=True)]
    shapes.extend(_connector(100 + i) for i in range(CONNECTOR_COUNT_THRESHOLD))

    slide = _slide(7, shapes)
    assert is_visually_complex(slide) is True


def test_ac3_커넥터가_임계_미만이면_그_신호만으로는_복잡하지_않음() -> None:
    """AC3 경계: 커넥터 수가 임계(N) 미만이면 커넥터 신호는 발동하지 않는다."""
    shapes: list[ShapeIR] = [_text(1, is_title=True)]
    shapes.extend(_connector(200 + i) for i in range(CONNECTOR_COUNT_THRESHOLD - 1))

    slide = _slide(70, shapes)
    assert is_visually_complex(slide) is False


def test_ac3_차트_1개_존재만으로_True() -> None:
    """AC3: 차트(GraphicFrame) 1개만 있어도 True (S12 KPI·도넛 재현)."""
    shapes: list[ShapeIR] = [_text(1, is_title=True), _chart()]

    slide = _slide(12, shapes)
    assert is_visually_complex(slide) is True


def test_ac3_smartart_1개_존재만으로_True() -> None:
    """AC3: SmartArt(DIAGRAM GraphicFrame) 1개만 있어도 True."""
    shapes: list[ShapeIR] = [_text(1, is_title=True), _smartart()]

    slide = _slide(13, shapes)
    assert is_visually_complex(slide) is True


def test_ac3_mso_shape_type_상수_구성() -> None:
    """AC3: 커넥터/차트 판별에 쓰이는 명명 상수 집합이 기대한 라벨을 포함한다."""
    assert "LINE" in CONNECTOR_MSO_SHAPE_TYPES
    assert "CHART" in CHART_SMARTART_MSO_SHAPE_TYPES
    assert "DIAGRAM" in CHART_SMARTART_MSO_SHAPE_TYPES


# ---------------------------------------------------------------------------
# AC4 — 순수 텍스트(제목 + 단일 텍스트프레임 선형 불릿) -> False
# ---------------------------------------------------------------------------


def test_ac4_제목과_단일_텍스트프레임_선형_불릿은_False() -> None:
    """AC4 (정상-부정): 제목 + 단일 텍스트프레임 선형 불릿 -> 텍스트 경로 유지."""
    title = _text(1, is_title=True, top=0, text="Title")
    body = TextShapeIR(
        shape_id=2,
        name="Body",
        kind=ShapeKind.TEXT,
        top=_EMU_PER_INCH,
        paragraphs=[
            ParagraphIR(text="Bullet one", level=0),
            ParagraphIR(text="Bullet two", level=0),
            ParagraphIR(text="Sub-bullet", level=1),
        ],
    )
    slide = _slide(0, [title, body])

    assert is_visually_complex(slide) is False
    assert complexity_score(slide) < COMPLEXITY_THRESHOLD


# ---------------------------------------------------------------------------
# AC5 — 빈 슬라이드(도형 0개) -> False
# ---------------------------------------------------------------------------


def test_ac5_빈_슬라이드는_False() -> None:
    """AC5 (경계): 도형 0개인 슬라이드는 항상 False."""
    slide = _slide(0, [])
    assert is_visually_complex(slide) is False
    assert complexity_score(slide) == 0


# ---------------------------------------------------------------------------
# AC6 — 명명 상수 노출 + 튜닝 가능성 + 재현율 편향(경계 시 복잡=True)
# ---------------------------------------------------------------------------


def test_ac6_임계값이_명명_상수로_노출된다() -> None:
    """AC6: 판정 임계값은 명명 상수(COMPLEXITY_THRESHOLD)로 모듈에 노출된다."""
    assert isinstance(COMPLEXITY_THRESHOLD, int)
    assert COMPLEXITY_THRESHOLD > 0


def test_ac6_점수가_정확히_임계값이면_복잡으로_판정된다() -> None:
    """AC6: 재현율 편향 — score == COMPLEXITY_THRESHOLD 경계는 True (과탐 허용)."""
    # score == W_CHART_SMARTART == COMPLEXITY_THRESHOLD (경계값)
    shapes: list[ShapeIR] = [_text(1, is_title=True), _chart()]
    slide = _slide(99, shapes)

    assert complexity_score(slide) == COMPLEXITY_THRESHOLD
    assert is_visually_complex(slide) is True


# ---------------------------------------------------------------------------
# 표/그룹 구조 신호 (§3.2 신호표 — 표/그룹 signal 보완 커버리지)
# ---------------------------------------------------------------------------


def test_table_and_group_signal_각각은_임계_미만이지만_결합시_복잡() -> None:
    """표 신호 + 그룹 신호가 결합되면 임계에 도달해 True 로 판정된다."""
    table = TableShapeIR(
        shape_id=2,
        name="Table 1",
        kind=ShapeKind.TABLE,
        rows=[["A", "B"], ["C", "D"]],
        n_rows=2,
        n_cols=2,
    )
    group = GroupShapeIR(shape_id=3, name="Group 1", kind=ShapeKind.GROUP, children=[])

    slide_table_only = _slide(1, [_text(1, is_title=True), table])
    slide_both = _slide(2, [_text(1, is_title=True), table, group])

    assert is_visually_complex(slide_table_only) is False
    assert is_visually_complex(slide_both) is True


# ---------------------------------------------------------------------------
# 도형 수 신호 (§3.2 신호표 보완 커버리지) — 단독으로는 임계 미달, 점수만 반영
# ---------------------------------------------------------------------------


def test_shape_count_초과시_점수에_반영되지만_단독으로는_임계_미달() -> None:
    """SHAPE_COUNT_THRESHOLD 초과 시 W_SHAPE_COUNT_OVER 만큼 점수가 오르되,
    다른 신호가 없으면 여전히 False (단독 신호로는 부족)."""
    shapes: list[ShapeIR] = [_text(1, is_title=True, top=0)]
    # 서로 다른 top 밴드(세로 스택, 동일 left)에 8개 배치 -> 멀티컬럼 신호 없음
    for i in range(8):
        shapes.append(_text(10 + i, left=0, top=(i + 1) * 2 * _EMU_PER_INCH))

    slide = _slide(50, shapes)
    assert len(shapes) > 8
    assert complexity_score(slide) == 1
    assert is_visually_complex(slide) is False


# ---------------------------------------------------------------------------
# 멀티컬럼 밴드 클러스터링 — 여러 밴드 중 앞쪽 밴드에서 조기 반환되는 경로 커버
# ---------------------------------------------------------------------------


def test_멀티컬럼_밴드가_여러개일때_앞쪽_밴드에서_조기_판정된다() -> None:
    """첫 번째 top 밴드가 멀티컬럼이면, 나머지 밴드 순회 전에 True 로 조기 판정된다."""
    band1_top = 1 * _EMU_PER_INCH
    shapes: list[ShapeIR] = [_text(1, is_title=True, top=0)]
    for i, left_in in enumerate([0, 2, 4]):
        shapes.append(_text(20 + i, left=left_in * _EMU_PER_INCH, top=band1_top))
    # 두 번째(별개) 밴드 — 첫 밴드보다 한참 아래(밴드 종료 트리거)
    shapes.append(_text(30, left=0, top=5 * _EMU_PER_INCH))

    slide = _slide(51, shapes)
    assert is_visually_complex(slide) is True


# ---------------------------------------------------------------------------
# AC5(FR-31 결정성 표현, INV-1) — 동일 IR 2회 입력 -> 동일 결과
# ---------------------------------------------------------------------------


def test_ac5_동일_ir_2회_입력은_동일_결과_결정성() -> None:
    """결정성(INV-1): 동일 SlideIR 값을 2회 판정해도 bool/score 가 동일하다."""
    shapes: list[ShapeIR] = [
        _text(1, is_title=True),
        _connector(2),
        _connector(3),
        _chart(),
    ]
    slide = _slide(3, shapes)

    score_1 = complexity_score(slide)
    score_2 = complexity_score(slide)
    bool_1 = is_visually_complex(slide)
    bool_2 = is_visually_complex(slide)

    assert score_1 == score_2
    assert bool_1 == bool_2


def test_ac5_동일_구조의_별개_slideir_인스턴스도_동일_결과() -> None:
    """결정성(INV-1): 동일 값을 갖는 별개의 SlideIR 인스턴스도 동일 판정을 낸다."""

    def _build() -> SlideIR:
        return _slide(
            8,
            [
                _text(1, is_title=True),
                _text(2, left=0, top=_EMU_PER_INCH),
                _text(3, left=2 * _EMU_PER_INCH, top=_EMU_PER_INCH),
                _text(4, left=4 * _EMU_PER_INCH, top=_EMU_PER_INCH),
            ],
        )

    slide_a = _build()
    slide_b = _build()

    assert complexity_score(slide_a) == complexity_score(slide_b)
    assert is_visually_complex(slide_a) == is_visually_complex(slide_b)
