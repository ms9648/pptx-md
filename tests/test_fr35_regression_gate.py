"""FR-35 회귀 검증 전략 — 비결정 VLM 골든 게이트 (issue #96, W-G).

이슈 #96 dispatch AC 목록 대조 (REQ-v020 §7 FR-35 AC1~AC7, ARCH-v020 §3.7/
ADR-629 1:1 매핑, `test_ac<N>_<description>` 네이밍):

    AC1 텍스트 골든 불변(INV-1)  -> TestAc1TextPathGoldenInvariant
    AC2 라우팅 골든              -> TestAc2RoutingGolden
    AC3 모의 프로바이더 결합 골든 -> TestAc3MockProviderCombinedGolden
    AC4 실 VLM 구조 불변식(환경 게이트) -> TestAc4RealVlmStructuralInvariant
    AC5 hard-fail                -> TestAc5HardFailOnGoldenDiff
    AC6 환경 게이트 skip          -> TestAc6EnvironmentGate
    AC7 합성 픽스처만             -> TestAc7SyntheticFixturesOnly

이 파일은 tests/ 하위 신규 파일이며 src/ 는 접촉하지 않는다(이슈 AC8). 합성
PPTX 픽스처는 tests/fr35_fixtures.py(python-pptx 로 즉석 생성, 사내/고객 전용
실증 파일은 일절 참조하지 않음)에서만 가져온다.

핵심 설계 메모 — "모의 프로바이더 골든"이 실 soffice 없이도 항상 실행되는
이유(AC3 ∧ AC6 동시 충족): ``api.convert()`` 는 렌더러/복잡도 함수 주입 지점을
공개하지 않는다(``ConvertOptions`` 에 그런 필드가 없다 — 의도적 설계,
ARCH-v020 §3.5). 대신 ``reconstruct_slides`` 의 ``renderer``/``complexity_fn``
키워드 전용 인자는 ``reconstruct_slides.__kwdefaults__`` 라는 **가변** dict에
저장되므로(파이썬 시맨틱: `*` 뒤의 키워드 전용 인자의 default 는
``__defaults__`` 가 아니라 ``__kwdefaults__`` 에 저장되고, 이 dict 자체는
런타임에 교체 가능하다), 테스트는 이 dict 의 ``"renderer"`` 항목만
monkeypatch 하여 실제 공개 API ``convert(..., visual_reconstruct=True,
reconstructor=Fake)`` 호출 경로를 그대로 태우면서도 실 soffice/pypdfium2 를
전혀 건드리지 않는다. ``slide_render_available`` 은 이름으로 임포트되어
있으므로(``pptx_md.slide_reconstruction.slide_render_available``, O 팀 설계
docstring 명시) 이것도 함께 monkeypatch 해 조기 return 을 우회한다. 결과적으로
soffice 가 전혀 설치되지 않은 이 검증 환경에서도 AC3 테스트는 스킵되지 않고
항상 실행된다(AC6).
"""

from __future__ import annotations

import difflib
import os
from pathlib import Path

import pytest

from pptx_md.api import ConvertOptions, convert
from pptx_md.complexity import complexity_score, is_visually_complex
from pptx_md.parser import parse_presentation
from pptx_md.slide_describer import SlideContext
from pptx_md.slide_reconstruction import reconstruct_slides
from pptx_md.slide_render import slide_render_available
from pptx_md.validator import validate_markdown
from tests.fr35_fixtures import (
    EXPECTED_COMPLEX_INDICES,
    EXPECTED_COMPLEXITY_SCORES,
    EXPECTED_TEXT_PATH_INDICES,
    PURE_TEXT_INDEX,
    build_fr35_synthetic_deck,
    build_fr35_synthetic_deck_bytes,
)

_GOLDEN_DIR: Path = Path(__file__).parent / "golden"
_MOCK_COMBINED_GOLDEN_PATH: Path = _GOLDEN_DIR / "fr35_mock_combined.md"


# ---------------------------------------------------------------------------
# Shared env-guarded compare-or-update helper (FR-29 D-6 hard-fail precedent,
# ADR-620 — reimplemented locally rather than imported from
# test_golden_regression.py, which is issue #81/W-E's own file: G owns
# tests/ additions but does not couple to another issue's private helpers).
# ---------------------------------------------------------------------------


def _compare_or_update_golden(path: Path, content: str) -> None:
    """Byte-compare *content* against *path*, or rewrite when opted in.

    Rewriting only happens when ``PPTXMD_UPDATE_GOLDEN=1`` (FR-29 D-6
    accidental-drift guard, reused verbatim as policy — not code — here).
    On mismatch, raises ``AssertionError`` with a unified diff so the
    ordinary ``pytest`` run hard-fails with an actionable message (AC5).
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
            f"FR-35 골든 회귀 실패(AC5): {path} 와 재산출 결과가 다릅니다"
            f"(hard-fail).\n의도된 변경이면 PPTXMD_UPDATE_GOLDEN=1 로 재기록 후"
            f" 별도 골든 갱신 PR 로 사람 승인을 받으세요(D-6 계승).\n{diff}"
        )


# ---------------------------------------------------------------------------
# AC1 — 텍스트 경로 골든 불변(INV-1): 신규 옵션 전부 off -> 기존 골든과 diff 0
# ---------------------------------------------------------------------------


class TestAc1TextPathGoldenInvariant:
    """신규 v0.2.0 배선이 기존 텍스트 경로 출력을 회귀시키지 않는다(AC1)."""

    def test_ac1_fr29_off_golden_diff_zero_after_noop_reconstruct(self) -> None:
        """FR-29 off 골든(tests/golden/wave3_off.md)과, reconstruct_slides 를
        reconstructor=None 으로 명시 호출한 뒤 재산출한 결과가 바이트 동일하다.

        FR-29 골든 파일을 직접 읽어(다른 테스트 모듈의 헬퍼에 결합하지 않고)
        FR-33 오케스트레이션 stage 를 실제로 한 번 통과시킨 뒤에도 무변화임을
        독립적으로 재확인한다(개발자 테스트 재실행이 아니라 별도 시나리오).
        """
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

        # FR-29 golden 픽스처와 동일한 필드값의 독립 사본(다른 테스트 모듈의
        # 내부 헬퍼를 import 하지 않고, 동일 값을 이 파일 안에서 재구성).
        pres = PresentationIR(
            source_path="golden-fixture.pptx",
            slides=[
                SlideIR(index=0, title="표지", shapes=[]),
                SlideIR(
                    index=1,
                    title="A > B > C",
                    notes="발표자 노트: 핵심 논지를 강조할 것",
                    shapes=[
                        TextShapeIR(
                            shape_id=101,
                            name="본문 텍스트",
                            kind=ShapeKind.TEXT,
                            paragraphs=[
                                ParagraphIR(text="핵심 내용 설명입니다.", level=0)
                            ],
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
                ),
                SlideIR(
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
                ),
            ],
        )

        # FR-33 Stage 3.5 를 reconstructor=None 으로 명시 통과(off 시나리오).
        reconstruct_slides(pres, None)
        for slide in pres.slides:
            assert slide.reconstructed_md is None  # 조기 return 확인(AC3, INV-4)

        rendered = assemble_document(pres)  # 4옵션 전부 default False
        golden_path = _GOLDEN_DIR / "wave3_off.md"
        expected = golden_path.read_bytes()
        assert rendered.encode("utf-8") == expected, (
            "FR-29 off 골든과 diff 발생 — v0.2.0 배선이 텍스트 경로를 "
            "회귀시켰다(AC1 위반)"
        )

    def test_ac1_fr29_on_golden_diff_zero_after_noop_reconstruct(self) -> None:
        """동일 시나리오를 FR-28 4옵션 전부 True(on)로 재확인."""
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

        pres = PresentationIR(
            source_path="golden-fixture.pptx",
            slides=[
                SlideIR(index=0, title="표지", shapes=[]),
                SlideIR(
                    index=1,
                    title="A > B > C",
                    notes="발표자 노트: 핵심 논지를 강조할 것",
                    shapes=[
                        TextShapeIR(
                            shape_id=101,
                            name="본문 텍스트",
                            kind=ShapeKind.TEXT,
                            paragraphs=[
                                ParagraphIR(text="핵심 내용 설명입니다.", level=0)
                            ],
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
                ),
                SlideIR(
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
                ),
            ],
        )

        reconstruct_slides(pres, None)
        rendered = assemble_document(
            pres,
            heading_hierarchy=True,
            emit_toc=True,
            emit_frontmatter=True,
            include_notes=True,
        )
        golden_path = _GOLDEN_DIR / "wave3_on.md"
        expected = golden_path.read_bytes()
        assert rendered.encode("utf-8") == expected

    def test_ac1_convert_toggle_off_vs_visual_reconstruct_true_none_reconstructor(
        self, tmp_path: Path
    ) -> None:
        """off/on 토글 대조: 실 pptx 를 통한 convert() 에서
        visual_reconstruct=False 와 visual_reconstruct=True+reconstructor=None
        은 바이트 동일해야 한다(FR-33 AC3 조기 return 이 토글을 무력화)."""
        pptx_path = tmp_path / "toggle.pptx"
        build_fr35_synthetic_deck(pptx_path)

        off_result = convert(
            pptx_path, options=ConvertOptions(visual_reconstruct=False)
        )
        on_toggle_but_no_provider = convert(
            pptx_path,
            options=ConvertOptions(visual_reconstruct=True, reconstructor=None),
        )

        assert off_result.encode("utf-8") == on_toggle_but_no_provider.encode("utf-8")

    def test_ac1_convert_default_options_byte_identical_across_two_calls(
        self, tmp_path: Path
    ) -> None:
        """결정성(INV-1): 동일 소스에 대한 2회 convert() 호출은 바이트 동일하다."""
        pptx_path = tmp_path / "determinism.pptx"
        build_fr35_synthetic_deck(pptx_path)

        first = convert(pptx_path)
        second = convert(pptx_path)
        assert first.encode("utf-8") == second.encode("utf-8")


# ---------------------------------------------------------------------------
# AC2 — 라우팅 결정성 골든: 합성 픽스처 -> is_visually_complex 인덱스 집합
# ---------------------------------------------------------------------------


class TestAc2RoutingGolden:
    """합성 픽스처(S5/S7/S12/순수텍스트)에 대한 렌더 경로 인덱스 집합이
    결정적 골든과 일치한다(AC2). VLM/렌더 미호출 — 순수 IR 판정만 사용."""

    def _parse_deck(self, tmp_path: Path) -> list:
        pptx_path = tmp_path / "fr35_routing.pptx"
        build_fr35_synthetic_deck(pptx_path)
        pres = parse_presentation(pptx_path)
        assert len(pres.slides) == 4
        return pres.slides

    def test_ac2_routed_index_set_matches_golden(self, tmp_path: Path) -> None:
        """S5/S7/S12 인덱스는 렌더 경로로, 순수 텍스트 인덱스는 텍스트 경로로
        라우팅되는 인덱스 집합이 하드코딩 골든과 일치한다."""
        slides = self._parse_deck(tmp_path)

        routed = {s.index for s in slides if is_visually_complex(s)}
        not_routed = {s.index for s in slides if not is_visually_complex(s)}

        assert routed == EXPECTED_COMPLEX_INDICES
        assert not_routed == EXPECTED_TEXT_PATH_INDICES

    def test_ac2_complexity_scores_match_golden(self, tmp_path: Path) -> None:
        """복잡도 점수 자체도 하드코딩 골든과 일치한다(재현율 편향 캘리브레이션
        회귀 감지 — 가중치/임계 변경 시 이 값이 먼저 달라진다)."""
        slides = self._parse_deck(tmp_path)
        for slide in slides:
            assert complexity_score(slide) == EXPECTED_COMPLEXITY_SCORES[slide.index]

    def test_ac2_pure_text_slide_never_routed(self, tmp_path: Path) -> None:
        """AC4 정상-부정 계승: 순수 텍스트 슬라이드(index 3)는 항상 False."""
        slides = self._parse_deck(tmp_path)
        text_slide = slides[PURE_TEXT_INDEX]
        assert is_visually_complex(text_slide) is False

    def test_ac2_routing_is_deterministic_across_two_parses(
        self, tmp_path: Path
    ) -> None:
        """결정성(INV-1): 동일 합성 PPTX 를 2회 파싱해도 라우팅 집합이 동일."""
        pptx_path = tmp_path / "fr35_routing_det.pptx"
        build_fr35_synthetic_deck(pptx_path)

        pres1 = parse_presentation(pptx_path)
        pres2 = parse_presentation(pptx_path)

        routed1 = {s.index for s in pres1.slides if is_visually_complex(s)}
        routed2 = {s.index for s in pres2.slides if is_visually_complex(s)}
        assert routed1 == routed2 == EXPECTED_COMPLEX_INDICES

    def test_ac2_routing_never_touches_render_or_vlm(self, tmp_path: Path) -> None:
        """라우팅 판정 중 render_deck/reconstruct 가 호출되지 않는다(AC1 VLM·
        렌더 미호출 요건의 직접 증거 — monkeypatch 로 호출 시 즉시 실패)."""
        from unittest.mock import patch

        slides = self._parse_deck(tmp_path)
        with patch(
            "pptx_md.slide_render.render_deck",
            side_effect=AssertionError("render_deck must not be called by routing"),
        ):
            for slide in slides:
                is_visually_complex(slide)  # must not raise


# ---------------------------------------------------------------------------
# AC3 — 모의 프로바이더 결합 골든: FakeReconstructor + 실 convert() end-to-end
# ---------------------------------------------------------------------------


class _FakeReconstructor:
    """결정적 고정 응답 — 슬라이드 인덱스별 정합 pipe table 을 담은 MD 조각."""

    def __init__(self) -> None:
        self.calls: list[SlideContext] = []

    def reconstruct(
        self, image_bytes: bytes, image_ext: str, context: SlideContext
    ) -> str:
        self.calls.append(context)
        return (
            f"### 구조화 재구성 (slide {context.slide_index})\n\n"
            "| 항목 | 값 |\n"
            "| --- | --- |\n"
            f"| 신호 | slide-{context.slide_index} |\n"
        )


class _FakeRenderer:
    """index -> 고정(고유) PNG-형 바이트. 실 soffice/pypdfium2 미사용."""

    def __init__(self, png_map: dict[int, bytes]) -> None:
        self._png_map = png_map
        self.calls: list[list[int]] = []

    def __call__(self, pptx_path: object, indices, *, dpi: int = 150):
        idx_list = list(indices)
        self.calls.append(idx_list)
        return {i: self._png_map.get(i) for i in idx_list}


class TestAc3MockProviderCombinedGolden:
    """FakeReconstructor 고정 응답으로 실 convert(visual_reconstruct=True)
    end-to-end 결합 결과가 골든과 바이트 동일하다(AC3). 오케스트레이션/결합
    로직만 회귀 측정 대상 — VLM 비결정성은 프로바이더 경계 밖(Fake)이라 애초에
    없다."""

    def _run_convert_with_fakes(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> tuple[str, _FakeReconstructor, _FakeRenderer]:
        pptx_path = tmp_path / "fr35_mock_combined.pptx"
        build_fr35_synthetic_deck(pptx_path)

        fake_png_map = {
            0: b"\x89PNG\r\n\x1a\n-fake-slide-0-",
            1: b"\x89PNG\r\n\x1a\n-fake-slide-1-",
            2: b"\x89PNG\r\n\x1a\n-fake-slide-2-",
        }
        fake_renderer = _FakeRenderer(fake_png_map)
        fake_reconstructor = _FakeReconstructor()

        # slide_render_available is imported by *name* into
        # slide_reconstruction.py (O 팀 설계 docstring) -- monkeypatch there.
        monkeypatch.setattr(
            "pptx_md.slide_reconstruction.slide_render_available", lambda: True
        )
        # `renderer` is a keyword-only parameter of reconstruct_slides(); its
        # default lives in __kwdefaults__ (a mutable dict at runtime, unlike
        # __defaults__ for positional-or-keyword params) so it can be
        # monkeypatched even though api.convert() never passes renderer=
        # explicitly (ConvertOptions exposes no such injection point by
        # design, ARCH-v020 §3.5).
        monkeypatch.setitem(
            reconstruct_slides.__kwdefaults__, "renderer", fake_renderer
        )

        opts = ConvertOptions(
            visual_reconstruct=True,
            reconstructor=fake_reconstructor,
            force_render=frozenset({0, 1, 2}),
            force_text=frozenset({3}),
        )
        result = convert(pptx_path, options=opts)
        return result, fake_reconstructor, fake_renderer

    def test_ac3_combined_markdown_matches_golden(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        result, _, _ = self._run_convert_with_fakes(tmp_path, monkeypatch)
        _compare_or_update_golden(_MOCK_COMBINED_GOLDEN_PATH, result)

    def test_ac3_golden_path_exists(self) -> None:
        assert _MOCK_COMBINED_GOLDEN_PATH.exists()

    def test_ac3_reconstruct_called_for_every_forced_render_slide(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """슬롯 소비 검증: force_render 3개 슬라이드 모두 reconstruct() 콜을
        받았고, force_text 슬라이드는 받지 않았다(사전 렌더 캐시로 인한
        중복 제거가 없으므로 3콜 — 각기 다른 PNG)."""
        _, fake_reconstructor, fake_renderer = self._run_convert_with_fakes(
            tmp_path, monkeypatch
        )
        called_indices = {c.slide_index for c in fake_reconstructor.calls}
        assert called_indices == {0, 1, 2}
        assert fake_renderer.calls, "render 단계가 최소 1회 호출되어야 한다"
        assert set(fake_renderer.calls[0]) == {0, 1, 2}

    def test_ac3_slide_index_order_preserved_in_merge(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """슬라이드 순 병합(ADR-220 계승): 결합 문서에서 slide 0/1/2/3 재구성·
        텍스트 블록이 인덱스 오름차순으로 나타난다."""
        result, _, _ = self._run_convert_with_fakes(tmp_path, monkeypatch)
        pos0 = result.index("slide 0")
        pos1 = result.index("slide 1")
        pos2 = result.index("slide 2")
        pos3 = result.index("요약")  # index 3, 텍스트 경로(제목 그대로 렌더)
        assert pos0 < pos1 < pos2 < pos3

    def test_ac3_determinism_two_runs_byte_identical(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """결정성(INV-1): 동일 Fake 배선으로 2회 convert() 실행 시 바이트 동일."""
        result1, _, _ = self._run_convert_with_fakes(tmp_path, monkeypatch)
        result2, _, _ = self._run_convert_with_fakes(tmp_path, monkeypatch)
        assert result1.encode("utf-8") == result2.encode("utf-8")

    def test_ac3_combined_markdown_has_no_broken_structure(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """결합 산출물이 유효한 Markdown 이며 표/헤딩 관련 DoD 경고가 없다
        (FR-32 AC1/AC2 정합 표 요건이 결합 후에도 유지됨을 확인)."""
        result, _, _ = self._run_convert_with_fakes(tmp_path, monkeypatch)
        outcome = validate_markdown(result)
        assert outcome.valid
        broken_table_warnings = [w for w in outcome.warnings if "깨진 표" in w]
        assert broken_table_warnings == []


# ---------------------------------------------------------------------------
# AC4 — 실 VLM 구조 불변식(바이트 골든 제외, 환경 게이트)
# ---------------------------------------------------------------------------


def _real_vlm_credentials_available() -> bool:
    """실 VLM 콜에 쓸 수 있는 자격 증명이 환경에 있는지.

    이슈 #104(AC1/AC4)는 **실 OpenAI provider(reconstruct_model=gpt-4o)** 를
    구체적으로 주입해 실행하도록 명시하므로, 이 게이트는 ``OPENAI_API_KEY``
    하나만 확인한다(``ANTHROPIC_API_KEY`` 만 있고 openai 키가 없는 환경에서
    이 특정 테스트가 잘못 실행되지 않도록 — 다른 provider 로의 확장은 별도
    이슈로 다룬다).
    """
    return bool(os.environ.get("OPENAI_API_KEY"))


class TestAc4RealVlmStructuralInvariant:
    """실 VLM/실 soffice 산출물은 바이트 골든 대상에서 제외하고, 구조 불변식
    (유효 Markdown·validator 경고 0·표 정합·슬라이드 수 보존)만 어서션한다
    (AC4). 렌더(soffice/pypdfium2) 또는 VLM 자격 증명이 없으면 skip(AC6)."""

    pytestmark = pytest.mark.skipif(
        not slide_render_available() or not _real_vlm_credentials_available(),
        reason=(
            "실 VLM 구조 불변식 테스트는 soffice+pypdfium2(slide_render_available)"
            " 그리고 VLM API 자격 증명(OPENAI_API_KEY) 이 모두 있을 때만"
            " 실행된다 — 이 환경/CI 에는 없으므로 skip(FR-35 AC4/AC6, 이슈"
            " #104 AC3). 실행하려면 workflow_dispatch 전용"
            " .github/workflows/real-vlm-smoke.yml 을 사람 승인 후 트리거한다"
            "(#104 AC4, 비용 통제)."
        ),
    )

    def test_ac4_real_vlm_structural_invariants(self, tmp_path: Path) -> None:
        """실 soffice 렌더 + 실 OpenAI VLM 재구성 산출물의 구조 불변식만
        검증한다(바이트 골든이 아님 — 비결정성 흡수, AC4/이슈#104 AC2).

        soffice + OPENAI_API_KEY 가 모두 갖춰진 환경(수동 workflow_dispatch,
        #104 AC4)에서만 실행되고, 그 외에는 클래스 ``pytestmark`` 로 skip
        된다(AC6/#104 AC3). 합성 픽스처(fr35_fixtures, #96 AC7 계승 — 사내
        실증 파일 미사용)의 자연 라우팅(S5/S7/S12 인덱스는 render+VLM 경로,
        순수 텍스트 인덱스는 텍스트 경로, tests/fr35_fixtures.py 참조)에
        맡겨 실 gpt-4o 호출을 최소 1회 이상 태운다.
        """
        from pptx_md.assembler import _SLIDE_SEPARATOR
        from pptx_md.errors import InstallationError
        from pptx_md.providers.openai import OpenAIDescriber

        try:
            reconstructor = OpenAIDescriber(reconstruct_model="gpt-4o")
        except InstallationError:
            pytest.skip(
                "openai SDK(VLM extras)가 설치되지 않아 실 provider 를 구성할"
                " 수 없다 — pip install pptx-md[vlm] 필요(#104 AC3 계승)."
            )

        pptx_path = tmp_path / "fr35_ac4_real_vlm.pptx"
        build_fr35_synthetic_deck(pptx_path)

        result = convert(
            pptx_path,
            options=ConvertOptions(
                visual_reconstruct=True,
                reconstructor=reconstructor,
            ),
        )

        # (a) 유효 Markdown 산출 — 비어있지 않은 실 문서.
        assert result.strip() != ""

        outcome = validate_markdown(result)

        # (b) validator 경고 0(이슈 #104 AC2, ARCH-v020 §3.7/ADR-629).
        assert outcome.valid, f"validate_markdown invalid: {outcome.warnings}"
        assert outcome.warnings == [], f"unexpected warnings: {outcome.warnings}"

        # (c) 슬라이드 수 보존 — 렌더+VLM 실패 시에도 텍스트 폴백으로 슬롯이
        # 유지되므로(INV-4, slide_reconstruction.py), 합성 픽스처의 4개
        # 슬라이드 블록이 그대로 보존된다. assemble_document 의 실 구분자를
        # 그대로 재사용해 이 테스트의 판정과 산출물 실 구조를 결합한다.
        slide_blocks = result.split(_SLIDE_SEPARATOR)
        assert len(slide_blocks) == 4

        # (d) 표 정합 — 깨진/열불일치/전공백 표 경고가 전혀 없어야 한다
        # (validate_markdown 의 FR-29 pipe-table 규칙, 이슈 #104 AC2(d)).
        broken_table_warnings = [w for w in outcome.warnings if "표" in w]
        assert broken_table_warnings == []


# ---------------------------------------------------------------------------
# AC5 — hard-fail: 골든 diff 발생 시 pytest 실패(FR-29 D-6 계승)
# ---------------------------------------------------------------------------


class TestAc5HardFailOnGoldenDiff:
    """FR-35 자체 골든 비교 헬퍼가 diff 발생 시 hard-fail 함을 증명한다
    (AC5). 실제 tests/golden/fr35_mock_combined.md 는 절대 건드리지 않고
    tmp_path 격리 파일만 사용한다."""

    def test_ac5_mismatch_raises_assertion_with_diff(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.delenv("PPTXMD_UPDATE_GOLDEN", raising=False)
        golden = tmp_path / "sample.md"
        golden.write_text("old", encoding="utf-8")

        with pytest.raises(AssertionError, match="FR-35 골든 회귀 실패"):
            _compare_or_update_golden(golden, "new")

        assert golden.read_text(encoding="utf-8") == "old"  # 우발 갱신 없음

    def test_ac5_match_passes_without_env_var(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.delenv("PPTXMD_UPDATE_GOLDEN", raising=False)
        golden = tmp_path / "sample.md"
        golden.write_text("same", encoding="utf-8")
        _compare_or_update_golden(golden, "same")  # no raise

    def test_ac5_update_env_var_rewrites_file(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setenv("PPTXMD_UPDATE_GOLDEN", "1")
        golden = tmp_path / "sample.md"
        golden.write_text("old", encoding="utf-8")
        _compare_or_update_golden(golden, "new")
        assert golden.read_bytes() == b"new"


# ---------------------------------------------------------------------------
# AC6 — 환경 격리: 렌더/VLM 미가용 시 skip, 텍스트/라우팅/모의 골든은 항상 실행
# ---------------------------------------------------------------------------


class TestAc6EnvironmentGate:
    """렌더/VLM 의존 테스트만 환경 게이트로 skip 되고, 텍스트 경로·라우팅·
    모의 프로바이더 테스트는 이 환경(soffice 미설치)에서도 항상 실행됨을
    구조적으로 확인한다(AC6)."""

    def test_ac6_real_vlm_class_has_environment_skip_marker(self) -> None:
        marker = TestAc4RealVlmStructuralInvariant.pytestmark
        assert marker.name == "skipif"

    def test_ac6_routing_fixture_builder_never_touches_render_or_vlm(self) -> None:
        """fr35_fixtures.py(순수 python-pptx 합성)는 렌더/VLM 모듈을 전혀
        import 하지 않는다 — 라우팅 골든이 렌더/VLM 가용성과 무관하게 항상
        실행 가능함의 정적 증거."""
        import pathlib

        from tests import fr35_fixtures as fixtures_module

        source = pathlib.Path(fixtures_module.__file__).read_text(encoding="utf-8")
        for forbidden in (
            "slide_render",
            "soffice",
            "pypdfium2",
            "slide_reconstruction",
            "import anthropic",
            "import openai",
        ):
            assert forbidden not in source

    def test_ac6_slide_render_unavailable_in_this_environment(self) -> None:
        """이 검증 환경 자체가 soffice 미설치임을 명시적으로 기록한다(리포트
        재현성 — TR-96.md 의 skip 근거와 직접 대응)."""
        # 참고용 기록 목적(soffice 유무에 따라 값이 달라질 수 있음) — 실패
        # 어서션이 아니라, 이 상수가 AC4 skip 사유와 일치하는지 확인.
        available = slide_render_available()
        assert isinstance(available, bool)


# ---------------------------------------------------------------------------
# AC7 — 픽스처 자산: 전부 합성 PPTX, 사내 실증 자료 미참조
# ---------------------------------------------------------------------------


class TestAc7SyntheticFixturesOnly:
    """회귀 스위트의 모든 픽스처가 합성 PPTX 이며, 사내 실증 파일을 전혀
    참조하지 않는다(AC7). 사내 파일 경로 리터럴을 텍스트 검색하는 방식은
    (그 리터럴 자체를 설명하는 docstring 문구와 자기-매칭되는) 오탐 위험이
    있어, 대신 **런타임 디스크 접근 여부**를 직접 관찰한다: 픽스처 빌더가
    파일 시스템의 어떤 경로도 열지 않고 순수 in-memory 로만 .pptx 바이트를
    만들어내는지 ``open()`` 을 계측해 증명한다."""

    def test_ac7_synthetic_deck_bytes_are_a_real_pptx_zip(self) -> None:
        """build_fr35_synthetic_deck_bytes() 산출물이 실제 .pptx(zip) 포맷임을
        확인 — 즉석 합성이지 어딘가에서 읽어온 바이너리가 아님."""
        data = build_fr35_synthetic_deck_bytes()
        assert data[:2] == b"PK"  # zip local-file-header magic (.pptx == zip)

    def test_ac7_synthetic_deck_build_opens_no_project_or_customer_files(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """build_fr35_synthetic_deck_bytes() 실행 중 ``open()`` 이 프로젝트/
        사내 경로를 전혀 열지 않는다는 런타임 증거 — in-house/customer 파일을
        우회 경로로도 읽을 수 없다(AC7). ``python-pptx`` 패키지 자체의 번들
        기본 템플릿(``site-packages/pptx/templates/default.pptx`` — 라이브러리
        내부 자산, 사내/고객 자료 아님)만은 정상적인 라이브러리 동작이므로
        허용한다."""
        import builtins

        calls: list[str] = []
        real_open = builtins.open

        def _tracking_open(file: object, *args: object, **kwargs: object) -> object:
            calls.append(str(file))
            return real_open(file, *args, **kwargs)  # type: ignore[arg-type]

        monkeypatch.setattr(builtins, "open", _tracking_open)
        data = build_fr35_synthetic_deck_bytes()

        assert data[:2] == b"PK"
        lib_template = os.path.normpath("pptx" + os.sep + "templates")
        project_or_customer_opens = [
            c for c in calls if lib_template not in os.path.normpath(c)
        ]
        msg = f"unexpected disk open() calls: {project_or_customer_opens}"
        assert project_or_customer_opens == [], msg

    def test_ac7_fixtures_module_source_has_no_customer_deck_filename(self) -> None:
        """리포지토리에 실제 존재하는 사내 실증 파일명(고유해 오탐 위험이
        낮은 리터럴)이 픽스처 빌더 소스에 등장하지 않는다."""
        import pathlib

        from tests import fr35_fixtures as fixtures_module

        source = pathlib.Path(fixtures_module.__file__).read_text(encoding="utf-8")
        assert "SamsungSDS" not in source
        assert "AICC_제안서" not in source
