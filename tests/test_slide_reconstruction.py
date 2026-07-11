"""Tests for src/pptx_md/slide_reconstruction.py — FR-33 하이브리드 오케스트레이션
(이슈 #95, ARCH-v020 §3.4~3.6, ADR-627/628).

이슈 #95 WBS-O AC 목록 대조 (dispatch 본문 기준, 이하 test_ac<N>_* 로 매핑):
    AC1  라우팅(complexity)->캡->렌더(1회)->sha256 캐시->재구성->슬롯 충전, 결정적
    AC2  SlideIR.reconstructed_md 슬롯 default None (하위호환) -- tests/test_ir.py 및
         본 파일의 결정성 테스트에서 간접 커버
    AC3  ConvertOptions 신규 필드 -- tests/test_api_fr33.py
    AC4  INV-3 바이트 동일 -- tests/test_api_fr33.py
    AC5  reconstructor=None/렌더 미가용 -> 조기 return
    AC6  INV-4 격리(슬라이드 1장 실패 -> 텍스트 폴백, 예외 미전파)
    AC7  max_vlm_slides 캡(top-N)·sha256 캐시·embed_rendered_image 병행
    AC8  어셈블러 슬롯 소비 -- tests/test_assembler_fr33.py
    AC9  INV-5 격리 -- 본 파일(grep) + tests/test_assembler_fr33.py
    AC10 mypy/ruff/black exit 0 + 결정성/완료순서 무관성/캡/캐시/폴백 커버

FakeRenderer(index -> 고정 PNG bytes)와 FakeReconstructor(bytes -> 고정 MD)를 사용해
실 soffice/VLM 호출 없이 오케스트레이션 로직만 결정적으로 검증한다.

Test naming convention: test_ac<N>_<description>.
"""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import pytest

import pptx_md.slide_reconstruction as sr_module
from pptx_md.ir import (
    GroupShapeIR,
    ParagraphIR,
    PresentationIR,
    ShapeKind,
    SlideIR,
    TableShapeIR,
    TextShapeIR,
)
from pptx_md.slide_describer import SlideContext
from pptx_md.slide_reconstruction import reconstruct_slides

# ---------------------------------------------------------------------------
# Fakes (ARCH-v020 §7 test strategy: FakeRenderer / FakeReconstructor)
# ---------------------------------------------------------------------------


class FakeRenderer:
    """index -> fixed PNG bytes. Records every call (path, indices, dpi)."""

    def __init__(self, png_map: dict[int, bytes | None]) -> None:
        self._png_map = png_map
        self.calls: list[tuple[Any, list[int], int]] = []

    @property
    def call_count(self) -> int:
        return len(self.calls)

    def __call__(
        self, pptx_path: Any, indices: Any, *, dpi: int = 150
    ) -> dict[int, bytes | None]:
        idx_list = list(indices)
        self.calls.append((pptx_path, idx_list, dpi))
        return {i: self._png_map.get(i) for i in idx_list}


class FakeReconstructor:
    """bytes -> fixed markdown. Records every SlideContext it was called with.

    ``fail_bytes``: raises for these PNG byte values (INV-4 isolation test).
    ``delays``: optional per-bytes sleep (seconds) to shuffle completion
    order for the AC10 완료순서 무관성 test, without affecting the
    (deterministic) result mapping.
    """

    def __init__(
        self,
        *,
        responses: dict[bytes, str] | None = None,
        default: str = "## Reconstructed\n\nstructured body",
        fail_bytes: frozenset[bytes] = frozenset(),
        empty_bytes: frozenset[bytes] = frozenset(),
        delays: dict[bytes, float] | None = None,
    ) -> None:
        self._responses = responses or {}
        self._default = default
        self._fail_bytes = fail_bytes
        self._empty_bytes = empty_bytes
        self._delays = delays or {}
        self.calls: list[SlideContext] = []

    @property
    def call_count(self) -> int:
        return len(self.calls)

    def reconstruct(
        self, image_bytes: bytes, image_ext: str, context: SlideContext
    ) -> str:
        self.calls.append(context)
        delay = self._delays.get(image_bytes)
        if delay:
            time.sleep(delay)
        if image_bytes in self._fail_bytes:
            raise RuntimeError("simulated VLM reconstruct failure")
        if image_bytes in self._empty_bytes:
            return ""
        return self._responses.get(image_bytes, self._default)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _text_slide(index: int, title: str = "", text: str = "body") -> SlideIR:
    """A plain text-only slide (complexity_score == 0, not complex)."""
    return SlideIR(
        index=index,
        title=title,
        shapes=[
            TextShapeIR(
                shape_id=1000 + index,
                name="body",
                kind=ShapeKind.TEXT,
                paragraphs=[ParagraphIR(text=text, level=0)],
            )
        ],
    )


def _complex_slide(index: int, title: str = "") -> SlideIR:
    """A slide with real complexity signals (table + group -> score 4, complex)."""
    return SlideIR(
        index=index,
        title=title,
        shapes=[
            TableShapeIR(
                shape_id=2000 + index,
                name="table",
                kind=ShapeKind.TABLE,
                rows=[["a", "b"], ["c", "d"]],
                n_rows=2,
                n_cols=2,
            ),
            GroupShapeIR(shape_id=2100 + index, name="group", kind=ShapeKind.GROUP),
        ],
    )


def _presentation(slides: list[SlideIR]) -> PresentationIR:
    return PresentationIR(source_path="fake-deck.pptx", slides=slides)


def _always_available(monkeypatch: pytest.MonkeyPatch) -> None:
    """monkeypatch slide_render_available()=True so FakeRenderer path runs
    regardless of whether real soffice/pypdfium2 are installed locally."""
    monkeypatch.setattr(sr_module, "slide_render_available", lambda: True)


# ===========================================================================
# AC5 — reconstructor=None / 렌더 미가용 -> 조기 return (무동작)
# ===========================================================================


def test_ac5_reconstructor_none_no_op(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC5: reconstructor=None -> stage 조기 return, 전 슬라이드
    reconstructed_md=None."""
    _always_available(monkeypatch)
    slide = _complex_slide(0)
    pres = _presentation([slide])
    renderer = FakeRenderer({0: b"png-bytes"})

    reconstruct_slides(pres, None, renderer=renderer)

    assert slide.reconstructed_md is None
    assert (
        renderer.call_count == 0
    ), "reconstructor=None이면 렌더러조차 호출되지 않아야 함"


def test_ac5_render_unavailable_no_op(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC5: slide_render_available()==False -> 조기 return, 예외 미전파(core-only
    안전)."""
    monkeypatch.setattr(sr_module, "slide_render_available", lambda: False)
    slide = _complex_slide(0)
    pres = _presentation([slide])
    renderer = FakeRenderer({0: b"png-bytes"})
    reconstructor = FakeReconstructor()

    reconstruct_slides(pres, reconstructor, renderer=renderer)

    assert slide.reconstructed_md is None
    assert renderer.call_count == 0
    assert reconstructor.call_count == 0


def test_ac5_no_complex_slides_renderer_not_called(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """라우팅 결과 후보가 0건이면 렌더러/재구성 모두 미호출(비용 0)."""
    _always_available(monkeypatch)
    slide = _text_slide(0)
    pres = _presentation([slide])
    renderer = FakeRenderer({0: b"png"})
    reconstructor = FakeReconstructor()

    reconstruct_slides(pres, reconstructor, renderer=renderer)

    assert slide.reconstructed_md is None
    assert renderer.call_count == 0
    assert reconstructor.call_count == 0


# ===========================================================================
# AC1 — 라우팅 -> 렌더(1회) -> 재구성 -> 슬롯 충전, 결정적 병합(인덱스 순 무관)
# ===========================================================================


def test_ac1_complex_slides_routed_others_text_path(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """AC1: complexity_fn=True 인 슬라이드만 reconstructed_md 충전, 나머지는 None."""
    _always_available(monkeypatch)
    complex0 = _complex_slide(0)
    text1 = _text_slide(1)
    complex2 = _complex_slide(2)
    pres = _presentation([complex0, text1, complex2])

    renderer = FakeRenderer({0: b"png0", 2: b"png2"})
    reconstructor = FakeReconstructor(
        responses={
            b"png0": "## Slide 0\n\nreconstructed",
            b"png2": "## Slide 2\n\n|a|b|",
        }
    )

    reconstruct_slides(pres, reconstructor, renderer=renderer)

    assert complex0.reconstructed_md == "## Slide 0\n\nreconstructed"
    assert complex2.reconstructed_md == "## Slide 2\n\n|a|b|"
    assert text1.reconstructed_md is None
    # soffice-cost amortisation: exactly one render_deck call for the whole deck
    assert renderer.call_count == 1
    assert renderer.calls[0][1] == [0, 2]


def test_ac1_render_called_exactly_once_regardless_of_slide_count(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """AC1: 여러 복잡 슬라이드가 있어도 renderer는 정확히 1회 호출된다."""
    _always_available(monkeypatch)
    slides = [_complex_slide(i) for i in range(5)]
    pres = _presentation(slides)
    renderer = FakeRenderer({i: f"png{i}".encode() for i in range(5)})
    reconstructor = FakeReconstructor()

    reconstruct_slides(pres, reconstructor, renderer=renderer)

    assert renderer.call_count == 1
    assert all(s.reconstructed_md is not None for s in slides)


def test_ac1_real_complexity_fn_default_routing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """AC1: complexity_fn 기본값(real is_visually_complex)이 실제로 사용된다."""
    _always_available(monkeypatch)
    complex_slide = _complex_slide(0)  # table+group -> score 4 -> complex True
    text_slide = _text_slide(1)  # score 0 -> not complex
    pres = _presentation([complex_slide, text_slide])
    renderer = FakeRenderer({0: b"png0"})
    reconstructor = FakeReconstructor()

    reconstruct_slides(
        pres, reconstructor, renderer=renderer
    )  # complexity_fn 기본값 사용

    assert complex_slide.reconstructed_md is not None
    assert text_slide.reconstructed_md is None


def test_ac1_deterministic_two_runs_identical(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC1/INV-1: 동일 입력을 2회 실행해도 동일한 reconstructed_md 매핑이 나온다."""
    _always_available(monkeypatch)

    def _run() -> dict[int, str | None]:
        slides = [_complex_slide(0), _text_slide(1), _complex_slide(2)]
        pres = _presentation(slides)
        renderer = FakeRenderer({0: b"png0", 2: b"png2"})
        reconstructor = FakeReconstructor(responses={b"png0": "## A", b"png2": "## B"})
        reconstruct_slides(pres, reconstructor, renderer=renderer)
        return {s.index: s.reconstructed_md for s in slides}

    first = _run()
    second = _run()
    assert first == second == {0: "## A", 1: None, 2: "## B"}


# ===========================================================================
# AC7 — force_render/force_text 오버라이드
# ===========================================================================


def test_ac7_force_render_overrides_heuristic(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC7: force_render는 complexity_fn=False인 슬라이드도 렌더+VLM 경로로 보낸다."""
    _always_available(monkeypatch)
    slide = _text_slide(0)  # complexity_fn(가짜)이 False를 반환해도
    pres = _presentation([slide])
    renderer = FakeRenderer({0: b"png0"})
    reconstructor = FakeReconstructor()

    reconstruct_slides(
        pres,
        reconstructor,
        renderer=renderer,
        complexity_fn=lambda _s: False,
        force_render=frozenset({0}),
    )

    assert slide.reconstructed_md is not None


def test_ac7_force_text_overrides_complexity_and_force_render(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """AC7: force_text는 complexity_fn=True/force_render 지정과 무관하게 항상 이긴다."""
    _always_available(monkeypatch)
    slide = _complex_slide(0)
    pres = _presentation([slide])
    renderer = FakeRenderer({0: b"png0"})
    reconstructor = FakeReconstructor()

    reconstruct_slides(
        pres,
        reconstructor,
        renderer=renderer,
        force_render=frozenset({0}),
        force_text=frozenset({0}),
    )

    assert slide.reconstructed_md is None
    assert renderer.call_count == 0


# ===========================================================================
# AC7 (FR-34 AC2/AC4) — max_vlm_slides 캡: (score desc, index asc) top-N
# ===========================================================================


def test_ac7_cap_top_n_by_score_desc_index_asc(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """FR-34 AC2: max_slides 지정 시 (complexity_score desc, index asc) top-N만 렌더."""
    _always_available(monkeypatch)
    scores = {0: 1, 1: 5, 2: 5, 3: 3}
    monkeypatch.setattr(sr_module, "complexity_score", lambda s: scores[s.index])

    slides = [_text_slide(i) for i in range(4)]
    pres = _presentation(slides)
    renderer = FakeRenderer({i: f"png{i}".encode() for i in range(4)})
    reconstructor = FakeReconstructor()

    reconstruct_slides(
        pres,
        reconstructor,
        renderer=renderer,
        complexity_fn=lambda _s: True,  # 전부 후보
        max_slides=2,
    )

    # top-2 by score desc: idx1(5), idx2(5) tie -> index asc keeps idx1 before idx2;
    # idx3(3) and idx0(1) exceed the cap -> text fallback.
    assert slides[1].reconstructed_md is not None
    assert slides[2].reconstructed_md is not None
    assert slides[0].reconstructed_md is None
    assert slides[3].reconstructed_md is None
    assert renderer.calls[0][1] == [1, 2]


def test_ac7_max_vlm_slides_zero_no_render_no_calls(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """FR-34 AC4: max_vlm_slides=0 -> 렌더 0건, VLM 콜 0."""
    _always_available(monkeypatch)
    slide = _complex_slide(0)
    pres = _presentation([slide])
    renderer = FakeRenderer({0: b"png0"})
    reconstructor = FakeReconstructor()

    reconstruct_slides(pres, reconstructor, renderer=renderer, max_slides=0)

    assert slide.reconstructed_md is None
    assert renderer.call_count == 0
    assert reconstructor.call_count == 0


def test_ac7_max_vlm_slides_none_no_cap(monkeypatch: pytest.MonkeyPatch) -> None:
    """FR-34 AC4: 상한 미지정(None) -> 후보 전체가 렌더+재구성된다(캡 없음)."""
    _always_available(monkeypatch)
    slides = [_complex_slide(i) for i in range(3)]
    pres = _presentation(slides)
    renderer = FakeRenderer({i: f"png{i}".encode() for i in range(3)})
    reconstructor = FakeReconstructor()

    reconstruct_slides(pres, reconstructor, renderer=renderer, max_slides=None)

    assert all(s.reconstructed_md is not None for s in slides)


# ===========================================================================
# AC7 (FR-34 AC3) — sha256 렌더 캐시: 동일 PNG -> 재구성 1회
# ===========================================================================


def test_ac7_identical_png_reconstructed_once(monkeypatch: pytest.MonkeyPatch) -> None:
    """FR-34 AC3: 두 슬라이드의 렌더 PNG가 동일(sha256 동일)하면 reconstruct는 1회만."""
    _always_available(monkeypatch)
    same_png = b"identical-render-bytes"
    slides = [_complex_slide(0), _complex_slide(1), _complex_slide(2)]
    pres = _presentation(slides)
    renderer = FakeRenderer({0: same_png, 1: same_png, 2: b"different-bytes"})
    reconstructor = FakeReconstructor(
        responses={same_png: "## Shared", b"different-bytes": "## Unique"}
    )

    reconstruct_slides(pres, reconstructor, renderer=renderer)

    assert reconstructor.call_count == 2, "고유 PNG 2종 -> 재구성 콜 2회(중복 dedup)"
    assert slides[0].reconstructed_md == "## Shared"
    assert slides[1].reconstructed_md == "## Shared"
    assert slides[2].reconstructed_md == "## Unique"


# ===========================================================================
# AC6 (INV-4) — 격리: 슬라이드 1장 렌더/VLM 실패 -> 텍스트 폴백, 나머지 정상
# ===========================================================================


def test_ac6_render_failure_isolated_to_one_slide(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """INV-4: 한 슬라이드의 render_deck 결과가 None -> 그 슬라이드만 폴백."""
    _always_available(monkeypatch)
    ok_slide = _complex_slide(0)
    failed_slide = _complex_slide(1)
    pres = _presentation([ok_slide, failed_slide])
    renderer = FakeRenderer({0: b"png0", 1: None})  # index 1 render failed
    reconstructor = FakeReconstructor(responses={b"png0": "## OK"})

    reconstruct_slides(pres, reconstructor, renderer=renderer)

    assert ok_slide.reconstructed_md == "## OK"
    assert failed_slide.reconstructed_md is None


def test_ac6_reconstruct_exception_isolated_no_propagation(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """INV-4: reconstruct() 예외 -> 해당 슬라이드만 폴백, 예외는 밖으로 전파되지
    않는다."""
    _always_available(monkeypatch)
    ok_slide = _complex_slide(0)
    bad_slide = _complex_slide(1)
    pres = _presentation([ok_slide, bad_slide])
    renderer = FakeRenderer({0: b"png-ok", 1: b"png-bad"})
    reconstructor = FakeReconstructor(
        responses={b"png-ok": "## OK"}, fail_bytes=frozenset({b"png-bad"})
    )

    reconstruct_slides(pres, reconstructor, renderer=renderer)  # must not raise

    assert ok_slide.reconstructed_md == "## OK"
    assert bad_slide.reconstructed_md is None


def test_ac6_empty_reconstruct_response_falls_back(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """INV-4 보강: reconstruct()가 빈 문자열을 반환하면 텍스트 폴백(None 유지)."""
    _always_available(monkeypatch)
    slide = _complex_slide(0)
    pres = _presentation([slide])
    renderer = FakeRenderer({0: b"png-empty"})
    reconstructor = FakeReconstructor(empty_bytes=frozenset({b"png-empty"}))

    reconstruct_slides(pres, reconstructor, renderer=renderer)

    assert slide.reconstructed_md is None


def test_ac6_multiple_failures_do_not_affect_successful_slides(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """INV-4: 여러 슬라이드 중 일부만 실패해도 나머지는 전부 정상 완료된다."""
    _always_available(monkeypatch)
    slides = [_complex_slide(i) for i in range(4)]
    pres = _presentation(slides)
    renderer = FakeRenderer({0: b"png0", 1: None, 2: b"png2", 3: b"png3-fail"})
    reconstructor = FakeReconstructor(
        responses={b"png0": "## S0", b"png2": "## S2"},
        fail_bytes=frozenset({b"png3-fail"}),
    )

    reconstruct_slides(pres, reconstructor, renderer=renderer)

    assert slides[0].reconstructed_md == "## S0"
    assert slides[1].reconstructed_md is None  # render failure
    assert slides[2].reconstructed_md == "## S2"
    assert slides[3].reconstructed_md is None  # reconstruct failure


# ===========================================================================
# AC7 (FR-34 AC1) — embed_rendered_image: base64 인라인 병행
# ===========================================================================


def test_ac7_embed_prepends_base64_data_uri(monkeypatch: pytest.MonkeyPatch) -> None:
    """FR-34 AC1: embed=True -> reconstructed_md 앞에 base64 PNG data URI가 붙는다."""
    _always_available(monkeypatch)
    slide = _complex_slide(0)
    pres = _presentation([slide])
    renderer = FakeRenderer({0: b"\x89PNG-fake-bytes"})
    reconstructor = FakeReconstructor(responses={b"\x89PNG-fake-bytes": "## Body"})

    reconstruct_slides(pres, reconstructor, renderer=renderer, embed=True)

    assert slide.reconstructed_md is not None
    assert slide.reconstructed_md.startswith("![")
    assert "data:image/png;base64," in slide.reconstructed_md
    assert slide.reconstructed_md.endswith("## Body")


def test_ac7_embed_false_default_no_data_uri(monkeypatch: pytest.MonkeyPatch) -> None:
    """FR-34 AC1: embed=False(기본) -> 재구성 텍스트만 출력, data URI 없음."""
    _always_available(monkeypatch)
    slide = _complex_slide(0)
    pres = _presentation([slide])
    renderer = FakeRenderer({0: b"png"})
    reconstructor = FakeReconstructor(responses={b"png": "## Body"})

    reconstruct_slides(pres, reconstructor, renderer=renderer)  # embed=False 기본

    assert slide.reconstructed_md == "## Body"
    assert "data:image/png;base64," not in slide.reconstructed_md


# ===========================================================================
# AC10 — 완료 순서 무관성(ThreadPoolExecutor 동시성 결과 결정성)
# ===========================================================================


def test_ac10_completion_order_independence(monkeypatch: pytest.MonkeyPatch) -> None:
    """AC10: reconstruct() 완료 순서가 뒤섞여도 최종 매핑은 결정적으로 동일하다."""
    _always_available(monkeypatch)
    slides = [_complex_slide(i) for i in range(4)]
    pres = _presentation(slides)
    png_map = {i: f"png{i}".encode() for i in range(4)}
    renderer = FakeRenderer(png_map)
    # Slowest job first (reverse of submission order) forces mixed completion.
    delays = {png_map[0]: 0.06, png_map[1]: 0.04, png_map[2]: 0.02, png_map[3]: 0.0}
    responses = {png_map[i]: f"## Slide {i}" for i in range(4)}
    reconstructor = FakeReconstructor(responses=responses, delays=delays)

    reconstruct_slides(pres, reconstructor, renderer=renderer, max_workers=4)

    for i, slide in enumerate(slides):
        assert slide.reconstructed_md == f"## Slide {i}"


# ===========================================================================
# AC9 (INV-5) — SDK/pypdfium2 top-level import 0 (slide_reconstruction.py 소스)
# ===========================================================================


def test_ac9_no_vlm_sdk_or_pypdfium2_import_in_source() -> None:
    """INV-5: slide_reconstruction.py에 anthropic/openai/pypdfium2 import가 없다(AST).

    (docstring 산문에서 pypdfium2를 *언급*하는 것은 허용 — 실제 ``import``
    구문이 없는지만 검사한다, slide_render.py AC7 테스트와 동일한 접근.)
    """
    import ast

    src_path = (
        Path(__file__).parent.parent / "src" / "pptx_md" / "slide_reconstruction.py"
    )
    tree = ast.parse(src_path.read_text(encoding="utf-8"))

    imported_roots: list[str] = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                imported_roots.append(alias.name.split(".")[0])
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported_roots.append(node.module.split(".")[0])

    disallowed = {"anthropic", "openai", "pypdfium2"}
    assert not (disallowed & set(imported_roots)), (
        f"slide_reconstruction.py에 금지된 import 발견: "
        f"{disallowed & set(imported_roots)}"
    )


def test_ac9_core_only_import_succeeds() -> None:
    """INV-5: import pptx_md.slide_reconstruction은 항상 성공한다(SDK 미주입)."""
    import importlib

    mod = importlib.import_module("pptx_md.slide_reconstruction")
    assert mod is not None


# ===========================================================================
# AC10 — mypy/ruff/black exit 0 (프로파일 §4 명령)
# ===========================================================================


def test_ac10_mypy_exit_0() -> None:
    """AC10: mypy src/ 가 exit 0 (slide_reconstruction.py 포함)."""
    result = subprocess.run(
        [sys.executable, "-m", "mypy", "src/"],
        capture_output=True,
        text=True,
        cwd=str(Path(__file__).parent.parent),
    )
    assert (
        result.returncode == 0
    ), f"mypy exit {result.returncode}:\n{result.stdout}\n{result.stderr}"


def test_ac10_ruff_check_exit_0() -> None:
    """AC10: ruff check가 slide_reconstruction.py에 대해 exit 0."""
    result = subprocess.run(
        [sys.executable, "-m", "ruff", "check", "src/pptx_md/slide_reconstruction.py"],
        capture_output=True,
        text=True,
        cwd=str(Path(__file__).parent.parent),
    )
    assert (
        result.returncode == 0
    ), f"ruff exit {result.returncode}:\n{result.stdout}\n{result.stderr}"


def test_ac10_black_check_exit_0() -> None:
    """AC10: black --check가 slide_reconstruction.py에 대해 exit 0."""
    result = subprocess.run(
        [
            sys.executable,
            "-m",
            "black",
            "--check",
            "src/pptx_md/slide_reconstruction.py",
        ],
        capture_output=True,
        text=True,
        cwd=str(Path(__file__).parent.parent),
    )
    assert (
        result.returncode == 0
    ), f"black exit {result.returncode}:\n{result.stdout}\n{result.stderr}"
