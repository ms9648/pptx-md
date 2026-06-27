# ARCH-M1 — 패키지 구조 초기화 & CI 파이프라인

> 범위: M1 (FR-01 패키지 구조 초기화, FR-02 CI 파이프라인)
> 전제: `docs/00-charter/project-profile.md`, `docs/00-charter/charter.md`, `docs/10-requirements/REQ-core.md`
> 스택 스킬: `.claude/skills/stack-python-packaging`
> 작성: architect / 2026-06-27
> 상태: 설계 초안 (reviewer 리뷰 / 사람 승인 전 — 아키텍처 게이트 대상)

---

## 1. 아키텍처에 영향을 주는 요구사항 추출

| 출처 | 항목 | 설계 영향 |
|------|------|-----------|
| FR-01 AC1 | hatchling 백엔드, MIT, `hatch build` 로 wheel+sdist 생성 | `[build-system]`, `[tool.hatch.build.targets]`, sdist include 명시 |
| FR-01 AC2 | `import pptx_md` 성공 | `src/pptx_md/__init__.py` 존재 + src layout 패키지 등록 |
| FR-01 AC3 | core=python-pptx·Pillow, VLM SDK 는 `[vlm]` extras 한정 | `[project.dependencies]` vs `[project.optional-dependencies].vlm` 분리 |
| FR-01 AC4 | `ruff check . && black --check . && mypy src/` exit 0 | `[tool.ruff]`/`[tool.black]`/`[tool.mypy]` 통합 + 통과 가능한 초기 트리 |
| FR-01 AC5 | Python 3.10 이하 설치 거부 | `requires-python = ">=3.11"` |
| FR-02 AC1 | PR/push 트리거, 4개 잡 실행 | `.github/workflows/ci.yml` 트리거 + step 구성 |
| FR-02 AC2 | Python 3.11 매트릭스 | `strategy.matrix.python-version: ["3.11"]` |
| FR-02 AC3/AC4 | 전 단계 통과=green, 1개라도 실패=failure | step 순차 실행 + `set -e` 기본 동작 (fail-fast) |
| NFR-01 | 측정 환경 = GitHub Actions ubuntu-latest (2-core) | CI runner 를 `ubuntu-latest` 로 고정 (성능 측정 기준선) |
| NFR-02 | 신규 코드 라인 커버리지 ≥ 75% | `pytest-cov` + `addopts` 등록 (M1 은 코드 미존재 → 강제 임계는 M6 게이트, ADR-004) |
| NFR-03 | `mypy src/` exit 0 | `[tool.mypy] strict = true` |
| NFR-04 | ruff + black exit 0 | `[tool.ruff]`/`[tool.black]` line-length 88 통일 |
| NFR-07 | 3.11+ 만 동작 | `requires-python`, matrix 3.11 |
| NFR-08 | core 설치 시 VLM/LibreOffice 없이 import·동작 | VLM SDK 를 dependencies 에서 제외, extras 격리 |

> 통합 지점: M1 단계에는 외부 시스템 통합 없음(로컬 라이브러리 + GitHub Actions). 데이터 볼륨/영속 저장소 없음 → ERD 불필요(2~3절 N/A 표기).

---

## 2. 컨텍스트 / 컴포넌트 구조

M1 은 런타임 컴포넌트가 아니라 **빌드·배포·검증 인프라**를 구성하는 단계다. 산출물 간 관계는 아래와 같다.

```mermaid
graph TD
    subgraph repo["Git 저장소 (pptx-md)"]
        PP[pyproject.toml<br/>빌드·린트·타입·테스트 설정 정본]
        SRC["src/pptx_md/__init__.py<br/>패키지 진입점 (M2+ 모듈 re-export 자리)"]
        TST["tests/<br/>conftest.py + smoke test"]
        CI[".github/workflows/ci.yml"]
    end

    subgraph local["로컬 개발자 환경"]
        DEV[developer]
    end

    subgraph gha["GitHub Actions (ubuntu-latest, py3.11)"]
        J1[ruff check .]
        J2[black --check .]
        J3[mypy src/]
        J4[pytest]
    end

    subgraph pypi["PyPI (퍼블릭) — M6 에서 실제 배포"]
        DIST[dist/ .whl + .tar.gz]
    end

    DEV -->|pip install -e .[dev]| PP
    DEV -->|hatch build| DIST
    PP -.설정 제공.-> J1 & J2 & J3 & J4
    SRC -.검사 대상.-> J3 & J4
    TST -.실행 대상.-> J4
    CI -->|push main / PR 트리거| J1 --> J2 --> J3 --> J4
    J4 -->|all pass| GREEN([CI success / green])
```

### 컴포넌트 책임

| 산출물 | 책임 | M1 범위에서의 상태 |
|--------|------|-------------------|
| `pyproject.toml` | 메타데이터·의존성·빌드·린트·타입·테스트 설정의 **단일 정본** | 전체 확정 |
| `src/pptx_md/__init__.py` | 패키지 진입점, 향후 공개 API re-export 위치 | 골격(주석만), `import` 가능 보장 |
| `tests/` | 테스트 루트, 픽스처 집중 위치(`conftest.py`) | smoke 테스트 1개로 CI 파이프라인 가동 보증 |
| `.github/workflows/ci.yml` | lint→type→test 자동 게이트 | 전체 확정 |

---

## 3. 데이터 모델 (ERD)

**N/A** — M1 은 영속 데이터 모델/DB 가 없는 빌드 인프라 단계다. pptx-md 전체로도 stateless 변환 라이브러리(파일 입력 → 문자열 출력)이며 DB 를 사용하지 않는다. 내부 IR(중간 표현) 데이터 구조는 FR-05(M2)에서 정의한다.

---

## 4. 인터페이스 명세

M1 단계의 "인터페이스"는 HTTP API 가 아니라 **빌드/검증 명령 계약**과 **CI 트리거 계약**이다. 프로파일 §4 명령과 1:1 일치한다.

### 4.1 명령 계약 (developer / CI 공통)

| 명령 | 입력 | 정상 출력 | 실패 신호(오류 경로) |
|------|------|-----------|----------------------|
| `hatch build` | 저장소 루트 | `dist/*.whl` 1개 + `dist/*.tar.gz` 1개, exit 0 | 메타데이터 누락/패키지 미발견 시 exit ≠ 0 |
| `pip install .` | 저장소 루트 | core 의존성만 설치(anthropic/openai 미설치) | 빌드 백엔드 오류 시 비-0 |
| `pip install .[vlm]` | 저장소 루트 | core + VLM SDK 설치 | extras 미정의 시 경고/실패 |
| `ruff check .` | 소스 트리 | exit 0 (위반 0) | 위반 존재 시 exit 1 + 위반 목록 |
| `black --check .` | 소스 트리 | exit 0 (포맷 일치) | 미포맷 파일 존재 시 exit 1 |
| `mypy src/` | `src/` | exit 0 (에러 0) | 타입 에러 존재 시 exit 1 + 에러 목록 |
| `pytest` | `tests/` | exit 0 (전 테스트 통과) | 실패/수집오류 시 exit ≠ 0 |

### 4.2 CI 트리거 계약 (`.github/workflows/ci.yml`)

| 이벤트 | 조건 | 동작 |
|--------|------|------|
| `push` | `main` 브랜치 | 전체 잡 실행 |
| `pull_request` | 모든 대상 브랜치 | 전체 잡 실행 |
| (그 외) | — | 트리거 안 함 |

### 4.3 오류 코드 / 실패 모드 표

| 코드 | 단계 | 의미 | CI 결과 |
|------|------|------|---------|
| E-LINT | `ruff check .` | 린트 위반 | failure (이후 step skip, FR-02 AC4) |
| E-FMT | `black --check .` | 포맷 불일치 | failure |
| E-TYPE | `mypy src/` | 타입 에러 | failure |
| E-TEST | `pytest` | 테스트 실패/수집 오류 | failure |
| E-BUILD | `hatch build` | 빌드/메타데이터 오류 | (M1 CI 에는 build step 미포함 — 배포는 M6/사람 게이트) |
| OK | 전 단계 | 모두 exit 0 | success(green) (FR-02 AC3) |

> step 은 순차 실행이며 GitHub Actions 의 기본 fail-fast(앞 step 실패 시 후속 step 미실행)로 FR-02 AC4 를 충족한다. 어떤 step 이 비-0 을 반환하면 잡 전체가 failure → 머지 가능 신호(green) 미발생.

---

## 5. 횡단 관심사

| 관심사 | M1 설계 |
|--------|---------|
| 인증·인가 | **N/A** — 로컬 라이브러리, 권한 체계 없음(REQ §0). PyPI 배포 토큰은 M6/사람 게이트(`hatch publish` 사람 직접 실행, 헌장 §3). |
| 트랜잭션 경계 | **N/A** — DB·상태 없음. |
| 예외 처리 전략 | M1 코드 자체는 골격뿐. 라이브러리 전역 예외 정책(특히 VLM 미설치 시 `ImportError("pip install pptx-md[vlm]")`)은 스킬 §4 규약을 따르며 M4 에서 구현. M1 은 그 격리 구조(extras)만 마련. |
| 로깅·감사 | M1 에 런타임 로깅 없음. 마스킹 옵션 로그 정책(NFR-06)은 FR-15(M5) 구현 시 적용. |
| 비밀정보(NFR-05) | VLM API key 를 패키지/소스/CI 에 일절 포함하지 않음. CI 워크플로는 어떤 시크릿도 사용하지 않음(lint/type/test 만). PyPI 토큰은 M6 에서 GitHub Secrets/사람 직접 실행으로 분리. |

---

## 6. 기술 선택지 비교 (ADR 후보)

아래 결정들은 프로파일/스킬에 의해 대부분 **이미 확정**되어 있으나, 갈림이 있는 지점만 비교표로 정리한다. M1 의 결정들은 프로파일 범위 내 세부 구현이므로 NEEDS_DECISION 사유는 없다(7절 ADR 로 박제).

### 6.1 wheel 패키지 발견 방식

| 후보 | 설명 | 장점 | 단점 |
|------|------|------|------|
| A. 명시 `packages = ["src/pptx_md"]` | 스킬 §2 예시 방식 | 명확, 의도 고정 | 모듈 추가 시 신경 안 써도 됨(디렉터리 단위) |
| B. hatch 자동 탐지(`sources`) | hatchling src layout 관례 | 설정 최소 | 디렉터리 매핑 오해 소지 |

**권고: A** — 스킬 §2 와 일치, src layout 매핑이 명시적. (→ ADR-001)

### 6.2 커버리지 강제 임계값(`--cov-fail-under`) 적용 시점

| 후보 | 설명 | 장점 | 단점 |
|------|------|------|------|
| A. M1 부터 `--cov-fail-under=75` | 처음부터 강제 | 일관성 | M1 은 코드 거의 없음 → smoke 만으로 75% 의미 없음/오탐 위험 |
| B. M1 은 `--cov` 측정만, 강제는 M6 | 점진 적용 | 초기 오탐 회피 | 중간 단계 커버리지 게이트 부재 |

**권고: B** — M1 트리는 골격뿐이라 라인 커버리지 강제가 무의미. `addopts` 에 `--cov` 측정은 켜되 `--cov-fail-under` 는 M6(FR-17) 게이트에서 도입. (→ ADR-004)

---

## 7. 아키텍처 결정 기록 (ADR)

### ADR-001 빌드 백엔드 및 패키지 발견
**배경**: src layout 패키지를 PyPI 배포 가능한 형태로 빌드해야 함(FR-01 AC1/AC2).
**결정**: `hatchling` 백엔드 + `[tool.hatch.build.targets.wheel] packages = ["src/pptx_md"]` 명시. sdist 는 `src`·`tests`·`README`·`LICENSE` 포함.
**근거**: 프로파일 §1(hatch) + 스킬 §2 예시와 일치. 명시 매핑으로 src layout 의도 고정.
**대안과 기각 사유**: setuptools(프로파일이 hatch 지정, 스킬 §6 `setup.py` 금지), hatch 자동탐지(매핑 모호).
**영향**: `hatch build` 시 wheel 1 + sdist 1 생성(FR-01 AC1).

### ADR-002 VLM SDK 의존성 격리
**배경**: VLM SDK 는 core 에 포함하면 안 됨(NFR-08, 헌장 §3, 스킬 §4).
**결정**: `python-pptx`·`Pillow` 만 `[project.dependencies]`. `anthropic`·`openai` 는 `[project.optional-dependencies].vlm`. core 코드에서 VLM SDK 직접 import 금지.
**근거**: FR-01 AC3 / NFR-08 — extras 미지정 설치 시 VLM SDK 미설치 보장.
**대안과 기각 사유**: 단일 dependencies 에 포함(헌장·NFR 위반), 별도 패키지 분리(과설계 — 단일 패키지 + extras 로 충분).
**영향**: `pip install .` 와 `pip install .[vlm]` 의 설치 집합이 분리됨.

### ADR-003 린트·포맷·타입 설정의 단일 정본화
**배경**: NFR-03/04 충족과 도구 설정 분산 방지.
**결정**: ruff(line-length 88, select E/F/I/UP), black(88, py311), mypy(`strict=true`, `ignore_missing_imports=true`)를 모두 `pyproject.toml` 에 통합. `setup.cfg`·별도 `.flake8`·`mypy.ini` 미사용.
**근거**: 프로파일 §3, 스킬 §2·§6(`setup.cfg` 금지). `ignore_missing_imports=true` 로 python-pptx 등 미타이핑 외부 패키지의 strict 충돌 회피.
**대안과 기각 사유**: 도구별 분리 설정 파일(스킬 §6 금지, 정본 분산).
**영향**: `ruff check . && black --check . && mypy src/` 가 단일 설정으로 동작(FR-01 AC4).

### ADR-004 커버리지 측정 vs 강제의 단계 분리
**배경**: NFR-02(신규 코드 75%)는 코드가 쌓인 뒤라야 의미. M1 트리는 골격뿐.
**결정**: M1 은 `addopts` 에 `--cov=src/pptx_md --cov-report=term-missing`(측정)만 등록. `--cov-fail-under=75`(강제)는 M6(FR-17) 게이트에서 도입.
**근거**: M1 에서 강제 시 골격 코드로 인한 오탐/무의미 통과. 프로파일 §7 커버리지는 "신규 코드" 기준.
**대안과 기각 사유**: M1 부터 강제(오탐), 측정 미설정(M2 부터 가시성 상실).
**영향**: M1 CI 의 pytest 는 smoke 통과만으로 green. 커버리지 게이트는 M6 디스패치에서 추가 이슈로 처리.

### ADR-005 CI 잡 구조 — 단일 잡 순차 step vs 병렬 잡
**배경**: FR-02 AC1(4개 검사 실행), AC4(하나라도 실패 시 failure).
**결정**: ubuntu-latest, python 3.11 매트릭스의 **단일 잡** 내에서 ruff→black→mypy→pytest 를 순차 step 으로 실행.
**근거**: 검사 4종이 가볍고 동일 환경 의존(설치 1회 재사용). fail-fast 기본 동작으로 AC4 충족. 병렬 잡은 환경 설치 N배 비용 + 솔로 프로젝트엔 과설계. ubuntu-latest 고정은 NFR-01 측정 기준선과도 일치.
**대안과 기각 사유**: 검사별 병렬 잡(설치 중복·과설계), self-hosted runner(프로파일에 없음).
**영향**: 한 step 비-0 → 후속 step 미실행 → 잡 failure → green 미발생(FR-02 AC3/AC4).

---

## 8. 산출물 상세 초안

### 8.1 최종 디렉터리 트리 (M1 완료 시)

```
pptx-md/
├── .github/
│   ├── ISSUE_TEMPLATE/        # (기존)
│   ├── PULL_REQUEST_TEMPLATE.md  # (기존)
│   └── workflows/
│       └── ci.yml             # 신규 (FR-02)
├── src/
│   └── pptx_md/
│       └── __init__.py        # 신규 (FR-01 AC2)
├── tests/
│   ├── __init__.py            # 신규
│   ├── conftest.py            # 신규 (픽스처 집중 위치, 스킬 §5)
│   └── test_smoke.py          # 신규 (import 가능성 + 파이프라인 가동 보증)
├── docs/                      # (기존)
├── pyproject.toml             # 신규 (FR-01)
├── LICENSE                    # 신규 (MIT — sdist/메타데이터 필수)
└── README.md                  # 신규 (배포 메타데이터 long_description 근거)
```

> 비고: `README.md`·`LICENSE` 는 PyPI 배포 메타데이터상 필요하지만 본문(사용 가이드)은 FR-18(M6)에서 채운다. M1 에서는 최소 스텁만 둔다(빌드 메타데이터 충족용). 실제 작성 범위/분량은 8.6 WBS 의 W1 AC 참조.

### 8.2 `pyproject.toml` 전체 초안

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "pptx-md"
version = "0.0.0"
description = "Convert PPTX files into LLM-friendly Markdown (VLM-first image understanding)."
readme = "README.md"
license = { text = "MIT" }
requires-python = ">=3.11"
authors = [{ name = "ms9648" }]
keywords = ["pptx", "markdown", "powerpoint", "llm", "vlm"]
classifiers = [
    "Development Status :: 3 - Alpha",
    "Intended Audience :: Developers",
    "License :: OSI Approved :: MIT License",
    "Programming Language :: Python :: 3",
    "Programming Language :: Python :: 3.11",
    "Topic :: Text Processing :: Markup :: Markdown",
]
dependencies = [
    "python-pptx",
    "Pillow",
]

[project.optional-dependencies]
# VLM SDK 는 반드시 여기에만 (스킬 §4, ADR-002). core 에 포함 금지.
vlm = [
    "anthropic>=0.20",
    "openai>=1.0",
]
dev = [
    "pytest",
    "pytest-cov",
    "ruff",
    "black",
    "mypy",
]

[project.urls]
Homepage = "https://github.com/ms9648/pptx-md"
Repository = "https://github.com/ms9648/pptx-md"

[tool.hatch.build.targets.wheel]
packages = ["src/pptx_md"]

[tool.hatch.build.targets.sdist]
include = [
    "src/pptx_md",
    "tests",
    "README.md",
    "LICENSE",
    "pyproject.toml",
]

[tool.ruff]
line-length = 88
target-version = "py311"

[tool.ruff.lint]
select = ["E", "F", "I", "UP"]

[tool.black]
line-length = 88
target-version = ["py311"]

[tool.mypy]
python_version = "3.11"
strict = true
ignore_missing_imports = true
files = ["src"]

[tool.pytest.ini_options]
testpaths = ["tests"]
# M1: 측정만 (강제 임계 --cov-fail-under 는 M6/FR-17, ADR-004)
addopts = "--cov=src/pptx_md --cov-report=term-missing"
```

> 비고1: `select` 를 최신 ruff 권장 위치인 `[tool.ruff.lint]` 로 둠(스킬 예시는 `[tool.ruff]` 평면 배치였으나 최신 ruff 가 deprecation 경고 → 동일 의도, 위치만 정정). developer 가 설치된 ruff 버전이 평면 키만 지원하면 `[tool.ruff].select` 로 대체 가능(둘 다 프로파일 §3 "pyproject 통합" 충족).
> 비고2: `version = "0.0.0"` 은 M1 플레이스홀더. 배포 버전 확정은 M6(FR-19).
> 비고3: VLM extras 의 버전 하한(`anthropic>=0.20`, `openai>=1.0`)은 스킬 §2 주석 예시값을 채택. 실제 구현은 M4 이며, 그때 호환 검증 후 조정 가능(이 자체가 core 의존성은 아니므로 NFR-08 영향 없음).

### 8.3 `src/pptx_md/__init__.py` 골격

```python
"""pptx-md — Convert PPTX files into LLM-friendly Markdown.

Public API re-exports go here as modules land in later milestones, e.g.:

    # M6 / FR-16:
    # from pptx_md.api import convert, ConvertOptions
    # __all__ = ["convert", "ConvertOptions"]

Until then this module only guarantees that ``import pptx_md`` succeeds
(FR-01 AC2). Internal submodules must not be exposed directly (skill §1).
"""

__all__: list[str] = []
```

### 8.4 `tests/conftest.py` 골격

```python
"""Shared pytest fixtures (skill §5: fixtures concentrated here).

Empty for M1; M2+ adds sample-PPTX fixtures and VLM mock/stub fixtures.
"""
```

### 8.5 `tests/test_smoke.py` 골격

```python
"""M1 smoke test — guarantees the package imports and the CI pipeline runs.

Covers FR-01 AC2 (import succeeds) and gives pytest a target so the CI
``pytest`` step exercises a real (passing) test.
"""


def test_import_pptx_md() -> None:
    import pptx_md

    assert pptx_md is not None
```

### 8.6 `.github/workflows/ci.yml` 전체 초안

```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:

jobs:
  lint-type-test:
    runs-on: ubuntu-latest   # NFR-01 측정 기준선 / FR-02 AC1
    strategy:
      matrix:
        python-version: ["3.11"]   # FR-02 AC2 / NFR-07
    steps:
      - uses: actions/checkout@v4

      - name: Set up Python ${{ matrix.python-version }}
        uses: actions/setup-python@v5
        with:
          python-version: ${{ matrix.python-version }}

      - name: Install (editable + dev extras)
        run: |
          python -m pip install --upgrade pip
          pip install -e .[dev]

      - name: Ruff (lint)
        run: ruff check .

      - name: Black (format check)
        run: black --check .

      - name: Mypy (type check)
        run: mypy src/

      - name: Pytest
        run: pytest
```

> step 순차 실행 + GitHub Actions 기본 fail-fast → 어느 step 이라도 비-0 이면 잡 failure(FR-02 AC4). 전 step 통과 시에만 success/green(FR-02 AC3). 시크릿 미사용(NFR-05).

---

## 9. WBS — 구현 이슈 분해

각 단위는 developer 가 반나절~하루 내 독립 수행 가능한 크기.

| ID | 작업 | 참조 설계 절 | AC 초안 | 의존 |
|----|------|-------------|---------|------|
| W1 | 패키지 골격 생성: `pyproject.toml`, `src/pptx_md/__init__.py`, `LICENSE`(MIT), `README.md` 스텁 | 8.1, 8.2, 8.3 | `hatch build` → `dist/*.whl` 1 + `*.tar.gz` 1, exit 0(FR-01 AC1); `python -c "import pptx_md"` exit 0(AC2); `pip install .` 후 `python -c "import anthropic"` 실패=VLM 미설치 확인(AC3) | 없음 |
| W2 | 린트·타입·테스트 설정 + smoke 테스트: `[tool.ruff/black/mypy/pytest]`, `tests/__init__.py`·`conftest.py`·`test_smoke.py` | 8.2, 8.4, 8.5 | `ruff check . && black --check . && mypy src/` 모두 exit 0(FR-01 AC4); `pytest` exit 0(smoke 통과) | W1 |
| W3 | CI 워크플로: `.github/workflows/ci.yml` | 8.6 | PR 생성 시 ruff/black/mypy/pytest 4 step 실행(FR-02 AC1); 3.11 매트릭스(AC2); 전 통과=green(AC3); 의도적 위반 주입 시 failure 확인(AC4) | W2 |

> W1~W3 는 모두 FR-01·FR-02 의존(FR-02→FR-01) 순서를 따른다. 한 이슈(#M1)로 묶어 W1→W2→W3 순차 PR 가능하며, 분리 이슈로 쪼개도 무방(PM 판단). AC5(3.11 미만 거부)는 `requires-python` 선언으로 W1 에서 자동 충족되며 별도 작업 불요(검증은 reviewer/tester 가 메타데이터 확인).

---

## 10. 요구사항 추적표 (FR/NFR → 설계 요소)

### 10.1 M1 활성 FR

| FR/AC | 충족 설계 요소 |
|-------|---------------|
| FR-01 AC1 (build → whl+sdist) | ADR-001, 8.2 `[build-system]`/`[tool.hatch.build.targets]`, W1 |
| FR-01 AC2 (import 성공) | 8.1 src layout, 8.3 `__init__.py`, 8.5 smoke test, W1/W2 |
| FR-01 AC3 (extras 격리) | ADR-002, 8.2 dependencies vs optional-dependencies.vlm, W1 |
| FR-01 AC4 (ruff/black/mypy 통과) | ADR-003, 8.2 `[tool.ruff/black/mypy]`, W2 |
| FR-01 AC5 (3.10 거부) | 8.2 `requires-python = ">=3.11"`, W1 |
| FR-02 AC1 (4 잡 트리거) | ADR-005, 8.6 ci.yml steps + on:push/pull_request, W3 |
| FR-02 AC2 (3.11 매트릭스) | 8.6 `matrix.python-version: ["3.11"]`, W3 |
| FR-02 AC3 (전 통과=green) | 8.6 순차 step + 4.3 OK, W3 |
| FR-02 AC4 (1 실패=failure) | ADR-005, 4.3 실패 코드 표 + fail-fast, W3 |

### 10.2 NFR 충족 1줄 근거

| NFR | 이 설계가 충족하는 방법 |
|-----|------------------------|
| NFR-01 변환 성능 | CI runner 를 `ubuntu-latest`(2-core)로 고정 → NFR-01 측정 기준선 환경을 M1 부터 확보(실제 측정은 M2+). |
| NFR-02 커버리지 75% | `addopts` 에 `--cov` 측정 등록(8.2), 강제 임계는 M6 게이트(ADR-004). |
| NFR-03 타입 안전성 | `[tool.mypy] strict=true` + CI `mypy src/` step(ADR-003, 8.6). |
| NFR-04 코드 스타일 | `[tool.ruff]`+`[tool.black]` line-length 88 통일 + CI 2 step(ADR-003, 8.6). |
| NFR-05 비밀정보 | CI 워크플로 시크릿 미사용, VLM key 패키지/소스 비포함(§5, ADR-002). |
| NFR-06 로깅(PII) | M1 런타임 로깅 없음 — 정책은 FR-15(M5)에서 적용(§5). |
| NFR-07 런타임 호환 | `requires-python=">=3.11"` + matrix 3.11(8.2, 8.6). |
| NFR-08 의존성 격리 | core dependencies 에서 VLM SDK 제외, extras 분리(ADR-002, 8.2). |

### 10.3 M1 범위 밖 FR (M2~M6)

FR-03~FR-19 는 비활성 마일스톤(REQ §4, rolling-wave) — M1 설계 대상 아님. 단, M1 산출물이 후속을 막지 않음을 보장:
- `__init__.py` 가 FR-16 공개 API re-export 위치를 명시(8.3) → M6 진입점 확보.
- extras 구조가 FR-08/09 VLM plug-in 의 의존성 격리 토대(ADR-002).
- `tests/conftest.py` 가 M2+ 픽스처 집중 위치(8.4) → 테스트 확장 토대.

> M1 활성 FR(FR-01, FR-02) 누락 0. NFR 8건 전부 매핑(미적용 항목은 적용 마일스톤 명시).

---

## 11. 리스크 / 부채

| 구분 | 내용 | 대응 |
|------|------|------|
| 부채 | `version="0.0.0"` 플레이스홀더 | M6/FR-19 에서 정식 버전 확정 |
| 리스크 | ruff 버전에 따라 `[tool.ruff.lint].select` vs `[tool.ruff].select` 위치 상이 | 8.2 비고1 — developer 가 설치 버전에 맞게 택일(둘 다 프로파일 충족) |
| 부채 | 커버리지 강제 미적용(M1) | ADR-004 — M6 에서 `--cov-fail-under=75` 도입(별도 이슈) |
| 리스크 | python-pptx/Pillow 미타이핑 → mypy strict 충돌 | `ignore_missing_imports=true` 로 회피(ADR-003) |

> 프로파일에 없는 인프라/도구 도입 0. 환경 제약(PyPI 퍼블릭, 폐쇄망 아님)과 모순 없음. NEEDS_DECISION 사유 없음.
