---
name: stack-python-packaging
description: >
  Python 라이브러리 개발·PyPI 배포 규약.
  src layout, hatch 빌드, ruff+black+mypy+pytest 설정, PyPI 배포 절차, 안티패턴.
  pptx-md 프로젝트 전용 — project-profile.md §1~4 와 동일 명령 사용.
---

# Python 패키징 스택 스킬 — pptx-md

> 이 스킬은 project-profile.md 의 파생물이다. 충돌 시 프로파일 우선.
> 프로파일에 없는 도구·명령을 임의로 사용하지 않는다.

## 1. 프로젝트 구조 (src layout)

```
pptx-md/
├── src/
│   └── pptx_md/          # 배포 패키지 루트
│       ├── __init__.py
│       └── ...            # 기능 단위 서브모듈
├── tests/
│   └── ...
├── docs/                  # 설계·API 레퍼런스
├── pyproject.toml         # 빌드·린트·테스트 설정 통합
└── README.md
```

- 패키지 임포트명: `pptx_md` (snake_case)
- 테스트는 `tests/` 에 배치, `src/` 안에 두지 않는다
- `__init__.py` 에 공개 API 만 re-export (내부 모듈 직접 노출 금지)

## 2. pyproject.toml 핵심 구조

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "pptx-md"
requires-python = ">=3.11"
dependencies = [
    "python-pptx",
    "Pillow",
]

[project.optional-dependencies]
vlm = [
    # VLM SDK 는 여기에 — 예: "anthropic>=0.20", "openai>=1.0"
]
dev = ["pytest", "pytest-cov", "ruff", "black", "mypy"]

[tool.hatch.build.targets.wheel]
packages = ["src/pptx_md"]

[tool.ruff]
line-length = 88
select = ["E", "F", "I", "UP"]

[tool.black]
line-length = 88
target-version = ["py311"]

[tool.mypy]
python_version = "3.11"
strict = true
ignore_missing_imports = true

[tool.pytest.ini_options]
testpaths = ["tests"]
addopts = "--cov=src/pptx_md --cov-report=term-missing"
```

## 3. 빌드/테스트 명령 (profile §4 와 동일 — 여기가 정본)

| 목적 | 명령 |
|------|------|
| 단위 테스트 | `pytest` |
| 커버리지 포함 | `pytest --cov=src/pptx_md --cov-report=term-missing` |
| 린트 | `ruff check . && black --check .` |
| 타입 검사 | `mypy src/` |
| 빌드 | `hatch build` |
| PyPI 배포 | `hatch publish` (사람이 직접 실행 — PM 이 에이전트에 위임 금지) |

## 4. VLM extras 처리 규약

- VLM SDK 는 `[vlm]` extras 로만 선언한다 (`pip install pptx-md[vlm]`)
- Core 코드에서 VLM SDK 를 직접 import 하지 않는다
- VLM 기능 진입점은 반드시 `try/except ImportError` 로 감싸고
  미설치 시 `InstallationError("pip install pptx-md[vlm]")` 를 raise

```python
# 올바른 패턴
def describe_image(image: bytes, provider: str) -> str:
    try:
        from pptx_md._vlm import get_describer
    except ImportError:
        raise ImportError("VLM 기능은 `pip install pptx-md[vlm]` 이 필요합니다.")
    return get_describer(provider).describe(image)
```

## 5. 테스트 규약

- 커버리지 기대: 신규 코드 라인 **75%+** (CI 에서 강제)
- `pytest-cov` 로 측정, `pyproject.toml` 의 `addopts` 에 등록
- VLM 연동 테스트는 실제 API 호출 금지 — mock/stub 사용
- 픽스처는 `tests/conftest.py` 에 집중

## 6. 안티패턴 (금지)

- `setup.py` / `setup.cfg` 사용 금지 (hatch + pyproject.toml 로 통일)
- `src/` 밖에 패키지 코드 배치 금지
- `requirements.txt` 를 의존성 원본으로 사용 금지 (`pyproject.toml` 이 정본)
- VLM SDK 를 필수 의존성으로 선언 금지 (반드시 extras)
- `hatch publish` 를 에이전트가 자율 실행 금지 (릴리스는 사람 게이트)
- mypy `# type: ignore` 남발 금지 — 불가피한 경우 이유 주석 필수
