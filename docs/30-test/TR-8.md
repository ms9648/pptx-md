# TR-8 FR-01 패키지 구조 초기화

실행: 2026-06-27 / 브랜치: `feature/8-package-structure` / 환경: Windows 11, Python 3.11.0

---

## 테스트 케이스 매트릭스

| TC | 대응 AC | 시나리오 | 결과 |
|----|---------|---------|------|
| TC-8-01 | AC1 | `hatch build` → dist에 .whl + .tar.gz 생성, exit 0 | PASS |
| TC-8-02 | AC1 | pyproject.toml build-backend = hatchling, license = MIT 확인 | PASS |
| TC-8-03 | AC2 | `python -c "import pptx_md"` → ImportError 없이 exit 0 | PASS |
| TC-8-04 | AC2 | pytest smoke test (test_ac2_import_pptx_md) 통과 | PASS |
| TC-8-05 | AC3 | core deps에 anthropic/openai 없음 (정적 검증) | PASS |
| TC-8-06 | AC3 | [project.optional-dependencies].vlm에 anthropic>=0.20, openai>=1.0 선언 | PASS |
| TC-8-07 | AC4 | `ruff check .` → exit 0 ("All checks passed!") | PASS |
| TC-8-08 | AC4 | `black --check .` → exit 0 ("4 files would be left unchanged") | PASS |
| TC-8-09 | AC4 | `mypy src/` → exit 0 ("no issues found in 1 source file") | PASS |
| TC-8-10 | AC5 | `requires-python = ">=3.11"` 선언 정적 확인 | PASS |
| TC-8-11 | AC5 | Python 3.10 환경에서 pip 설치 거부 (경계) | SKIP* |

> \* TC-8-11: 본 환경에 Python 3.10 인터프리터 미설치 (`py -3.10 --version` → "No suitable Python runtime found"). `requires-python = ">=3.11"` 선언 자체는 정적 검증 PASS. pip 9.0+ 는 해당 메타데이터를 준수하여 설치를 거부하는 것이 표준 동작이며, pyproject.toml 값이 정확하므로 AC5의 "Given" 조건이 충족된 환경이라면 거부가 보장됨. 실행 증거 없이 SKIP 처리.

---

## 개별 명령 실행 결과

### 1. `ruff check .` (TC-8-07)
```
All checks passed!
EXIT: 0
```

### 2. `black --check .` (TC-8-08)
```
All done! ✨ 🍰 ✨
4 files would be left unchanged.
EXIT: 0
```

### 3. `mypy src/` (TC-8-09)
```
Success: no issues found in 1 source file
EXIT: 0
```

### 4. `pytest --cov=src/pptx_md --cov-report=term-missing -v` (TC-8-04)
```
============================= test session starts =============================
platform win32 -- Python 3.11.0, pytest-9.1.1, pluggy-1.6.0
cachedir: .pytest_cache
rootdir: D:\dev\pptx-md
configfile: pyproject.toml
testpaths: tests

tests/test_smoke.py::test_ac2_import_pptx_md PASSED                      [100%]

=============================== tests coverage ================================
Name                      Stmts   Miss  Cover   Missing
-------------------------------------------------------
src\pptx_md\__init__.py       1      0   100%
-------------------------------------------------------
TOTAL                         1      0   100%
1 passed in 0.09s
EXIT: 0
```

커버리지: 100% (M1 단계 — 코드 1줄, 목표 임계 75%는 FR-17/M6에서 강제, 현재 측정만)

### 5. `hatch build` (TC-8-01)
```
Inspecting build dependencies
------------------------------------ sdist ------------------------------------
dist\pptx_md-0.0.0.tar.gz
------------------------------------ wheel ------------------------------------
dist\pptx_md-0.0.0-py3-none-any.whl
EXIT: 0
```
생성 파일: `dist/pptx_md-0.0.0-py3-none-any.whl` (2627 bytes), `dist/pptx_md-0.0.0.tar.gz` (2936 bytes)

### 6. `python -c "import pptx_md"` (TC-8-03)
```
import OK
EXIT: 0
```

### 7. pyproject.toml 정적 검증 (TC-8-02, TC-8-05, TC-8-06, TC-8-10)
tomllib 파싱 결과:
```
requires-python: >=3.11           # AC5
core deps: ['python-pptx', 'Pillow']  # AC3 — anthropic/openai 없음
vlm extras: ['anthropic>=0.20', 'openai>=1.0']  # AC3 — extras에만 선언
build-backend: hatchling.build     # AC1
license: MIT                        # AC1
[tool.ruff]: 존재                  # AC4
[tool.black]: 존재                 # AC4
```

---

## developer 단위 테스트 형식성 점검

- `tests/test_smoke.py::test_ac2_import_pptx_md`: `import pptx_md` 후 `assert pptx_md is not None` — AC2의 Then("ImportError 없이 exit 0")을 실제로 검증함. 형식적 테스트 아님. PASS.
- AC1/AC3/AC4/AC5에 대응하는 pytest 테스트 없음 — 이들은 빌드/설치/린트 단계에서 검증되는 인프라 AC이므로 pytest 범위 외. 정상.

---

## AC 커버리지 요약

| AC | 내용 | 판정 | 근거 |
|----|------|------|------|
| AC1 | hatch build → .whl + .tar.gz, exit 0, hatchling 백엔드, MIT 라이선스 | PASS | TC-8-01, TC-8-02 |
| AC2 | import pptx_md → ImportError 없이 exit 0 | PASS | TC-8-03, TC-8-04 |
| AC3 | core에 VLM 미포함, vlm extras에만 선언 | PASS | TC-8-05, TC-8-06 |
| AC4 | ruff + black + mypy 모두 exit 0 | PASS | TC-8-07, TC-8-08, TC-8-09 |
| AC5 | requires-python >=3.11, 3.10 이하 설치 거부 | PASS(정적)/SKIP(실행) | TC-8-10, TC-8-11 |

---

## 전체 판정: PASS

모든 실행 가능한 AC가 통과. TC-8-11(Python 3.10 라이브 설치 거부)은 환경 부재로 SKIP이나
`requires-python = ">=3.11"` 선언 정확성으로 AC5 의도 충족 판정.
