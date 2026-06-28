# TR-9 — FR-02 CI 파이프라인 테스트 리포트

**이슈**: #9 FR-02
**브랜치**: feature/9-ci-pipeline
**PR**: #11
**날짜**: 2026-06-27
**판정**: PASS

---

## 테스트 케이스 매트릭스

| TC | 대응 AC | 시나리오 | 결과 |
|----|---------|---------|------|
| TC-9-01 | AC1 | ci.yml에 push/PR 트리거 + ruff/black/mypy/pytest step 4개 존재 (정적 검증) | PASS |
| TC-9-02 | AC1 | GHA 런 28285971298에서 4개 step 실제 실행됨 (CI 런 증거) | PASS |
| TC-9-03 | AC2 | matrix.python-version: ["3.11"] 선언 확인 | PASS |
| TC-9-04 | AC2 | GHA 잡명 "lint-type-test (3.11)" — 3.11 환경 실행 확인 | PASS |
| TC-9-05 | AC3 | GHA 런 결론 success, 전체 잡 1개 모두 success | PASS |
| TC-9-06 | AC3 | 로컬 ruff/black/mypy/pytest 4개 명령 모두 exit 0 | PASS |
| TC-9-07 | AC4 | continue-on-error / if: always() 설정 부재 — 기본 fail-fast 동작 확인 | PASS |
| TC-9-08 | AC4 | 경계: step 중 하나라도 실패 시 GHA 기본 동작으로 워크플로 failure 종료 (설정 분석) | PASS |

---

## AC 검증 결과

| AC | 판정 | 근거 |
|----|------|------|
| AC1 | PASS | ci.yml에 `on: pull_request` 트리거 및 Ruff/Black/Mypy/Pytest step 4개 명시적 존재. GHA 런 28285971298에서 step 5~8번으로 모두 실행·success 확인 (TC-9-01, TC-9-02) |
| AC2 | PASS | `matrix: python-version: ["3.11"]` 선언, GHA 잡명 "lint-type-test (3.11)"으로 3.11 환경 실행 확인 (TC-9-03, TC-9-04) |
| AC3 | PASS | GHA 런 conclusion=success, 로컬 4개 명령 모두 exit 0 (TC-9-05, TC-9-06) |
| AC4 | PASS | `continue-on-error: true` 및 `if: always()` 미사용 확인 (grep 결과 No matches). GHA 기본 동작상 step 실패 시 이후 step 스킵 및 workflow failure 보장 (TC-9-07, TC-9-08) |

---

## CI 런 증거

- 런 ID: 28285971298
- 결론: success
- Python 버전: 3.11
- 잡명: lint-type-test (3.11)
- 실행된 step:
  - step 5: Ruff (lint) — success (10:02:57 ~ 10:02:57)
  - step 6: Black (format check) — success (10:02:57 ~ 10:02:58)
  - step 7: Mypy (type check) — success (10:02:58 ~ 10:02:59)
  - step 8: Pytest — success (10:02:59 ~ 10:02:59)
- 잡 URL: https://github.com/ms9648/pptx-md/actions/runs/28285971298/job/83809883447

---

## 로컬 검증 결과

환경: Windows 11, Python 3.11.0, 브랜치 feature/9-ci-pipeline

| 명령 | 출력 요약 | exit code |
|------|----------|-----------|
| `ruff check .` | All checks passed! | 0 |
| `black --check .` | 4 files would be left unchanged | 0 |
| `mypy src/` | Success: no issues found in 1 source file | 0 |
| `pytest --cov=src/pptx_md --cov-report=term-missing` | 1 passed, coverage 100% | 0 |

### pytest 상세 출력
```
============================= test session starts =============================
platform win32 -- Python 3.11.0, pytest-9.1.1, pluggy-1.6.0
rootdir: D:\dev\pptx-md
configfile: pyproject.toml
testpaths: tests

tests\test_smoke.py .                                                    [100%]

Name                      Stmts   Miss  Cover   Missing
-------------------------------------------------------
src\pptx_md\__init__.py       1      0   100%
-------------------------------------------------------
TOTAL                         1      0   100%
1 passed in 0.10s
EXIT: 0
```

---

## developer 단위 테스트 형식성 점검

CI 워크플로는 인프라 산출물이므로 별도 pytest 단위 테스트 없음 — 정상.
AC 검증은 GHA 런 실행 로그와 ci.yml 정적 분석으로 수행.
ci.yml에 `# FR-02 AC1`, `# FR-02 AC2` 주석 삽입으로 AC 추적성 확보되어 있음.

---

## 결함 목록

없음

---

## 전체 판정: PASS

AC1~AC4 전체 통과. GHA 런 28285971298 success 증거 및 로컬 4개 명령 exit 0 확인.
AC4는 실제 failure 런이 없으나 ci.yml 정적 분석으로 실패 흡수 설정(continue-on-error, if: always()) 부재를 확인하여 GHA 기본 fail-fast 동작으로 충족됨을 검증.
