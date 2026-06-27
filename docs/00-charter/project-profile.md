# 프로젝트 프로파일 — pptx-md
> 모든 에이전트는 작업 전 이 문서를 읽고 준수한다. 충돌 발견 시 NEEDS_DECISION.
> 승인: ms9648 / 변경은 변경 요청 이슈로만.

## 1. 기술 스택
- Language: Python 3.11+
- Core 의존성: python-pptx, Pillow
- Optional 의존성: VLM SDK(anthropic, openai 등) — `[vlm]` extras로 옵셔널 처리
- 린터/포매터: ruff, black (설정은 pyproject.toml 통합)
- 타입 검사: mypy
- 테스트: pytest, pytest-cov
- 빌드/배포: hatch (pyproject.toml 기반)
- 라이선스: MIT

## 2. 환경 제약
- 망: 공개 인터넷 가능 (폐쇄망 아님)
- 프록시: 없음
- 레지스트리: PyPI 퍼블릭 (pypi.org)
- 배포 대상: PyPI 퍼블릭 (`hatch build && hatch publish`)

## 3. 코드 규약
- 패키지 구조: `src` layout — `src/pptx_md/` (하위 모듈은 기능 단위로 분리)
- 네이밍: snake_case (Python 표준)
- 린터 설정: `pyproject.toml` 내 `[tool.ruff]`, `[tool.black]` 절
- 금지 사항: 프로파일에 없는 외부 라이브러리 임의 추가 금지 (NEEDS_DECISION 으로 에스컬레이션)
- 기존 코드베이스: 없음 (신규 시작)

## 4. 빌드/테스트 명령 (에이전트가 실행할 정확한 명령)
- 빌드: `hatch build`
- 단위 테스트: `pytest`
- 커버리지 포함 테스트: `pytest --cov=src/pptx_md --cov-report=term-missing`
- 배포: `hatch publish`
- 린트: `ruff check . && black --check .`
- 타입 검사: `mypy src/`

## 5. Git 규약
- 브랜치: `feature/<이슈번호>-<슬러그>` (예: `feature/7-core-parser`)
- 커밋: Conventional Commits — `<type>(#이슈): 요약` (예: `feat(#7): add slide parser`)
  - type: feat / fix / refactor / test / docs / chore
- PR 정책: merge commit (솔로 프로젝트, self-merge 허용)
- main 브랜치 직접 push 금지 — PR 경유 필수

## 6. 산출물 요구
- SI 표준 산출물 불필요 (오픈소스 개인 프로젝트)
- 문서화: `README.md` (사용자 가이드) + `docs/` (설계·API 레퍼런스)
- 문서 도구: 미정 (MkDocs 또는 단순 Markdown — M6 착수 시 확정)

## 7. 보안/품질 요구
- 개인정보: PPTX 내 개인정보 포함 가능성 있음 — 마스킹 처리를 사전 옵션으로 제공 (opt-in)
- 커버리지 기대: 신규 코드 라인 75%+
- 타입 안전성: mypy 통과 필수
- 로깅: 개인정보 마스킹 옵션 활성 시 로그에 원본 텍스트 출력 금지

## 8. 활성 스택 스킬
- `skills/stack-python-packaging` — 목차: Python 라이브러리 구조(src layout) / hatch 빌드·배포 / ruff+black+mypy+pytest 규약 / PyPI 배포 절차 / 안티패턴
  ← 킥오프 §4.5.5 에서 확정. 이 절 승인 = 스킬 목록 승인, 승인 후 PM 이 생성

## 9. 마일스톤 골격
- M1: 패키지 구조 초기화 (pyproject.toml, src layout, GitHub Actions CI)
- M2: 코어 파서 (python-pptx 기반 슬라이드/도형 파싱, 중간 IR)
- M3: 이미지 분류기 (rule-based 4종 + LibreOffice optional)
- M4: VLM 연동 (ImageDescriber 프로토콜, shape_hint, 서브타입별 프롬프트)
- M5: Markdown 어셈블러 (Map-Reduce + Mermaid fallback + validate_markdown)
- M6: 테스트 + 문서화 + PyPI 배포
- 일정: PR 완료 기준 유동 (고정 마감 없음)
