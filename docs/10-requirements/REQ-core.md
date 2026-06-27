# 요구사항 명세 — pptx-md (core)

> 원천: `docs/10-requirements/intake/INTAKE-core.md` (작성 ms9648 / 2026-06-27)
> 정제: requirements-analyst / 2026-06-27
> 전제: `docs/00-charter/charter.md`, `docs/00-charter/project-profile.md`
> 정제 방식: **rolling-wave** (활성 마일스톤 M1 만 상세 AC, M2~M6 은 차기 정제 대상)
> 상태: 초안 (사람 승인 전 — 베이스라인 미확정)
> RFI-1 답변 반영: requirements-analyst / 2026-06-27 (FR-09, FR-13, NFR-01, FR-15)

---

## 0. 액터 (인테이크 B절)

| 액터 | 정의 | 규모 | 핵심 행위 |
|------|------|------|----------|
| 라이브러리 사용자 | Python 개발자 / LLM 앱 빌더 | 불특정 다수 | `pip install pptx-md` 후 코드에서 `convert()` 호출 |

> 헌장·인테이크 모두 단일 액터. 권한 체계 없음 (로컬 실행 라이브러리, 동시접속·인증 개념 없음).

---

## 1. 액터별 핵심 시나리오

**라이브러리 사용자 (변환 워크플로)**
1. `pip install pptx-md` (VLM 사용 시 `pip install pptx-md[vlm]`)로 설치한다.
2. 코드에서 `from pptx_md import convert` 후 `convert("deck.pptx")` 를 호출한다.
3. 내부적으로 PPTX 가 파싱되어 슬라이드·도형이 중간 IR 로 추출된다.
4. 이미지는 rule-based 4종 분류 후, VLM 이 설정된 경우 분류 힌트와 함께 설명 텍스트로 변환된다.
5. 슬라이드별 Markdown 블록이 Map-Reduce 로 합쳐져 단일 Markdown 문서가 반환된다.
6. (opt-in 시) 텍스트 내 개인정보가 마스킹된 상태로 출력된다.

---

## 2. 기능 요구 (FR) — 추적성 매핑

| FR | 기능명 | 원천 R-ID | 마일스톤 | 우선순위 | 정제 상태 |
|----|--------|-----------|----------|----------|-----------|
| FR-01 | 패키지 구조 초기화 | R-01 | M1 | Must | 상세 AC |
| FR-02 | CI 파이프라인 | R-02 | M1 | Must | 상세 AC |
| FR-03 | 슬라이드 파싱 | R-03 | M2 | Must | 차기 정제 대상 |
| FR-04 | 도형 파싱 | R-04 | M2 | Must | 차기 정제 대상 |
| FR-05 | 중간 IR 정의 | R-05 | M2 | Must | 차기 정제 대상 |
| FR-06 | 이미지 분류 (rule-based) | R-06 | M3 | Must | 차기 정제 대상 |
| FR-07 | LibreOffice 변환 (optional) | R-07 | M3 | Should | 차기 정제 대상 |
| FR-08 | ImageDescriber 프로토콜 | R-08 | M4 | Must | 차기 정제 대상 |
| FR-09 | VLM 제공자 구현 | R-09 | M4 | Must | 차기 정제 대상 |
| FR-10 | shape_hint 지원 | R-10 | M4 | Must | 차기 정제 대상 |
| FR-11 | Markdown 어셈블러 | R-11 | M5 | Must | 차기 정제 대상 |
| FR-12 | Map-Reduce 파이프라인 | R-12 | M5 | Must | 차기 정제 대상 |
| FR-13 | Mermaid fallback | R-13 | M5 | Should | 차기 정제 대상 |
| FR-14 | validate_markdown | R-14 | M5 | Should | 차기 정제 대상 |
| FR-15 | 개인정보 마스킹 | R-15 | M5 | Should | 차기 정제 대상 |
| FR-16 | 공개 API | R-16 | M6 | Must | 차기 정제 대상 |
| FR-17 | 테스트 스위트 | R-17 | M6 | Must | 차기 정제 대상 |
| FR-18 | README + 사용 예시 | R-18 | M6 | Must | 차기 정제 대상 |
| FR-19 | PyPI 배포 | R-19 | M6 | Must | 차기 정제 대상 |

> 전수 커버: R-01 ~ R-19 (19개) → FR-01 ~ FR-19 (19개), 누락 0. 분해는 1:1.

---

## 3. 활성 마일스톤 (M1) — 상세 명세

### FR-01 패키지 구조 초기화
**스토리**: 라이브러리 사용자(개발자)로서, `pip install pptx-md` 로 설치 가능한 src layout 패키지를 구성하고 싶다. 배포·임포트가 표준 방식으로 동작하게 하기 위해서다.
**AC**:
- [ ] AC1: Given `pyproject.toml` 이 hatch(hatchling) 빌드 백엔드와 라이선스 MIT 로 설정됨 When `hatch build` 실행 Then `dist/` 에 `.whl` 과 `.tar.gz`(sdist) 가 각 1개 생성되고 exit code 0 이다
- [ ] AC2: Given `src/pptx_md/__init__.py` 존재 When `python -c "import pptx_md"` 실행 Then ImportError 없이 exit code 0 으로 종료한다
- [ ] AC3: Given core 의존성이 `python-pptx`, `Pillow` 로 선언되고 VLM SDK(anthropic, openai)는 `[project.optional-dependencies]` 의 `vlm` extras 에만 선언됨 When `pip install .` (extras 미지정) 실행 Then anthropic/openai 패키지는 설치되지 않는다
- [ ] AC4: Given `pyproject.toml` 에 `[tool.ruff]`, `[tool.black]` 설정이 통합됨 When `ruff check . && black --check . && mypy src/` 실행 Then 모두 exit code 0 으로 통과한다
- [ ] AC5 (경계): Given Python 3.10 이하 환경 When 설치 시도 Then `requires-python = ">=3.11"` 제약에 의해 pip 가 설치를 거부한다
**우선순위**: Must / **의존**: 없음

### FR-02 CI 파이프라인
**스토리**: 라이브러리 사용자(개발자)로서, PR 마다 lint·type·test 가 자동 실행되길 원한다. 머지 전 회귀를 자동 차단하기 위해서다.
**AC**:
- [ ] AC1: Given `.github/workflows/` 에 CI 워크플로가 존재하고 PR(또는 push) 트리거가 설정됨 When PR 을 생성 Then `ruff check .`, `black --check .`, `mypy src/`, `pytest` 잡이 모두 실행된다
- [ ] AC2: Given 워크플로 매트릭스에 Python 3.11 이 포함됨 When CI 실행 Then 3.11 환경에서 전체 잡이 수행된다
- [ ] AC3: Given 모든 단계가 통과함 When CI 종료 Then 워크플로 전체 상태가 success(green) 로 보고된다
- [ ] AC4 (예외): Given lint/type/test 중 하나라도 0 이 아닌 exit code 를 반환 When CI 실행 Then 워크플로가 failure 로 종료되어 머지 가능 신호를 주지 않는다
**우선순위**: Must / **의존**: FR-01(패키지 구조)

---

## 4. 비활성 마일스톤 (M2~M6) — 스토리 골격 (차기 정제 대상)

> rolling-wave: 아래 FR 은 활성화되는 마일스톤 착수 시점에 상세 AC 로 정제한다. 지금 AC 를 확정하지 않는다(낡음 방지).

### FR-03 슬라이드 파싱 (M2)
**스토리**: 사용자로서, PPTX 에서 슬라이드 목록과 메타데이터를 추출하고 싶다 (python-pptx 기반).
**AC**: 차기 정제 대상 (MS-2)

### FR-04 도형 파싱 (M2)
**스토리**: 사용자로서, 각 슬라이드의 텍스트·표·이미지·그룹 도형을 추출하고 싶다.
**AC**: 차기 정제 대상 (MS-2)

### FR-05 중간 IR 정의 (M2)
**스토리**: 사용자로서, 파서 출력과 어셈블러 입력을 연결하는 내부 데이터 구조를 정의하고 싶다 (dataclass 또는 TypedDict).
**AC**: 차기 정제 대상 (MS-2)

### FR-06 이미지 분류 rule-based (M3)
**스토리**: 사용자로서, 이미지를 텍스트/도표/사진/로고 4종으로 자동 분류하고 싶다 (Pillow 휴리스틱).
**AC**: 차기 정제 대상 (MS-3)

### FR-07 LibreOffice 변환 optional (M3)
**스토리**: 사용자로서, LibreOffice 가 설치된 환경에서 EMF/WMF 도형을 PNG 로 변환하고, 없으면 skip 하고 싶다 (런타임 감지).
**AC**: 차기 정제 대상 (MS-3)

### FR-08 ImageDescriber 프로토콜 (M4)
**스토리**: 사용자로서, VLM 제공자를 교체 가능하도록 추상 인터페이스(Protocol 또는 ABC)를 정의하고 싶다.
**AC**: 차기 정제 대상 (MS-4)

### FR-09 VLM 제공자 구현 (M4)
**스토리**: 사용자로서, anthropic / openai provider 를 기본 제공받고 싶다 (`[vlm]` extras).
**설명** (RFI-1 Q1 답변 반영): anthropic·openai 는 **참조 구현(reference implementation)** 으로 제공한다. 라이브러리는 FR-08 의 `ImageDescriber` Protocol(공통 인터페이스)만을 계약으로 삼으며, 사용자가 이 Protocol 을 구현한 외부 provider(예: Google Gemini 등)를 직접 작성하여 **plug-in 으로 주입**할 수 있다. 즉 anthropic/openai 외 provider 의 라이브러리 내 구현은 v1 범위 밖이며, 프로토콜 확장 가능성만 보장한다.
**AC**: 차기 정제 대상 (MS-4)

### FR-10 shape_hint 지원 (M4)
**스토리**: 사용자로서, 이미지 분류 결과를 VLM 프롬프트에 힌트로 전달하고 싶다.
**AC**: 차기 정제 대상 (MS-4)

### FR-11 Markdown 어셈블러 (M5)
**스토리**: 사용자로서, IR 을 받아 슬라이드별 Markdown 블록을 생성하고 싶다.
**AC**: 차기 정제 대상 (MS-5)

### FR-12 Map-Reduce 파이프라인 (M5)
**스토리**: 사용자로서, 슬라이드를 병렬 변환(Map)한 뒤 하나의 문서로 합치고(Reduce) 싶다.
**AC**: 차기 정제 대상 (MS-5)

### FR-13 Mermaid fallback (M5)
**스토리**: 사용자로서, 표/구조가 복잡한 경우 Mermaid 다이어그램으로 대체하고 싶다.
**AC**: 차기 정제 대상 (MS-5) — 복잡도 임계값은 **고정값 없이 M5 개발 중 실험으로 결정**한다 (RFI-1 Q2 확정). M5 착수 시 planner 가 AC 작성 시점에 확정한다.

### FR-14 validate_markdown (M5)
**스토리**: 사용자로서, 생성된 Markdown 의 구조적 유효성을 검사하고 싶다.
**AC**: 차기 정제 대상 (MS-5)

### FR-15 개인정보 마스킹 (M5)
**스토리**: 사용자로서, opt-in 옵션을 켜면 텍스트 내 개인정보를 (정규식 기반) 마스킹하고 싶다.
**설명** (RFI-1 Q4 답변 반영): 라이브러리는 **정규식 마스킹 엔진**만 제공한다. 이메일·전화번호 패턴을 **built-in defaults** 로 동봉하되, 사용자가 **커스텀 정규식 패턴을 등록**하여 마스킹 대상을 확장할 수 있다. 특정 PII 카테고리(주민번호·카드번호 등)를 라이브러리가 직접 식별·보장하지는 않으며, 사용자 커스텀 패턴으로 대응한다.
**AC**: 차기 정제 대상 (MS-5)

### FR-16 공개 API (M6)
**스토리**: 사용자로서, `convert()`, `ConvertOptions` 등 공개 인터페이스를 확정받고 싶다.
**AC**: 차기 정제 대상 (MS-6)

### FR-17 테스트 스위트 (M6)
**스토리**: 사용자로서, 커버리지 75%+ 를 달성하는 pytest 테스트 스위트를 갖추고 싶다.
**AC**: 차기 정제 대상 (MS-6)

### FR-18 README + 사용 예시 (M6)
**스토리**: 사용자로서, pip install 방법과 기본 사용 예시가 담긴 README 를 원한다.
**AC**: 차기 정제 대상 (MS-6)

### FR-19 PyPI 배포 (M6)
**스토리**: 사용자로서, `hatch publish` 로 pypi.org 에 정식 배포된 패키지를 원한다.
**AC**: 차기 정제 대상 (MS-6) — 헌장 제약상 `hatch publish` 실행은 사람이 직접 수행

---

## 5. 비기능 요구 (NFR) — 인테이크 D절 기반

| NFR | 항목 | 기준 (수치) | 근거 |
|-----|------|-------------|------|
| NFR-01 | 변환 성능 | 20슬라이드 PPTX 변환이 VLM 호출 제외 기준 p95 < 5초 (GitHub Actions ubuntu-latest 2-core 환경) | 인테이크 D "5초 이내" → **확정**: 측정 환경 GitHub Actions ubuntu-latest (2-core), p95 기준 (RFI-1 Q3 답변) |
| NFR-02 | 테스트 커버리지 | 신규 코드 라인 커버리지 ≥ 75% (`pytest --cov`) | 인테이크 R-17, 프로파일 7절 |
| NFR-03 | 타입 안전성 | `mypy src/` exit code 0 (에러 0건) | 프로파일 7절 |
| NFR-04 | 코드 스타일 | `ruff check .` + `black --check .` exit code 0 | 프로파일 3·4절 |
| NFR-05 | 보안 (비밀정보) | VLM API key 는 패키지/소스/로그에 일절 포함하지 않고 사용자가 환경변수로 제공 | 인테이크 D, 헌장 3절 |
| NFR-06 | 로깅 (개인정보) | 마스킹 옵션 활성 시 로그에 원본 텍스트 0건 출력 | 프로파일 7절 |
| NFR-07 | 런타임 호환 | Python 3.11 이상에서만 설치·동작 (3.10 이하 설치 거부) | 인테이크 F, 헌장 3절 |
| NFR-08 | 의존성 격리 | core 설치 시 VLM SDK·LibreOffice 미설치 환경에서도 import 및 비-VLM 변환이 동작 | 인테이크 F, 헌장 3절 |

> **NFR-01 확정 (RFI-1 Q3)**: 기존 "가정(p95/머신 사양 미정)" → 측정 환경을 GitHub Actions ubuntu-latest (2-core) 로 확정. 통계 기준 p95 < 5초 (VLM 호출 제외).

---

## 6. 범위 제외 (Out of Scope) — 인테이크 G절 + 헌장 2절

- GUI / 웹 UI
- PPTX 생성·편집 (변환 단방향)
- PDF, DOCX 등 PPTX 이외 포맷 지원
- 실시간 / 스트리밍 변환
- 클라우드 호스팅 서비스
- (헌장 추가) core 에 VLM API key 포함 — 사용자 환경변수 제공만 지원
- (RFI-1 Q1) anthropic/openai 외 VLM provider 의 라이브러리 내 구현 — Protocol plug-in 으로만 지원, v1 직접 구현 범위 밖
- (RFI-1 Q4) 특정 PII 카테고리(주민번호·카드번호 등) 식별 보장 — 사용자 커스텀 정규식 패턴으로 대응

---

## 7. FR 간 의존관계

| FR | 의존 대상 | 비고 |
|----|----------|------|
| FR-01 | 없음 | 기반 |
| FR-02 | FR-01 | CI 가 빌드/테스트 대상 패키지 전제 |
| FR-03 | FR-01 | |
| FR-04 | FR-03 | 슬라이드 파싱 결과를 도형 단위로 분해 |
| FR-05 | FR-03, FR-04 | 파서 출력을 담는 구조 |
| FR-06 | FR-04, FR-05 | 이미지 도형(IR) 입력 |
| FR-07 | FR-04 | EMF/WMF 도형 입력 (optional) |
| FR-08 | FR-05 | IR 의 이미지 항목을 입력으로 |
| FR-09 | FR-08 | 프로토콜 구현체 (anthropic/openai 참조 구현 + 사용자 plug-in) |
| FR-10 | FR-06, FR-08 | 분류 결과를 프롬프트 힌트로 |
| FR-11 | FR-05 | IR → Markdown |
| FR-12 | FR-09, FR-11 | Map(슬라이드별 변환)+Reduce(병합) |
| FR-13 | FR-04, FR-11 | 복잡 표/구조 대체 |
| FR-14 | FR-11 | 생성 결과 검증 |
| FR-15 | FR-11 | 텍스트 마스킹 (opt-in) |
| FR-16 | FR-12 | 파이프라인을 공개 API 로 노출 |
| FR-17 | FR-03~FR-16 | 전 기능 대상 테스트 |
| FR-18 | FR-16 | 공개 API 기준 사용 예시 |
| FR-19 | FR-01, FR-17, FR-18 | 빌드·테스트·문서 완비 후 배포 |

---

## 8. RFI 질문 목록 (인테이크 H절 `?` + 분석 중 발견)

> **회차 1 (RFI-1) 전 4건 모두 답변 완료** (2026-06-27, `docs/10-requirements/intake/RFI-1.md`). 본문 해당 FR/NFR 에 반영 완료. 진행 차단급 없음.

### RFI-1 (인테이크 H절 `?`) — 추가 VLM provider 지원 범위 — [해결됨]
- **질문**: VLM 제공자를 anthropic/openai 외 Google Gemini 등으로 확장할 계획이 있는가? 있다면 v1 범위인가 차기인가?
- **답변 (RFI-1 Q1 / C 확정)**: M4 에서 `ImageDescriber` Protocol 만 정의하고 anthropic·openai 는 참조 구현으로 제공. 사용자가 직접 구현체를 만들어 plug-in 가능. → FR-09 설명·Out of Scope 반영 완료.

### RFI-2 (인테이크 H절 `?`) — Mermaid fallback 복잡도 임계값 — [해결됨]
- **질문**: FR-13 Mermaid fallback 을 발동하는 "복잡도" 의 판정 기준은? (예: 표 셀 수, 병합 셀 존재, 중첩 그룹 깊이)
- **답변 (RFI-1 Q2 / C 확정)**: 고정 임계값 없이 M5 개발 중 실험으로 결정. M5 착수 전 planner 가 AC 작성 시점에 확정. → FR-13 AC 비고 반영 완료.

### RFI-3 (분석 발견 — NFR-01) — 성능 기준의 측정 조건 — [해결됨]
- **질문**: NFR-01 "20슬라이드 5초"의 측정 기준 머신 사양(CPU/메모리)과 통계 기준(평균 vs p95)은?
- **답변 (RFI-1 Q3 / A 확정)**: GitHub Actions ubuntu-latest (2-core) 기준, VLM 제외 5초 이내. → NFR-01 기준·근거를 "확정" 으로 업데이트 완료.

### RFI-4 (분석 발견 — FR-15) — 개인정보 마스킹 대상 카테고리 — [해결됨]
- **질문**: 정규식 마스킹이 다룰 개인정보 카테고리는 무엇인가? (예: 이메일, 전화번호, 주민번호, 카드번호)
- **답변 (RFI-1 Q4 / C 확정)**: 라이브러리는 정규식 마스킹 엔진만 제공. 이메일·전화번호 패턴을 built-in defaults 로 동봉하되 사용자 커스텀 패턴 등록 가능. → FR-15 설명·Out of Scope 반영 완료.

---

## 9. 제안 (P-) — 인테이크에 없음, 사람 승인 후에만 FR 승격

> 아래는 발명이 아니라 분석 중 떠오른 보완 제안. FR 목록에 섞지 않음. 승인 전까지 비공식.

- **P-01 변환 캐시**: 동일 PPTX/옵션 재변환 시 VLM 호출 결과 캐시로 비용·시간 절감. (근거: VLM 호출이 성능·과금 병목 가능 — 인테이크 D)
- **P-02 부분 실패 리포트**: 특정 슬라이드/도형 변환 실패 시 전체 중단 대신 해당 블록만 경고 표기하고 계속 진행하는 모드. (근거: 헌장 리스크 "python-pptx 미지원 요소 누락 가능")
- **P-03 CLI 진입점**: `pptx-md convert file.pptx` 형태 콘솔 스크립트. (근거: 라이브러리 사용성 향상 — 단, 인테이크 G "GUI 제외"와는 별개 영역이나 범위 확장이므로 승인 필요)
