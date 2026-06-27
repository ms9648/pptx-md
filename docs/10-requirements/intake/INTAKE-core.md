# 요구사항 인테이크 — pptx-md (core)
> 작성: ms9648 / 2026-06-27 / 회수 후 PM·planner 가 분석하여 질문 회신서를 드립니다.
>
> 작성 안내:
> - 모르는 항목은 비워두지 말고 `?` 로 표기하세요. 빈칸과 `?` 는 다르게 취급됩니다
>   (`?` = 함께 정해야 할 항목, 빈칸 = 해당 없음).
> - 완벽한 문장이 아니어도 됩니다. 키워드·예시도 좋습니다.
> - 우선순위: M(필수) / S(중요) / C(있으면 좋음) / ?(미정)
> - **작성을 마치면 Claude Code 에서 `/cb:intake` 를 실행하세요** — 이 파일을
>   회수해 분석을 시작합니다.

## A. 배경과 목표
- 이 시스템이 해결하려는 문제: PPTX 파일을 LLM 에 입력할 때 이미지·구조 정보가 소실되는 문제
- 성공의 모습: `pptx_md.convert("file.pptx")` 한 줄로 슬라이드 전체가 구조화된 Markdown 으로 변환됨. 이미지는 VLM 설명 텍스트로 포함됨.
- 기존 시스템/프로세스: ssine/pptx2md 등 기존 도구는 이미지를 경로만 저장 — 신규 설계

## B. 사용자와 액터
| 액터 | 누구인가 | 대략 인원 | 주로 하는 일 |
|------|---------|----------|-------------|
| 라이브러리 사용자 | Python 개발자, LLM 앱 빌더 | 불특정 다수 | pip install 후 코드에서 convert() 호출 |

## C. 기능 요구 목록  ← 가장 중요한 표. 행은 자유롭게 추가
| ID | 기능명 | 설명 (누가 무엇을 왜) | 우선순위 | 비고/예시 |
|----|--------|----------------------|---------|----------|
| R-01 | 패키지 초기화 | 개발자가 pip install 가능한 구조로 패키지를 구성한다 | M | pyproject.toml, src layout, MIT |
| R-02 | CI 파이프라인 | PR 마다 lint+type+test 가 자동 실행된다 | M | GitHub Actions |
| R-03 | 슬라이드 파싱 | PPTX 파일에서 슬라이드 목록과 메타데이터를 추출한다 | M | python-pptx 기반 |
| R-04 | 도형 파싱 | 각 슬라이드의 텍스트·표·이미지·그룹 도형을 추출한다 | M | |
| R-05 | 중간 IR 정의 | 파서 출력과 어셈블러 입력을 연결하는 내부 데이터 구조를 정의한다 | M | dataclass 또는 TypedDict |
| R-06 | 이미지 분류 (rule-based) | 이미지를 텍스트/도표/사진/로고 4종으로 자동 분류한다 | M | Pillow 기반 휴리스틱 |
| R-07 | LibreOffice 변환 (optional) | LibreOffice 가 설치된 환경에서 EMF/WMF 도형을 PNG 로 변환한다 | S | 런타임 감지, 없으면 skip |
| R-08 | ImageDescriber 프로토콜 | VLM 제공자를 교체 가능하도록 추상 인터페이스를 정의한다 | M | Protocol 또는 ABC |
| R-09 | VLM 제공자 구현 | anthropic / openai provider 를 기본 제공한다 | M | [vlm] extras |
| R-10 | shape_hint 지원 | 이미지 유형(분류 결과)을 VLM 프롬프트에 힌트로 전달한다 | M | 분류 결과 → 프롬프트 파라미터 |
| R-11 | Markdown 어셈블러 | IR 을 받아 슬라이드별 Markdown 블록을 생성한다 | M | |
| R-12 | Map-Reduce 파이프라인 | 슬라이드를 병렬 변환(Map)한 뒤 하나의 문서로 합친다(Reduce) | M | |
| R-13 | Mermaid fallback | 표/구조가 복잡한 경우 Mermaid 다이어그램으로 대체한다 | S | |
| R-14 | validate_markdown | 생성된 Markdown 의 구조적 유효성을 검사한다 | S | |
| R-15 | 개인정보 마스킹 | 사용자가 opt-in 옵션을 켜면 텍스트 내 개인정보를 마스킹한다 | S | |
| R-16 | 공개 API | convert(), ConvertOptions 등 공개 인터페이스를 확정한다 | M | |
| R-17 | 테스트 스위트 | 커버리지 75%+ 를 달성하는 pytest 테스트를 작성한다 | M | |
| R-18 | README + 사용 예시 | pip install 방법과 기본 사용 예시를 README 에 작성한다 | M | |
| R-19 | PyPI 배포 | hatch publish 로 pypi.org 에 정식 배포한다 | M | |

## D. 비기능 요구 (아는 만큼만)
- 사용자 규모/동시접속: 라이브러리 (동시접속 개념 없음)
- 성능 기대: 20슬라이드 PPTX 변환이 VLM 제외 기준 5초 이내 (추정)
- 보안/권한: VLM API key 는 라이브러리에 포함하지 않음 — 사용자가 환경변수로 제공
- 가용 시간대: 해당 없음 (로컬 실행 라이브러리)

## E. 연동과 데이터
- 연동 시스템: VLM API (anthropic, openai) — 옵셔널
- LibreOffice — 옵셔널 (런타임 감지)
- 마이그레이션: 없음

## F. 제약과 일정
- Python 3.11+ 필수
- 고정 마감 없음 (PR 완료 기준 유동)
- VLM SDK 는 반드시 [vlm] extras 로만 선언

## G. 명시적 범위 제외 (이번에 안 하는 것)
- GUI / 웹 UI
- PPTX 생성·편집
- PDF, DOCX 등 PPTX 이외 포맷
- 실시간/스트리밍 변환
- 클라우드 호스팅 서비스

## H. 미정/논의 필요 사항 (자유 기술)
- VLM 제공자 외 추가 provider (예: Google Gemini) 지원 여부: ?
- Mermaid fallback 의 적용 기준(복잡도 임계값): ?
- 개인정보 마스킹 방식 상세: 정규식 기반 (별도 라이브러리 없음)
