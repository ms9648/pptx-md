# 프로젝트 헌장 — pptx-md

> 승인: ms9648 / 변경은 변경 요청 이슈로만.

## 1. 목표

PPTX 파일을 LLM 이 이해하기 쉬운 Markdown 으로 변환하는 Python 라이브러리를 개발하고,
PyPI 에 오픈소스(MIT)로 배포한다.

**핵심 차별점**: 기존 도구(ssine/pptx2md 등)는 이미지를 파일 경로로만 저장한다.
pptx-md 는 VLM-first 설계로 이미지의 의미(텍스트, 다이어그램, 사진 등)까지
Markdown 에 포함한다.

## 2. 범위

### In Scope
- PPTX → Markdown 변환 (슬라이드, 텍스트, 표, 이미지)
- 이미지 분류 (rule-based 4종: 텍스트/도표/사진/로고 + LibreOffice optional)
- VLM 연동 ImageDescriber (anthropic, openai 등 — `[vlm]` extras)
- 개인정보 마스킹 옵션 (opt-in)
- PyPI 배포 (`pip install pptx-md`)
- GitHub Actions CI (lint + type-check + test)

### Out of Scope
- GUI / 웹 UI
- PPTX 생성·편집 (변환 단방향)
- PDF, DOCX 등 PPTX 이외 포맷 지원
- 실시간/스트리밍 변환

## 3. 제약

- Python 3.11+ (하위 버전 지원 없음)
- 외부 의존성 추가는 사람 승인 필수 (NEEDS_DECISION)
- VLM SDK 는 반드시 옵셔널 extras 처리 (core 에 API key 불포함)
- `hatch publish` (PyPI 배포)는 사람이 직접 실행

## 4. 마일스톤 골격

| 마일스톤 | 목표 | Done 기준 |
|---------|------|----------|
| M1 | 패키지 구조 초기화 | pyproject.toml, src layout, CI 파이프라인 green |
| M2 | 코어 파서 | python-pptx 기반 슬라이드·도형 파싱 + 중간 IR 정의, 테스트 75%+ |
| M3 | 이미지 분류기 | rule-based 4종 분류 동작, LibreOffice optional 플래그 |
| M4 | VLM 연동 | ImageDescriber 프로토콜 + provider별 구현, mock 테스트 통과 |
| M5 | Markdown 어셈블러 | Map-Reduce 파이프라인, Mermaid fallback, validate_markdown |
| M6 | 릴리스 | 커버리지 75%+, README + docs/, PyPI 배포 완료 |

일정: PR 완료 기준 유동 (고정 마감 없음)

## 5. Done 의 정의 (프로젝트 전체)

- 모든 마일스톤 이슈가 Done 상태
- PyPI 에 `pptx-md` 패키지가 정상 배포되고 `pip install pptx-md` 로 설치 가능
- README 에 기본 사용 예시 포함
- CI (lint + mypy + pytest) 모두 green

## 6. 가정 및 리스크

| 구분 | 내용 |
|------|------|
| 가정 | VLM API key 는 사용자가 직접 제공 (라이브러리에 포함 안 함) |
| 가정 | LibreOffice 설치 여부는 런타임 감지 (optional) |
| 리스크 | VLM provider API 변경 시 ImageDescriber 인터페이스 영향 가능 |
| 리스크 | python-pptx 가 지원하지 않는 PPTX 요소는 변환 누락 가능 |
