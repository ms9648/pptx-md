# API 레퍼런스

`pptx_md` 패키지의 공개 심볼 레퍼런스입니다.

```python
from pptx_md import (
    convert,
    ConvertOptions,
    MaskingOptions,
    MASK_TOKEN,
    validate_markdown,
    ValidationResult,
    ImageDescriber,
    get_describer,
    PptxMdError,
    ParseError,
    DescribeError,
    InstallationError,
)
```

공개 API 이외의 심볼(예: `parse_presentation`, `enrich_images`)은 내부 모듈로,
SemVer 보증 범위 밖입니다.

---

## 함수

### `convert`

```python
def convert(
    source: str | os.PathLike[str],
    *,
    options: ConvertOptions | None = None,
) -> str
```

PPTX 파일을 Markdown 문서로 변환합니다 (FR-16).

M2~M5 파이프라인(parse → enrich_images → enrich_descriptions → assemble)을
단일 호출로 오케스트레이션합니다.

**매개변수**

| 이름 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `source` | `str \| os.PathLike[str]` | — | 변환할 `.pptx` 파일 경로 |
| `options` | `ConvertOptions \| None` | `None` | 변환 옵션; `None`이면 기본값(VLM 없음, 마스킹 없음, 검증 없음) 사용 |

**반환값**: `str` — Markdown 문서 전체.

**예외**

- `ParseError`: 파일이 없거나 유효한 PPTX가 아닐 때 (file-level fail-fast).
  도형·이미지 단위 실패는 격리되어 예외가 전파되지 않음.

**예시**

```python
from pptx_md import convert

md = convert("deck.pptx")
```

---

### `validate_markdown`

```python
def validate_markdown(md: str) -> ValidationResult
```

Markdown 문자열의 구조적 무결성을 검증합니다 (FR-14).

**검사 항목** (순서 고정):
1. 빈 문서 — `valid=False`
2. 닫히지 않은 코드 펜스 — `valid=False`
3. 헤딩 순서 이상(h1/h2 미시작, 레벨 점프) — 경고(`valid` 유지)

**매개변수**

| 이름 | 타입 | 설명 |
|------|------|------|
| `md` | `str` | 검증할 Markdown 문자열 |

**반환값**: `ValidationResult` — 검증 결과. 절대 예외를 발생시키지 않음.

**예시**

```python
from pptx_md import validate_markdown

result = validate_markdown("# 제목\n\n본문")
assert result.valid
```

---

### `get_describer`

```python
def get_describer(name: str, **config: object) -> ImageDescriber
```

이름으로 VLM 제공자 인스턴스를 생성해 반환합니다 (FR-09).

SDK는 지연 import되므로 미설치 제공자의 SDK 부재가 다른 제공자 사용에 영향을 주지
않습니다.

**매개변수**

| 이름 | 타입 | 설명 |
|------|------|------|
| `name` | `str` | 제공자 이름: `"anthropic"` 또는 `"openai"` |
| `**config` | `object` | 제공자 생성자에 전달할 키워드 인자 (예: `api_key`, `model`) |

**반환값**: `ImageDescriber` 프로토콜을 만족하는 인스턴스.

**예외**

- `ValueError`: 알 수 없는 제공자 이름.
- `InstallationError`: 해당 SDK가 설치되어 있지 않을 때.

**예시**

```python
import os
from pptx_md import get_describer

describer = get_describer("anthropic", api_key=os.environ["ANTHROPIC_API_KEY"])
```

---

## 클래스 / 데이터클래스

### `ConvertOptions`

```python
@dataclass(frozen=True)
class ConvertOptions:
    describer: ImageDescriber | None = None
    masking: MaskingOptions | None = None
    validate: bool = False
```

`convert()` 변환 옵션을 집약하는 불변 값 객체 (FR-16, ADR-602).

**필드**

| 이름 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `describer` | `ImageDescriber \| None` | `None` | VLM 제공자. `None`이면 VLM 미사용 (NFR-08) |
| `masking` | `MaskingOptions \| None` | `None` | PII 마스킹 옵션. `None`이면 마스킹 안 함 (FR-15 opt-in) |
| `validate` | `bool` | `False` | `True`이면 변환 후 `validate_markdown()` 실행 (로깅만, 반환값은 항상 `str`) |

frozen dataclass이므로 여러 번 재사용해도 안전합니다.

---

### `MaskingOptions`

```python
@dataclass
class MaskingOptions:
    enabled: bool = False
    patterns: list[re.Pattern[str]] = field(default_factory=...)
```

개인정보 마스킹 옵션 (FR-15).

**필드**

| 이름 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `enabled` | `bool` | `False` | `True`이면 마스킹 활성화 |
| `patterns` | `list[re.Pattern[str]]` | 기본 이메일·전화번호 패턴 | 적용할 정규식 패턴 목록 |

**예시**

```python
import re
from pptx_md import MaskingOptions

masking = MaskingOptions(
    enabled=True,
    patterns=[re.compile(r"\d{6}-\d{7}")],
)
```

---

### `ValidationResult`

```python
@dataclass
class ValidationResult:
    valid: bool
    warnings: list[str] = field(default_factory=list)
```

`validate_markdown()` 반환값.

**필드**

| 이름 | 타입 | 설명 |
|------|------|------|
| `valid` | `bool` | `False`: 빈 문서 또는 닫히지 않은 코드 펜스 발견 |
| `warnings` | `list[str]` | 경고 메시지 목록 (헤딩 순서 이상 등). `valid=True`여도 비어있지 않을 수 있음 |

---

## 프로토콜

### `ImageDescriber`

```python
@runtime_checkable
class ImageDescriber(Protocol):
    def describe(
        self,
        image_bytes: bytes,
        image_ext: str,
        shape_hint: str | None,
    ) -> str: ...
```

VLM 이미지 설명 제공자 프로토콜 (FR-08, ADR-217).

명목 상속 없이 구조적 타이핑으로 만족됩니다.
외부 VLM 제공자를 플러그인으로 사용하려면 이 메서드를 구현하세요.

**`describe` 매개변수**

| 이름 | 타입 | 설명 |
|------|------|------|
| `image_bytes` | `bytes` | 이미지 원본 바이트 (PNG, JPEG 등) |
| `image_ext` | `str` | 이미지 확장자 힌트 (`"png"`, `"jpeg"` 등) |
| `shape_hint` | `str \| None` | 이미지 분류 힌트. `None`이면 제공자가 범용 프롬프트 사용 |

**반환값**: 자연어 설명 문자열. 빈 문자열은 허용되나 설명 실패를 의미합니다.

**예외**: `DescribeError` — API/SDK 실패 시.

---

## 예외 계층

### `PptxMdError`

모든 pptx-md 예외의 베이스 클래스.

```python
class PptxMdError(Exception): ...
```

---

### `ParseError`

```python
class ParseError(PptxMdError): ...
```

PPTX 파일이 없거나 파싱할 수 없을 때 발생합니다 (file-level fail-fast).

`convert()`에서만 전파됩니다. 도형·이미지 단위 실패는 격리됩니다.

---

### `DescribeError`

```python
class DescribeError(PptxMdError): ...
```

VLM 제공자 API 호출이 실패할 때 발생합니다 (FR-09).
`convert()` 내에서는 이미지 단위로 격리되어 부분 산출로 이어집니다.

---

### `InstallationError`

```python
class InstallationError(PptxMdError): ...
```

VLM SDK가 설치되지 않은 환경에서 provider를 생성하려 할 때 발생합니다.

메시지에 `pip install pptx-md[vlm]` 안내가 포함됩니다.

---

## 상수

### `MASK_TOKEN`

```python
MASK_TOKEN: str = "[REDACTED]"
```

PII 마스킹 치환 토큰. `MaskingOptions`가 활성화된 경우 모든 패턴 매치를
이 값으로 대체합니다.
