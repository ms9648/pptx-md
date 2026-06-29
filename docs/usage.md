# 사용 가이드

pptx-md는 PowerPoint(PPTX) 파일을 LLM-friendly Markdown으로 변환하는 라이브러리입니다.
이 문서는 설치부터 고급 사용법까지 단계별로 안내합니다.

---

## 설치

### Core 설치 (텍스트 변환)

```bash
pip install pptx-md
```

### VLM 지원 설치 (이미지 설명 생성)

```bash
pip install pptx-md[vlm]
```

VLM extras는 `anthropic` 및 `openai` SDK를 포함합니다.

---

## 기본 사용

### 최소 예시

```python
from pptx_md import convert

md = convert("presentation.pptx")
print(md)
```

`convert()`는 파싱 → 이미지 분류 → Markdown 어셈블 파이프라인을 자동으로 실행합니다.

### `pathlib.Path` 사용

```python
from pathlib import Path
from pptx_md import convert

path = Path("/data/slides/deck.pptx")
md = convert(path)
```

### 결과 파일 저장

```python
from pptx_md import convert

md = convert("deck.pptx")
with open("output.md", "w", encoding="utf-8") as f:
    f.write(md)
```

---

## ConvertOptions — 변환 옵션

`ConvertOptions`는 변환 파이프라인의 모든 옵션을 하나의 불변 객체로 집약합니다.

```python
from pptx_md import convert, ConvertOptions

opts = ConvertOptions(
    describer=None,   # VLM 제공자 (기본값: None = VLM 미사용)
    masking=None,     # PII 마스킹 옵션 (기본값: None = 마스킹 안 함)
    validate=False,   # Markdown 검증 여부 (기본값: False)
)
md = convert("deck.pptx", options=opts)
```

`ConvertOptions`는 `frozen=True` dataclass이므로 생성 후 재사용이 안전합니다.

---

## VLM 이미지 설명

VLM(Vision Language Model)을 연동하면 이미지 슬라이드에 자연어 설명을 생성합니다.

### Anthropic Claude

```python
import os
from pptx_md import convert, ConvertOptions, get_describer

describer = get_describer("anthropic", api_key=os.environ["ANTHROPIC_API_KEY"])
opts = ConvertOptions(describer=describer)
md = convert("deck.pptx", options=opts)
```

### OpenAI GPT-4o

```python
import os
from pptx_md import convert, ConvertOptions, get_describer

describer = get_describer("openai", api_key=os.environ["OPENAI_API_KEY"])
opts = ConvertOptions(describer=describer)
md = convert("deck.pptx", options=opts)
```

API 키는 **반드시 환경변수**로 전달해야 합니다. 코드에 평문으로 포함하지 마세요.

### 커스텀 VLM 제공자

`ImageDescriber` 프로토콜을 구현하면 어떤 VLM 제공자도 사용할 수 있습니다.

```python
from pptx_md import convert, ConvertOptions, ImageDescriber


class GeminiDescriber:
    def describe(
        self,
        image_bytes: bytes,
        image_ext: str,
        shape_hint: str | None,
    ) -> str:
        # Gemini API 호출 로직
        return "이미지 설명"


opts = ConvertOptions(describer=GeminiDescriber())
md = convert("deck.pptx", options=opts)
```

---

## 개인정보 마스킹

마스킹 기능은 **기본값 비활성(opt-in)**입니다. 활성화하면 이메일·전화번호 등을
`[REDACTED]`로 치환합니다.

### 기본 패턴 활성화

```python
from pptx_md import convert, ConvertOptions, MaskingOptions

opts = ConvertOptions(masking=MaskingOptions(enabled=True))
md = convert("deck.pptx", options=opts)
```

기본 내장 패턴:
- 이메일 주소 (RFC 5321)
- 한국 전화번호 (010-xxxx-xxxx 등)

### 커스텀 패턴 추가

```python
import re
from pptx_md import convert, ConvertOptions, MaskingOptions

masking = MaskingOptions(
    enabled=True,
    patterns=[
        re.compile(r"\d{6}-\d{7}"),    # 주민등록번호
        re.compile(r"사번\s*:\s*\d+"),  # 사번
        re.compile(r"PAT-\w{16,}"),    # 개인 액세스 토큰
    ],
)
opts = ConvertOptions(masking=masking)
md = convert("deck.pptx", options=opts)
```

마스킹 치환 토큰(`[REDACTED]`)은 `MASK_TOKEN` 상수로 참조할 수 있습니다.

```python
from pptx_md import MASK_TOKEN

print(MASK_TOKEN)  # "[REDACTED]"
```

---

## Markdown 검증

### 변환 후 독립 검증

```python
from pptx_md import convert, validate_markdown

md = convert("deck.pptx")
result = validate_markdown(md)

print(result.valid)     # True / False
print(result.warnings)  # 경고 목록
```

### 변환과 동시에 검증 로깅

```python
import logging
from pptx_md import convert, ConvertOptions

logging.basicConfig(level=logging.INFO)

opts = ConvertOptions(validate=True)
md = convert("deck.pptx", options=opts)
# 검증 결과는 pptx_md.api 로거로 INFO/WARNING 출력
# 반환값은 항상 str
```

---

## 예외 처리

| 예외 | 발생 조건 |
|------|----------|
| `ParseError` | PPTX 파일이 없거나 손상됨 (file-level fail-fast) |
| `DescribeError` | VLM 제공자 API 호출 실패 |
| `InstallationError` | VLM SDK 미설치 상태에서 provider 생성 시도 |

```python
from pptx_md import convert, ParseError, DescribeError, InstallationError

try:
    md = convert("deck.pptx")
except ParseError as e:
    print(f"파일을 읽을 수 없습니다: {e}")
except InstallationError as e:
    print(f"VLM SDK가 설치되지 않았습니다: {e}")
    print("힌트: pip install pptx-md[vlm]")
```

도형·이미지 단위 실패는 격리되어 부분 산출로 처리되며 예외가 전파되지 않습니다.

---

## 고급: 내부 모듈 직접 접근

공개 API로 해결할 수 없는 경우 내부 모듈을 직접 import할 수 있습니다.
단, **SemVer 보증 범위 밖**이므로 버전 업데이트 시 깨질 수 있습니다.

```python
# SemVer 보증 밖 — 주의해서 사용
from pptx_md.parser import parse_presentation
from pptx_md.assembler import assemble_document
```

공개 API 목록은 [api.md](api.md)를 참고하세요.

---

## 요구사항

- Python 3.11+
- Core: `python-pptx`, `Pillow`
- VLM: `pip install pptx-md[vlm]`
