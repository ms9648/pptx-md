# pptx-md

[![CI](https://github.com/ms9648/pptx-md/actions/workflows/ci.yml/badge.svg)](https://github.com/ms9648/pptx-md/actions/workflows/ci.yml)
[![PyPI version](https://img.shields.io/pypi/v/pptx-md.svg)](https://pypi.org/project/pptx-md/)
[![Python](https://img.shields.io/pypi/pyversions/pptx-md.svg)](https://pypi.org/project/pptx-md/)

PPTX 파일을 LLM-friendly Markdown으로 변환하는 Python 라이브러리.

VLM(Vision Language Model) 기반 이미지 이해와 개인정보 마스킹을 지원합니다.

---

## 설치

### Core (텍스트 변환만)

```bash
pip install pptx-md
```

### VLM 지원 포함 (이미지 설명 생성)

```bash
pip install pptx-md[vlm]
```

VLM extras는 `anthropic` 및 `openai` SDK를 함께 설치합니다.

---

## Quick Start

### 기본 변환 (core-only)

```python
from pptx_md import convert

md = convert("deck.pptx")
print(md)
```

`convert()`는 파싱 → 이미지 분류 → Markdown 어셈블 파이프라인을 실행하고
Markdown 문자열을 반환합니다.

### 옵션과 함께

```python
from pptx_md import convert, ConvertOptions

opts = ConvertOptions(validate=True)
md = convert("deck.pptx", options=opts)
```

---

## VLM 이미지 설명

VLM을 사용하면 이미지 슬라이드에 자연어 설명을 자동으로 생성합니다.
API 키는 **반드시 환경변수**로 전달합니다 (NFR-05).

```python
import os
from pptx_md import convert, ConvertOptions, get_describer

describer = get_describer("anthropic", api_key=os.environ["ANTHROPIC_API_KEY"])
opts = ConvertOptions(describer=describer)
md = convert("deck.pptx", options=opts)
```

OpenAI를 사용하려면:

```python
import os
from pptx_md import convert, ConvertOptions, get_describer

describer = get_describer("openai", api_key=os.environ["OPENAI_API_KEY"])
opts = ConvertOptions(describer=describer)
md = convert("deck.pptx", options=opts)
```

---

## 개인정보 마스킹 (opt-in)

이메일·전화번호 등 PII를 `[REDACTED]`로 치환합니다. 기본값은 **비활성**입니다.

### 기본 패턴 활성화

```python
from pptx_md import convert, ConvertOptions, MaskingOptions

opts = ConvertOptions(masking=MaskingOptions(enabled=True))
md = convert("deck.pptx", options=opts)
```

### 커스텀 패턴 추가

```python
import re
from pptx_md import convert, ConvertOptions, MaskingOptions

custom_masking = MaskingOptions(
    enabled=True,
    patterns=[
        re.compile(r"\d{6}-\d{7}"),   # 주민등록번호
        re.compile(r"사번\s*:\s*\d+"), # 사번
    ],
)
opts = ConvertOptions(masking=custom_masking)
md = convert("deck.pptx", options=opts)
```

---

## Markdown 검증

```python
from pptx_md import convert, validate_markdown

md = convert("deck.pptx")
result = validate_markdown(md)

if not result.valid:
    print("검증 실패:", result.warnings)
elif result.warnings:
    print("경고:", result.warnings)
```

`convert(validate=True)`를 사용하면 변환과 동시에 검증 결과를 로그로 출력합니다
(반환값은 항상 `str`).

---

## 커스텀 VLM 제공자 (플러그인)

`ImageDescriber` 프로토콜을 구현하면 어떤 VLM 제공자도 플러그인으로 사용할 수
있습니다.

```python
from pptx_md import convert, ConvertOptions, ImageDescriber


class MyDescriber:
    def describe(
        self,
        image_bytes: bytes,
        image_ext: str,
        shape_hint: str | None,
    ) -> str:
        return "이미지에 대한 설명"


opts = ConvertOptions(describer=MyDescriber())
md = convert("deck.pptx", options=opts)
```

---

## 예외 처리

```python
from pptx_md import convert, ParseError, DescribeError

try:
    md = convert("deck.pptx")
except ParseError as e:
    print(f"PPTX 파일을 읽을 수 없습니다: {e}")
```

---

## 내부 모듈 직접 접근 (SemVer 보증 밖)

```python
# 아래 import는 가능하지만 공개 API가 아니므로
# SemVer 보증 범위 밖입니다. 버전 업데이트 시 깨질 수 있습니다.
from pptx_md.parser import parse_presentation
```

공개 API 목록은 [docs/api.md](docs/api.md)를 참고하세요.

---

## 전체 API 레퍼런스

[docs/api.md](docs/api.md) 에 공개 심볼 전체 레퍼런스가 있습니다.

상세 사용 가이드는 [docs/usage.md](docs/usage.md)를 참고하세요.

---

## 요구사항

- Python 3.11+
- Core: `python-pptx`, `Pillow`
- VLM 지원: `pip install pptx-md[vlm]` (`anthropic` 또는 `openai` SDK)

---

## 라이선스

[MIT](LICENSE)
