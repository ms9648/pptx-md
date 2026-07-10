# 요구사항 명세 — pptx-md v0.1.2 Wave 3 (구조·품질 게이트)

> 원천: `.upload/v011-과업.md` §Wave 3 (M12 헤딩 계층 & 문서 메타데이터 / M13 Validator 강화 & 골든 회귀 게이트)
> 정제: planner / 2026-07-05
> 전제: `docs/00-charter/project-profile.md`, `docs/10-requirements/REQ-core.md`(FR-01~19 베이스라인), `docs/20-design/ARCH-M12.md`(FR-27 설계 서식·ADR 컨벤션)
> 골든 픽스처: `.upload/output_no_vlm.md`(현행 산출), DoD 표(§DoD)
> 상태: **베이스라인 확정 (2026-07-05, ms9648 승인)**. §7 결정 로그 전건 확정
> 번호 규약: 마지막 FR 은 FR-27 → 신규는 **FR-28부터**. ADR 은 ARCH-M12 의 ADR-613 이후로 연속(설계 단계에서 ADR-614~ 부여 예정, 본 문서는 요구만 정의)

---

## 0. 전제·불변식 (모든 신규 FR 공통)

Wave 3 는 **동작하는 어셈블러·validator 골격을 바꾸지 않고** 구조화 출력과 품질 게이트를 얹는다. 아래 불변식은 모든 AC 의 암묵 전제이며 위반 시 그 AC 는 미충족으로 본다.

- **INV-1 결정성(ADR-218 계승)**: 동일 IR → 동일 Markdown. 신규 로직에 난수·set/dict 순회 의존·시각 의존 0. `assemble_document` 는 슬라이드 인덱스 순 병합(ADR-220) 유지.
- **INV-2 IR 하위호환**: `ir.py` 에 필드 추가 시 **항상 default 유지**(ADR-202: `""`/`0`/`None`, ADR-205 확장 슬롯). 기존 호출·기존 테스트 무손상. (D-3 확정: Wave 3 는 IR 무변경)
- **INV-3 옵션 하위호환**: 신규 `ConvertOptions` 필드는 frozen dataclass 에 **default 추가**(ADR-602). 기본값은 **현행 출력과 동등**(신규 구조화는 전부 opt-in). `ConvertOptions()`·`convert(src)` 무손상.
- **INV-4 격리(ADR-204/214 계승)**: 슬라이드/도형 1개 실패가 전체 변환을 중단시키지 않는다. validator 는 무예외(항상 `ValidationResult` 반환, FR-14 AC7 계승).
- **INV-5 의존 격리(NFR-08 계승)**: `assembler.py`·`validator.py` 최상단 VLM SDK·python-pptx·Pillow import 0. 신규 import 는 stdlib(`re` 등)·내부 모듈만.
- **INV-6 프로파일 준수**: 신규 외부 라이브러리 도입 0(프로파일 §1). frontmatter/YAML 도 **stdlib 문자열 조립**으로 산출(D-1 확정: PyYAML 미도입).

---

## 1. 기능 요구(FR) 추적성 매핑

| FR | 기능명 | 원천(과업지서) | 마일스톤 | 우선순위 | 정제 상태 |
|----|--------|----------------|----------|----------|-----------|
| FR-28 | 헤딩 계층 & 문서 메타데이터 | Wave 3 §M12 (Agent A/B/C) | M13 | Should | 상세 AC |
| FR-29 | Validator 강화 & 골든 회귀 게이트 | Wave 3 §M13 (Agent A/B/C) | M14 | Must | 상세 AC |

> GitHub 마일스톤 정정: M11=FR-26, M12=FR-27 이 이미 존재하므로 Wave 3 는 **M13(FR-28)/M14(FR-29)**. (과업지서의 "M12/M13" 라벨은 문서 내부 절 번호이며 GitHub 마일스톤과 별개)
> 전수 커버: 과업지서 Wave 3 의 M12·M13 전 레인(각 3레인 + Reviewer) → FR-28·FR-29 로 흡수, 누락 0.
> 스코프 과대 시 하위 분할은 §5(하위 분할 제안) 참조 — PM 이 이슈 분해 시 채택 판단.

---

## 2. FR-28 헤딩 계층 & 문서 메타데이터 (M13)

**스토리**: RAG 파이프라인 빌더로서, 변환된 Markdown 이 `## Slide N` 평면 나열이 아니라 `>` 경로가 `##/###` 계층으로 분해되고 슬라이드별 메타(slide_index/section/has_diagram)와 발표자 노트를 담기를 원한다. 청크 단위 검색·추적과 문서 목차 복원을 위해서다.

**신규 `ConvertOptions` 필드** (D-2 확정: bool 4개 유지. 전부 opt-in, 기본값 = 현행 출력 유지, INV-3):

| 필드 | 타입 | 기본값 | 의미 |
|------|------|--------|------|
| `heading_hierarchy` | `bool` | `False` | 제목 내 경로 구분자(`>`)를 `##/###…` 계층 헤딩으로 분해(AC1~AC3) |
| `emit_frontmatter` | `bool` | `False` | 문서 최상단 YAML frontmatter 1회 + 슬라이드별 HTML 주석 메타(`slide_index`/`section`/`has_diagram`) 출력(AC5~AC7) |
| `emit_toc` | `bool` | `False` | 빈 구분(섹션 표지) 슬라이드를 TOC 항목/`---` 구분으로 변환(AC4) |
| `include_notes` | `bool` | `False` | `SlideIR.notes` 를 옵션 `> notes` 블록으로 출력(AC8~AC9) |

> D-3 확정: `section`/`has_diagram` 은 IR 확장 없이 **어셈블 시점 파생**한다(IR 무변경, INV-2). `has_diagram` 은 슬라이드 내 `ImageShapeIR.classification is ImageClass.DIAGRAM` 존재 여부로, `section` 은 현재 최상위 `##` 헤딩 텍스트로 계산.

**AC**:

*헤딩 계층 (Agent A)*
- [ ] AC1: Given `heading_hierarchy=True` 이고 슬라이드 제목이 경로 구분자를 포함(`"Ⅱ. 사업의 이해 > ⑦ 현황·문제점"`) When 어셈블 Then 최상위 세그먼트는 `##`, 후속 세그먼트는 `###`(그 다음은 `####`…)으로 분해되어 각 세그먼트가 별도 헤딩 라인이 된다. 세그먼트 텍스트는 앞뒤 공백 trim.
- [ ] AC2: Given `heading_hierarchy=True` 이고 제목에 구분자가 없는 슬라이드 When 어셈블 Then 기존과 동일하게 단일 `##` 헤딩으로 렌더(분해 미발동).
- [ ] AC3: Given `heading_hierarchy=False`(기본) When 어셈블 Then 출력은 현행과 **바이트 동일**(INV-3, opt-in 하위호환) — 회귀 diff 0.

*TOC / 구분 슬라이드 (Agent A)*
- [ ] AC4: Given `emit_toc=True` 이고 **본문 렌더 결과가 헤딩 라인 외 비공백 0 인 슬라이드**(D-4 확정 판정 기준) When 어셈블 Then 해당 슬라이드는 빈 `## Slide N` 블록 대신 TOC 항목(또는 `---` 섹션 구분)으로 변환되어 **빈 슬라이드 블록이 0개**가 된다(DoD "빈 슬라이드" 연계).

*메타데이터 / frontmatter (Agent B)*
- [ ] AC5: Given `emit_frontmatter=True` When 어셈블 Then 문서 최상단에 YAML frontmatter 블록(`---\n…\n---`)이 1회 출력되고, 각 슬라이드 블록에 `slide_index`(1-based 정수)·`section`(현재 최상위 `##` 헤딩 텍스트, 없으면 `""`)·`has_diagram`(bool) 3개 키를 담은 **HTML 주석 메타**(`<!-- slide_index: 1 | section: … | has_diagram: false -->`)가 결정적 순서로 부착된다.
- [ ] AC6: Given `emit_frontmatter=True` 이고 슬라이드에 `classification is DIAGRAM` 인 이미지 도형이 1개 이상 When 어셈블 Then 그 슬라이드 메타의 `has_diagram` 이 `true`, 그렇지 않으면 `false`.
- [ ] AC7: Given `emit_frontmatter=True` When validate_markdown 로 검사 Then 슬라이드별 메타는 **HTML 주석**(`<!-- -->`)으로만 출력되어 어셈블러 슬라이드 구분자 `---` 와 충돌하지 않으며(D-1 확정), heading-order·fence 규칙(FR-14)을 깨지 않는다(경고 0 증가). 문서 최상단 YAML frontmatter `---` 블록은 1회로 한정되어 슬라이드 구분자와 구별된다. PyYAML 미도입, stdlib 문자열 조립(INV-6).

*노트 (Agent C)*
- [ ] AC8: Given `include_notes=True` 이고 `SlideIR.notes != ""` When 어셈블 Then 해당 슬라이드 블록 말미에 `> notes` 블록(각 노트 줄이 `> ` 프리픽스)이 부착된다.
- [ ] AC9: Given `include_notes=True` 이고 `SlideIR.notes == ""`(기본) When 어셈블 Then notes 블록이 출력되지 않는다(빈 blockquote 미생성).

*공통 회귀 (Reviewer)*
- [ ] AC10: 4개 옵션 전부 활성 상태로 골든 픽스처 변환 시 validator 의 **헤딩 레벨 점프 경고 0**(FR-14 `_warn_heading_jump`), 중복 헤딩 0(DoD 연계). 결정성: 동일 IR·동일 옵션 2회 변환 결과 바이트 동일(INV-1).

**우선순위**: Should / **의존**: FR-11/FR-12(어셈블러·Map-Reduce, 기존), FR-14(validator, AC7/AC10 검증), FR-29(골든 게이트가 본 FR 출력 변화를 회귀 측정)

---

## 3. FR-29 Validator 강화 & 골든 회귀 게이트 (M14)

**스토리**: 라이브러리 사용자/CI 로서, validate_markdown 이 깨진 표·중복 헤딩·`\v` 잔존·빈 슬라이드 같은 구조 결함을 경고로 잡아내고, `output.md` 골든 대비 회귀를 CI 에서 자동 차단하기를 원한다. 품질 퇴행을 사람 리뷰 전에 막기 위해서다.

**배선**: D-5 확정 — validator 신규 규칙은 `validate_markdown` 을 **확장**(경고 추가)하며 기존 `ConvertOptions.validate` 플래그로 노출한다. 신규 규칙은 **상시 실행**(규칙 토글 옵션 신설 없음). INV-4(무예외·항상 `ValidationResult`) 계승. 신규 규칙은 heading/fence 규칙처럼 **경고**로 취급하며 `valid` 플래그 변경 여부는 규칙별로 AC 에 명시.

**AC**:

*깨진 pipe table 감지 (Agent A)*
- [ ] AC1: Given 헤더 행 + 구분 행(`|---|`)만 있고 데이터 행이 0개인 pipe table When validate_markdown Then "깨진 표: 데이터 행이 없습니다" 류 경고 1건이 `warnings` 에 추가된다(`valid` 불변). 코드펜스 내부의 `|` 라인은 표로 오인하지 않는다(FR-14 AC6 계승).
- [ ] AC2: Given 열 수가 헤더와 불일치하는 데이터 행(파이프 개수 상이)이 있는 pipe table When validate_markdown Then 열 수 불일치 경고 1건이 추가된다.

*중복·제어문자·빈 요소 규칙 (Agent B)*
- [ ] AC3: Given 인접한(사이에 본문 없이 연속된) 동일 텍스트 헤딩 2줄 When validate_markdown Then 중복 인접 헤딩 경고 1건이 추가된다. DoD "중복 헤딩 0" 을 이 규칙으로 계측한다.
- [ ] AC4: Given 문서에 제어문자 `\v`(U+000B)가 1개 이상 잔존 When validate_markdown Then `\v` 잔존 경고 1건이 추가되고 경고 메시지에 잔존 개수를 포함한다. DoD "`\v` 잔존 0" 계측.
- [ ] AC5: Given 모든 셀이 공백/빈 문자열인 pipe table When validate_markdown Then 전공백 표 경고 1건이 추가된다.
- [ ] AC6: Given **본문 렌더 결과가 헤딩 라인 외 비공백 0 인 슬라이드 블록**(D-4 확정 판정 기준 — FR-28 AC4 와 공통) When validate_markdown Then 빈 슬라이드 경고 1건이 추가된다.
- [ ] AC7: Given 결함이 전혀 없는 정상 문서 When validate_markdown Then 신규 규칙에 의한 경고가 0건이며 기존 FR-14 동작(빈 문서·미닫힘 펜스·heading order)은 회귀 없이 유지된다(무예외, INV-4).

*골든 회귀 CI 게이트 (Agent C)*
- [ ] AC8: 골든 픽스처를 **2종 박제**한다(D-7 확정): (a) off 골든 — FR-28 옵션 전부 off(현행 동등, 하위호환 회귀용), (b) on 골든 — FR-28 옵션(`heading_hierarchy`/`emit_toc`/`emit_frontmatter`/`include_notes`) 전부 on(구조화 품질 계측용). 각 골든의 경로·생성 `ConvertOptions` 조합을 테스트/문서에 명시한다.
- [ ] AC9: Given 어셈블러/validator 변경으로 골든 2종 중 하나라도 변환 결과가 달라짐 When CI 회귀 테스트 실행 Then diff 를 출력하고 테스트가 **실패**하여 머지가 차단된다(D-6 확정: hard fail). 의도된 변경은 별도 골든 갱신 PR 로 명시 승인한다. Given 변경 없음 When 실행 Then diff 0 으로 통과. 결정성(INV-1) 위반 시(동일 입력 2회 상이) 실패.
- [ ] AC10: DoD 지표(중복 헤딩 0 / `\v` 잔존 0 / 전공백·깨진 표 0)를 골든 2종 산출물에 대해 validate_markdown 경고 0 으로 검증하는 테스트가 존재한다(이미지 설명 커버리지 100% 는 VLM 경로/FR-27 소관이므로 본 FR 게이트에서는 no-VLM 골든 기준으로 N/A 명시).

**우선순위**: Must / **의존**: FR-14(validate_markdown 확장 대상), FR-28(본 게이트가 FR-28 출력 변화를 측정), FR-24/FR-25(`\v`·표 정규화가 경고 0 을 만족시키는 선행 조건)

---

## 4. 비기능 요구(NFR) — Wave 3 적용분

기존 v0.1.2 NFR 계승. Wave 3 신규 로직에 적용되는 수치 기준:

- **NFR-W3-1 성능**: FR-28 구조화 로직은 순수 문자열 처리(VLM 미개입). 20슬라이드 문서 어셈블 p95 < 5초(VLM 제외) 유지 — 기존 NFR-01 게이트 회귀 없음.
- **NFR-W3-2 커버리지**: 신규 코드 라인 커버리지 ≥ 75%(프로파일 §7). FR-28 옵션별 분기·FR-29 신규 규칙별 경고를 결정적 픽스처로 커버.
- **NFR-W3-3 타입 안전성**: mypy `src/` exit 0(프로파일 §4). 신규 `ConvertOptions` 필드·validator 반환 타입 명시.
- **NFR-W3-4 린트**: `ruff check . && black --check .` exit 0.
- **NFR-W3-5 결정성 게이트**: 동일 IR·동일 옵션 2회 변환 결과 바이트 동일(측정 가능, INV-1).

---

## 5. 하위 분할 제안 (PM 이슈 분해용)

과업지서는 M12/M13 각각을 3레인으로 분해했다. FR-28 은 관심사가 4개(계층/TOC/메타/노트)로 스코프가 크므로, PM 이 이슈를 나눌 경우 아래 분할을 권장한다(요구 번호는 유지, 이슈만 분할):

- FR-28 → 이슈 A(AC1~AC4: 헤딩 계층 + TOC, `assembler.py` 헤딩 경로), 이슈 B(AC5~AC7: frontmatter/메타), 이슈 C(AC8~AC9: notes). 세 이슈는 `assembler.py` 를 공유하므로 **파일 소유권 충돌 방지**를 위해 순차 또는 함수 경계 분리 필요(과업지서 §0 규칙).
- FR-29 → 이슈 A(AC1~AC2: 표 규칙, `validator.py`), 이슈 B(AC3~AC7: 중복/제어문자/빈요소 규칙, `validator.py`), 이슈 C(AC8~AC10: CI 골든 게이트, `tests/`·CI). A/B 는 같은 파일이므로 순차 권장.

> 대안: FR-28 메타/노트를 별도 FR-30 으로 승격하는 방안도 가능하나, 과업지서가 M12 를 단일 마일스톤으로 묶었으므로 **FR 은 2개 유지 + 이슈 분할**을 기본 권고.

---

## 6. Out of Scope (Wave 3 배제)

- COM 경로(`convert_via_com`) 구조화 — 코드베이스 미존재(ARCH-M12 §9.3 부채 계승). 신설 금지.
- 이미지 설명 커버리지 100%(DoD) 의 VLM 실호출 — FR-27(M11/M12 VLM 파이프라인) 소관. 본 Wave 는 no-VLM 골든 기준.
- 디스크·세션 간 캐시(REQ P-01) — 별도 승인 영역.
- 사용자 정의 heading 매핑/프롬프트 템플릿 — v1 미요구.
- MkDocs 등 문서 사이트 렌더링 — 산출물은 Markdown 문자열까지.

---

## 7. 결정 로그 (확정) — 2026-07-05 / 승인자 ms9648

> 초안 §7 의 NEEDS_DECISION 전건이 사람 게이트에서 확정됨. 아래 결정이 정본이며 위 AC 문구에 반영 완료.

| ID | 주제 | 확정 결정 | 결정일 / 승인자 |
|----|------|-----------|-----------------|
| D-1 | 메타 포맷 | **문서 최상단 YAML frontmatter 1회 + 슬라이드별 HTML 주석**(`<!-- slide_index: 1 … -->`). 슬라이드 구분자 `---` 와 충돌 회피. PyYAML 미도입, stdlib 문자열 조립 | 2026-07-05 / ms9648 |
| D-2 | 옵션 구조 | **bool 4개 유지**(`heading_hierarchy`/`emit_toc`/`emit_frontmatter`/`include_notes`) | 2026-07-05 / ms9648 |
| D-3 | has_diagram/section 산출 | **IR 확장 없이 어셈블 시점 파생**(IR 무변경, INV-2) | 2026-07-05 / ms9648 |
| D-4 | 빈/구분 슬라이드 판정 | **본문 렌더 결과가 헤딩 라인 외 비공백 0 인 슬라이드** — FR-28 AC4 와 FR-29 AC6 공통 정의 | 2026-07-05 / ms9648 |
| D-5 | validator 신규 필드 | **기존 `ConvertOptions.validate` 재사용, 신규 규칙 상시 실행**(신규 옵션 필드 없음) | 2026-07-05 / ms9648 |
| D-6 | CI 게이트 정책 | **골든 diff 발생 시 CI hard fail(머지 차단) + 의도된 변경은 골든 갱신 PR 로 명시 승인**. FR-29 AC9 판정문 = "실패" | 2026-07-05 / ms9648 |
| D-7 | 골든 조합 | **off 골든(하위호환 회귀용, 현행 동등) + on 골든(FR-28 옵션 전부 on, 구조화 품질 계측용) 2종 박제**. FR-29 AC8 을 골든 2종 기준으로 확정 | 2026-07-05 / ms9648 |

---

## 8. 제안 (P-) — 인테이크(과업지서) 밖, 승인 후에만 FR 승격

- **P-W3-1**: validator 경고에 **머신 리더블 코드**(예: `W_BROKEN_TABLE`) 부여 — CI 스코어카드·필터링 용이. *현재 validator 는 한글 문자열 경고만.*
- **P-W3-2**: `validate_markdown` 결과를 JSON 리포트로 덤프하는 CLI/함수 — CI 아티팩트화. (과업지서 "품질 스코어카드"(M13 Reviewer)의 자동화 확장)

> 위 제안은 과업지서에 명시되지 않았으므로 FR 목록에 미포함. PM/사람 승인 시 별도 FR 로 승격.
