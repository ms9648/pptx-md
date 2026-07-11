# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.2] - 2026-07-11

### Added
- IR coordinate fields (position/size) and reading-order sort for parsed shapes (#58)
- `heading_hierarchy`, `emit_toc`, `emit_frontmatter`, `include_notes` document
  metadata options, all default `False` (#75)
- Validator rules for table integrity, duplicate headings, control characters,
  and empty slides, plus a golden regression gate (#76)

### Changed
- Assembler normalization: remove duplicate titles, control characters, empty
  elements, and repeated boilerplate labels (#59)
- Table rendering: escape pipe characters and cell line breaks, and apply a
  Mermaid code-block prefix for diagram fallback tables (#60)
- Parser/COM coverage extended to Chart/SmartArt `fallback_text` and recursive
  Group shape traversal (#67)
- VLM pipeline integration: sha256-based image hash caching, ThreadPoolExecutor
  concurrency (default 4 workers), and opt-in `diagram_mermaid` handling (#68)

## [0.1.1] - 2026-06-30

### Fixed
- Use actual slide title as heading when `is_title` placeholder exists, instead of always
  outputting `## Slide N` (#53)
- Filter out footer, slide-number, and date placeholders from Markdown output to remove
  repeated boilerplate text (#54)
- Standardize no-VLM image markers to `![슬라이드 N 이미지 M]` format with slide and
  image position, replacing inconsistent `_[image]_` variants (#55)

## [0.1.0] - 2026-06-29

### Added
- Initial release: PPTX → Markdown conversion pipeline
- Rule-based image classifier (diagram, photo, logo, text)
- VLM image description support (Anthropic, OpenAI providers)
- Mermaid fallback for complex tables
- PII masking option
