# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
