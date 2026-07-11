"""Slide-to-PNG rendering via LibreOffice + pypdfium2 (FR-30, ADR-622/623).

Public interface:
    RENDER_TIMEOUT_S: int
    RENDER_DPI_DEFAULT: int
    slide_render_available() -> bool
    render_deck(pptx_path, indices, *, dpi=RENDER_DPI_DEFAULT)
        -> dict[int, bytes | None]
    render_slide(pptx_path, n, *, dpi=RENDER_DPI_DEFAULT) -> bytes | None

Design (vector.py graceful-skip/timeout conventions carried over — ADR-622/623):
    - Only stdlib imports at module level (io, logging, os, shutil, subprocess,
      tempfile, pathlib). ``pypdfium2`` is imported lazily *inside* functions
      (INV-5 / AC7): ``import pptx_md.slide_render`` must succeed even when
      pypdfium2 is not installed (it ships in the optional ``[render]`` extra).
    - LibreOffice (``soffice``) is an *external executable* — never imported as
      a Python module. ``slide_render_available()`` uses ``shutil.which`` at
      call time (not import time).
    - The deck is converted to PDF via ``soffice --convert-to pdf`` **exactly
      once** per ``render_deck()`` call; requested pages are then rasterised
      individually with pypdfium2 (ADR-623: soffice startup cost dominates).
    - ``render_deck()`` never raises (INV-4): missing soffice/pypdfium2,
      corrupt PPTX, out-of-range indices, subprocess failure, and timeouts all
      map the affected indices to ``None``.
    - shell=False (list args) — no shell injection surface.
    - Temporary files are cleaned up via TemporaryDirectory context manager.
    - PII (NFR-06): never log original file paths, slide text, or PNG bytes.
      Only indices, success/failure, and byte lengths are logged.
"""

from __future__ import annotations

import io
import logging
import os
import shutil
import subprocess
import tempfile
from collections.abc import Sequence
from pathlib import Path
from typing import Any

__all__ = [
    "RENDER_TIMEOUT_S",
    "RENDER_DPI_DEFAULT",
    "slide_render_available",
    "render_deck",
    "render_slide",
]

_logger = logging.getLogger("pptx_md.slide_render")

# Candidate executable names (platform variations, vector.py parity)
_SOFFICE_CANDIDATES: tuple[str, ...] = ("soffice", "soffice.bin", "libreoffice")

# Subprocess timeout in seconds — deck->PDF conversion (ADR-622/623).
# vector.py's VECTOR_TIMEOUT_S (30s) is for single EMF/WMF shapes; a full deck
# with many slides needs headroom, hence the higher default.
RENDER_TIMEOUT_S: int = 90

# Default target DPI for PNG rasterisation (16:9 slide -> ~2000x1125px @150dpi)
RENDER_DPI_DEFAULT: int = 150

# Fixed internal filenames inside the temp workdir — never derived from the
# caller-supplied path, so log messages never leak the original filename (PII).
_DECK_IN_NAME = "deck.pptx"
_DECK_PDF_NAME = "deck.pdf"


def _find_soffice() -> str | None:
    """Return the first available soffice executable path, or None."""
    for candidate in _SOFFICE_CANDIDATES:
        found = shutil.which(candidate)
        if found is not None:
            return found
    return None


def _pypdfium2_importable() -> bool:
    """Return True iff ``pypdfium2`` can be imported (lazy check, INV-5)."""
    try:
        import pypdfium2  # noqa: F401
    except Exception:
        return False
    return True


def slide_render_available() -> bool:
    """Return ``True`` iff LibreOffice ``soffice`` AND ``pypdfium2`` are both
    available. Never raises. Neither check spawns a subprocess or renders
    anything — safe to call cheaply and repeatedly.
    """
    return _find_soffice() is not None and _pypdfium2_importable()


def _render_page_to_png(pdf_doc: Any, index: int, scale: float) -> bytes | None:
    """Rasterise a single PDF page to PNG bytes. Never raises."""
    try:
        page = pdf_doc[index]
        bitmap = page.render(scale=scale)
        pil_image = bitmap.to_pil()
        buf = io.BytesIO()
        pil_image.save(buf, format="PNG")
        png_bytes = buf.getvalue()
    except Exception as exc:
        _logger.warning(
            "render_deck: page %d render failed (%s)", index, type(exc).__name__
        )
        return None

    if not png_bytes:
        return None
    _logger.debug("render_deck: page %d rendered (%d bytes)", index, len(png_bytes))
    return png_bytes


def render_deck(
    pptx_path: str | os.PathLike[str],
    indices: Sequence[int],
    *,
    dpi: int = RENDER_DPI_DEFAULT,
) -> dict[int, bytes | None]:
    """Render selected slide indices of *pptx_path* to PNG bytes.

    Converts the whole deck to PDF via ``soffice --convert-to pdf`` **exactly
    once**, then rasterises only the requested page indices via pypdfium2
    (ADR-623). Returns a dict keyed by every value in *indices* (deduplicated
    by first occurrence); each value is PNG bytes (``b"\\x89PNG"...``) on
    success or ``None`` on graceful skip/failure.

    Never raises (INV-4). All of the following degrade to ``None`` for the
    affected indices instead of propagating an exception:
    - LibreOffice not installed / pypdfium2 not importable (all indices None)
    - unreadable/corrupt input path (all indices None)
    - soffice non-zero exit / missing PDF output (all indices None)
    - soffice timeout after :data:`RENDER_TIMEOUT_S` (all indices None)
    - negative or out-of-range page index (that index only)
    - a single page's rasterisation failing (that index only)

    Args:
        pptx_path: Path to the source PPTX file.
        indices: Zero-based slide indices to render.
        dpi: Target DPI; internally converted to a pypdfium2 render
            ``scale = dpi / 72`` (72pt = 1 inch).

    Returns:
        Mapping of requested index -> PNG bytes or None.
    """
    result: dict[int, bytes | None] = dict.fromkeys(indices)
    if not result:
        return result

    soffice_path = _find_soffice()
    if soffice_path is None:
        _logger.warning(
            "render_deck: LibreOffice not found -> graceful skip (all None)"
        )
        return result

    try:
        import pypdfium2 as pdfium
    except Exception:
        _logger.warning(
            "render_deck: pypdfium2 not importable -> graceful skip (all None)"
        )
        return result

    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            in_file = tmp_path / _DECK_IN_NAME

            try:
                shutil.copyfile(os.fspath(pptx_path), in_file)
            except OSError as exc:
                _logger.warning(
                    "render_deck: failed to read input deck (%s) -> "
                    "graceful skip (all None)",
                    type(exc).__name__,
                )
                return result

            cmd = [
                soffice_path,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmp_path),
                str(in_file),
            ]
            proc = subprocess.run(
                cmd,
                capture_output=True,
                timeout=RENDER_TIMEOUT_S,
                check=False,
            )

            if proc.returncode != 0:
                _logger.warning(
                    "render_deck: soffice exited with code %d -> "
                    "graceful skip (all None)",
                    proc.returncode,
                )
                return result

            pdf_file = tmp_path / _DECK_PDF_NAME
            if not pdf_file.exists():
                _logger.warning(
                    "render_deck: expected PDF output not found -> "
                    "graceful skip (all None)"
                )
                return result

            try:
                pdf_doc = pdfium.PdfDocument(str(pdf_file))
            except Exception as exc:
                _logger.warning(
                    "render_deck: PDF open failed (%s) -> graceful skip (all None)",
                    type(exc).__name__,
                )
                return result

            try:
                page_count = len(pdf_doc)
                scale = dpi / 72
                for idx in result:
                    if idx < 0 or idx >= page_count:
                        _logger.debug(
                            "render_deck: index %d out of range (page_count=%d)",
                            idx,
                            page_count,
                        )
                        continue
                    result[idx] = _render_page_to_png(pdf_doc, idx, scale)
            finally:
                pdf_doc.close()

        return result

    except subprocess.TimeoutExpired:
        _logger.warning(
            "render_deck: soffice timed out after %ds -> all indices None",
            RENDER_TIMEOUT_S,
        )
        return dict.fromkeys(indices)
    except Exception as exc:
        _logger.warning(
            "render_deck: unexpected error (%s) -> all indices None",
            type(exc).__name__,
        )
        return dict.fromkeys(indices)


def render_slide(
    pptx_path: str | os.PathLike[str], n: int, *, dpi: int = RENDER_DPI_DEFAULT
) -> bytes | None:
    """Convenience wrapper: render a single slide index.

    Equivalent to ``render_deck(pptx_path, [n], dpi=dpi)[n]``.
    """
    return render_deck(pptx_path, [n], dpi=dpi)[n]
