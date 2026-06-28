"""LibreOffice-based vector image conversion (FR-07, M3).

Public interface:
    VECTOR_EXTS: frozenset[str]
    libreoffice_available() -> bool
    convert_vector_to_png(image_bytes, image_ext) -> bytes | None

Design:
    - Only stdlib imports at module level (subprocess, shutil, tempfile, logging).
    - LibreOffice is an *external executable* — never imported as a Python module.
    - libreoffice_available() uses shutil.which at call time (not import time).
    - convert_vector_to_png() never raises (ADR-209): unsupported ext, missing
      LibreOffice, subprocess failure, or timeout all return None.
    - shell=False (list args) — no shell injection surface.
    - Temporary files are cleaned up via TemporaryDirectory context manager.
"""

from __future__ import annotations

import logging
import shutil
import subprocess
import tempfile
from pathlib import Path

__all__ = ["VECTOR_EXTS", "libreoffice_available", "convert_vector_to_png"]

_logger = logging.getLogger("pptx_md.vector")

# Supported vector extensions (lowercase, no leading dot — M2 normalisation)
VECTOR_EXTS: frozenset[str] = frozenset({"emf", "wmf"})

# Candidate executable names (platform variations)
_SOFFICE_CANDIDATES: tuple[str, ...] = ("soffice", "soffice.bin", "libreoffice")

# Subprocess timeout in seconds (ADR-209)
VECTOR_TIMEOUT_S: int = 30


def libreoffice_available() -> bool:
    """Return ``True`` iff a LibreOffice ``soffice`` executable is on PATH.

    Uses :func:`shutil.which` only — no subprocess spawn, no LibreOffice import.
    Called at function invocation time (never at import time).
    """
    for candidate in _SOFFICE_CANDIDATES:
        if shutil.which(candidate) is not None:
            return True
    return False


def convert_vector_to_png(image_bytes: bytes, image_ext: str) -> bytes | None:
    """Convert EMF/WMF bytes to PNG bytes via LibreOffice (FR-07).

    Returns PNG bytes on success, or ``None`` if:
    - *image_ext* is not in :data:`VECTOR_EXTS`, or
    - LibreOffice is not installed (graceful skip), or
    - the subprocess fails or times out.

    Never raises (ADR-209).

    Args:
        image_bytes: Raw vector image bytes (EMF or WMF blob).
        image_ext: Lowercase extension without leading dot (e.g. ``"emf"``).

    Returns:
        PNG bytes starting with ``b"\\x89PNG"``, or ``None``.
    """
    if image_ext.lower() not in VECTOR_EXTS:
        _logger.debug(
            "convert_vector_to_png: ext %r not in VECTOR_EXTS -> None", image_ext
        )
        return None

    # Detect LibreOffice at call time
    soffice_path: str | None = None
    for candidate in _SOFFICE_CANDIDATES:
        found = shutil.which(candidate)
        if found is not None:
            soffice_path = found
            break

    if soffice_path is None:
        _logger.debug(
            "convert_vector_to_png: LibreOffice not found -> graceful skip (None)"
        )
        return None

    try:
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            in_file = tmp_path / f"in.{image_ext.lower()}"
            in_file.write_bytes(image_bytes)

            cmd = [
                soffice_path,
                "--headless",
                "--convert-to",
                "png",
                "--outdir",
                str(tmp_path),
                str(in_file),
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                timeout=VECTOR_TIMEOUT_S,
                check=False,  # handle returncode manually
            )

            if result.returncode != 0:
                _logger.warning(
                    "convert_vector_to_png: soffice exited with code %d",
                    result.returncode,
                )
                return None

            out_file = tmp_path / "in.png"
            if not out_file.exists():
                _logger.warning(
                    "convert_vector_to_png: expected output %s not found", out_file.name
                )
                return None

            png_bytes = out_file.read_bytes()

        return png_bytes if png_bytes else None

    except subprocess.TimeoutExpired:
        _logger.warning(
            "convert_vector_to_png: soffice timed out after %ds", VECTOR_TIMEOUT_S
        )
        return None
    except Exception as exc:
        _logger.warning(
            "convert_vector_to_png: unexpected error (%s): %s",
            type(exc).__name__,
            exc,
        )
        return None
