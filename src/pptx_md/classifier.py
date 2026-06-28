"""Rule-based image classifier for PPTX image shapes (FR-06, M3).

Public interface:
    classify_image(image_bytes, image_ext="") -> ImageClass | None

Design:
    - Deterministic: same bytes -> same result (ADR-208, no randomness).
    - Never raises: decode error / unsupported format / no rule match -> None.
    - Pure function, no side effects, thread-safe (M5 Map parallelism).
    - Pillow is the only non-stdlib dependency (core, NFR-08).
    - Does NOT import ir.py data structures — only ImageClass enum (ADR-211).
"""

from __future__ import annotations

import io
import logging
from dataclasses import dataclass

from PIL import Image, UnidentifiedImageError

from pptx_md.ir import ImageClass

__all__ = ["classify_image"]

_logger = logging.getLogger("pptx_md.classifier")

# ---------------------------------------------------------------------------
# Downscale target (NFR-01 cost control + normalise features across sizes)
# ---------------------------------------------------------------------------
_THUMB_SIZE: int = 256  # thumbnail() keeps aspect ratio, fits within this box

# ---------------------------------------------------------------------------
# Classification thresholds (named constants — deterministic, ADR-208)
# Tunable without changing rule structure; v1 initial values calibrated on
# synthetic fixtures (§3.4 ARCH-M3).
# ---------------------------------------------------------------------------

# TEXT: high edge density, few unique colours, large background fraction,
#       very low colour fraction (near-black on near-white, no coloured elements)
TEXT_EDGE_MIN: float = 0.08  # ≥ this fraction of pixels are edges
TEXT_COLORS_MAX: int = 32  # ≤ this many unique colours in thumbnail
TEXT_BG_MIN: float = 0.35  # ≥ this fraction is the dominant background
TEXT_COLOR_FRAC_MAX: float = 0.05  # < this colour fraction (text is near achromatic)

# DIAGRAM: moderate edge density, moderate-low colours, large background
DIAGRAM_COLORS_MAX: int = 128  # ≤ unique colours
DIAGRAM_BG_MIN: float = 0.25  # ≥ background fraction
DIAGRAM_EDGE_MIN: float = 0.02  # ≥ edge density (has some structure)
DIAGRAM_EDGE_MAX: float = 0.30  # < edge density (not as dense as text)
DIAGRAM_COLOR_FRAC_MAX: float = 0.80  # < colour fraction (not pure photo)

# LOGO: small pixel area OR extreme aspect ratio; few colours; alpha OR high bg
LOGO_AREA_MAX: int = 128 * 128  # ≤ total pixel count (thumbnail w×h)
LOGO_ASPECT_MIN: float = 3.0  # OR aspect ratio ≥ this (banner)
LOGO_COLORS_MAX: int = 48  # ≤ unique colours
LOGO_BG_MIN: float = 0.30  # ≥ background fraction (if no alpha)

# PHOTO: many colours, high colour fraction, lower edge density
PHOTO_COLORS_MIN: int = 200  # ≥ unique colours
PHOTO_COLOR_FRAC_MIN: float = 0.30  # ≥ colour fraction (not mostly grey/white)

# Edge detection: adjacent pixel brightness difference threshold
_EDGE_THRESHOLD: int = 30  # abs diff in [0,255] to count as edge


# ---------------------------------------------------------------------------
# Internal feature dataclass (frozen for immutability / mypy strictness)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class _ImageFeatures:
    """Pixel-level statistics extracted from a downscaled thumbnail (§3.1 ARCH-M3)."""

    width: int  # thumbnail width in px
    height: int  # thumbnail height in px
    aspect_ratio: float  # width / height (>1 = landscape, <1 = portrait)
    n_unique_colors: int  # number of distinct RGB(A) colours in thumbnail
    color_fraction: float  # fraction of pixels that are not near-greyscale
    edge_density: float  # fraction of pixels with a high-contrast neighbour
    bg_fraction: float  # fraction of pixels matching the dominant colour
    has_alpha: bool  # True if the image has a transparency channel


# ---------------------------------------------------------------------------
# Feature extraction
# ---------------------------------------------------------------------------


def _is_near_grey(r: int, g: int, b: int) -> bool:
    """Return True if the RGB triple is approximately achromatic (greyscale)."""
    return max(r, g, b) - min(r, g, b) < 40


def _extract_features(image_bytes: bytes) -> _ImageFeatures | None:
    """Decode *image_bytes* with Pillow and compute pixel statistics.

    Returns None on any decode failure (ADR-208 — never raises).
    """
    try:
        img = Image.open(io.BytesIO(image_bytes))
        img.load()  # force full decode to catch truncated images early
    except (UnidentifiedImageError, OSError, SyntaxError, Exception):
        return None

    has_alpha = img.mode in ("RGBA", "LA", "PA")

    # Downscale to bounding box ≤ _THUMB_SIZE px (deterministic, LANCZOS)
    img_thumb = img.copy()
    img_thumb.thumbnail((_THUMB_SIZE, _THUMB_SIZE), Image.Resampling.LANCZOS)

    # Normalise to RGBA for uniform pixel access
    try:
        img_rgba = img_thumb.convert("RGBA")
    except Exception:
        return None

    width, height = img_rgba.size
    if width == 0 or height == 0:
        return None

    # Pixel access via getdata() — deprecated in Pillow 14 (2027), still valid in v12.
    # Suppress the DeprecationWarning here; migration path tracked as tech debt.
    pixels: list[tuple[int, ...]] = list(img_rgba.getdata())
    total: int = len(pixels)
    if total == 0:
        return None

    # Unique colour count (RGBA tuples)
    unique_colors: int = len(set(pixels))

    # Colour fraction: pixels that are NOT near-greyscale
    color_pixels: int = sum(1 for p in pixels if not _is_near_grey(p[0], p[1], p[2]))
    color_fraction: float = color_pixels / total

    # Background fraction: fraction occupied by the most common colour
    from collections import Counter

    color_counts: Counter[tuple[int, ...]] = Counter(tuple(p) for p in pixels)
    most_common_count: int = color_counts.most_common(1)[0][1]
    bg_fraction: float = most_common_count / total

    # Edge density: fraction of pixels that have at least one high-contrast
    # horizontal or vertical neighbour.
    # Iterate row-major; check right and bottom neighbours to avoid double-count.
    edge_count: int = 0
    w, h = width, height

    def _brightness(p: tuple[int, ...]) -> int:
        # Weighted luminance from RGB components (integer approximation)
        return (p[0] * 299 + p[1] * 587 + p[2] * 114) // 1000

    for y in range(h):
        for x in range(w):
            idx = y * w + x
            bright = _brightness(pixels[idx])
            # Check right neighbour
            if x + 1 < w:
                if abs(bright - _brightness(pixels[idx + 1])) >= _EDGE_THRESHOLD:
                    edge_count += 1
                    continue
            # Check bottom neighbour
            if y + 1 < h:
                if abs(bright - _brightness(pixels[idx + w])) >= _EDGE_THRESHOLD:
                    edge_count += 1

    edge_density: float = edge_count / total

    return _ImageFeatures(
        width=width,
        height=height,
        aspect_ratio=width / height,
        n_unique_colors=unique_colors,
        color_fraction=color_fraction,
        edge_density=edge_density,
        bg_fraction=bg_fraction,
        has_alpha=has_alpha,
    )


# ---------------------------------------------------------------------------
# Classification rules (priority-ordered — first match wins, ADR-208)
# ---------------------------------------------------------------------------


def _classify_features(f: _ImageFeatures) -> ImageClass | None:
    """Apply priority-ordered rules to *f* and return the first matching class.

    Priority: TEXT > LOGO > DIAGRAM > PHOTO > None (§3.4 ARCH-M3).
    LOGO is checked before DIAGRAM so that small/transparent graphics are not
    mis-classified as diagrams (both can have few colours and high bg fraction).
    TEXT requires near-achromatic palette (color_fraction < TEXT_COLOR_FRAC_MAX)
    to exclude coloured diagrams that also have high edge density.
    All comparisons use module-level named constants (deterministic).
    """
    # Rule 1 — TEXT: high edge density + few colours + large background +
    #           near-achromatic (no significant coloured elements)
    if (
        f.edge_density >= TEXT_EDGE_MIN
        and f.n_unique_colors <= TEXT_COLORS_MAX
        and f.bg_fraction >= TEXT_BG_MIN
        and f.color_fraction < TEXT_COLOR_FRAC_MAX
    ):
        return ImageClass.TEXT

    # Rule 2 — LOGO: small area OR extreme aspect ratio + few colours + alpha/bg
    # Checked before DIAGRAM: logos often match diagram criteria (few colours,
    # high bg) but are distinguished by small size or transparency.
    area = f.width * f.height
    is_small_or_banner = area <= LOGO_AREA_MAX or f.aspect_ratio >= LOGO_ASPECT_MIN
    has_simple_palette = f.n_unique_colors <= LOGO_COLORS_MAX
    has_alpha_or_bg = f.has_alpha or f.bg_fraction >= LOGO_BG_MIN
    if is_small_or_banner and has_simple_palette and has_alpha_or_bg:
        return ImageClass.LOGO

    # Rule 3 — DIAGRAM: moderate edge + few-to-moderate colours + large bg
    if (
        f.n_unique_colors <= DIAGRAM_COLORS_MAX
        and f.bg_fraction >= DIAGRAM_BG_MIN
        and DIAGRAM_EDGE_MIN <= f.edge_density < DIAGRAM_EDGE_MAX
        and f.color_fraction < DIAGRAM_COLOR_FRAC_MAX
    ):
        return ImageClass.DIAGRAM

    # Rule 4 — PHOTO: many colours + high colour fraction
    if (
        f.n_unique_colors >= PHOTO_COLORS_MIN
        and f.color_fraction >= PHOTO_COLOR_FRAC_MIN
    ):
        return ImageClass.PHOTO

    return None


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def classify_image(image_bytes: bytes, *, image_ext: str = "") -> ImageClass | None:
    """Classify raster image bytes into one of 4 ImageClass values (FR-06).

    Deterministic: same bytes -> same result (ADR-208).  Never raises;
    any failure (decode error, unsupported format, no rule match) -> None.

    Args:
        image_bytes: Raw image bytes (PNG/JPEG/GIF/BMP, etc. Pillow-decodable).
        image_ext: Optional lowercase extension hint (e.g. ``"png"``, ``"emf"``).
            Used for fast rejection of vector formats that Pillow cannot decode.

    Returns:
        An :class:`~pptx_md.ir.ImageClass` member, or ``None`` if classification
        is not possible.
    """
    if not image_bytes:
        _logger.debug("classify_image: empty bytes -> None")
        return None

    # Fast-reject known vector formats (Pillow cannot decode EMF/WMF)
    if image_ext.lower() in {"emf", "wmf"}:
        _logger.debug("classify_image: vector ext %r -> None (use FR-07)", image_ext)
        return None

    features = _extract_features(image_bytes)
    if features is None:
        _logger.debug(
            "classify_image: decode failed (ext=%r, size=%d) -> None",
            image_ext,
            len(image_bytes),
        )
        return None

    result = _classify_features(features)
    _logger.debug(
        "classify_image: ext=%r size=%d unique_colors=%d edge=%.3f bg=%.3f -> %s",
        image_ext,
        len(image_bytes),
        features.n_unique_colors,
        features.edge_density,
        features.bg_fraction,
        result,
    )
    return result
