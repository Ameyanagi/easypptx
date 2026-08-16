"""Position and size arithmetic for EasyPPTX.

All layout math in the library goes through these helpers.

Position values can be:
- a number (int or float): a percentage of the slide dimension, 0-100
  (``x=10`` means 10% — same as ``x="10%"``)
- a percentage string: ``"10%"``
- an absolute length created with :func:`in_` (inches), e.g. ``in_(1.5)``
"""

from __future__ import annotations

import warnings

from pptx.util import Inches, Length

from easypptx.common import EMU_PER_INCH

# Accepts a percentage (number or "10%" string) or an absolute Length (via in_())
PositionType = float | str | Length


def in_(inches: float) -> Length:
    """Create an absolute length in inches.

    Use this to position or size in physical units instead of percentages:
    ``slide.add_image("logo.png", x=in_(0.5), width=in_(2))``.
    """
    return Inches(inches)


def is_percent(value: PositionType | None) -> bool:
    """Return True if value is a percentage string like "10%"."""
    return isinstance(value, str) and value.endswith("%")


def parse_percent(value: PositionType) -> float:
    """Extract the numeric part of a percentage string ("12.5%" -> 12.5).

    Bare numbers are returned as floats (they are already percentages).
    """
    if isinstance(value, str):
        return float(value.strip().removesuffix("%"))
    return float(value)


def pct(value: float) -> str:
    """Format a number as a percentage string ("12.50%")."""
    return f"{value:.2f}%"


def _clamped_percent(value: PositionType) -> float:
    """Parse a percentage and clamp it to 0-100, warning when clamping changes it."""
    percent = parse_percent(value)
    clamped = max(0.0, min(100.0, percent))
    if abs(clamped - percent) > 1e-9:
        warnings.warn(
            f"Position {value!r} is outside 0-100% and was clamped to {pct(clamped)}",
            stacklevel=4,
        )
    return clamped


def to_percent(value: PositionType, total_emu: int) -> float:
    """Convert a position value to a percentage of a slide dimension.

    Numbers and percentage strings are already percentages; absolute
    lengths (from :func:`in_`) are converted using the slide dimension.
    """
    if isinstance(value, Length):
        return (int(value) / total_emu) * 100
    return parse_percent(value)


def to_inches(value: PositionType, total_emu: int) -> float:
    """Convert a position value to inches.

    Percentages (numbers or "10%" strings) are resolved against the slide
    dimension in EMUs; absolute lengths pass through. Percentages outside
    0-100 are clamped with a warning (use in_() for off-slide placement).
    """
    if isinstance(value, Length):
        return int(value) / EMU_PER_INCH
    return (_clamped_percent(value) / 100) * (total_emu / EMU_PER_INCH)


def shift_band(
    y: PositionType,
    height: PositionType,
    band_height: PositionType,
    total_emu: int,
) -> tuple[str, str]:
    """Reserve a horizontal band (e.g. a title area) at the top of a region.

    Returns the (y, height) of the remaining region below the band, as
    percentage strings. Inputs may be percentages or absolute lengths.

    Args:
        y: Top of the region
        height: Height of the region
        band_height: Height of the band to reserve
        total_emu: Slide height in EMUs, for converting absolute lengths
    """
    y_p = to_percent(y, total_emu)
    height_p = to_percent(height, total_emu)
    band_p = to_percent(band_height, total_emu)
    return pct(y_p + band_p), pct(height_p - band_p)


def resolve_padding(
    padding: PositionType | None,
    axis_padding: PositionType | None,
) -> PositionType | None:
    """Resolve a padding value: the general padding wins over the axis one."""
    return padding if padding is not None else axis_padding


def apply_content_padding(
    y: PositionType,
    height: PositionType,
    padding: PositionType | None,
    axis_padding: PositionType | None,
    total_emu: int,
) -> tuple[PositionType, PositionType]:
    """Shift a content region down by the resolved vertical padding.

    Returns (y, height) with the padding applied, or the inputs unchanged
    when no padding is specified.
    """
    pad = resolve_padding(padding, axis_padding)
    if pad is None:
        return y, height
    return shift_band(y, height, pad, total_emu)
