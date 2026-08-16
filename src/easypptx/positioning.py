"""Position and size arithmetic for EasyPPTX.

All layout math in the library goes through these helpers. Positions are
either percentage strings ("10%") relative to the slide, or absolute
inches as floats. Helpers here convert between the two so downstream
arithmetic works for both conventions.
"""

from __future__ import annotations

import warnings

from easypptx.common import EMU_PER_INCH

# Accepts either a percentage string ("10%") or absolute inches (10.0)
PositionType = float | str


def is_percent(value: PositionType | None) -> bool:
    """Return True if value is a percentage string like "10%"."""
    return isinstance(value, str) and value.endswith("%")


def parse_percent(value: PositionType) -> float:
    """Extract the numeric part of a percentage string ("12.5%" -> 12.5).

    Bare numbers are returned as floats for convenience.
    """
    if isinstance(value, str):
        return float(value.strip().removesuffix("%"))
    return float(value)


def pct(value: float) -> str:
    """Format a number as a percentage string ("12.50%")."""
    return f"{value:.2f}%"


def to_percent(value: PositionType, total_emu: int) -> float:
    """Convert a position value to a percentage of a slide dimension.

    Percentage strings pass through; floats are treated as inches and
    converted using the slide dimension in EMUs.
    """
    if is_percent(value):
        return parse_percent(value)
    return (float(value) * EMU_PER_INCH / total_emu) * 100


def to_inches(value: PositionType, total_emu: int) -> float:
    """Convert a position value to inches.

    Percentage strings are resolved against the slide dimension in EMUs;
    floats are already inches and pass through. Percentages outside 0-100
    are clamped with a warning (off-slide placement should use inches).
    """
    if is_percent(value):
        percent = parse_percent(value)
        clamped = max(0.0, min(100.0, percent))
        if abs(clamped - percent) > 1e-9:
            warnings.warn(
                f"Position {value!r} is outside 0-100% and was clamped to {pct(clamped)}",
                stacklevel=3,
            )
        return (clamped / 100) * (total_emu / EMU_PER_INCH)
    return float(value)


def shift_band(
    y: PositionType,
    height: PositionType,
    band_height: PositionType,
    total_emu: int,
) -> tuple[str, str]:
    """Reserve a horizontal band (e.g. a title area) at the top of a region.

    Returns the (y, height) of the remaining region below the band, as
    percentage strings. All inputs may be percentage strings or inches.

    Args:
        y: Top of the region
        height: Height of the region
        band_height: Height of the band to reserve
        total_emu: Slide height in EMUs, for converting inch values
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
