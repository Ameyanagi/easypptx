"""Shared constants and plumbing for EasyPPTX.

This module is the single home for color/alignment constants and the
template-default merging logic that Slide and Grid share. Importing from
here (instead of from ``easypptx.presentation``) avoids circular imports.
"""

from __future__ import annotations

import sys
import warnings
from typing import Any, cast

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

# English Metric Units per inch (python-pptx base unit)
EMU_PER_INCH = 914400

# Default font used across the library
DEFAULT_FONT = "Meiryo"

# Named colors usable anywhere a color parameter is accepted
COLORS: dict[str, RGBColor] = {
    "black": RGBColor(0x10, 0x10, 0x10),
    "darkgray": RGBColor(0x40, 0x40, 0x40),
    "gray": RGBColor(0x80, 0x80, 0x80),
    "lightgray": RGBColor(0xD0, 0xD0, 0xD0),
    "red": RGBColor(0xFF, 0x40, 0x40),
    "green": RGBColor(0x40, 0xFF, 0x40),
    "blue": RGBColor(0x40, 0x40, 0xFF),
    "white": RGBColor(0xFF, 0xFF, 0xFF),
    "yellow": RGBColor(0xFF, 0xD7, 0x00),
    "cyan": RGBColor(0x00, 0xE5, 0xFF),
    "magenta": RGBColor(0xFF, 0x00, 0xFF),
    "orange": RGBColor(0xFF, 0xA5, 0x00),
}

# Horizontal text alignment
ALIGN: dict[str, PP_ALIGN] = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
}

# Vertical text anchoring
VERTICAL: dict[str, MSO_ANCHOR] = {
    "top": MSO_ANCHOR.TOP,
    "middle": MSO_ANCHOR.MIDDLE,
    "bottom": MSO_ANCHOR.BOTTOM,
}


def normalize_color(
    color: str | tuple[int, int, int] | list[int] | None,
) -> str | tuple[int, int, int] | None:
    """Normalize a color value, converting 3-element lists to tuples.

    Template files (TOML/JSON) deserialize colors as lists; the rest of the
    library works with tuples or named-color strings.
    """
    if isinstance(color, list) and len(color) == 3:
        return (color[0], color[1], color[2])
    return cast("str | tuple[int, int, int] | None", color)


def resolve_color(color: str | tuple[int, int, int] | list[int] | None) -> RGBColor | None:
    """Resolve a color name or RGB tuple/list to an RGBColor, or None."""
    color = normalize_color(color)
    if isinstance(color, str):
        return COLORS.get(color)
    if isinstance(color, tuple) and len(color) == 3:
        return RGBColor(*color)
    return None


def blend_colors(a: tuple[int, int, int], b: tuple[int, int, int], t: float) -> tuple[int, int, int]:
    """Blend color a toward color b by factor t (0 = a, 1 = b)."""
    return (
        round(a[0] + (b[0] - a[0]) * t),
        round(a[1] + (b[1] - a[1]) * t),
        round(a[2] + (b[2] - a[2]) * t),
    )


def apply_shadow(shape: Any) -> None:
    """Apply the library's standard drop shadow to a shape."""
    shadow = shape.shadow
    shadow.inherit = False
    shadow.visible = True
    shadow.blur_radius = 5
    shadow.distance = 3
    shadow.angle = 45


def is_dataframe(obj: Any) -> bool:
    """Return True if obj is a pandas DataFrame, without importing pandas.

    If pandas has not been imported, obj cannot be a DataFrame, so this
    check is safe and keeps pandas an optional dependency.
    """
    pd = sys.modules.get("pandas")
    return pd is not None and isinstance(obj, pd.DataFrame)


def warn_ignored_kwargs(method: str, kwargs: dict[str, Any]) -> None:
    """Warn about parameters a method received but does not support.

    Unknown parameters used to be silently discarded, which hid typos
    (``font_szie=24``) and unsupported options from the caller.
    """
    if kwargs:
        warnings.warn(
            f"{method} ignored unsupported parameter(s): {', '.join(sorted(kwargs))}",
            stacklevel=3,
        )


def filter_to_signature(func: Any, merged: dict[str, Any], explicit: dict[str, Any]) -> dict[str, Any]:
    """Drop template-default keys a target method does not accept.

    Template defaults can contain keys meant for other methods (e.g. a
    global ``title_font_size`` intended for slide factories). Keys the
    caller passed explicitly are always kept, so typos still fail loudly.

    Args:
        func: The target callable whose signature defines accepted keys
        merged: Merged kwargs (template defaults + explicit arguments)
        explicit: The kwargs the caller passed explicitly

    Returns:
        merged, minus default-sourced keys not accepted by func
    """
    import inspect

    try:
        params = inspect.signature(func).parameters
    except (TypeError, ValueError):
        return merged
    named = {
        name
        for name, p in params.items()
        if p.kind in (inspect.Parameter.POSITIONAL_OR_KEYWORD, inspect.Parameter.KEYWORD_ONLY)
    }
    if not named:
        # Opaque callable (e.g. only *args/**kwargs): can't filter meaningfully
        return merged
    return {k: v for k, v in merged.items() if k in named or k in explicit}


def merge_defaults(
    template_defaults: dict[str, dict[str, Any]],
    method_type: str,
    kwargs: dict[str, Any],
) -> dict[str, Any]:
    """Merge provided arguments with template defaults.

    Precedence (lowest to highest): global defaults, method-type defaults,
    explicitly provided non-None kwargs.

    Args:
        template_defaults: Mapping of element type to its default kwargs
        method_type: The type of method ("text", "image", etc.)
        kwargs: Keyword arguments provided to the method

    Returns:
        Dictionary with merged arguments
    """
    result: dict[str, Any] = dict(template_defaults.get("global", {}))
    result.update(template_defaults.get(method_type, {}))
    for key, value in kwargs.items():
        if value is not None:
            result[key] = value
    return result
