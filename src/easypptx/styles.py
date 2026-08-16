"""Reusable style objects and theme presets for EasyPPTX.

Style objects bundle the formatting keywords that would otherwise be
repeated on every call:

    heading = TextStyle(font_size=28, font_bold=True, color="white")
    slide.add_text("Results", style=heading)

A :class:`Theme` groups styles with a background color and can be applied
to a whole presentation:

    pres = Presentation(theme="dark")            # built-in preset
    pres = Presentation(theme=Theme(bg_color=(10, 20, 40), ...))  # custom
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any

ColorType = str | tuple[int, int, int]


@dataclass
class TextStyle:
    """Formatting for text content (add_text, add_bullets, titles)."""

    font_name: str | None = None
    font_size: int | None = None
    font_bold: bool | None = None
    font_italic: bool | None = None
    color: ColorType | None = None
    align: str | None = None
    vertical: str | None = None

    def to_kwargs(self) -> dict[str, Any]:
        """Return the non-None fields as keyword arguments."""
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class TableStyle:
    """Formatting for tables."""

    has_header: bool | None = None
    style_id: int | None = None

    def to_kwargs(self) -> dict[str, Any]:
        """Return the non-None fields as keyword arguments."""
        kwargs: dict[str, Any] = {}
        if self.has_header is not None:
            kwargs["has_header"] = self.has_header
        if self.style_id is not None:
            kwargs["style"] = self.style_id
        return kwargs


@dataclass
class ChartStyle:
    """Formatting for charts."""

    chart_type: str | None = None
    has_legend: bool | None = None
    legend_position: str | None = None

    def to_kwargs(self) -> dict[str, Any]:
        """Return the non-None fields as keyword arguments."""
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class Theme:
    """A presentation-wide look: background, title style, and body style.

    Built-in presets are available by name: Presentation(theme="dark").
    Custom themes are plain instances of this class; every field is
    optional and unset fields fall back to library defaults.
    """

    name: str = "custom"
    bg_color: ColorType | None = None
    title: TextStyle = field(default_factory=TextStyle)
    body: TextStyle = field(default_factory=TextStyle)
    accent_color: ColorType | None = None
    palette: list[ColorType] | None = None
    title_accent: bool = True
    table: dict[str, Any] | None = None

    def to_template(self) -> dict[str, Any]:
        """Express the theme in the template-defaults structure.

        The result plugs into the same machinery TOML templates use, so
        themed defaults cascade exactly like template defaults.
        """
        template: dict[str, Any] = {"defaults": {"global": self.body.to_kwargs()}}
        chart_defaults: dict[str, Any] = {}
        if self.palette is not None:
            chart_defaults["palette"] = list(self.palette)
        if self.body.color is not None:
            chart_defaults["font_color"] = self.body.color
        if chart_defaults:
            template["defaults"]["chart"] = chart_defaults
        if self.table is not None:
            template["defaults"]["table"] = dict(self.table)
        if self.bg_color is not None:
            template["bg_color"] = self.bg_color
        return template


# Built-in theme presets
THEMES: dict[str, Theme] = {
    "light": Theme(
        name="light",
        bg_color="white",
        title=TextStyle(font_size=30, font_bold=True, color=(0x1A, 0x1A, 0x1A), align="left"),
        body=TextStyle(color=(0x33, 0x33, 0x33)),
        accent_color=(0x2E, 0x75, 0xB6),
        palette=[(0x2E, 0x75, 0xB6), (0xED, 0x7D, 0x31), (0x54, 0x9E, 0x39), (0xBF, 0x90, 0x00), (0x7A, 0x5C, 0xA8)],
        table={
            "header_fill": (0x2E, 0x75, 0xB6),
            "header_color": (0xFF, 0xFF, 0xFF),
            "band_fills": [(0xFF, 0xFF, 0xFF), (0xF2, 0xF5, 0xF9)],
            "body_color": (0x33, 0x33, 0x33),
            "font_size": 12,
            "header_font_size": 12,
        },
    ),
    "dark": Theme(
        name="dark",
        bg_color=(0x14, 0x14, 0x16),
        title=TextStyle(font_size=30, font_bold=True, color=(0xF2, 0xF2, 0xF2), align="left"),
        body=TextStyle(color=(0xDC, 0xDC, 0xDC)),
        accent_color=(0x4F, 0xC3, 0xF7),
        palette=[(0x4F, 0xC3, 0xF7), (0xFF, 0xB7, 0x4D), (0x81, 0xC7, 0x84), (0xE5, 0x73, 0x73), (0xBA, 0x68, 0xC8)],
        table={
            "header_fill": (0x2A, 0x2A, 0x2E),
            "header_color": (0xF2, 0xF2, 0xF2),
            "band_fills": [(0x1B, 0x1B, 0x1E), (0x22, 0x22, 0x26)],
            "body_color": (0xDC, 0xDC, 0xDC),
            "font_size": 12,
            "header_font_size": 12,
        },
    ),
    "corporate": Theme(
        name="corporate",
        bg_color=(0x0F, 0x2A, 0x44),
        title=TextStyle(font_size=30, font_bold=True, color=(0xFF, 0xFF, 0xFF), align="left"),
        body=TextStyle(color=(0xD8, 0xE2, 0xEC)),
        accent_color=(0xF5, 0xA6, 0x23),
        palette=[(0xF5, 0xA6, 0x23), (0x5B, 0x9B, 0xD5), (0x8F, 0xB8, 0x6B), (0xC9, 0x6A, 0x50), (0x9E, 0x86, 0xC8)],
        table={
            "header_fill": (0x1E, 0x4A, 0x73),
            "header_color": (0xFF, 0xFF, 0xFF),
            "band_fills": [(0x14, 0x33, 0x52), (0x18, 0x3B, 0x5E)],
            "body_color": (0xD8, 0xE2, 0xEC),
            "font_size": 12,
            "header_font_size": 12,
        },
    ),
}


def resolve_theme(theme: str | Theme | None) -> Theme | None:
    """Resolve a theme given by name or instance.

    Raises:
        ValueError: If a theme name is not a built-in preset
    """
    if theme is None or isinstance(theme, Theme):
        return theme
    if theme in THEMES:
        return THEMES[theme]
    raise ValueError(f"Unknown theme: {theme!r}. Built-in themes: {', '.join(sorted(THEMES))}")
