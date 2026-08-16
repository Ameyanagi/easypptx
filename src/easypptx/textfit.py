"""Heuristic text measurement and fitting for EasyPPTX.

python-pptx can only set PowerPoint's autofit *flag* — the actual font
scaling happens when the file is opened in PowerPoint, so other renderers
show overflowing text. These helpers estimate line wrapping with simple
glyph-width heuristics and compute a font size that fits the box, so the
saved file renders correctly everywhere.

The estimate is deliberately conservative (slightly over-estimates text
width) and treats CJK characters as full-width.
"""

from __future__ import annotations

import math

# Average glyph width as a fraction of the font size (points)
_LATIN_WIDTH = 0.55
_WIDE_WIDTH = 1.05  # CJK and other full-width characters

# Line height as a multiple of the font size
_LINE_HEIGHT = 1.25

POINTS_PER_INCH = 72.0


def _char_width_factor(char: str) -> float:
    """Width of one character as a fraction of the font size."""
    code = ord(char)
    # CJK Unified Ideographs, Hiragana/Katakana, Hangul, full-width forms
    if (
        0x1100 <= code <= 0x11FF
        or 0x2E80 <= code <= 0x9FFF
        or 0xAC00 <= code <= 0xD7AF
        or 0xF900 <= code <= 0xFAFF
        or 0xFF00 <= code <= 0xFF60
    ):
        return _WIDE_WIDTH
    return _LATIN_WIDTH


def text_width_points(text: str, font_size: float) -> float:
    """Estimate the rendered width of a line of text, in points."""
    return sum(_char_width_factor(c) for c in text) * font_size


def estimate_lines(text: str, font_size: float, box_width_inches: float) -> int:
    """Estimate how many lines a paragraph wraps into inside a box."""
    box_points = max(box_width_inches, 0.1) * POINTS_PER_INCH
    lines = 0
    for paragraph in text.splitlines() or [""]:
        lines += max(1, math.ceil(text_width_points(paragraph, font_size) / box_points))
    return lines


def estimate_height_inches(text: str, font_size: float, box_width_inches: float) -> float:
    """Estimate the rendered height of wrapped text, in inches."""
    lines = estimate_lines(text, font_size, box_width_inches)
    return lines * font_size * _LINE_HEIGHT / POINTS_PER_INCH


def fit_font_size(
    paragraphs: list[str],
    box_width_inches: float,
    box_height_inches: float,
    font_size: float,
    min_size: float = 8.0,
) -> int:
    """Find the largest font size (<= font_size) whose text fits the box.

    Args:
        paragraphs: The text content, one string per paragraph
        box_width_inches: Available width in inches
        box_height_inches: Available height in inches
        font_size: The requested (maximum) font size in points
        min_size: The smallest size to shrink to (default: 8)

    Returns:
        A font size in whole points between min_size and font_size
    """
    text = "\n".join(paragraphs)
    size = float(font_size)
    while size > min_size:
        if estimate_height_inches(text, size, box_width_inches) <= box_height_inches:
            break
        size -= 1
    return max(int(size), int(min_size))
