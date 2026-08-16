# Percentage-Based Positioning

Positions and sizes in EasyPPTX are percentages of the slide's dimensions. As of 0.7.0, **bare numbers are percentages too**: `x=10` means the same thing as `x="10%"`. Absolute positioning in inches is available through the `in_()` helper.

1. **Percentage-based positioning** (the default): bare numbers (`x=10`) or percentage strings (`x="10%"`)
2. **Absolute positioning**: exact lengths in inches via `in_()` (`x=in_(1.5)`)

Percentage-based positioning makes layouts responsive to different slide sizes and aspect ratios, similar to the approach used in CSS for web development.

> **Changed in 0.7.0:** Floats no longer mean inches. In earlier versions `x=1.0` placed an element 1 inch from the left; it now means 1% of the slide width. Use `in_(1.0)` for the old behavior. See [Migrating to 0.7.0](migration.md).

## How It Works

When using percentage values, the position and size are calculated as a percentage of the slide's total width or height:

- `10` (or `"10%"`) of width means 10% of the slide's width
- `50` (or `"50%"`) of height means 50% of the slide's height
- `100` (or `"100%"`) represents the full slide dimension

## Usage Examples

### Text Positioning

```python
from easypptx import Presentation, Text, in_

# Create a presentation
pres = Presentation()
slide = pres.add_slide()
text = Text(slide)

# Add text using percentage-based positioning (bare numbers are percentages)
text.add_title("Centered Title", x=10, y=5, width=80, height=15)

# Percentage strings work identically
text.add_paragraph(
    "This text is positioned at 20% from the left and 30% from the top.",
    x="20%",
    y="30%",
    width="60%",
    height="10%"
)

# Mix absolute inches (via in_) with percentages
text.add_paragraph(
    "This text uses absolute x (2 inches) and percentage y (70%).",
    x=in_(2.0),
    y="70%",
    width="50%",
    height=in_(1.0)
)

# Save the presentation
pres.save("percentage_positioning.pptx")
```

### Image Positioning

```python
from easypptx import Presentation, Image

# Create a presentation
pres = Presentation()
slide = pres.add_slide()
img = Image(slide)

# Add image at 10% from left, 30% from top, with 80% width
img.add(
    "logo.png",
    x=10,
    y=30,
    width=80
)
```

### Shape Positioning

```python
from easypptx import Presentation
from pptx.enum.shapes import MSO_SHAPE

# Create a presentation
pres = Presentation()
slide = pres.add_slide()

# Add a rectangle shape using percentage positioning
slide.add_shape(
    shape_type=MSO_SHAPE.RECTANGLE,
    x=25,
    y=60,
    width=50,
    height=10,
    fill_color="blue"
)
```

## Absolute Positioning with in_()

When you need physical units — for example a logo that must be exactly 2 inches wide — use `in_()`:

```python
from easypptx import Presentation, in_

pres = Presentation()
slide = pres.add_slide()

# 0.5 inches from the left and top, 2 inches wide
slide.add_image("logo.png", x=in_(0.5), y=in_(0.5), width=in_(2))
```

`in_()` values can be mixed freely with percentages, and they are not clamped, so they can also place content partially off-slide when needed.

## Responsive Layouts and Centering

Percentage-based positioning is the responsive mechanism in EasyPPTX: elements keep their relative position and size when the slide dimensions or aspect ratio change.

To center content, use `align="center"` together with a symmetric `x`/`width` pair. For example, `x=10` with `width=80` leaves 10% on each side, so the element stays centered at any aspect ratio:

```python
slide.add_text(
    text="Centered Title",
    x=10,
    y=5,
    width=80,
    height=10,
    align="center"
)
```

The same principle works for images and shapes: give them a symmetric x/width (e.g. `x=20`, `width=60`) to keep them horizontally centered.

## Default Positions

Content methods default to sensible percentage positions, so you can omit coordinates entirely:

| Method | x | y | width | height |
| --- | --- | --- | --- | --- |
| `slide.add_text` | 5 | 5 | 90 | 10 |
| `slide.add_image` | 5 | 5 | auto | auto |
| `slide.add_shape` | 5 | 5 | 40 | 10 |
| `slide.add_table` | 5 | 20 | auto | auto |
| `slide.add_chart` | 10 | 20 | 60 | 60 |
| `slide.add_bullets` | 5 | 20 | 90 | 70 |
| `slide.add_pyplot` | 5 | 15 | 90 | 75 |

## Validation and Warnings

- Unknown parameters passed to content methods trigger a warning instead of being silently ignored.
- Out-of-range percentages (e.g. `150` or `"150%"`) are clamped to 0-100 with a warning. Use `in_()` when you intentionally need off-slide placement.

## Benefits of Percentage-Based Positioning

1. **Responsive layouts**: Elements maintain their relative positions regardless of the presentation's aspect ratio
2. **Easier layout adjustments**: Change slide dimensions without needing to recalculate all positions
3. **Simpler scaling**: Create presentations that work well with different display sizes
4. **Layout consistency**: Maintain the same visual layout across slides with different aspect ratios

## Implementation Details

The conversion from percentages to absolute dimensions happens automatically for bare numbers and `"%"` strings. The conversion formula is:

```
absolute_value = (percentage / 100) × slide_dimension
```

Where:

- `percentage` is the numeric value (without the % sign)
- `slide_dimension` is either the slide width or height (depending on the axis)

Values created with `in_()` are `pptx.util.Length` objects and pass through unchanged. The conversion helpers live in the `easypptx.positioning` module.
