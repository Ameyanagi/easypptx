# Styling and Formatting

EasyPPTX provides a comprehensive set of styling and formatting options to customize the appearance of your presentations.

## Style Objects

*New in 0.7.0.* The `TextStyle`, `TableStyle`, and `ChartStyle` dataclasses (importable from `easypptx` or `easypptx.styles`) bundle formatting that would otherwise be repeated on every call. Pass them via the `style=` parameter of `add_text`, `add_bullets`, `add_table`, and `add_chart`:

```python
from easypptx import Presentation, TextStyle, TableStyle, ChartStyle

pres = Presentation()
slide = pres.add_slide()

# Define a style once...
heading = TextStyle(font_size=28, font_bold=True, color="blue", align="center")

# ...and reuse it
slide.add_text("Results", style=heading)
slide.add_text("Outlook", y=20, style=heading)

# Bulleted lists take the same TextStyle
body = TextStyle(font_size=20, color="black")
slide.add_bullets(["Point one", "Point two"], style=body)

# Tables and charts have their own style dataclasses
slide.add_table(
    [["A", "B"], [1, 2]],
    style=TableStyle(has_header=True, style_id=5),
)
slide.add_chart(
    categories=["A", "B", "C"],
    values=[10, 20, 30],
    style=ChartStyle(chart_type="bar", has_legend=False),
)
```

Explicit keyword arguments always beat the style object, so a style acts as a set of defaults you can override per call:

```python
# heading sets color="blue", but this call renders red
slide.add_text("Warning", style=heading, color="red")
```

The fields:

- `TextStyle`: `font_name`, `font_size`, `font_bold`, `font_italic`, `color`, `align`, `vertical`
- `TableStyle`: `has_header`, `style_id`
- `ChartStyle`: `chart_type`, `has_legend`, `legend_position`

All fields are optional; unset fields fall back to the normal defaults.

## Themes

*New in 0.7.0.* A `Theme` groups a background color with title and body text styles and applies to the whole presentation. Three presets are built in — `"light"`, `"dark"`, and `"corporate"` — selectable by name:

```python
from easypptx import Presentation

pres = Presentation(theme="dark")
slide = pres.add_slide(title="Styled by the theme")
slide.add_text("Body text inherits the theme's color", y=25)
```

A theme sets the default slide background, cascades its body `TextStyle` to text content through the template-defaults machinery, and styles titles created with `add_slide(title=...)`. Explicit arguments (and slide-level `bg_color`) still win.

Custom themes are plain `Theme` instances; every field is optional:

```python
from easypptx import Presentation, TextStyle, Theme

brand = Theme(
    bg_color=(10, 20, 40),
    title=TextStyle(font_size=30, font_bold=True, color="white", align="left"),
    body=TextStyle(color="lightgray"),
    accent_color="orange",
    palette=[(245, 166, 35), (91, 155, 213), (143, 184, 107)],
)
pres = Presentation(theme=brand)
```

Themes also work with [markdown decks](markdown.md): `Presentation.from_markdown("deck.md", theme="corporate")` or a `theme:` key in the frontmatter.

### Theme Palettes

*New in 0.8.0.* The `palette` field lists series colors (names or RGB tuples) for native PowerPoint charts. When a themed presentation adds a native chart, the series are colored from the theme's palette automatically, cycling if a chart has more series than colors. The built-in `light`, `dark`, and `corporate` themes all ship palettes tuned to their backgrounds.

```python
pres = Presentation(theme="dark")
slide = pres.add_slide(title="Q3")

# Series colors come from the dark theme's palette
slide.add_chart(data=df, chart_type="column")

# An explicit palette= always beats the theme
slide.add_chart(data=df, chart_type="column", palette=["red", "blue"])
```

See [Native Chart Styling](plots.md#native-chart-styling) for the other chart styling options (`show_values`, `number_format`, axis titles and limits).

## Default Font and Colors

EasyPPTX comes with sensible defaults:

- **Default font**: Meiryo
- **Default color scheme**:
  - Black: RGB(0x10, 0x10, 0x10)
  - Red: RGB(0xFF, 0x40, 0x40)
  - Green: RGB(0x40, 0xFF, 0x40)
  - Blue: RGB(0x40, 0x40, 0xFF)
  - White: RGB(0xFF, 0xFF, 0xFF)

These constants are defined in `easypptx.common` (`COLORS`, `ALIGN`, `VERTICAL`, `DEFAULT_FONT`). The `Presentation.COLORS` alias still works.

## Text Formatting

### Font Properties

```python
from easypptx import Presentation, Text

# Create a presentation
pres = Presentation()
slide = pres.add_slide()
text = Text(slide)

# Add a title with custom font properties
text.add_title(
    "Styled Title",
    font_size=48,                    # Font size in points
    font_name="Meiryo",              # Font name
    color="blue",                    # Color name from default colors
    align="center"                   # Text alignment
)

# Add a paragraph with custom styling
text.add_paragraph(
    "This is a formatted paragraph",
    font_size=24,
    font_bold=True,
    font_italic=True,
    font_name="Arial",
    color=(128, 0, 128),             # Custom RGB color (purple)
    align="left",
    vertical="middle"
)

pres.save("styled_text.pptx")
```

### Text Alignment

You can control both horizontal and vertical alignment:

**Horizontal alignment options**:
- `"left"`: Align text to the left
- `"center"`: Center text horizontally
- `"right"`: Align text to the right

**Vertical alignment options**:
- `"top"`: Align text to the top
- `"middle"`: Center text vertically
- `"bottom"`: Align text to the bottom

```python
# Center-aligned text
text.add_paragraph(
    "This text is centered",
    align="center",
    vertical="middle"
)
```

### Text Fitting

*New in 0.8.0.* `add_text` and `add_bullets` take a `fit=` parameter controlling how text relates to its box:

- `fit="shrink"` (the default) estimates how the text wraps — using glyph-width heuristics that treat CJK characters as full-width — and writes a font size that fits the box into the file, in addition to setting PowerPoint's autofit flag. Because the fitting size is computed at save time (not just flagged for PowerPoint to compute later), the deck renders correctly in LibreOffice, quick-look previews, and other renderers too, not only after opening it in PowerPoint. Text never shrinks below 8 pt.
- `fit="resize"` grows the box to fit the text (PowerPoint's `SHAPE_TO_FIT_TEXT`).
- `fit="none"` leaves both the font size and the box alone.

```python
slide.add_text("A long paragraph...", width=30, height=10)              # shrinks to fit
slide.add_text("A long paragraph...", width=30, height=10, fit="resize")
slide.add_text("A long paragraph...", width=30, height=10, fit="none")
```

Any other value raises a `ValueError`. The estimation helpers live in `easypptx.textfit`.

### Color Specification

You can specify colors in two ways:

1. **Named colors** from the default color dictionary:
   ```python
   text.add_paragraph("Red text", color="red")
   text.add_paragraph("Blue text", color="blue")
   ```

2. **RGB tuples** for custom colors:
   ```python
   text.add_paragraph("Purple text", color=(128, 0, 128))
   text.add_paragraph("Orange text", color=(255, 165, 0))
   ```

## Shape Styling

When adding shapes, you can specify fill colors:

```python
from easypptx import Presentation
from pptx.enum.shapes import MSO_SHAPE

# Create a presentation
pres = Presentation()
slide = pres.add_slide()

# Add a colored rectangle
slide.add_shape(
    shape_type=MSO_SHAPE.RECTANGLE,
    x="10%",
    y="20%",
    width="80%",
    height="10%",
    fill_color="blue"  # Named color
)

# Add an oval with custom RGB color
slide.add_shape(
    shape_type=MSO_SHAPE.OVAL,
    x="10%",
    y="40%",
    width="80%",
    height="10%",
    fill_color=(255, 165, 0)  # Orange
)
```

## Formatting Existing Text Frames

You can also format existing text frames using the static `format_text_frame` method:

```python
from easypptx import Presentation, Text
from easypptx.text import Text

# Create a presentation
pres = Presentation()
slide = pres.add_slide()

# Add a textbox
text_box = slide.add_text("This text will be formatted")
text_frame = text_box.text_frame

# Apply formatting to the text frame
Text.format_text_frame(
    text_frame,
    font_size=24,
    font_bold=True,
    font_italic=True,
    font_name="Calibri",
    color="green",
    align="center",
    vertical="middle"
)

pres.save("formatted_text_frame.pptx")
```

## Using Template Styles

EasyPPTX allows you to use styles from existing PowerPoint templates:

```python
from easypptx import Presentation, Text

# Create a presentation from a template
pres = Presentation(template_path="template.pptx")

# Add a slide
slide = pres.add_slide()

# Add content that will use the template's styles
text = Text(slide)
text.add_title("Title Using Template Style")
text.add_paragraph("This text will use the template's default styling")

pres.save("template_styled.pptx")
```

## Implementation Details

The styling options are implemented using python-pptx's underlying API, with EasyPPTX providing a more intuitive interface.

Colors are converted to `RGBColor` objects, and font properties are applied to paragraph objects in the text frames.

## Professional Defaults (0.9.0)

Themed decks apply a designed look automatically:

- Slide titles are left-aligned with a short accent bar underneath
  (`Theme(title_accent=False)` disables the bar).
- Tables get the theme's header fill, subtle row banding, cell padding,
  and automatic right-alignment for numeric columns. Custom themes can
  provide their own spec via `Theme(table={...})` with keys
  `header_fill`, `header_color`, `band_fills`, `body_color`,
  `font_size`, and `header_font_size`.
- Bullets use paragraph spacing and step nested levels down in size.
- Chart gridlines fade toward the background (luminance-aware).
