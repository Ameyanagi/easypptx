# Styling and Formatting

EasyPPTX provides a comprehensive set of styling and formatting options to customize the appearance of your presentations.

## Default Font and Colors

EasyPPTX comes with sensible defaults:

- **Default font**: Meiryo
- **Default color scheme**:
  - Black: RGB(0x40, 0x40, 0x40)
  - Red: RGB(0xFF, 0x40, 0x40)
  - Green: RGB(0x40, 0xFF, 0x40)
  - Blue: RGB(0x40, 0x40, 0xFF)
  - White: RGB(0xFF, 0xFF, 0xFF)

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
