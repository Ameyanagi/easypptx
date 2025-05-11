# Percentage-Based Positioning

EasyPPTX supports two positioning methods for slide elements:

1. **Absolute positioning**: Specifying exact dimensions in inches
2. **Percentage-based positioning**: Using percentages of the slide's dimensions

Percentage-based positioning makes layouts responsive to different slide sizes and aspect ratios, similar to the approach used in CSS for web development.

## How It Works

When using percentage values, the position and size are calculated as a percentage of the slide's total width or height:

- "10%" of width means 10% of the slide's width
- "50%" of height means 50% of the slide's height
- "100%" represents the full slide dimension

## Usage Examples

### Text Positioning

```python
from easypptx import Presentation, Text

# Create a presentation
pres = Presentation()
slide = pres.add_slide()
text = Text(slide)

# Add text using percentage-based positioning
text.add_title("Centered Title", x="10%", y="5%", width="80%", height="15%")

# Add paragraphs using percentage positioning
text.add_paragraph(
    "This text is positioned at 20% from the left and 30% from the top.",
    x="20%",
    y="30%",
    width="60%",
    height="10%"
)

# Mix percentages with absolute values
text.add_paragraph(
    "This text uses absolute x (2 inches) and percentage y (70%).",
    x=2.0,
    y="70%",
    width="50%",
    height=1.0
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
    x="10%",
    y="30%",
    width="80%"
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
    x="25%",
    y="60%",
    width="50%",
    height="10%",
    fill_color="blue"
)
```

## Benefits of Percentage-Based Positioning

1. **Responsive layouts**: Elements maintain their relative positions regardless of the presentation's aspect ratio
2. **Easier layout adjustments**: Change slide dimensions without needing to recalculate all positions
3. **Simpler scaling**: Create presentations that work well with different display sizes
4. **Layout consistency**: Maintain the same visual layout across slides with different aspect ratios

## Implementation Details

The conversion from percentages to absolute dimensions happens automatically when you specify a position or size value with a "%" suffix. The conversion formula is:

```
absolute_value = (percentage / 100) × slide_dimension
```

Where:
- `percentage` is the numeric value (without the % sign)
- `slide_dimension` is either the slide width or height (depending on the axis)
