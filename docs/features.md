# EasyPPTX Features

EasyPPTX provides a comprehensive set of features for creating and manipulating PowerPoint presentations programmatically. This guide provides an overview of all available features.

## Core Features

### Creating Presentations

```python
from easypptx import Presentation

# Create a new presentation with default 16:9 aspect ratio
pres = Presentation()

# Create a presentation with 4:3 aspect ratio
pres = Presentation(aspect_ratio="4:3")

# Create a presentation with custom dimensions
pres = Presentation(width_inches=12, height_inches=9)

# Create a presentation from a template
pres = Presentation(template_path="template.pptx")

# Open an existing presentation
pres = Presentation.open("existing.pptx")

# Save a presentation
pres.save("output.pptx")
```

### Working with Slides

```python
from easypptx import Presentation

# Create a presentation
pres = Presentation()

# Add a slide
slide = pres.add_slide()

# Add a slide with a specific layout
slide = pres.add_slide(layout_index=1)  # Content layout

# Add a slide with a title (title_color sets the title color)
slide = pres.add_slide(title="Section Title", title_color="blue")

# Opt a single slide out of the presentation's default TOML template
slide = pres.add_slide(template_toml=False)

# Get all slides
slides = pres.slides

# Clear a slide (remove all shapes)
slide.clear()

# Get or set the slide title
title = slide.title
slide.title = "New Title"
```

## Text Manipulation

```python
from easypptx import Presentation, Text

# Create a presentation
pres = Presentation()
slide = pres.add_slide()
text = Text(slide)

# Add a title
text.add_title("Presentation Title")

# Add a paragraph
text.add_paragraph("This is a paragraph of text")

# Add formatted text
text.add_paragraph(
    "Formatted text",
    font_size=24,
    font_bold=True,
    font_italic=True,
    font_name="Meiryo",
    color="blue",
    align="center",
    vertical="middle"
)
```

## Working with Images

```python
from easypptx import Presentation, Image

# Create a presentation
pres = Presentation()
slide = pres.add_slide()
img = Image(slide)

# Add an image
img.add("image.png", x=1, y=2, width=4)

# Add an image with maintenance of aspect ratio
img.add("image.png", x=1, y=2, width=4)  # Height calculated automatically

# Add an image with specific dimensions
img.add("image.png", x=1, y=2, width=4, height=3)
```

## Creating Tables

> pandas DataFrame support requires the optional extra: `pip install "easypptx[dataframe]"`.

```python
from easypptx import Presentation

# Create a presentation
pres = Presentation()
slide = pres.add_slide()

# Add a simple table
data = [["Header 1", "Header 2"], ["Cell A1", "Cell A2"], ["Cell B1", "Cell B2"]]
slide.add_table(data, x=1, y=1, width=4, height=2)

# Add a table from a pandas DataFrame
import pandas as pd
df = pd.DataFrame({"Name": ["John", "Jane"], "Score": [85, 92]})
slide.add_table(df, x=1, y=3, width=4, height=2)

# Include the DataFrame index as the first column
from easypptx import Table
Table(slide).from_dataframe(df, x=1, y=5, include_index=True)

# Add a table with a header row and style
slide.add_table(
    data,
    x=1,
    y=5,
    width=4,
    height=2,
    has_header=True,
    style=5
)
```

## Creating Charts

```python
from easypptx import Presentation

# Create a presentation
pres = Presentation()
slide = pres.add_slide()

# Add a simple chart from explicit categories and values
slide.add_chart(
    chart_type="pie",
    categories=["A", "B", "C"],
    values=[10, 20, 30],
    x=1, y=1, width=4, height=3,
    title="Chart Title"
)

# Add a chart from a pandas DataFrame
import pandas as pd
df = pd.DataFrame({"Category": ["A", "B", "C"], "Value": [10, 20, 30]})
slide.add_chart(
    data=df,
    chart_type="column",
    category_column="Category",
    value_columns="Value",
    x=6,
    y=1,
    title="DataFrame Chart"
)
```

## Direct Object APIs

EasyPPTX provides a set of content methods on the Slide class to manipulate slide content without creating separate object instances.

> **Deprecated:** The pass-through variants on `Presentation` that take a slide as the first argument (`pres.add_text(slide, ...)`, `pres.add_image(slide, ...)`, `pres.add_shape(slide, ...)`, `pres.add_table(slide, ...)`, `pres.add_chart(slide, ...)`, `pres.add_pyplot(slide, ...)`) still work but emit a `DeprecationWarning`. Use the `Slide` methods shown below instead.

```python
from easypptx import Presentation

# Create a presentation
pres = Presentation()
slide = pres.add_slide()

# Add text directly
slide.add_text(
    text="Hello World",
    x="10%",
    y="20%",
    font_size=24,
    font_bold=True,
    color="blue"
)

# Add an image directly
slide.add_image(
    image_path="image.png",
    x="10%",
    y="30%",
    width="30%"
)

# Add a shape directly (string shape names work in addition to MSO_SHAPE enums)
slide.add_shape(
    shape_type="ROUNDED_RECTANGLE",
    x="50%",
    y="30%",
    width="40%",
    height="15%",
    fill_color="blue"
)

# Add a table directly
data = [["Name", "Value"], ["Item 1", 100], ["Item 2", 200]]
slide.add_table(
    data=data,
    x="10%",
    y="50%",
    width="80%",
    height="20%"
)

# Add a chart directly
import pandas as pd
df = pd.DataFrame({"Category": ["A", "B", "C"], "Value": [10, 20, 30]})
slide.add_chart(
    data=df,
    chart_type="column",
    x="10%",
    y="75%",
    width="80%",
    height="20%",
    category_column="Category",
    value_columns="Value"
)

# Add a matplotlib figure (embedded via an in-memory stream, no temp files)
import matplotlib.pyplot as plt
from easypptx import Pyplot

plt.figure(figsize=(10, 6))
plt.plot([1, 2, 3, 4], [1, 4, 9, 16])
plt.title('Sample Plot')

Pyplot.add(
    slide=slide,
    figure=plt.gcf(),
    position={"x": "60%", "y": "50%", "width": "30%", "height": "20%"}
)
```

Unknown parameters passed to these methods now trigger a warning instead of being silently ignored. Out-of-range percentages (e.g. `"150%"`) are clamped with a warning.

## Advanced Features

### Grid Layout System

```python
from easypptx import Presentation

# Create a presentation
pres = Presentation()

# Create a slide with a 2x2 grid in one step
slide, grid = pres.add_grid_slide(
    rows=2,
    cols=2,
    title="Grid Example",
    padding=5.0
)

# Add content to a cell using grid[row, col] syntax
grid[0, 0].add_text(
    text="Top Left Cell",
    font_size=24,
    align="center",
    vertical="middle"
)

grid[0, 1].add_image(image_path="image.png")

grid[1, 0].add_table(data=[["A", "B"], [1, 2]])

grid[1, 1].add_text(
    text="Bottom Right Cell",
    font_size=24,
    align="center",
    vertical="middle"
)
```

See [Grid Layout](grid_layout.md) for more details.

### Percentage-Based Positioning

```python
from easypptx import Presentation, Text

pres = Presentation()
slide = pres.add_slide()
text = Text(slide)

text.add_title("Percentage Positioning", x="10%", y="5%", width="80%", height="15%")
text.add_paragraph("Positioned using percentages", x="20%", y="30%", width="60%", height="10%")
```

See [Percentage-Based Positioning](percentage_positioning.md) for more details.

### Dark Theme Support

```python
from easypptx import Presentation, Text

# Create a presentation with black background
pres = Presentation(default_bg_color="black")
slide = pres.add_slide()
text = Text(slide)

# Add high-contrast text
text.add_title("Dark Theme Presentation", color="cyan")
text.add_paragraph("High contrast text for readability", color="white")

# Add a slide with custom background
slide2 = pres.add_slide(bg_color=(0, 20, 40))  # Dark blue
```

See [Dark Theme](dark_theme.md) for more details.

### Auto-Alignment of Multiple Objects

```python
from easypptx import Presentation
from pptx.enum.shapes import MSO_SHAPE

pres = Presentation()
slide = pres.add_slide()

objects = [
    {"type": "text", "text": "Item 1", "color": "black"},
    {"type": "text", "text": "Item 2", "color": "red"},
    {"type": "shape", "shape_type": MSO_SHAPE.RECTANGLE, "fill_color": "green"}
]

slide.add_multiple_objects(
    objects_data=objects,
    layout="grid",
    padding_percent=5.0,
    start_x="10%",
    start_y="20%",
    width="80%",
    height="60%"
)
```

See [Auto-Alignment of Multiple Objects](auto_alignment.md) for more details.

### Custom Styling and Formatting

```python
from easypptx import Presentation, Text

pres = Presentation()
slide = pres.add_slide()
text = Text(slide)

text.add_title("Styled Title", font_name="Meiryo", color="blue", align="center")
```

See [Styling and Formatting](styling.md) for more details.

### PowerPoint Templates

```python
from easypptx import Presentation

pres = Presentation(template_path="template.pptx")
slide = pres.add_slide()
```

See [Using PowerPoint Templates](templates.md) for more details.

### Aspect Ratio Options

```python
from easypptx import Presentation

# Default 16:9 widescreen
pres = Presentation()

# 4:3 standard
pres = Presentation(aspect_ratio="4:3")

# 16:10 widescreen alternative
pres = Presentation(aspect_ratio="16:10")

# A4 paper size
pres = Presentation(aspect_ratio="A4")

# US Letter paper size
pres = Presentation(aspect_ratio="LETTER")

# Custom dimensions
pres = Presentation(width_inches=12, height_inches=9)
```

## Examples

EasyPPTX comes with several example scripts demonstrating various features:

- **quick_start.py**: Basic introduction to EasyPPTX
- **basic_demo.py**: Introduction to basic features
- **comprehensive_example.py**: Full-featured business presentation
- **aspect_ratio_example.py**: Demonstration of aspect ratio options
- **extended_features_example.py**: Showcase of percentage-based positioning and auto-alignment
- **plot_example.py**: Examples of using matplotlib and seaborn plots
- **object_api_example.py**: Demonstrates the direct object APIs for slide manipulation
- **Grid Examples**:
  - **001_basic_grid.py**: Basic grid creation and usage
  - **002_grid_indexing.py**: Grid indexing and iteration features
  - **003_nested_grid.py**: Nested grids and cell merging
  - **004_autogrid.py**: Automatic grid layout features

See the [Examples](https://github.com/Ameyanagi/EasyPPTX/tree/main/examples) directory for the full source code.
