# EasyPPTX

[![Release](https://img.shields.io/github/v/release/Ameyanagi/EasyPPTX)](https://img.shields.io/github/v/release/Ameyanagi/EasyPPTX)
[![Build status](https://img.shields.io/github/actions/workflow/status/Ameyanagi/EasyPPTX/main.yml?branch=main)](https://github.com/Ameyanagi/EasyPPTX/actions/workflows/main.yml?query=branch%3Amain)
[![codecov](https://codecov.io/gh/Ameyanagi/EasyPPTX/branch/main/graph/badge.svg)](https://codecov.io/gh/Ameyanagi/EasyPPTX)
[![Commit activity](https://img.shields.io/github/commit-activity/m/Ameyanagi/EasyPPTX)](https://img.shields.io/github/commit-activity/m/Ameyanagi/EasyPPTX)
[![License](https://img.shields.io/github/license/Ameyanagi/EasyPPTX)](https://img.shields.io/github/license/Ameyanagi/EasyPPTX)

A Python library for easily creating and manipulating PowerPoint presentations programmatically with simple APIs, designed to be easy for both humans and AI assistants to use.

- **Github repository**: <https://github.com/Ameyanagi/EasyPPTX/>
- **Documentation** <https://Ameyanagi.github.io/EasyPPTX/>

## Features

- Simple, intuitive API for PowerPoint manipulation
- Create slides with text, images, tables, and charts
- Format elements with easy-to-use styling options
- Default 16:9 aspect ratio with support for multiple ratio options
- Percentage-based positioning for responsive layouts
- Auto-alignment of multiple objects (grid, horizontal, vertical)
- Advanced Grid layout system with convenience methods
- Grid iteration, indexing, and nested grid capabilities
- Dark theme support with custom background colors
- Expanded color palette for modern designs
- Default font settings with Meiryo
- Support for reference PowerPoint templates
- Optimized for use with AI assistants and LLMs
- Built on top of python-pptx with a more user-friendly interface

## Installation

```bash
pip install easypptx
```

## Quick Start

```python
from easypptx import Presentation
import pandas as pd

# Create a new presentation (uses 16:9 aspect ratio by default)
pres = Presentation()

# Add a slide with a title
slide = pres.add_slide(title="EasyPPTX Demo")

# Add text directly to the slide
slide.add_text(
    text="This presentation was created with EasyPPTX",
    x="10%",
    y="20%",
    width="80%",
    height="10%",
    font_size=24
)

# Add an image
slide.add_image(
    image_path="path/to/image.png",
    x="10%",
    y="35%",
    width="40%"
)

# Create a table
data = [["Name", "Value"], ["Item 1", 100], ["Item 2", 200]]
slide.add_table(
    data=data,
    x="60%",
    y="35%",
    width="30%",
    has_header=True
)

# Add a slide with a chart from pandas DataFrame
chart_slide = pres.add_slide(title="Chart Example")

df = pd.DataFrame({"Category": ["A", "B", "C"], "Value": [10, 20, 30]})
chart_slide.add_chart(
    data=df,
    chart_type="pie",
    category_column="Category",
    value_column="Value",
    x="20%",
    y="20%",
    width="60%",
    height="60%",
    title="Sample Chart"
)

# Save the presentation
pres.save("example.pptx")
```

## Aspect Ratio Options

EasyPPTX supports multiple aspect ratios for presentations:

```python
# Default 16:9 widescreen presentation
pres = Presentation()

# Standard 4:3 presentation
pres = Presentation(aspect_ratio="4:3")

# Other supported options: "16:10", "A4", "LETTER"
pres = Presentation(aspect_ratio="16:10")

# Custom dimensions (width and height in inches)
pres = Presentation(width_inches=12, height_inches=9)
```

## Custom Reference PPTX Files

You can use custom PPTX files as references for your presentations:

```python
# Use a custom reference PPTX file (keeping all layouts and styles)
pres = Presentation(reference_pptx="path/to/your/reference.pptx")

# Specify a custom blank layout index (if the default auto-detection doesn't work)
pres = Presentation(
    reference_pptx="path/to/your/reference.pptx",
    blank_layout_index=2  # Use the third layout as blank
)

# When opening existing presentations, you can also specify blank layout
pres = Presentation.open(
    "path/to/existing.pptx",
    blank_layout_index=4  # Use the fifth layout as blank
)
```

### TOML Template Reference PPTX Specification

You can also specify reference PPTX files in TOML template files:

```toml
# Specify the reference PPTX file path (absolute or relative to the TOML file)
reference_pptx = "../references/my_template.pptx"

# Optionally specify the blank layout index
blank_layout_index = 3

# Rest of your template content
[title]
text = "Presentation Title"
position = { x = "10%", y = "30%", width = "80%", height = "20%" }
font = { name = "Meiryo", size = 44, bold = true }
align = "center"
# ...
```

Then use the template in your code:

```python
from easypptx import Presentation
from easypptx.template import TemplateManager

# Load the template file
template_manager = TemplateManager()
template_name = template_manager.load("path/to/template.toml")

# Create a presentation and use the template
pres = Presentation()
slide = pres.add_slide_from_template(template_name)
# The reference PPTX specified in the TOML is automatically loaded
```

## Percentage-Based Positioning

Position and size elements using percentages of the slide dimensions:

```python
# Add text at 20% from the left, 30% from the top, 60% width, 10% height
text.add_paragraph("Positioned with percentages", x="20%", y="30%", width="60%", height="10%")

# Add an image at 10% from the left, 50% from the top, 40% width
img.add("image.png", x="10%", y="50%", width="40%")

# Add a shape using percentages
slide.add_shape(
    x="70%",
    y="50%",
    width="20%",
    height="20%",
    fill_color="blue"
)
```

## Auto-Alignment of Multiple Objects

Easily align multiple objects in a grid, horizontal, or vertical layout:

```python
# Define objects to be added
objects = [
    {"type": "text", "text": "Item 1", "color": "black"},
    {"type": "text", "text": "Item 2", "color": "red"},
    {"type": "text", "text": "Item 3", "color": "blue"},
    {"type": "shape", "shape_type": MSO_SHAPE.RECTANGLE, "fill_color": "green"}
]

# Add objects in a grid layout (2x2)
slide.add_multiple_objects(
    objects_data=objects,
    layout="grid",
    padding_percent=5.0,
    start_x="10%",
    start_y="30%",
    width="80%",
    height="60%"
)

# Add objects in a horizontal layout (row)
slide.add_multiple_objects(
    objects_data=objects,
    layout="horizontal",
    start_y="50%",
    height="20%"
)
```

## Enhanced Grid Layout System

EasyPPTX provides a powerful Grid layout system for creating complex and responsive layouts:

### Creating Grids with Add Grid Slide

Create a slide with a grid layout in one step:

```python
# Create a slide with a 2x2 grid
slide, grid = pres.add_grid_slide(
    title="Grid Layout Example",
    rows=2,
    cols=2,
    title_height="10%",
    padding=5.0
)

# Add content directly to grid cells using the enhanced access API
grid[0, 0].add_text(
    text="Top Left Cell",
    font_size=18,
    align="center",
    vertical="middle"
)

grid[0, 1].add_image(
    image_path="path/to/image.jpg",
    maintain_aspect_ratio=True
)

# Generate a matplotlib figure
import matplotlib.pyplot as plt
fig, ax = plt.subplots()
ax.plot([1, 2, 3, 4], [1, 4, 2, 3])
ax.set_title("Sample Plot")

# Add the matplotlib figure to a grid cell
grid[1, 0].add_pyplot(figure=fig, dpi=150)

# Add a table to the remaining cell
grid[1, 1].add_table(
    data=[["A", "B"], [1, 2]],
    has_header=True
)
```

### Grid Iteration and Indexing

Easily access and iterate through grid cells:

```python
# Access cells using indexing
cell1 = grid[0, 0]  # Row 0, Column 0
cell2 = grid[3]     # Flat index (row-major order)

# Iterate through all cells
for cell in grid:
    print(f"Cell at {cell.row}, {cell.col}")

# Use the flat property for flat iteration
for cell in grid.flat:
    print(f"Cell content: {cell.content}")

# Merge cells to create larger areas
merged_cell = grid.merge_cells(0, 0, 1, 1)  # 2x2 merged area
```

### Row-Based Grid Access

Easily add content to grid rows without specifying exact column indices:

```python
# Create a slide with a 3x3 grid
slide, grid = pres.add_grid_slide(
    title="Row-Based Grid Access",
    rows=3,
    cols=3,
    padding=5.0
)

# Add content to the first row using row-level access
# This automatically adds each item to the next available cell in the row
grid[0].add_text(
    text="First item in row 0",
    font_size=16,
    align="center"
)

grid[0].add_text(
    text="Second item in row 0",
    font_size=16,
    align="center"
)

grid[0].add_text(
    text="Third item in row 0",
    font_size=16,
    align="center"
)

# Add content to the second row
grid[1].add_text(
    text="First item in row 1",
    font_size=16,
    align="center"
)

grid[1].add_text(
    text="Second item in row 1",
    font_size=16,
    align="center"
)
```

## Templates

EasyPPTX supports multiple template formats for consistent presentation design.

### Reference PowerPoint Templates

Use existing PowerPoint files as templates:

```python
# Create a presentation using an existing template
pres = Presentation(template_path="template.pptx")

# Add a slide with a title
slide = pres.add_slide(title="Presentation with Template")

# Add content to the slide
slide.add_text(
    text="Content using the template styles",
    x="10%",
    y="30%",
    width="80%",
    height="30%",
    font_size=24
)

# Save the presentation
pres.save("output.pptx")
```

### TOML-Based Templates

Create, share, and reuse templates using human-readable TOML files:

```python
from easypptx import Presentation
from easypptx.template import TemplateManager

# Initialize template manager with template directory
tm = TemplateManager(template_dir="templates")

# Load a template from a TOML file
template_name = tm.load("templates/business_title.toml")

# Create a presentation
pres = Presentation()

# Create a slide using the loaded template
slide = pres.add_slide_from_template(template_name)

# Add content to the templated slide
slide.add_text(
    text="Quarterly Business Review",
    x="10%",
    y="30%",
    width="80%",
    height="20%",
    font_size=44,
    font_bold=True,
    align="center",
    color="white"
)

slide.add_text(
    text="Q2 2025 Financial Results",
    x="10%",
    y="55%",
    width="80%",
    height="10%",
    font_size=24,
    align="center",
    color="#66ccff"
)

# Save the presentation
pres.save("output.pptx")
```

Sample TOML template (business_title.toml):

```toml
# Business Template - Title Slide
bg_color = "#003366"  # Dark blue background

[title]
text = "Presentation Title"
position = { x = "10%", y = "30%", width = "80%", height = "20%" }
font = { name = "Meiryo", size = 44, bold = true }
align = "center"
vertical = "middle"
color = "white"

[subtitle]
text = "Subtitle or Presenter Information"
position = { x = "10%", y = "55%", width = "80%", height = "10%" }
font = { name = "Meiryo", size = 24, bold = false }
align = "center"
vertical = "middle"
color = "#66ccff"  # Light blue for subtitle
```

## Dark Theme Support

Create modern presentations with dark backgrounds and vibrant colors:

```python
# Create a presentation with black background
pres = Presentation(default_bg_color="black")

# Add a slide with default black background and a title
slide1 = pres.add_slide(
    title="Dark Theme Example",
    title_color="cyan"
)

# Add high-contrast text directly to the slide
slide1.add_text(
    text="High contrast text on dark background",
    x="10%",
    y="30%",
    width="80%",
    height="20%",
    font_size=24,
    color="white"
)

# Add a slide with a custom background color
slide2 = pres.add_slide(
    title="Custom Dark Background",
    title_color="white",
    bg_color=(0, 20, 40)  # Dark blue
)

# Add content with vibrant colors
slide2.add_text(
    text="Text with vibrant color",
    x="10%",
    y="30%",
    width="80%",
    height="20%",
    font_size=24,
    color="lime"
)

# Add a shape with custom color
slide2.add_shape(
    shape_type="ROUNDED_RECTANGLE",
    x="30%",
    y="60%",
    width="40%",
    height="15%",
    fill_color="purple"
)

# Set background color for an existing slide
slide1.set_background_color("darkgray")
```

## Getting started with development

### 1. Create a New Repository

First, create a repository on GitHub with the same name as this project, and then run the following commands:

```bash
git init -b main
git add .
git commit -m "init commit"
git remote add origin git@github.com:Ameyanagi/EasyPPTX.git
git push -u origin main
```

### 2. Set Up Your Development Environment

Then, install the environment and the pre-commit hooks with

```bash
make install
```

This will also generate your `uv.lock` file

### 3. Run tests

```bash
uv run pytest
```

## Project Structure

- `src/easypptx/` - Main package
  - `presentation.py` - Core presentation handling
  - `slide.py` - Slide creation and manipulation
  - `text.py` - Text elements and formatting
  - `image.py` - Image handling
  - `table.py` - Table creation from data
  - `chart.py` - Chart generation
  - `grid.py` - Grid layout system for complex arrangements
  - `pyplot.py` - Integration with matplotlib plots
  - `template.py` - Template management and utilities
- `examples/` - Example scripts demonstrating usage
  - `quick_start.py` - Basic usage example
  - `basic_demo.py` - Introduction to basic features
  - `comprehensive_example.py` - Full-featured business presentation
  - `aspect_ratio_example.py` - Demonstration of aspect ratio options
  - `grid/` - Grid layout examples
    - `001_basic_grid.py` - Basic Grid usage
    - `002_grid_indexing.py` - Grid indexing and iteration
    - `003_nested_grid.py` - Nested grids and merged cells
    - `004_autogrid.py` - Automatic grid layout
    - `005_enhanced_grid.py` - Enhanced Grid with convenience methods
  - `extended_features_example.py` - Showcase of percentage-based positioning, auto-alignment, and more

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

---

Repository initiated with [fpgmaas/cookiecutter-uv](https://github.com/fpgmaas/cookiecutter-uv).
