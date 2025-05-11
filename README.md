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
from easypptx import Presentation, Slide, Text, Image, Table, Chart
import pandas as pd

# Create a new presentation (uses 16:9 aspect ratio by default)
pres = Presentation()

# Add a slide
slide = pres.add_slide()

# Add title
text = Text(slide)
text.add_title("EasyPPTX Demo")

# Add text
text.add_paragraph("This presentation was created with EasyPPTX",
                  x=1, y=2, font_size=24)

# Add an image
img = Image(slide)
img.add("path/to/image.png", x=1, y=3, width=4)

# Create a table
tbl = Table(slide)
data = [["Name", "Value"], ["Item 1", 100], ["Item 2", 200]]
tbl.add(data, x=6, y=2)

# Add a chart from pandas DataFrame
df = pd.DataFrame({"Category": ["A", "B", "C"], "Value": [10, 20, 30]})
chart = Chart(slide)
chart.from_dataframe(df, chart_type="pie",
                    category_column="Category",
                    value_column="Value",
                    x=6, y=4, title="Sample Chart")

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

## Reference PowerPoint Templates

Use existing PowerPoint files as templates:

```python
# Create a presentation using an existing template
pres = Presentation(template_path="template.pptx")

# Add a slide
slide = pres.add_slide()

# Add content and save
text = Text(slide)
text.add_title("Presentation with Template")
pres.save("output.pptx")
```

## Dark Theme Support

Create modern presentations with dark backgrounds and vibrant colors:

```python
# Create a presentation with black background
pres = Presentation(default_bg_color="black")

# Add a slide with default black background
slide = pres.add_slide()

# Add a slide with a custom background color
slide = pres.add_slide(bg_color=(0, 20, 40))  # Dark blue

# Add high-contrast text
text = Text(slide)
text.add_title("Dark Theme", color="cyan", align="center")
text.add_paragraph("High contrast text", color="white")

# Set background color for an existing slide
slide.set_background_color("darkgray")
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
- `examples/` - Example scripts demonstrating usage
  - `quick_start.py` - Basic usage example
  - `basic_demo.py` - Introduction to basic features
  - `comprehensive_example.py` - Full-featured business presentation
  - `aspect_ratio_example.py` - Demonstration of aspect ratio options
  - `extended_features_example.py` - Showcase of percentage-based positioning, auto-alignment, and more

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

---

Repository initiated with [fpgmaas/cookiecutter-uv](https://github.com/fpgmaas/cookiecutter-uv).
