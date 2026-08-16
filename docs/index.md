# EasyPPTX

[![Release](https://img.shields.io/github/v/release/Ameyanagi/EasyPPTX)](https://img.shields.io/github/v/release/Ameyanagi/EasyPPTX)
[![Build status](https://img.shields.io/github/actions/workflow/status/Ameyanagi/EasyPPTX/main.yml?branch=main)](https://github.com/Ameyanagi/EasyPPTX/actions/workflows/main.yml?query=branch%3Amain)
[![Commit activity](https://img.shields.io/github/commit-activity/m/Ameyanagi/EasyPPTX)](https://img.shields.io/github/commit-activity/m/Ameyanagi/EasyPPTX)
[![License](https://img.shields.io/github/license/Ameyanagi/EasyPPTX)](https://img.shields.io/github/license/Ameyanagi/EasyPPTX)

A Python library for easily creating and manipulating PowerPoint presentations programmatically with simple APIs, designed to be easy for both humans and AI assistants to use.

## Features

- Simple, intuitive API for PowerPoint manipulation
- Create slides with text, images, tables, and charts
- Format elements with easy-to-use styling options
- Default 16:9 aspect ratio with support for multiple ratio options
- Percentage-based positioning for responsive layouts
- Auto-alignment of multiple objects (grid, horizontal, vertical)
- Default color scheme and Meiryo font
- Support for reference PowerPoint templates
- Optimized for use with AI assistants and LLMs
- Built on top of python-pptx with a more user-friendly interface

## Installation

```bash
pip install easypptx
```

pandas and matplotlib support are optional extras:

```bash
pip install "easypptx[dataframe]"  # pandas DataFrame support (tables, charts)
pip install "easypptx[plot]"       # matplotlib figure embedding
pip install "easypptx[all]"        # both
```

seaborn is not a dependency. seaborn plots still work: pass their figure to `add_pyplot_slide` or `Pyplot.add`.

## Quick Start

```python
from easypptx import Presentation
import pandas as pd

# Create a new presentation (uses 16:9 aspect ratio by default)
pres = Presentation()

# Add a slide with a title
slide = pres.add_slide(title="EasyPPTX Demo")

# Add text
slide.add_text("This presentation was created with EasyPPTX",
               x="10%", y="30%", font_size=24)

# Add an image
slide.add_image("path/to/image.png", x="10%", y="40%", width="40%")

# Create a table
data = [["Name", "Value"], ["Item 1", 100], ["Item 2", 200]]
slide.add_table(data, x="60%", y="30%", width="30%", height="15%")

# Add a chart from a pandas DataFrame
df = pd.DataFrame({"Category": ["A", "B", "C"], "Value": [10, 20, 30]})
slide.add_chart(data=df, chart_type="pie",
                category_column="Category",
                value_columns="Value",
                x="60%", y="50%", title="Sample Chart")

# Save the presentation
pres.save("example.pptx")
```

## Documentation

- [Features Overview](features.md)
- [User Guide](percentage_positioning.md)
- [API Reference](api_reference.md)
- [Examples](https://github.com/Ameyanagi/EasyPPTX/tree/main/examples)
