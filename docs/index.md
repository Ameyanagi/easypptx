# EasyPPTX

[![Release](https://img.shields.io/github/v/release/Ameyanagi/EasyPPTX)](https://img.shields.io/github/v/release/Ameyanagi/EasyPPTX)
[![Build status](https://img.shields.io/github/actions/workflow/status/Ameyanagi/EasyPPTX/main.yml?branch=main)](https://github.com/Ameyanagi/EasyPPTX/actions/workflows/main.yml?query=branch%3Amain)
[![Commit activity](https://img.shields.io/github/commit-activity/m/Ameyanagi/EasyPPTX)](https://img.shields.io/github/commit-activity/m/Ameyanagi/EasyPPTX)
[![License](https://img.shields.io/github/license/Ameyanagi/EasyPPTX)](https://img.shields.io/github/license/Ameyanagi/EasyPPTX)

A Python library for easily creating and manipulating PowerPoint presentations programmatically with simple APIs, designed to be easy for both humans and AI assistants to use.

## Features

- Simple, intuitive API for PowerPoint manipulation
- Create slides with text, images, tables, charts, and bulleted lists
- Tables and charts accept pandas/polars DataFrames, pandas Series, numpy arrays, dicts, and lists — plus a `df.pptx` accessor for one-line DataFrame-to-slide
- Native PowerPoint charts stay editable; heatmap/histogram/box/violin render via matplotlib automatically
- Build whole decks from markdown with `Presentation.from_markdown`
- Reusable styles (`TextStyle`, `TableStyle`, `ChartStyle`) and built-in themes (`light`, `dark`, `corporate`)
- Default 16:9 aspect ratio with support for multiple ratio options
- Percentage-based positioning for responsive layouts (`x=10` means 10%; use `in_()` for inches)
- Grid layout system with slice spans, weighted tracks, and auto-flow
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

seaborn is not a dependency. seaborn plots still work: pass their figure to `slide.add_pyplot` or `pres.add_pyplot_slide`.

## Migrating to 0.7.0

Version 0.7.0 has two breaking changes:

1. **Bare numbers are percentages.** `x=10` now means 10% of the slide (same as `x="10%"`); floats no longer mean inches. Use `in_(1.5)` (from `easypptx import in_`) for absolute inches.
2. **Deprecated `Presentation` methods removed.** The `pres.add_text(slide, ...)`-style pass-throughs and `add_matplotlib_slide` / `add_seaborn_slide` / `add_plot` / `add_image_slide` are gone; use the `Slide` methods and `pres.add_pyplot_slide` / `pres.add_image_gen_slide`.

See the [migration guide](migration.md) for before/after examples.

## Quick Start

```python
from easypptx import Presentation
import pandas as pd

# Create a new presentation (uses 16:9 aspect ratio by default)
pres = Presentation()

# Add a slide with a title
slide = pres.add_slide(title="EasyPPTX Demo")

# Add text (positions are percentages of the slide; x=10 == "10%")
slide.add_text("This presentation was created with EasyPPTX",
               x=10, y=30, font_size=24)

# Add an image
slide.add_image("path/to/image.png", x=10, y=40, width=40)

# Create a table
data = [["Name", "Value"], ["Item 1", 100], ["Item 2", 200]]
slide.add_table(data, x=60, y=30, width=30, height=15)

# Add a chart from a pandas DataFrame
df = pd.DataFrame({"Category": ["A", "B", "C"], "Value": [10, 20, 30]})
slide.add_chart(data=df, chart_type="pie",
                category_column="Category",
                value_columns="Value",
                x=60, y=50, title="Sample Chart")

# Save the presentation
pres.save("example.pptx")
```

*New in 0.8.0*, DataFrames can send themselves to a slide with the `df.pptx` accessor, and tables/charts accept data in whatever shape you have it (Series, polars, numpy, dicts, lists):

```python
slide2 = pres.add_slide(title="Sales")
df = pd.DataFrame({"Region": ["East", "West"], "Sales": [1200, 950]})

df.pptx.table(slide2, y=20, height=30, number_format="{:,.0f}")
df.pptx.chart(slide2, kind="column", y=55, height=40, show_values=True)

# dicts and numpy arrays work directly too
slide2.add_chart(data={"Rev": [100, 120], "Cost": [80, 85]}, chart_type="line")
```

See [Data to Slides](data.md) for the full story.

Or write the deck in markdown and convert it:

```python
from easypptx import Presentation

pres = Presentation.from_markdown("deck.md", theme="dark")
pres.save("deck.pptx")
```

## Documentation

- [Features Overview](features.md)
- [User Guide](percentage_positioning.md)
- [Markdown to Presentation](markdown.md)
- [Data to Slides](data.md)
- [Migrating to 0.7.0](migration.md)
- [API Reference](api_reference.md)
- [Examples](https://github.com/Ameyanagi/EasyPPTX/tree/main/examples)
