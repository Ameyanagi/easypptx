"""
008_quick_slide_deck.py - Quick Slide Deck Creation with Enhanced Row Access API

This example demonstrates using the enhanced grid[row].add_xxx() row-based API
to quickly create a simple slide deck, automatically placing content in rows
without having to specify column indices.
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

from easypptx import Presentation

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# -------------------------------------------------------------------------
# Slide 1: Title
# -------------------------------------------------------------------------
slide1 = pres.add_slide()

# Create a simple 3-row grid for the title slide
title_grid = pres.add_autogrid(
    slide=slide1,
    content_funcs=None,  # Empty grid
    rows=3,
    cols=1,
    x="10%",
    y="15%",
    width="80%",
    height="70%",
    padding=5.0,
)

# Add content sequentially to rows without specifying column indices
title_grid[0].add_text(
    text="Quick Slide Deck Creation",
    font_size=44,
    font_bold=True,
    align="center",
    vertical="middle",
)

title_grid[1].add_text(
    text="Using Enhanced Row-Based API",
    font_size=32,
    align="center",
    vertical="top",
)

title_grid[2].add_text(
    text="Created with EasyPPTX v0.0.3",
    font_size=20,
    font_italic=True,
    align="center",
    vertical="middle",
)

# -------------------------------------------------------------------------
# Slide 2: Feature Overview with Row-Based Content Addition
# -------------------------------------------------------------------------
slide2 = pres.add_slide()

slide2.add_text(
    text="Key Features",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=36,
    font_bold=True,
    align="center",
)

# Create a grid with 4 rows and 2 columns
features_grid = pres.add_autogrid(
    slide=slide2,
    content_funcs=None,
    rows=4,
    cols=2,
    x="5%",
    y="17%",
    width="90%",
    height="78%",
    padding=5.0,
)

# Add feature items to each row without manually tracking columns
# For row 0, this will add to (0,0) then (0,1)
features_grid[0].add_text(
    text="Simplified Syntax",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

features_grid[0].add_text(
    text="Intuitive and concise API",
    font_size=18,
    align="center",
    vertical="middle",
)

# For row 1, this will add to (1,0) then (1,1)
features_grid[1].add_text(
    text="Automatic Placement",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

features_grid[1].add_text(
    text="Content flows naturally across rows",
    font_size=18,
    align="center",
    vertical="middle",
)

# For row 2, this will add to (2,0) then (2,1)
features_grid[2].add_text(
    text="Reduced Code",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

features_grid[2].add_text(
    text="Less typing, fewer parameters",
    font_size=18,
    align="center",
    vertical="middle",
)

# For row 3, this will add to (3,0) then (3,1)
features_grid[3].add_text(
    text="All Content Types",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

features_grid[3].add_text(
    text="Works with text, images, tables, charts",
    font_size=18,
    align="center",
    vertical="middle",
)

# -------------------------------------------------------------------------
# Slide 3: Chart Example with Row-Based API
# -------------------------------------------------------------------------
slide3 = pres.add_slide()

slide3.add_text(
    text="Data Visualization Example",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=36,
    font_bold=True,
    align="center",
)

# Create a grid with 2 rows, 2 columns
chart_grid = pres.add_autogrid(
    slide=slide3,
    content_funcs=None,
    rows=2,
    cols=2,
    x="5%",
    y="17%",
    width="90%",
    height="78%",
    padding=8.0,
)

# Create some sample visualization data
x = np.linspace(0, 10, 100)
y1 = np.sin(x)
y2 = np.cos(x)

# Create charts for the presentation
fig1 = plt.figure(figsize=(5, 3))
plt.plot(x, y1, "b-", linewidth=2)
plt.title("Sine Wave")
plt.grid(True)
plt.tight_layout()

fig2 = plt.figure(figsize=(5, 3))
plt.plot(x, y2, "r-", linewidth=2)
plt.title("Cosine Wave")
plt.grid(True)
plt.tight_layout()

# Add the charts to the grid using row-based API
chart_grid[0].add_text(
    text="Sine Wave Example",
    font_size=20,
    font_bold=True,
    align="center",
)

chart_grid[0].add_pyplot(
    figure=fig1,
    dpi=150,
)

chart_grid[1].add_text(
    text="Cosine Wave Example",
    font_size=20,
    font_bold=True,
    align="center",
)

chart_grid[1].add_pyplot(
    figure=fig2,
    dpi=150,
)

# -------------------------------------------------------------------------
# Slide 4: Table Data Example
# -------------------------------------------------------------------------
slide4 = pres.add_slide()

slide4.add_text(
    text="Table Data Example",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=36,
    font_bold=True,
    align="center",
)

# Create a grid for the table slide
table_grid = pres.add_autogrid(
    slide=slide4,
    content_funcs=None,
    rows=2,
    cols=1,
    x="10%",
    y="17%",
    width="80%",
    height="78%",
    padding=5.0,
)

# Add an explanation text
table_grid[0].add_text(
    text="Example of using grid[row].add_table() without specifying column indices:",
    font_size=20,
    align="center",
    vertical="middle",
)

# Create a simple data table
data = [
    ["Feature", "Description", "Benefits"],
    ["Row Access", "Add items sequentially to rows", "Easy, streamlined workflow"],
    ["Cell Access", "Direct cell targeting", "Fine-grained control"],
    ["Autogrid", "Automatic grid creation", "Quick layout development"],
]

# Add the table to the next row
table_grid[1].add_table(
    data=data,
    has_header=True,
)

# Save the presentation
output_path = output_dir / "008_quick_slide_deck.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
