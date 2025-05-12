"""
007_grid_slide_examples.py - Grid Slide Examples

This example demonstrates the use of the add_grid_slide method to create slides with
various grid layouts and headers.
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

# --------------------------------------------------------------------------
# Example 1: Basic Grid Slide with Title
# --------------------------------------------------------------------------
slide1, grid1 = pres.add_grid_slide(
    rows=2,
    cols=2,
    title="Basic Grid Slide with Title",
    title_height="15%",
    padding=5.0,
)

# Add content to the cells
grid1[0, 0].add_text(
    text="Top Left",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid1[0, 1].add_text(
    text="Top Right",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid1[1, 0].add_text(
    text="Bottom Left",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid1[1, 1].add_text(
    text="Bottom Right",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

# --------------------------------------------------------------------------
# Example 2: Grid Slide with Title and Subtitle
# --------------------------------------------------------------------------
slide2, grid2 = pres.add_grid_slide(
    rows=3,
    cols=2,
    title="Grid Slide with Title and Subtitle",
    subtitle="Detailed Information About Grid Layouts",
    title_height="10%",
    subtitle_height="5%",
    padding=8.0,
)

# Add content to the grid using row-level access
grid2[0].add_text(
    text="Feature 1",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid2[0].add_text(
    text="Description of Feature 1. This demonstrates using row-level access to automatically fill cells left-to-right.",
    font_size=14,
    align="left",
    vertical="top",
)

grid2[1].add_text(
    text="Feature 2",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid2[1].add_text(
    text="Description of Feature 2. You don't need to manually specify which column to use.",
    font_size=14,
    align="left",
    vertical="top",
)

grid2[2].add_text(
    text="Feature 3",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid2[2].add_text(
    text="Description of Feature 3. The content is automatically placed in the next available cell in the row.",
    font_size=14,
    align="left",
    vertical="top",
)

# --------------------------------------------------------------------------
# Example 3: Custom Grid Positioning and Dimensions
# --------------------------------------------------------------------------
slide3, grid3 = pres.add_grid_slide(
    rows=2,
    cols=3,
    title="Custom Grid Positioning",
    x="10%",
    y="20%",
    width="80%",
    height="70%",
    padding=5.0,
)

# Create some sample data
x = np.linspace(0, 10, 100)
y = np.sin(x)

# Create a matplotlib figure
fig = plt.figure(figsize=(4, 3))
plt.plot(x, y)
plt.title("Sample Plot")
plt.grid(True)

# Add different types of content to the cells
grid3[0, 0].add_pyplot(figure=fig, dpi=150)

grid3[0, 1].add_text(
    text="Text Cell",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

data = [["Header 1", "Header 2"], [100, 200], [300, 400]]
grid3[0, 2].add_table(data=data, has_header=True)

# Add content to the second row using row-level access
grid3[1].add_text(
    text="Row 1, Col 0",
    font_size=14,
    align="center",
    vertical="middle",
)

grid3[1].add_text(
    text="Row 1, Col 1",
    font_size=14,
    align="center",
    vertical="middle",
)

grid3[1].add_text(
    text="Row 1, Col 2",
    font_size=14,
    align="center",
    vertical="middle",
)

# Save the presentation
output_path = output_dir / "007_grid_slide_examples.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
