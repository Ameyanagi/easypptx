"""
004_autogrid.py - Automatic Grid Layout Example

This example demonstrates the auto-grid features:
1. Using autogrid to automatically arrange content
2. Using autogrid_pyplot to automatically arrange matplotlib plots
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

from easypptx import Presentation
from easypptx.grid import Grid

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# -----------------------------------------------------
# Example 1: Basic AutoGrid with Content Functions
# -----------------------------------------------------
slide1 = pres.add_slide()

# Add a title
slide1.add_text(
    text="004 - AutoGrid Example",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
)


# Create content functions
def create_text1():
    return slide1.add_text(
        text="Content 1",
        font_size=24,
        align="center",
        vertical="middle",
    )


def create_text2():
    return slide1.add_text(
        text="Content 2",
        font_size=24,
        align="center",
        vertical="middle",
        font_bold=True,
    )


def create_shape():
    return slide1.add_shape(
        shape_type=1,  # Rectangle
        fill_color="blue",
    )


def create_text_on_shape():
    shape = slide1.add_shape(
        shape_type=1,  # Rectangle
        fill_color="green",
    )
    slide1.add_text(
        text="Text on Shape",
        font_size=18,
        align="center",
        vertical="middle",
    )
    return shape


# Use add_autogrid method
content_funcs = [create_text1, create_text2, create_shape, create_text_on_shape]
pres.add_autogrid(
    slide=slide1,
    content_funcs=content_funcs,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    padding=5.0,
)

# -----------------------------------------------------
# Example 2: AutoGrid with Matplotlib Plots
# -----------------------------------------------------
slide2 = pres.add_slide()
slide2.add_text(
    text="004 - AutoGrid with Plots",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Create matplotlib figures
fig1 = plt.figure(figsize=(4, 3))
categories = ["A", "B", "C", "D", "E"]
values = [23, 45, 56, 78, 32]
plt.bar(categories, values)
plt.title("Bar Chart")

fig2 = plt.figure(figsize=(4, 3))
x = np.random.rand(50)
y = np.random.rand(50)
plt.scatter(x, y)
plt.title("Scatter Plot")
plt.grid(True)

fig3 = plt.figure(figsize=(4, 3))
data = np.random.randn(1000)
plt.hist(data, bins=30)
plt.title("Histogram")

fig4 = plt.figure(figsize=(4, 3))
labels = ["A", "B", "C", "D"]
sizes = [15, 30, 45, 10]
plt.pie(sizes, labels=labels, autopct="%1.1f%%")
plt.title("Pie Chart")

# Use autogrid_pyplot to create a grid with the plots
figures = [fig1, fig2, fig3, fig4]
grid = Grid.autogrid_pyplot(
    parent=slide2,
    figures=figures,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    rows=2,
    cols=2,
    padding=5.0,
    title="Matplotlib Plots in Grid",
    title_height="5%",
    dpi=300,
)

# Save the presentation
output_path = output_dir / "004_autogrid.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
