"""
Example demonstrating the new grid methods in EasyPPTX.

This example shows how to use the new grid methods added to the Presentation class:
- add_grid
- add_grid_slide
- add_autogrid
- add_autogrid_slide
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

# -----------------------------------------------------
# Example 1: add_grid method
# -----------------------------------------------------
slide1 = pres.add_slide()

# Add a title
slide1.add_text(
    text="add_grid Method Example",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Use add_grid method to create a grid on the slide
grid = pres.add_grid(
    slide=slide1,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    rows=2,
    cols=2,
    padding=5.0,
)

# Add content to each cell
grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide1.add_text,
    text="Top Left Cell",
    font_size=24,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=0,
    col=1,
    content_func=slide1.add_text,
    text="Top Right Cell",
    font_size=24,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=1,
    col=0,
    content_func=slide1.add_text,
    text="Bottom Left Cell",
    font_size=24,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=1,
    col=1,
    content_func=slide1.add_text,
    text="Bottom Right Cell",
    font_size=24,
    align="center",
    vertical="middle",
)

# -----------------------------------------------------
# Example 2: add_grid_slide method
# -----------------------------------------------------

# Use add_grid_slide method to create a slide with a grid
slide2, grid = pres.add_grid_slide(
    rows=3,
    cols=3,
    title="add_grid_slide Method Example",
    title_height="15%",
    padding=3.0,
)

# Add content to the grid cells
colors = ["red", "green", "blue", "yellow", "orange", "purple", "cyan", "magenta", "gray"]

for row in range(3):
    for col in range(3):
        index = row * 3 + col

        # Add a shape with a different color
        grid.add_to_cell(
            row=row,
            col=col,
            content_func=slide2.add_shape,
            shape_type=1,  # Rectangle
            fill_color=colors[index],
        )

        # Add text on top of the shape
        grid.add_to_cell(
            row=row,
            col=col,
            content_func=slide2.add_text,
            text=f"Cell {row},{col}",
            font_size=18,
            font_bold=True,
            align="center",
            vertical="middle",
            color="white" if colors[index] in ["red", "blue", "purple", "gray"] else "black",
        )

# -----------------------------------------------------
# Example 3: add_autogrid method
# -----------------------------------------------------
slide3 = pres.add_slide()

# Add a title
slide3.add_text(
    text="add_autogrid Method Example",
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
    return slide3.add_text(
        text="Content 1",
        font_size=24,
        align="center",
        vertical="middle",
    )


def create_text2():
    return slide3.add_text(
        text="Content 2",
        font_size=24,
        align="center",
        vertical="middle",
        font_bold=True,
    )


def create_shape():
    return slide3.add_shape(
        shape_type=1,  # Rectangle
        fill_color="blue",
    )


def create_text_on_shape():
    shape = slide3.add_shape(
        shape_type=1,  # Rectangle
        fill_color="green",
    )
    slide3.add_text(
        text="Text on Shape",
        font_size=18,
        align="center",
        vertical="middle",
    )
    return shape


# Use add_autogrid method
content_funcs = [create_text1, create_text2, create_shape, create_text_on_shape]
pres.add_autogrid(
    slide=slide3,
    content_funcs=content_funcs,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    padding=5.0,
)

# -----------------------------------------------------
# Example 4: add_autogrid_slide method
# -----------------------------------------------------

# Create matplotlib figures
figures = []

# Figure 1: Sine Function
fig1 = plt.figure(figsize=(5, 4))
x = np.linspace(0, 10, 100)
y = np.sin(x)
plt.plot(x, y)
plt.title("Sine Function")
plt.grid(True)
figures.append(fig1)

# Figure 2: Cosine Function
fig2 = plt.figure(figsize=(5, 4))
x = np.linspace(0, 10, 100)
y = np.cos(x)
plt.plot(x, y, "r-")
plt.title("Cosine Function")
plt.grid(True)
figures.append(fig2)

# Figure 3: Exponential Function
fig3 = plt.figure(figsize=(5, 4))
x = np.linspace(0, 5, 100)
y = np.exp(x)
plt.plot(x, y, "g-")
plt.title("Exponential Function")
plt.grid(True)
figures.append(fig3)


# Create functions to add these figures to a slide
def add_functions_to_autogrid():
    slide, _ = pres.add_autogrid_slide(
        content_funcs=[],  # We'll create our own functions below
        title="add_autogrid_slide Method Example",
        rows=2,
        cols=2,
    )

    # Create functions to add the figures
    def create_fig1_func():
        return pres.add_pyplot(slide, figures[0])

    def create_fig2_func():
        return pres.add_pyplot(slide, figures[1])

    def create_fig3_func():
        return pres.add_pyplot(slide, figures[2])

    def create_text_func():
        return slide.add_text(
            text="Mathematical Functions\nAutomatically arranged",
            font_size=18,
            align="center",
            vertical="middle",
        )

    # Create the content functions
    content_funcs = [create_fig1_func, create_fig2_func, create_fig3_func, create_text_func]

    # Use add_autogrid directly on the slide (add_autogrid_slide already created a slide)
    grid = pres.add_autogrid(
        slide=slide,
        content_funcs=content_funcs,
        x="5%",
        y="20%",
        width="90%",
        height="75%",
        padding=5.0,
    )

    return slide, grid


# Call the function to create the slide
add_functions_to_autogrid()

# -----------------------------------------------------
# Example 5: add_autogrid_slide with pyplot content
# -----------------------------------------------------

# Create matplotlib figures for grid layout
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

# Create a slide for the pyplot examples
slide5 = pres.add_slide()
slide5.add_text(
    text="add_autogrid with Plot Functions",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
)


# Create functions that will add the plots to the slide
def add_bar_chart():
    return pres.add_pyplot(slide5, fig1)


def add_scatter_plot():
    return pres.add_pyplot(slide5, fig2)


def add_histogram():
    return pres.add_pyplot(slide5, fig3)


def add_pie_chart():
    return pres.add_pyplot(slide5, fig4)


# Use add_autogrid to arrange the plots
plot_funcs = [add_bar_chart, add_scatter_plot, add_histogram, add_pie_chart]
grid = pres.add_autogrid(
    slide=slide5,
    content_funcs=plot_funcs,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    rows=2,
    cols=2,
    padding=5.0,
)

# Save the presentation
output_path = output_dir / "grid_methods_example.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
print("Open the presentation to see the grid method examples.")
