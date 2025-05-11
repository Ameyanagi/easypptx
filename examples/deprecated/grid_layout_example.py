"""
Example demonstrating the Grid layout feature in EasyPPTX.

This example shows how to create responsive grid layouts for slides,
including nested grids and merged cells.
"""

from pathlib import Path

from easypptx import Presentation
from easypptx.grid import Grid

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# ----------------------------------------------------------
# Slide 1: Basic Grid Layout
# ----------------------------------------------------------
slide1 = pres.add_slide()

# Add a title
slide1.add_text(
    text="Basic Grid Layout (2x2)",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",
)

# Create a 2x2 grid layout that takes up most of the slide
grid = Grid(
    parent=slide1,
    x="5%",
    y="15%",
    width="90%",
    height="80%",
    rows=2,
    cols=2,
    padding=5.0,
)

# Add content to each cell
grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide1.add_text,
    text="Top Left",
    font_size=24,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=0,
    col=1,
    content_func=slide1.add_text,
    text="Top Right",
    font_size=24,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=1,
    col=0,
    content_func=slide1.add_text,
    text="Bottom Left",
    font_size=24,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=1,
    col=1,
    content_func=slide1.add_text,
    text="Bottom Right",
    font_size=24,
    align="center",
    vertical="middle",
)

# ----------------------------------------------------------
# Slide 2: Grid with Shapes and Colors
# ----------------------------------------------------------
slide2 = pres.add_slide()

# Add a title
slide2.add_text(
    text="Grid with Shapes and Colors (3x3)",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",
)

# Create a 3x3 grid layout
grid = Grid(
    parent=slide2,
    x="5%",
    y="15%",
    width="90%",
    height="80%",
    rows=3,
    cols=3,
    padding=5.0,
)

# Add shapes with different colors to each cell
colors = ["red", "green", "blue", "yellow", "orange", "purple", "cyan", "magenta", "gray"]

for row in range(3):
    for col in range(3):
        index = row * 3 + col
        grid.add_to_cell(
            row=row,
            col=col,
            content_func=slide2.add_shape,
            shape_type=1,  # Rectangle
            fill_color=colors[index],
        )

        # Add cell labels
        grid.add_to_cell(
            row=row,
            col=col,
            content_func=slide2.add_text,
            text=f"Cell {row},{col}",
            font_size=14,
            font_bold=True,
            align="center",
            vertical="middle",
            color="white" if colors[index] in ["red", "blue", "purple", "gray"] else "black",
        )

# ----------------------------------------------------------
# Slide 3: Merged Cells
# ----------------------------------------------------------
slide3 = pres.add_slide()

# Add a title
slide3.add_text(
    text="Grid with Merged Cells (4x4)",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",
)

# Create a 4x4 grid layout
grid = Grid(
    parent=slide3,
    x="5%",
    y="15%",
    width="90%",
    height="80%",
    rows=4,
    cols=4,
    padding=3.0,
)

# Merge some cells
grid.merge_cells(0, 0, 1, 1)  # 2x2 top-left
grid.merge_cells(0, 2, 0, 3)  # 1x2 top-right
grid.merge_cells(2, 0, 3, 0)  # 2x1 bottom-left
grid.merge_cells(1, 2, 3, 3)  # 3x2 bottom-right

# Add content to the merged cells
grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide3.add_shape,
    shape_type=1,  # Rectangle
    fill_color="blue",
)

grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide3.add_text,
    text="2x2 Merged",
    font_size=20,
    align="center",
    vertical="middle",
    color="white",
)

grid.add_to_cell(
    row=0,
    col=2,
    content_func=slide3.add_shape,
    shape_type=1,  # Rectangle
    fill_color="green",
)

grid.add_to_cell(
    row=0,
    col=2,
    content_func=slide3.add_text,
    text="1x2 Merged",
    font_size=18,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=2,
    col=0,
    content_func=slide3.add_shape,
    shape_type=1,  # Rectangle
    fill_color="orange",
)

grid.add_to_cell(
    row=2,
    col=0,
    content_func=slide3.add_text,
    text="2x1 Merged",
    font_size=18,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=1,
    col=2,
    content_func=slide3.add_shape,
    shape_type=1,  # Rectangle
    fill_color="purple",
)

grid.add_to_cell(
    row=1,
    col=2,
    content_func=slide3.add_text,
    text="3x2 Merged",
    font_size=24,
    align="center",
    vertical="middle",
    color="white",
)

# ----------------------------------------------------------
# Slide 4: Nested Grids
# ----------------------------------------------------------
slide4 = pres.add_slide()

# Add a title
slide4.add_text(
    text="Nested Grid Layouts",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",
)

# Create a 2x2 main grid
main_grid = Grid(
    parent=slide4,
    x="5%",
    y="15%",
    width="90%",
    height="80%",
    rows=2,
    cols=2,
    padding=5.0,
)

# Add a label for the main grid
main_grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide4.add_text,
    text="Main Grid (2x2)",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="top",
)

# Add a nested 2x2 grid to the top-right cell
nested_grid1 = main_grid.add_grid_to_cell(
    row=0,
    col=1,
    rows=2,
    cols=2,
    padding=10.0,
)

# Add content to the nested grid
for row in range(2):
    for col in range(2):
        nested_grid1.add_to_cell(
            row=row,
            col=col,
            content_func=slide4.add_shape,
            shape_type=1,  # Rectangle
            fill_color="blue",
        )

        nested_grid1.add_to_cell(
            row=row,
            col=col,
            content_func=slide4.add_text,
            text=f"Nested 1\nCell {row},{col}",
            font_size=12,
            align="center",
            vertical="middle",
            color="white",
        )

# Add a nested 3x3 grid to the bottom-left cell
nested_grid2 = main_grid.add_grid_to_cell(
    row=1,
    col=0,
    rows=3,
    cols=3,
    padding=5.0,
)

# Add content to the nested grid
for row in range(3):
    for col in range(3):
        nested_grid2.add_to_cell(
            row=row,
            col=col,
            content_func=slide4.add_shape,
            shape_type=1,  # Rectangle
            fill_color="green",
        )

        nested_grid2.add_to_cell(
            row=row,
            col=col,
            content_func=slide4.add_text,
            text=f"N2\n{row},{col}",
            font_size=10,
            align="center",
            vertical="middle",
        )

# Create a double-nested grid in the bottom-right cell
nested_grid3 = main_grid.add_grid_to_cell(
    row=1,
    col=1,
    rows=2,
    cols=2,
    padding=5.0,
)

# Add a double-nested grid to the top-left cell of nested_grid3
double_nested = nested_grid3.add_grid_to_cell(
    row=0,
    col=0,
    rows=2,
    cols=2,
    padding=5.0,
)

# Add content to the double-nested grid
for row in range(2):
    for col in range(2):
        double_nested.add_to_cell(
            row=row,
            col=col,
            content_func=slide4.add_shape,
            shape_type=1,  # Rectangle
            fill_color="red",
        )

        double_nested.add_to_cell(
            row=row,
            col=col,
            content_func=slide4.add_text,
            text=f"DN\n{row},{col}",
            font_size=8,
            align="center",
            vertical="middle",
            color="white",
        )

# Add content to other cells of nested_grid3
labels = ["Double\nNested", "Cell\n0,1", "Cell\n1,0", "Cell\n1,1"]
colors = ["purple", "orange", "cyan", "magenta"]
positions = [(0, 0), (0, 1), (1, 0), (1, 1)]

for _, (position, label, color) in enumerate(zip(positions[1:], labels[1:], colors[1:], strict=False)):
    row, col = position
    nested_grid3.add_to_cell(
        row=row,
        col=col,
        content_func=slide4.add_shape,
        shape_type=1,  # Rectangle
        fill_color=color,
    )

    nested_grid3.add_to_cell(
        row=row,
        col=col,
        content_func=slide4.add_text,
        text=label,
        font_size=12,
        align="center",
        vertical="middle",
        color="white" if color in ["purple", "magenta"] else "black",
    )

# ----------------------------------------------------------
# Slide 5: Dashboard Layout
# ----------------------------------------------------------
slide5 = pres.add_slide()

# Add a title
slide5.add_text(
    text="Dashboard Layout Example",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",
)

# Create a flexible dashboard layout using a grid
dashboard = Grid(
    parent=slide5,
    x="5%",
    y="15%",
    width="90%",
    height="80%",
    rows=3,
    cols=4,
    padding=2.0,
)

# Create header area (spans the entire width)
dashboard.merge_cells(0, 0, 0, 3)
dashboard.add_to_cell(
    row=0,
    col=0,
    content_func=slide5.add_shape,
    shape_type=1,  # Rectangle
    fill_color="blue",
)
dashboard.add_to_cell(
    row=0,
    col=0,
    content_func=slide5.add_text,
    text="Sales Dashboard - FY 2023",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
    color="white",
)

# Create sidebar (spans two rows)
dashboard.merge_cells(1, 0, 2, 0)
dashboard.add_to_cell(
    row=1,
    col=0,
    content_func=slide5.add_shape,
    shape_type=1,  # Rectangle
    fill_color="gray",
)
dashboard.add_to_cell(
    row=1,
    col=0,
    content_func=slide5.add_text,
    text="Navigation\n\n• Overview\n• Products\n• Regions\n• Customers\n• Forecast",
    font_size=14,
    align="left",
    vertical="top",
    color="white",
)

# Create main KPI area (spans 2 columns)
dashboard.merge_cells(1, 1, 1, 2)
dashboard.add_to_cell(
    row=1,
    col=1,
    content_func=slide5.add_shape,
    shape_type=1,  # Rectangle
    fill_color="green",
)
dashboard.add_to_cell(
    row=1,
    col=1,
    content_func=slide5.add_text,
    text="Revenue: $4.2M\nUp 15% from last year",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

# Create supplementary KPI
dashboard.add_to_cell(
    row=1,
    col=3,
    content_func=slide5.add_shape,
    shape_type=1,  # Rectangle
    fill_color="orange",
)
dashboard.add_to_cell(
    row=1,
    col=3,
    content_func=slide5.add_text,
    text="Customer Growth\n+22%",
    font_size=16,
    align="center",
    vertical="middle",
)

# Create bottom data areas
regions = ["East", "West", "Central"]
perf = ["$1.2M", "$1.5M", "$1.5M"]
colors = ["purple", "magenta", "cyan"]
performance = perf  # Rename to avoid variable shadowing

for i, (region, _, color) in enumerate(zip(regions, performance, colors, strict=False)):
    col = i + 1
    dashboard.add_to_cell(
        row=2,
        col=col,
        content_func=slide5.add_shape,
        shape_type=1,  # Rectangle
        fill_color=color,
    )

    dashboard.add_to_cell(
        row=2,
        col=col,
        content_func=slide5.add_text,
        text=f"Region: {region}\nRevenue: {perf}",
        font_size=14,
        align="center",
        vertical="middle",
        color="white" if color in ["purple", "magenta"] else "black",
    )

# Save the presentation
output_path = output_dir / "grid_layout_example.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
print("Open the presentation to see the grid layout examples.")
