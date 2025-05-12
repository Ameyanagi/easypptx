"""
Comprehensive example demonstrating all Grid capabilities in EasyPPTX.

This example shows the complete set of Grid features:
1. Basic grid creation
2. Cell access using [row, col] syntax
3. Cell access using flat index (grid[n])
4. Iteration through all cells (for cell in grid)
5. Flat iteration (for cell in grid.flat)
6. Cell content assignment
7. Nested grids
8. Cell merging
"""

from pathlib import Path

from easypptx import Presentation

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# ---------------------------------------------------------
# Slide 1: Grid Indexing and Iteration
# ---------------------------------------------------------
slide1 = pres.add_slide()

# Add a title
slide1.add_text(
    text="Grid Indexing & Iteration",
    x="0%",
    y="2%",
    width="100%",
    height="8%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Create a 3x3 grid
grid = pres.add_grid(
    slide=slide1,
    x="5%",
    y="15%",
    width="90%",
    height="40%",
    rows=3,
    cols=3,
    padding=5.0,
)

# Method 1: Using the grid[row, col] syntax
# First add a shape for background
shape = slide1.add_shape(
    x=grid[0, 0].x,
    y=grid[0, 0].y,
    width=grid[0, 0].width,
    height=grid[0, 0].height,
    shape_type=1,  # Rectangle
    fill_color="lightblue",
)
# Then add text on top
text = slide1.add_text(
    text="grid[0, 0]\nTuple indexing",
    x=grid[0, 0].x,
    y=grid[0, 0].y,
    width=grid[0, 0].width,
    height=grid[0, 0].height,
    font_size=14,
    align="center",
    vertical="middle",
)
grid[0, 0].content = text

# Method 2: Using the flat index syntax
# First add a shape for background
shape = slide1.add_shape(
    x=grid[4].x,
    y=grid[4].y,
    width=grid[4].width,
    height=grid[4].height,
    shape_type=1,  # Rectangle
    fill_color="lightgreen",
)
# Then add text on top
text = slide1.add_text(  # Center cell (row 1, col 1) is index 4 in flattened grid
    text="grid[4]\nFlat indexing\n(Center)",
    x=grid[4].x,
    y=grid[4].y,
    width=grid[4].width,
    height=grid[4].height,
    font_size=14,
    align="center",
    vertical="middle",
)
grid[4].content = text

# Method 3: Using iteration
for i, cell in enumerate(grid):
    if i in [2, 6]:  # Only add content to specific cells
        # Add background shape
        slide1.add_shape(
            x=cell.x,
            y=cell.y,
            width=cell.width,
            height=cell.height,
            shape_type=1,  # Rectangle
            fill_color="lightyellow",
        )
        # Add text on top
        cell.content = slide1.add_text(
            text=f"grid iteration\nindex {i}\n[{cell.row}, {cell.col}]",
            x=cell.x,
            y=cell.y,
            width=cell.width,
            height=cell.height,
            font_size=14,
            align="center",
            vertical="middle",
        )

# Method 4: Using flat iterator to identify specific cells
for cell in grid.flat:
    # Add content to specific cells based on their row/col
    if (cell.row == 0 and cell.col == 1) or (cell.row == 2 and cell.col == 2):
        # Add background shape
        slide1.add_shape(
            x=cell.x,
            y=cell.y,
            width=cell.width,
            height=cell.height,
            shape_type=1,  # Rectangle
            fill_color="lightpink",
        )
        # Add text on top
        cell.content = slide1.add_text(
            text=f"grid.flat\n[{cell.row}, {cell.col}]",
            x=cell.x,
            y=cell.y,
            width=cell.width,
            height=cell.height,
            font_size=14,
            align="center",
            vertical="middle",
        )

# Method 5: Using add_to_cell with shape for background
grid.add_to_cell(
    row=1,
    col=2,
    content_func=slide1.add_shape,
    shape_type=1,  # Rectangle
    fill_color="lavender",
)

# Add text on top
grid.add_to_cell(
    row=1, col=2, content_func=slide1.add_text, text="add_to_cell(1,2)", font_size=14, align="center", vertical="middle"
)

# ---------------------------------------------------------
# Slide 2: Nested Grids and Cell Merging
# ---------------------------------------------------------
slide2 = pres.add_slide()

# Add a title
slide2.add_text(
    text="Nested Grids & Cell Merging",
    x="0%",
    y="2%",
    width="100%",
    height="8%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Create a 2x2 grid
main_grid = pres.add_grid(
    slide=slide2,
    x="5%",
    y="15%",
    width="90%",
    height="80%",
    rows=2,
    cols=2,
    padding=5.0,
)

# Add content to top left cell
# First add a shape for background
slide2.add_shape(
    x=main_grid[0, 0].x,
    y=main_grid[0, 0].y,
    width=main_grid[0, 0].width,
    height=main_grid[0, 0].height,
    shape_type=1,  # Rectangle
    fill_color="lightblue",
)
# Then add text on top
main_grid[0, 0].content = slide2.add_text(
    text="Main Grid [0,0]",
    x=main_grid[0, 0].x,
    y=main_grid[0, 0].y,
    width=main_grid[0, 0].width,
    height=main_grid[0, 0].height,
    font_size=18,
    align="center",
    vertical="middle",
)

# Add a nested 3x3 grid to the top right cell
nested_grid = main_grid.add_grid_to_cell(
    row=0,
    col=1,
    rows=3,
    cols=3,
    padding=3.0,
)

# Add title to the nested grid's top row
for col in range(3):
    # Add background shape
    slide2.add_shape(
        x=nested_grid[0, col].x,
        y=nested_grid[0, col].y,
        width=nested_grid[0, col].width,
        height=nested_grid[0, col].height,
        shape_type=1,  # Rectangle
        fill_color="lightgreen",
    )
    # Add text on top
    nested_grid[0, col].content = slide2.add_text(
        text=f"Nested [{0},{col}]",
        x=nested_grid[0, col].x,
        y=nested_grid[0, col].y,
        width=nested_grid[0, col].width,
        height=nested_grid[0, col].height,
        font_size=12,
        align="center",
        vertical="middle",
    )

# Merge cells in the nested grid's middle row
merged_cell = nested_grid.merge_cells(1, 0, 1, 2)
# Add background shape
slide2.add_shape(
    x=merged_cell.x,
    y=merged_cell.y,
    width=merged_cell.width,
    height=merged_cell.height,
    shape_type=1,  # Rectangle
    fill_color="lightyellow",
)
# Add text on top
merged_cell.content = slide2.add_text(
    text="Merged cells in nested grid\n[1,0]-[1,2]",
    x=merged_cell.x,
    y=merged_cell.y,
    width=merged_cell.width,
    height=merged_cell.height,
    font_size=14,
    align="center",
    vertical="middle",
)

# Use the flat iterator to access bottom row cells in nested grid
for cell in nested_grid.flat:
    if cell.row == 2:  # Bottom row
        # Add background shape
        slide2.add_shape(
            x=cell.x,
            y=cell.y,
            width=cell.width,
            height=cell.height,
            shape_type=1,  # Rectangle
            fill_color="lightpink",
        )
        # Add text on top
        cell.content = slide2.add_text(
            text=f"flat [{cell.row},{cell.col}]",
            x=cell.x,
            y=cell.y,
            width=cell.width,
            height=cell.height,
            font_size=10,
            align="center",
            vertical="middle",
        )

# Merge cells in the main grid's bottom row
merged_main_cell = main_grid.merge_cells(1, 0, 1, 1)
# Add background shape
slide2.add_shape(
    x=merged_main_cell.x,
    y=merged_main_cell.y,
    width=merged_main_cell.width,
    height=merged_main_cell.height,
    shape_type=1,  # Rectangle
    fill_color="lavender",
)
# Add text on top
merged_main_cell.content = slide2.add_text(
    text="Merged main grid cells [1,0]-[1,1]\nSpans bottom row",
    x=merged_main_cell.x,
    y=merged_main_cell.y,
    width=merged_main_cell.width,
    height=merged_main_cell.height,
    font_size=18,
    align="center",
    vertical="middle",
)

# ---------------------------------------------------------
# Slide 3: Grid Indexing and Iteration (Another example)
# ---------------------------------------------------------
slide3 = pres.add_slide()

# Add a title
slide3.add_text(
    text="Grid Access Methods Summary",
    x="0%",
    y="2%",
    width="100%",
    height="8%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Add a description
slide3.add_text(
    text="Various ways to access and manipulate grid cells",
    x="5%",
    y="12%",
    width="90%",
    height="5%",
    font_size=16,
    align="center",
)

# Create a 4x4 grid
grid = pres.add_grid(
    slide=slide3,
    x="10%",
    y="20%",
    width="80%",
    height="70%",
    rows=4,
    cols=4,
    padding=3.0,
)

# Different ways to access and style cells
methods = [
    {"row": 0, "col": 0, "text": "Tuple Access\ngrid[0,0]", "color": "lightblue"},
    {"row": 0, "col": 1, "text": "Flat Access\ngrid[1]", "color": "lightgreen"},
    {"row": 0, "col": 2, "text": "get_cell()\ngrid.get_cell(0,2)", "color": "lightyellow"},
    {"row": 0, "col": 3, "text": "add_to_cell()\nMethod", "color": "lightpink"},
    {"row": 1, "col": 0, "text": "Basic\nIteration\nfor cell in grid", "color": "lavender"},
    {"index": 5, "text": "Flat Index\ngrid[5]", "color": "peachpuff"},  # [1,1]
    {"row": 1, "col": 2, "text": "Flat Iteration\nfor cell in grid.flat", "color": "lightsalmon"},
    {"row": 1, "col": 3, "text": "Cell Content\ncell.content = obj", "color": "lightcyan"},
    {"index": 8, "text": "Grid[2,0]\nEnumerate\nfor i, cell in enumerate(grid)", "color": "thistle"},
    {"index": 9, "text": "Grid[2,1]\nProperties\nx, y, width, height", "color": "lightsteelblue"},
    {"index": 10, "text": "Grid[2,2]\nPosition\nrow=2, col=2", "color": "mistyrose"},
    {"index": 11, "text": "Grid[2,3]\nContent Access\ngrid[2,3].content", "color": "palegreen"},
    {"row": 3, "col": 0, "text": "Grid[3,0]\nMerging\nmerge_cells()", "color": "wheat"},
    {"row": 3, "col": 1, "text": "Grid[3,1]\nNesting\nadd_grid_to_cell()", "color": "honeydew"},
    {"row": 3, "col": 2, "text": "Grid[3,2]\nSpanning\nspan_rows, span_cols", "color": "lemonchiffon"},
    {"row": 3, "col": 3, "text": "Grid[3,3]\nLast Cell\nflat[-1]", "color": "aliceblue"},
]

# Apply all methods
for method in methods:
    # Use ternary operator for cell access
    cell = grid[method["index"]] if "index" in method else grid[method["row"], method["col"]]

    # Add background shape
    slide3.add_shape(
        x=cell.x,
        y=cell.y,
        width=cell.width,
        height=cell.height,
        shape_type=1,  # Rectangle
        fill_color=method["color"],
    )

    # Add text on top
    cell.content = slide3.add_text(
        text=method["text"],
        x=cell.x,
        y=cell.y,
        width=cell.width,
        height=cell.height,
        font_size=12,
        align="center",
        vertical="middle",
    )

# Save the presentation
output_path = output_dir / "grid_complete_example.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
print("Open the presentation to see all Grid features in action.")
