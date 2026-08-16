"""
002_grid_indexing.py - Grid Indexing and Iteration Example

This example demonstrates the different ways to access grid cells:
1. Using [row, col] indexing: grid[0, 1]
2. Using flat index: grid[2]
3. Using iteration: for cell in grid
4. Using flat iteration: for cell in grid.flat
"""

from pathlib import Path

from easypptx import Presentation

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# Add a slide
slide = pres.add_slide()

# Add a title
slide.add_text(
    text="002 - Grid Indexing & Iteration",
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
    slide=slide,
    x="5%",
    y="15%",
    width="90%",
    height="80%",
    rows=3,
    cols=3,
    padding=5.0,
)

# Method 1: Using the grid[row, col] syntax
# Style the cell background, then add text on top
grid[0, 0].style(fill=(173, 216, 230)).add_text(
    "grid[0, 0]\nTuple indexing",
    font_size=14,
    align="center",
    vertical="middle",
)

# Method 2: Using the flat index syntax
# Center cell (row 1, col 1) is index 4 in flattened grid
grid[4].style(fill=(144, 238, 144)).add_text(
    "grid[4]\nFlat indexing\n(Center)",
    font_size=14,
    align="center",
    vertical="middle",
)

# Method 3: Using iteration
for i, cell in enumerate(grid):
    if i in [2, 6]:  # Only add content to specific cells
        grid[cell.row, cell.col].style(fill=(255, 255, 224)).add_text(
            f"grid iteration\nindex {i}\n[{cell.row}, {cell.col}]",
            font_size=14,
            align="center",
            vertical="middle",
        )

# Method 4: Using flat iterator to identify specific cells
for cell in grid.flat:
    # Add content to specific cells based on their row/col
    if (cell.row == 0 and cell.col == 1) or (cell.row == 2 and cell.col == 2):
        grid[cell.row, cell.col].style(fill=(255, 182, 193)).add_text(
            f"grid.flat\n[{cell.row}, {cell.col}]",
            font_size=14,
            align="center",
            vertical="middle",
        )

# Method 5: Using add_to_cell with shape for background
grid.add_to_cell(
    row=1,
    col=2,
    content_func=slide.add_shape,
    shape_type=1,  # Rectangle
    fill_color=(230, 230, 250),
)

# Add text on top
grid.add_to_cell(
    row=1, col=2, content_func=slide.add_text, text="add_to_cell(1,2)", font_size=14, align="center", vertical="middle"
)

# Save the presentation
output_path = output_dir / "002_grid_indexing.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
