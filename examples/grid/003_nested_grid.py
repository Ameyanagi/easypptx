"""
003_nested_grid.py - Nested Grid and Cell Merging Example

This example demonstrates:
1. Creating nested grids (grids within grid cells)
2. Merging cells in both main and nested grids
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
    text="003 - Nested Grids & Cell Merging",
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
    slide=slide,
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
slide.add_shape(
    x=main_grid[0, 0].x,
    y=main_grid[0, 0].y,
    width=main_grid[0, 0].width,
    height=main_grid[0, 0].height,
    shape_type=1,  # Rectangle
    fill_color="lightblue",
)
# Then add text on top
main_grid[0, 0].content = slide.add_text(
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
    slide.add_shape(
        x=nested_grid[0, col].x,
        y=nested_grid[0, col].y,
        width=nested_grid[0, col].width,
        height=nested_grid[0, col].height,
        shape_type=1,  # Rectangle
        fill_color="lightgreen",
    )
    # Add text on top
    nested_grid[0, col].content = slide.add_text(
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
slide.add_shape(
    x=merged_cell.x,
    y=merged_cell.y,
    width=merged_cell.width,
    height=merged_cell.height,
    shape_type=1,  # Rectangle
    fill_color="lightyellow",
)
# Add text on top
merged_cell.content = slide.add_text(
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
        slide.add_shape(
            x=cell.x,
            y=cell.y,
            width=cell.width,
            height=cell.height,
            shape_type=1,  # Rectangle
            fill_color="lightpink",
        )
        # Add text on top
        cell.content = slide.add_text(
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
slide.add_shape(
    x=merged_main_cell.x,
    y=merged_main_cell.y,
    width=merged_main_cell.width,
    height=merged_main_cell.height,
    shape_type=1,  # Rectangle
    fill_color="lavender",
)
# Add text on top
merged_main_cell.content = slide.add_text(
    text="Merged main grid cells [1,0]-[1,1]\nSpans bottom row",
    x=merged_main_cell.x,
    y=merged_main_cell.y,
    width=merged_main_cell.width,
    height=merged_main_cell.height,
    font_size=18,
    align="center",
    vertical="middle",
)

# Save the presentation
output_path = output_dir / "003_nested_grid.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
