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

# Add content to top left cell: style the background, then add text
main_grid[0, 0].style(fill=(173, 216, 230)).add_text(
    "Main Grid [0,0]",
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
    nested_grid[0, col].style(fill=(144, 238, 144)).add_text(
        f"Nested [{0},{col}]",
        font_size=12,
        align="center",
        vertical="middle",
    )

# Merge cells in the nested grid's middle row using slice syntax
nested_grid[1, :].style(fill=(255, 255, 224)).add_text(
    "Merged cells in nested grid\n[1,0]-[1,2]",
    font_size=14,
    align="center",
    vertical="middle",
)

# Use the flat iterator to access bottom row cells in nested grid
for cell in nested_grid.flat:
    if cell.row == 2:  # Bottom row
        nested_grid[cell.row, cell.col].style(fill=(255, 182, 193)).add_text(
            f"flat [{cell.row},{cell.col}]",
            font_size=10,
            align="center",
            vertical="middle",
        )

# Merge cells in the main grid's bottom row using slice syntax
main_grid[1, :].style(fill=(230, 230, 250)).add_text(
    "Merged main grid cells [1,0]-[1,1]\nSpans bottom row",
    font_size=18,
    align="center",
    vertical="middle",
)

# Save the presentation
output_path = output_dir / "003_nested_grid.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
