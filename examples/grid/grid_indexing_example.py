"""
Grid Indexing Example

This example demonstrates the different ways to access cells in a grid layout:
1. Using row and column indices grid[row, col]
2. Using row access grid[row], which gives access to the next available cell in that row
3. Using flat indexing grid[idx], which accesses cells in a flattened array

Key features demonstrated:
- Different ways to access grid cells
- How flat indexing works with negative indices
- Showing the visual layout of accessed cells
"""

from easypptx import Presentation

# Create a new presentation
pres = Presentation()

# Create a slide with a title and a 3x3 grid
slide, grid = pres.add_grid_slide(rows=3, cols=3, title="Grid Indexing Demo", title_align="center")

# Method 1: Access cells by [row, col]
grid[0, 0].add_text(
    text="grid[0, 0]", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="lightgray"
)

grid[0, 1].add_text(
    text="grid[0, 1]", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="lightblue"
)

grid[1, 1].add_text(
    text="grid[1, 1]", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="lightgreen"
)

# Method 2: Create a new slide with a title and a 3x3 grid for row access
slide2, grid2 = pres.add_grid_slide(rows=3, cols=3, title="Row Access Demo", title_align="center")

# Access using row indexing (automatically uses next available cell in the row)
grid2[0].add_text(
    text="grid[0] - first", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="pink"
)

grid2[0].add_text(
    text="grid[0] - second", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="lightblue"
)

grid2[0].add_text(
    text="grid[0] - third", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="lightgreen"
)

grid2[1].add_text(
    text="grid[1] - first", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="orange"
)

# Create a third slide with a 3x3 grid for flat indexing
slide3, grid3 = pres.add_grid_slide(rows=3, cols=3, title="Flat Indexing Demo", title_align="center")

# Method 3: Access cells using flat indexing
# Flat indexing accesses cells in row-major order (0 is top-left, then across the row)

# Demonstrating flat indexing with all cells
for i in range(9):
    grid3[i].add_text(
        text=f"grid[{i}]",
        font_size=18,
        font_bold=True,
        align="center",
        vertical="middle",
        bg_color=[
            "lightgray",
            "lightblue",
            "lightgreen",
            "pink",
            "yellow",
            "lightcyan",
            "orange",
            "lavender",
            "lightsalmon",
        ][i],
    )

# Create a fourth slide with a 3x3 grid for negative indexing
slide4, grid4 = pres.add_grid_slide(rows=3, cols=3, title="Negative Flat Indexing Demo", title_align="center")

# Method 4: Access cells using negative flat indexing
grid4[-1].add_text(text="grid[-1]", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="pink")

grid4[-3].add_text(
    text="grid[-3]", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="lightblue"
)

grid4[-9].add_text(
    text="grid[-9]", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="lightgreen"
)

# Save the presentation
pres.save("grid_indexing_example.pptx")
print("Presentation saved as 'grid_indexing_example.pptx'")
