"""Example demonstrating Grid indexing and iteration features in EasyPPTX."""

from easypptx import Presentation

# Create a new presentation
presentation = Presentation()

# Add a slide
slide = presentation.add_slide()

# Add a title
slide.add_text(
    text="Grid Indexing & Iteration Example",
    x="0%",
    y="2%",
    width="100%",
    height="8%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Create a 3x3 grid
grid = presentation.add_grid(
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
grid[0, 0].content = slide.add_text(
    text="grid[0, 0]",
    x=grid[0, 0].x,
    y=grid[0, 0].y,
    width=grid[0, 0].width,
    height=grid[0, 0].height,
    font_size=18,
    align="center",
    vertical="middle",
)

# Method 2: Using the flat index syntax
grid[4].content = slide.add_text(  # Center cell (row 1, col 1) is index 4 in flattened grid
    text="grid[4]\n(Center)",
    x=grid[4].x,
    y=grid[4].y,
    width=grid[4].width,
    height=grid[4].height,
    font_size=18,
    align="center",
    vertical="middle",
)

# Method 3: Using iteration
for i, cell in enumerate(grid):
    if i in [2, 6, 8]:  # Only add content to specific cells
        cell.content = slide.add_text(
            text=f"grid[{i}]\nIteration",
            x=cell.x,
            y=cell.y,
            width=cell.width,
            height=cell.height,
            font_size=18,
            align="center",
            vertical="middle",
        )

# Method 4: Using flat iterator to identify specific cells
for cell in grid.flat:
    # Add content to specific cells based on their row/col
    if (cell.row == 1 and cell.col == 0) or (cell.row == 1 and cell.col == 2):
        cell.content = slide.add_text(
            text=f"flat: [{cell.row}, {cell.col}]",
            x=cell.x,
            y=cell.y,
            width=cell.width,
            height=cell.height,
            font_size=18,
            align="center",
            vertical="middle",
        )

# Method 5: Using the standard add_to_cell method as a comparison
grid.add_to_cell(
    row=2,
    col=1,
    content_func=slide.add_text,
    text="add_to_cell(2,1)",
    font_size=18,
    align="center",
    vertical="middle",
)

# Save the presentation
presentation.save("output/grid_indexing_example.pptx")
print("Presentation saved as output/grid_indexing_example.pptx")
