from easypptx import Presentation

# Create a new presentation
pres = Presentation()

# Create a new slide with a title and grid layout
slide, grid = pres.add_grid_slide(title="Flat Indexing Example", cols=3, rows=2)

# Using flattened indexing (row-wise by default)
# In a 3x2 grid, the flattened index mapping is:
#  0  1  2
#  3  4  5

# Flat indexing - first row
grid[0].add_text(
    text="Index 0 (row 0, col 0)",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid[1].add_text(
    text="Index 1 (row 0, col 1)",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid[2].add_text(
    text="Index 2 (row 0, col 2)",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

# Flat indexing - second row
grid[3].add_text(
    text="Index 3 (row 1, col 0)",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid[4].add_text(
    text="Index 4 (row 1, col 1)",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid[5].add_text(
    text="Index 5 (row 1, col 2)",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

# Create a second slide to demonstrate row access
slide2, grid2 = pres.add_grid_slide(title="Row Access Example", cols=3, rows=2)

# Row access automatically adds content to the next available cell in the row
grid2[0].add_text(
    text="Row 0 - First Cell",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid2[0].add_text(
    text="Row 0 - Second Cell",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid2[0].add_text(
    text="Row 0 - Third Cell",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid2[1].add_text(
    text="Row 1 - First Cell",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

grid2[1].add_text(
    text="Row 1 - Second Cell",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

# Save the presentation
pres.save("output/flat_grid_example.pptx")
