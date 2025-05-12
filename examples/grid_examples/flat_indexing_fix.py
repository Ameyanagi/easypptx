import os

from easypptx import Presentation

# Create a new presentation
pres = Presentation()

# Create a new slide with a title and grid layout
slide, grid = pres.add_grid_slide(title="Key Features", cols=3, rows=2, height="33%")
# Add content to the grid cells using flat indexing

# Row 0: Cells 0, 1, 2
grid[0].add_text(
    text="1. Easy-to-Use API",
    font_size=24,
    font_bold=True,
    align="left",
    vertical="top",
)

grid[1].add_text(
    text="2. Customizable Layouts",
    font_size=24,
    font_bold=True,
    align="left",
    vertical="top",
)

grid[2].add_text(
    text="3. Flexible Content Management",
    font_size=24,
    font_bold=True,
    align="left",
    vertical="top",
)

# Row 1: Cells 3, 4, 5
grid[3].add_text(
    text="4. Enhanced Row-Based API",
    font_size=24,
    font_bold=True,
    align="left",
    vertical="top",
)

grid[4].add_text(
    text="5. Dynamic Grid Sizing",
    font_size=24,
    font_bold=True,
    align="left",
    vertical="top",
)

grid[5].add_text(
    text="6. Nested Grid Support",
    font_size=24,
    font_bold=True,
    align="left",
    vertical="top",
)

# Save the presentation
# Ensure output directory exists
os.makedirs("../../output", exist_ok=True)
pres.save("../../output/flat_indexing_fix.pptx")
