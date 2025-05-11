"""Example demonstrating nested grid layouts in EasyPPTX."""

from easypptx import Presentation

# Create a new presentation
presentation = Presentation()

# Add a slide
slide = presentation.add_slide()

# Add a title
slide.add_text(
    text="Nested Grid Example",
    x="0%",
    y="2%",
    width="100%",
    height="8%",
    font_size=32,
    font_bold=True,
    align="center",
)

# 1. Create a main grid with 2 rows, 2 columns
main_grid = presentation.add_grid(
    slide=slide,
    x="5%",
    y="15%",
    width="90%",
    height="80%",
    rows=2,
    cols=2,
    padding=5.0,
)

# Add content to the first cell (Top-Left)
main_grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,
    text="Main Grid (0,0)",
    font_size=24,
    align="center",
    vertical="middle",
)

# Create a nested grid in the second cell (Top-Right)
nested_grid1 = main_grid.add_grid_to_cell(
    row=0,
    col=1,
    rows=2,
    cols=1,
    padding=5.0,
)

# Add content to the nested grid cells
nested_grid1.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,
    text="Nested Grid 1 (0,0)",
    font_size=18,
    align="center",
    vertical="middle",
)

nested_grid1.add_to_cell(
    row=1,
    col=0,
    content_func=slide.add_text,
    text="Nested Grid 1 (1,0)",
    font_size=18,
    align="center",
    vertical="middle",
)

# Create another nested grid in the third cell (Bottom-Left)
nested_grid2 = main_grid.add_grid_to_cell(
    row=1,
    col=0,
    rows=1,
    cols=2,
    padding=5.0,
)

# Add content to the nested grid cells
nested_grid2.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,
    text="Nested Grid 2 (0,0)",
    font_size=18,
    align="center",
    vertical="middle",
)

nested_grid2.add_to_cell(
    row=0,
    col=1,
    content_func=slide.add_text,
    text="Nested Grid 2 (0,1)",
    font_size=18,
    align="center",
    vertical="middle",
)

# Create a deeply nested grid in the fourth cell (Bottom-Right)
nested_grid3 = main_grid.add_grid_to_cell(
    row=1,
    col=1,
    rows=2,
    cols=2,
    padding=5.0,
)

# Add content to the deeply nested grid cells
nested_grid3.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,
    text="Nested Grid 3 (0,0)",
    font_size=14,
    align="center",
    vertical="middle",
)

# And even further nesting
deeply_nested_grid = nested_grid3.add_grid_to_cell(
    row=0,
    col=1,
    rows=2,
    cols=1,
    padding=3.0,
)

deeply_nested_grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,
    text="Deep Nested (0,0)",
    font_size=10,
    align="center",
    vertical="middle",
)

deeply_nested_grid.add_to_cell(
    row=1,
    col=0,
    content_func=slide.add_text,
    text="Deep Nested (1,0)",
    font_size=10,
    align="center",
    vertical="middle",
)

# Add content to the remaining cells of nested_grid3
nested_grid3.add_to_cell(
    row=1,
    col=0,
    content_func=slide.add_text,
    text="Nested Grid 3 (1,0)",
    font_size=14,
    align="center",
    vertical="middle",
)

nested_grid3.add_to_cell(
    row=1,
    col=1,
    content_func=slide.add_text,
    text="Nested Grid 3 (1,1)",
    font_size=14,
    align="center",
    vertical="middle",
)

# Save the presentation
presentation.save("output/nested_grid_example.pptx")
print("Presentation saved as output/nested_grid_example.pptx")
