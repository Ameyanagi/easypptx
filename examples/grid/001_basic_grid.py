"""
001_basic_grid.py - Basic Grid Layout Example

This example demonstrates basic grid layout creation and usage.
"""

from pathlib import Path

from easypptx import Presentation

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# Add a title slide
slide = pres.add_slide()
slide.add_text(
    text="001 - Basic Grid Layout",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Create a 2x2 grid layout
grid = pres.add_grid(
    slide=slide,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    rows=2,
    cols=2,
    padding=5.0,
)

# Add content to each cell using add_to_cell method
grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,
    text="Top Left Cell",
    font_size=24,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=0,
    col=1,
    content_func=slide.add_text,
    text="Top Right Cell",
    font_size=24,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=1,
    col=0,
    content_func=slide.add_text,
    text="Bottom Left Cell",
    font_size=24,
    align="center",
    vertical="middle",
)

grid.add_to_cell(
    row=1,
    col=1,
    content_func=slide.add_text,
    text="Bottom Right Cell",
    font_size=24,
    align="center",
    vertical="middle",
)

# Save the presentation
output_path = output_dir / "001_basic_grid.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
