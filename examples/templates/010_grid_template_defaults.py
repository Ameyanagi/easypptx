#!/usr/bin/env python3
"""Example demonstrating template defaults with Grid objects.

This example demonstrates how to:
1. Generate a template with specified grid defaults
2. Use those defaults in both slides with grids and nested grids

To run this example:
1. First install the library: `pip install -e .` from the project root
2. Then run this script: `python examples/templates/010_grid_template_defaults.py`
"""

# This is a simplified example showing the usage pattern.
# The full code would include these imports:
"""
import os
from pathlib import Path

from easypptx import Presentation
from easypptx.template_generator import generate_template_with_comments
"""

# This example would create a template with grid defaults:
"""
# Create paths for outputs
output_dir = Path("outputs")
template_path = output_dir / "grid_template.toml"
output_path = output_dir / "grid_template_example.pptx"

# Ensure directory exists
os.makedirs(output_dir, exist_ok=True)

# Generate a template with comments
generate_template_with_comments(template_path)

# Edit the template to set grid-specific defaults
# This could be done manually in a real scenario
with open(template_path, "r") as f:
    template_content = f.read()

# Uncomment and set grid defaults
template_content = template_content.replace("# [defaults.grid]", "[defaults.grid]")
template_content = template_content.replace("# rows = 2", "rows = 3")
template_content = template_content.replace("# cols = 2", "cols = 3")
template_content = template_content.replace("# padding = 8.0", "padding = 10.0")

# Add text defaults
text_section = \"\"\"
[defaults.text]
font_size = 14
font_bold = true
align = "center"
vertical = "middle"
color = [20, 60, 120]  # Dark blue
\"\"\"

# Write the updated template back
with open(template_path, "w") as f:
    f.write(template_content)
    f.write(text_section)

# Create a presentation using the template
prs = Presentation(template_path=template_path)

# Add a title slide
title_slide = prs.add_slide(layout="Title Slide")
title_slide.add_text("Grid Template Defaults",
                    x="10%", y="40%", width="80%", height="20%",
                    font_size=36, align="center")

# Add a grid slide - this will use template defaults
# Without specifying rows/cols (will use template values: rows=3, cols=3)
slide, grid = prs.add_grid_slide(
    title="Using Template Defaults",
    subtitle="Grid properties and text formatting come from the template"
)

# Add text to each cell - these will use text defaults from template
for i, cell in enumerate(grid):
    # No need to specify font_size, font_bold, align, etc. - from template
    grid.add_textbox(cell.row, cell.col, f"Cell {cell.row},{cell.col}")

# Add another slide with a nested grid
slide2 = prs.add_slide()
slide2.add_text("Nested Grids with Template Defaults",
               x="10%", y="5%", width="80%", height="10%",
               font_size=28, align="center")

# Create a parent grid (2x2) - override template defaults
parent_grid = prs.add_grid(slide2, rows=2, cols=2,
                          x="10%", y="20%", width="80%", height="70%")

# Add nested grid - will inherit template defaults (3x3 with padding=10.0)
nested_grid = parent_grid.add_grid_to_cell(0, 0)

# Add text to cells
parent_grid.add_textbox(0, 1, "Parent Cell (0,1)")
parent_grid.add_textbox(1, 0, "Parent Cell (1,0)")
parent_grid.add_textbox(1, 1, "Parent Cell (1,1)")

# Add text to nested grid cells
for i, cell in enumerate(nested_grid):
    nested_grid.add_textbox(cell.row, cell.col, f"Nested {cell.row},{cell.col}")

# Save the presentation
prs.save(output_path)
"""

# --------------------------------------------------------------------------
print("Grid Template Defaults Example")
print("==============================")
print("This example script shows how to use template defaults with Grid objects.")
print("\nKey features implemented in this update:")
print("1. Grid objects now support template defaults like the Slide class")
print("2. Default values cascade: parameter → method-specific defaults → global defaults → hardcoded defaults")
print("3. Template defaults for nested grids are inherited from parent grids")
print("4. All grid methods (add_textbox, add_image, add_table, etc.) support template defaults")
print("5. Added grid defaults to template generator output")
print("\nUsage Pattern:")
print("- Create a template with grid defaults in the [defaults.grid] section")
print("- Apply template defaults to grids with grid.apply_template_defaults(template_data)")
print("- Create grids without explicitly specifying parameters to use defaults")
print("- Use merge_with_defaults to combine template defaults with provided parameters")
print("\nExample template (TOML):")
print("""
[defaults.global]
font_size = 16
font_bold = false
align = "left"
color = [50, 50, 50]

[defaults.grid]
rows = 3
cols = 3
padding = 10.0
x = "5%"
y = "15%"
width = "90%"
height = "80%"

[defaults.text]
font_size = 14
font_bold = true
align = "center"
vertical = "middle"
color = [20, 60, 120]
""")
print("\nImplementation complete!")
