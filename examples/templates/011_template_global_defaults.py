"""Example showing how to use template global defaults.

This example demonstrates how settings in the defaults.global section of a
TOML template are applied to all element types.
"""

import os
from pathlib import Path

from easypptx import Presentation

# Create paths for outputs
output_dir = Path("output/templates")
template_path = output_dir / "global_defaults.toml"
output_path = output_dir / "global_defaults_example.pptx"

# Ensure directory exists
os.makedirs(output_dir, exist_ok=True)

# Create a template with global defaults
template_content = """
bg_color = [255, 255, 255]

[defaults.global]
# Global defaults for all element types
align = "left"
font_name = "Arial"
font_size = 16
font_bold = false
color = [50, 50, 50]
title_font_size = 24
subtitle_font_size = 18

# Optional element-specific defaults can override global defaults
[defaults.grid_slide]
title_align = "center"  # This will override the global align for grid slide titles

[defaults.text]
font_size = 14  # This will override the global font_size for text elements
"""

# Write the template to file
with open(template_path, "w") as f:
    f.write(template_content)

# Create a presentation using the template
prs = Presentation(template_toml=template_path)

# Add a standard slide
std_slide = prs.add_slide(title="Standard Slide")
std_slide.add_text("This text inherits the left alignment from defaults.global.align")

# Add a grid slide
grid_slide, grid = prs.add_grid_slide(
    rows=2,
    cols=2,
    title="Grid Slide",
    subtitle="With global defaults",
)

# The title will be centered because defaults.grid_slide.title_align overrides defaults.global.align
# The subtitle will be left-aligned because it uses defaults.global.align

# Add content to the grid cells
grid[0, 0].add_text("Cell (0,0): Text aligned left from defaults.global.align")
grid[0, 1].add_text("Cell (0,1): Text can be overridden", align="right")
grid[1, 0].add_text("Cell (1,0): Font size from defaults.text")

# Save the presentation
prs.save(output_path)
print(f"Presentation saved to {output_path}")
