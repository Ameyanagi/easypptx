#!/usr/bin/env python
"""
001_template_basic.py - Basic TOML Template Usage

This example demonstrates:
1. Creating a presentation with a default TOML template
2. Adding slides that automatically use this template
3. Adding a slide with a different TOML template
"""

from pathlib import Path

from easypptx import Presentation

# Set up paths
CURRENT_DIR = Path(__file__).parent
ROOT_DIR = CURRENT_DIR.parent.parent
TEMPLATE_DIR = ROOT_DIR / "templates"
OUTPUT_DIR = ROOT_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# Template paths
business_template = str(TEMPLATE_DIR / "business_title.toml")
tech_template = str(TEMPLATE_DIR / "tech_dark.toml")

# Create a presentation with a default template
pres = Presentation(template_toml=business_template)

# Slide 1: Uses the default template
slide1 = pres.add_slide(title="Default Template")
slide1.add_text(
    text="This slide uses the business_title.toml template",
    x="10%",
    y="50%",
    width="80%",
    height="20%",
    font_size=24,
    align="center",
)

# Slide 2: Uses a different template
slide2 = pres.add_slide(title="Override Template", template_toml=tech_template)
slide2.add_text(
    text="This slide uses tech_dark.toml template",
    x="10%",
    y="50%",
    width="80%",
    height="20%",
    font_size=24,
    align="center",
    color="white",
)

# Slide 3: Standard slide without a template
slide3 = pres.add_slide(
    title="No Template",
    template_toml=None,  # Explicitly disable template
)
slide3.add_text(
    text="This is a standard slide without a template",
    x="10%",
    y="50%",
    width="80%",
    height="20%",
    font_size=24,
    align="center",
)

# Save the presentation
output_path = OUTPUT_DIR / "001_template_basic.pptx"
pres.save(output_path)
print(f"Saved: {output_path}")

if __name__ == "__main__":
    pass  # The code runs automatically when the module is executed
