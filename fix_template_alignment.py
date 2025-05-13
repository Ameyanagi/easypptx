import os
import sys

sys.path.insert(0, os.path.abspath("."))
from src.easypptx import Presentation


def fix_template_title_alignment():
    """Demonstrate how to fix the template title alignment issue."""
    # Method 1: Create the presentation with template_toml parameter
    pres = Presentation(template_toml="template.toml")

    # For add_grid_slide, we need to explicitly pass the title_align parameter
    # to override the default "center" alignment
    slide, grid = pres.add_grid_slide(
        title="Key Features",
        cols=3,
        rows=2,
        height="33%",
        title_align="left",  # This overrides the default center alignment
    )

    # Add content to grid cells
    grid[0].add_text(
        text="1. Easy-to-Use API",
        font_size=24,
        font_bold=True,
        align="left",
        vertical="top",
    )

    # Save the presentation
    pres.save("fixed_alignment_method1.pptx")

    # Method 2: Use add_slide_from_template instead
    # This fully applies the template styling
    pres2 = Presentation()
    # Load the template
    template_name = pres2.template_manager.load("template.toml")
    # Add a slide using the template
    slide = pres2.add_slide_from_template(template_name)

    # Now add a grid to this slide
    grid = pres2.add_grid(
        slide=slide,
        x="0%",
        y="50%",  # Position below the title
        width="100%",
        height="50%",
        rows=2,
        cols=3,
    )

    # Add content to grid
    grid[0].add_text(
        text="1. Easy-to-Use API",
        font_size=24,
        font_bold=True,
        align="left",
        vertical="top",
    )

    # Set the title text
    title_shapes = [shape for shape in slide.shapes if shape.has_text_frame]
    if title_shapes and len(title_shapes) > 0:
        title_shape = title_shapes[0]
        title_shape.text_frame.text = "Key Features"

    # Save the presentation
    pres2.save("fixed_alignment_method2.pptx")


if __name__ == "__main__":
    fix_template_title_alignment()
