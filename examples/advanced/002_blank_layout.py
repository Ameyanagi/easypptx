"""
Blank layout example for EasyPPTX.

This example demonstrates how EasyPPTX uses blank layouts by default
and how different slide layouts can be specified.
"""

from pathlib import Path

from easypptx import Presentation, Text

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)


def create_blank_layout_presentation():
    """Create a presentation demonstrating different slide layouts."""
    print("Creating blank layout example...")

    # Create a presentation
    pres = Presentation()

    # Add slide with default blank layout
    slide = pres.add_slide()
    text = Text(slide)
    text.add_title("Default Blank Layout")
    text.add_paragraph("This slide uses the default blank layout (layout 6)", y="30%", font_size=24, color="blue")

    # Add slide with title layout (index 0)
    slide = pres.add_slide(layout_index=0)
    text = Text(slide)
    text.add_paragraph(
        "This slide uses the title layout (layout 0)", y="60%", font_size=24, color="red", align="center"
    )

    # Add slide with title and content layout (index 1)
    slide = pres.add_slide(layout_index=1)
    text = Text(slide)
    text.add_paragraph(
        "This slide uses the title and content layout (layout 1)", y="50%", font_size=24, color="green", align="center"
    )

    # Add slide with blank layout and background color
    slide = pres.add_slide(bg_color="blue")
    text = Text(slide)
    text.add_title("Blank Layout with Background", color="white")
    text.add_paragraph("This slide has a custom background color", y="40%", font_size=24, color="white", align="center")

    # Save the presentation
    pres.save(output_dir / "blank_layout_example.pptx")
    print(f"Saved to {output_dir / 'blank_layout_example.pptx'}")


def create_from_template():
    """Create a presentation from a template and add blank slides."""
    print("\nCreating template example...")

    # Look for an existing PPTX in the output directory to use as a template
    template_files = list(output_dir.glob("*.pptx"))

    if template_files:
        # Use the first PPTX found as a template
        template_path = template_files[0]
        print(f"Using {template_path} as template")

        # Create presentation from template
        pres = Presentation(template_path=str(template_path))

        # Add slide with default blank layout
        slide = pres.add_slide()
        text = Text(slide)
        text.add_title("Slide Based on Template")
        text.add_paragraph(
            f"This presentation uses {template_path.name} as a template", y="30%", font_size=24, color="blue"
        )

        # Add slide with specific layout if available
        try:
            slide = pres.add_slide(layout_index=2)  # Section header layout if available
            text = Text(slide)
            text.add_title("Using Template Layout")
            text.add_paragraph("This slide uses layout index 2 from the template", y="40%", font_size=24, color="green")
        except IndexError:
            # If layout index doesn't exist, use blank layout
            slide = pres.add_slide()
            text = Text(slide)
            text.add_title("Fallback to Blank Layout")
            text.add_paragraph(
                "Layout index 2 not available, using blank layout instead", y="40%", font_size=24, color="red"
            )

        # Save the presentation
        output_path = output_dir / "template_with_blank_slides.pptx"
        pres.save(output_path)
        print(f"Saved to {output_path}")
    else:
        print("No template file found in output directory")


# Execute examples
if __name__ == "__main__":
    print("Creating blank layout examples...\n")

    create_blank_layout_presentation()
    create_from_template()

    print("\nAll blank layout examples created successfully!")
