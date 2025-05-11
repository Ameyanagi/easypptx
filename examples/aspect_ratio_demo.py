"""
Aspect Ratio Demonstration for EasyPPTX.

This script demonstrates how percentage-based positioning works with different aspect ratios
and identifies any potential issues with templates when switching aspect ratios.
"""

from pathlib import Path

from easypptx import Presentation

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)


def create_positioning_demo(aspect_ratio, output_filename):
    """Create a demo presentation with position markers for a specific aspect ratio."""
    print(f"Creating presentation with {aspect_ratio} aspect ratio...")

    # Create presentation with specified aspect ratio
    pres = Presentation(aspect_ratio=aspect_ratio)

    # Add a title slide showing the aspect ratio
    pres.add_title_slide(
        title=f"Aspect Ratio: {aspect_ratio}", subtitle="Demonstrating positioning with different aspect ratios"
    )

    # Add a slide with position markers at key percentages
    slide = pres.add_slide()
    pres.add_text(
        slide=slide,
        text="Position Markers",
        x="50%",
        y="5%",
        width="90%",
        height="10%",
        font_size=32,
        font_bold=True,
        align="center",
    )

    # Create position markers at each corner and center
    positions = [
        {"name": "Top-Left", "x": "0%", "y": "0%", "align": "left"},
        {"name": "Top-Right", "x": "90%", "y": "0%", "align": "right"},
        {"name": "Bottom-Left", "x": "0%", "y": "90%", "align": "left"},
        {"name": "Bottom-Right", "x": "90%", "y": "90%", "align": "right"},
        {"name": "Center", "x": "45%", "y": "45%", "align": "center"},
    ]

    for pos in positions:
        pres.add_text(
            slide=slide,
            text=pos["name"],
            x=pos["x"],
            y=pos["y"],
            width="10%",
            height="10%",
            font_size=14,
            font_bold=True,
            align=pos["align"],
            color="blue",
        )

    # Add horizontal and vertical center lines
    for i in range(0, 101, 10):
        # Vertical lines
        pres.add_text(
            slide=slide, text=f"{i}%", x=f"{i}%", y="50%", width="5%", height="5%", font_size=10, align="center"
        )

        # Horizontal lines
        pres.add_text(
            slide=slide, text=f"{i}%", x="50%", y=f"{i}%", width="5%", height="5%", font_size=10, align="center"
        )

    # Add a slide with a template - for example, content_slide
    template_slide = pres.add_content_slide("Template Demonstration")
    pres.add_text(
        slide=template_slide,
        text="This slide uses the content_slide template.\nCheck if content is aligned correctly with different aspect ratios.",
        x="10%",
        y="20%",
        width="80%",
        height="60%",
        font_size=24,
        align="left",
    )

    # Add a section slide to test full-width content
    pres.add_section_slide("Full Width Section Slide")

    # Add a comparison slide to test side-by-side content
    pres.add_comparison_slide(
        "Side by Side Comparison",
        [
            "This text should be on the left side.\nIt should take up roughly 42% of the slide width.",
            "This text should be on the right side.\nIt should also take up roughly 42% of the slide width.",
        ],
    )

    # Save the presentation
    output_path = OUTPUT_DIR / output_filename
    pres.save(output_path)
    print(f"Saved presentation to {output_path}")

    return output_path


# Create demonstrations for different aspect ratios
demos = [
    ("16:9", "16_9_presentation.pptx"),
    ("4:3", "4_3_presentation.pptx"),
    ("16:10", "16_10_presentation.pptx"),
    ("A4", "a4_presentation.pptx"),
]

for aspect, filename in demos:
    create_positioning_demo(aspect, filename)

print("Created all demonstration presentations in the output directory.")
print("Compare these presentations to identify positioning inconsistencies across aspect ratios.")
