"""
Example demonstrating responsive positioning with different aspect ratios.

This example shows how EasyPPTX handles different aspect ratios.
"""

from pathlib import Path

from easypptx import Presentation

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

# Create presentations with different aspect ratios
aspect_ratios = ["16:9", "4:3", "16:10", "A4"]

for aspect_ratio in aspect_ratios:
    # Create a presentation with the specified aspect ratio
    print(f"Creating presentation with {aspect_ratio} aspect ratio...")
    pres = Presentation(aspect_ratio=aspect_ratio)

    # Add a title slide
    title_slide = pres.add_title_slide(
        title=f"Responsive Positioning ({aspect_ratio})",
        subtitle="Demonstration of percentage-based positioning for different aspect ratios",
    )

    # Add a slide with centered content
    centered_slide = pres.add_slide()

    # Add centered title with percentage-based positioning
    centered_slide.add_text(
        text="Centered Title",
        x="50%",  # 50% from left
        y="5%",
        width="80%",
        height="10%",
        font_size=32,
        font_bold=True,
        align="center",  # Center text within the box
    )

    # Add a large shape
    centered_slide.add_shape(
        shape_type=1,  # Rectangle
        x="10%",
        y="20%",
        width="80%",
        height="40%",
        fill_color="blue",
    )

    # Add explanatory text
    centered_slide.add_text(
        text=(
            "This slide demonstrates percentage-based positioning that works "
            "across different aspect ratios. Using percentages instead of "
            "absolute positions ensures elements scale properly."
        ),
        x="10%",
        y="70%",
        width="80%",
        height="20%",
        font_size=18,
        align="center",
    )

    # Add a comparison slide
    comparison_slide = pres.add_slide()
    comparison_slide.add_text(
        text="Comparison: Absolute vs. Percentage",
        x="50%",
        y="5%",
        width="80%",
        height="10%",
        font_size=32,
        font_bold=True,
        align="center",
    )

    # Absolute positioning (in inches)
    comparison_slide.add_shape(
        shape_type=1,  # Rectangle
        x=1.0,  # Absolute inches
        y=2.0,  # Absolute inches
        width=3.0,
        height=2.0,
        fill_color="red",
    )

    comparison_slide.add_text(
        text="Absolute positioning\n(fixed inches, may look different)",
        x=1.0,
        y=4.5,
        width=3.0,
        height=1.0,
        font_size=14,
        align="center",
    )

    # Percentage-based positioning
    comparison_slide.add_shape(
        shape_type=1,  # Rectangle
        x="55%",
        y="20%",
        width="35%",
        height="30%",
        fill_color="green",
    )

    comparison_slide.add_text(
        text="Percentage positioning\n(scales with slide dimensions)",
        x="55%",
        y="50%",
        width="35%",
        height="10%",
        font_size=14,
        align="center",
    )

    # Add implementation explanation
    comparison_slide.add_text(
        text=(
            'Implementation: When using percentage values (like "50%"), '
            "the positions and sizes are calculated relative to the slide dimensions. "
            "This ensures elements maintain proper proportions across different slide sizes."
        ),
        x="10%",
        y="70%",
        width="80%",
        height="20%",
        font_size=14,
        align="center",
    )

    # Save the presentation
    output_path = OUTPUT_DIR / f"responsive_positioning_{aspect_ratio.replace(':', '_')}.pptx"
    pres.save(output_path)
    print(f"Saved presentation to {output_path}")

print(
    "\nCreated presentations with different aspect ratios to demonstrate responsive positioning."
    "\nOpen and compare these files to see how content positioning scales with different dimensions."
)
