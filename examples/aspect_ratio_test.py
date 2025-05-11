"""
Test the responsive positioning feature with different aspect ratios.

This script generates the same presentation with different aspect ratios
to demonstrate the responsive positioning feature.
"""

from pathlib import Path

from easypptx import Presentation

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# List of aspect ratios to test
aspect_ratios = ["16:9", "4:3", "16:10", "A4"]

for aspect_ratio in aspect_ratios:
    print(f"Creating presentation with {aspect_ratio} aspect ratio...")

    # Create presentation with the specific aspect ratio
    pres = Presentation(aspect_ratio=aspect_ratio)

    # Add title slide
    slide1 = pres.add_slide()
    slide1.add_text(
        text=f"Aspect Ratio Test ({aspect_ratio})",
        x="50%",
        y="20%",
        width="80%",
        height="15%",
        font_size=44,
        font_bold=True,
        align="center",
        h_align="center",  # Enable responsive positioning
    )
    slide1.add_text(
        text="This presentation demonstrates responsive positioning",
        x="50%",
        y="40%",
        width="70%",
        height="10%",
        font_size=24,
        align="center",
        h_align="center",  # Enable responsive positioning
    )

    # Add a slide showing standard vs. responsive positioning
    slide2 = pres.add_slide()
    slide2.add_text(
        text="Standard vs. Responsive Positioning",
        x="50%",
        y="10%",
        width="80%",
        height="10%",
        font_size=32,
        font_bold=True,
        align="center",
        h_align="center",  # Enable responsive positioning
    )

    # Standard positioning (without h_align)
    slide2.add_shape(
        shape_type=1,  # Rectangle
        x="10%",
        y="25%",
        width="35%",
        height="25%",
        fill_color="red",
        # No h_align, so standard positioning
    )
    slide2.add_text(
        text="Standard positioning\n(may shift in different aspect ratios)",
        x="10%",
        y="55%",
        width="35%",
        height="10%",
        font_size=14,
        align="center",
    )

    # Responsive positioning (with h_align)
    slide2.add_shape(
        shape_type=1,  # Rectangle
        x="55%",
        y="25%",
        width="35%",
        height="25%",
        fill_color="green",
        h_align="center",  # Enable responsive positioning
    )
    slide2.add_text(
        text="Responsive positioning\n(maintains proper alignment)",
        x="55%",
        y="55%",
        width="35%",
        height="10%",
        font_size=14,
        align="center",
        h_align="center",  # Enable responsive positioning
    )

    # Add explanation
    slide2.add_text(
        text=(
            "Responsive positioning automatically adjusts element positions based on the "
            "aspect ratio, ensuring consistent layouts in any presentation format."
        ),
        x="50%",
        y="75%",
        width="80%",
        height="15%",
        font_size=14,
        align="center",
        h_align="center",  # Enable responsive positioning
    )

    # Save the presentation with the aspect ratio in the filename
    output_path = output_dir / f"aspect_ratio_test_{aspect_ratio.replace(':', '_')}.pptx"
    pres.save(output_path)
    print(f"Saved to {output_path}")

print("Created test presentations with different aspect ratios.")
print("Open these files to see how responsive positioning adjusts layout across formats.")
