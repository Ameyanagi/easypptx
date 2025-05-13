"""
Aspect ratio example for EasyPPTX.

This example demonstrates how to create presentations with different aspect ratios
and how to view the slide dimensions of each presentation.
"""

from pathlib import Path

from easypptx import Presentation, Text

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)


# Function to create and save a presentation with a specific aspect ratio
def create_presentation(aspect_ratio, filename):
    """Create a presentation with the specified aspect ratio."""
    # Create the presentation
    pres = Presentation(aspect_ratio=aspect_ratio)

    # Add a title slide
    slide = pres.add_slide()
    text = Text(slide)
    text.add_title(f"{aspect_ratio} Aspect Ratio")

    # Add dimensions information
    width_emu = pres.pptx_presentation.slide_width
    height_emu = pres.pptx_presentation.slide_height
    ratio = width_emu / height_emu

    # EMU (English Metric Unit) is 1/914400 of an inch
    width_inches = width_emu / 914400
    height_inches = height_emu / 914400

    text.add_paragraph(f"Width: {width_inches:.2f} inches", y=2)
    text.add_paragraph(f"Height: {height_inches:.2f} inches", y=2.5)
    text.add_paragraph(f"Aspect Ratio: {ratio:.2f}", y=3)

    # Save the presentation
    output_path = output_dir / filename
    pres.save(output_path)
    print(f"Created {aspect_ratio} presentation: {output_path}")

    return width_inches, height_inches, ratio


# Create presentations with different aspect ratios
print("Creating presentations with different aspect ratios...\n")

# 16:9 (default)
w1, h1, r1 = create_presentation("16:9", "16_9_presentation.pptx")

# 4:3
w2, h2, r2 = create_presentation("4:3", "4_3_presentation.pptx")

# 16:10
w3, h3, r3 = create_presentation("16:10", "16_10_presentation.pptx")

# A4
w4, h4, r4 = create_presentation("A4", "a4_presentation.pptx")

# Custom dimensions
custom_pres = Presentation(width_inches=12, height_inches=9)
slide = custom_pres.add_slide()
text = Text(slide)
text.add_title("Custom Dimensions")

width_emu = custom_pres.pptx_presentation.slide_width
height_emu = custom_pres.pptx_presentation.slide_height
ratio = width_emu / height_emu

width_inches = width_emu / 914400
height_inches = height_emu / 914400

text.add_paragraph(f"Width: {width_inches:.2f} inches", y=2)
text.add_paragraph(f"Height: {height_inches:.2f} inches", y=2.5)
text.add_paragraph(f"Aspect Ratio: {ratio:.2f}", y=3)

custom_pres.save(output_dir / "custom_dimensions.pptx")
print(f"Created custom dimensions presentation: {output_dir / 'custom_dimensions.pptx'}")

# Print summary table
print("\nAspect Ratio Summary:")
print("-" * 60)
print(f"{'Aspect Ratio':<15} {'Width (in)':<15} {'Height (in)':<15} {'Ratio':<10}")
print("-" * 60)
print(f"{'16:9':<15} {w1:<15.2f} {h1:<15.2f} {r1:<10.2f}")
print(f"{'4:3':<15} {w2:<15.2f} {h2:<15.2f} {r2:<10.2f}")
print(f"{'16:10':<15} {w3:<15.2f} {h3:<15.2f} {r3:<10.2f}")
print(f"{'A4':<15} {w4:<15.2f} {h4:<15.2f} {r4:<10.2f}")
print(f"{'Custom (12x9)':<15} {12:<15.2f} {9:<15.2f} {12 / 9:<10.2f}")
print("-" * 60)
