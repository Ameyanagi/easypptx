"""
Example demonstrating responsive positioning with different aspect ratios.

This example shows how EasyPPTX handles different aspect ratios
while maintaining properly centered content.
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
        subtitle="Demonstration of automatic adjustment for different aspect ratios"
    )
    
    # Add a slide with centered content
    centered_slide = pres.add_slide()
    
    # Add centered title with responsive positioning
    centered_slide.add_text(
        text="Centered Title",
        x="50%",  # 50% from left
        y="5%",
        width="80%",
        height="10%",
        font_size=32,
        font_bold=True,
        align="center",
        h_align="center"  # This enables responsive centering
    )
    
    # Add a large shape to demonstrate centering
    centered_slide.add_shape(
        shape_type=1,  # Rectangle
        x="10%",
        y="20%",
        width="80%",
        height="40%",
        fill_color="blue",
        h_align="center"  # This enables responsive centering
    )
    
    # Add explanatory text
    centered_slide.add_text(
        text=(
            "This slide demonstrates responsive positioning that adjusts "
            "automatically for different aspect ratios. The title and rectangle "
            "remain properly centered regardless of the presentation's aspect ratio."
        ),
        x="10%",
        y="70%",
        width="80%",
        height="20%",
        font_size=18,
        align="center",
        h_align="center"  # This enables responsive centering
    )
    
    # Add a comparison slide
    comparison_slide = pres.add_slide()
    comparison_slide.add_text(
        text="Comparison: Standard vs. Responsive",
        x="50%",
        y="5%",
        width="80%",
        height="10%",
        font_size=32,
        font_bold=True,
        align="center",
        h_align="center"
    )
    
    # Standard positioning (without responsive alignment)
    comparison_slide.add_shape(
        shape_type=1,  # Rectangle
        x="10%",
        y="20%",
        width="35%",
        height="30%",
        fill_color="red"
        # No h_align, so standard positioning
    )
    
    comparison_slide.add_text(
        text="Standard positioning\n(may shift left in wider aspect ratios)",
        x="10%",
        y="50%",
        width="35%",
        height="10%",
        font_size=14,
        align="center"
    )
    
    # Responsive positioning
    comparison_slide.add_shape(
        shape_type=1,  # Rectangle
        x="55%",
        y="20%",
        width="35%",
        height="30%",
        fill_color="green",
        h_align="center"  # Responsive centering
    )
    
    comparison_slide.add_text(
        text="Responsive positioning\n(maintains proper positioning)",
        x="55%",
        y="50%",
        width="35%",
        height="10%",
        font_size=14,
        align="center",
        h_align="center"  # Responsive centering
    )
    
    # Add implementation explanation
    comparison_slide.add_text(
        text=(
            "Implementation: When h_align=\"center\" is specified, the positioning code "
            "automatically adjusts x-coordinates based on the current aspect ratio. "
            "This ensures elements remain visually balanced regardless of slide dimensions."
        ),
        x="10%",
        y="70%",
        width="80%",
        height="20%",
        font_size=14,
        align="center",
        h_align="center"
    )
    
    # Save the presentation
    output_path = OUTPUT_DIR / f"responsive_positioning_{aspect_ratio.replace(':', '_')}.pptx"
    pres.save(output_path)
    print(f"Saved presentation to {output_path}")

print(
    "\nCreated presentations with different aspect ratios to demonstrate responsive positioning."
    "\nOpen and compare these files to see how content positioning adjusts automatically."
)