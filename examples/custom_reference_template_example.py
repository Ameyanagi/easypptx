#!/usr/bin/env python
"""
Example of using a custom reference PPTX template and custom blank layout.

This example demonstrates:
1. Loading a custom reference PPTX file
2. Specifying a blank layout index
3. Creating slides using the custom reference
"""

from pathlib import Path

from easypptx import Presentation

# Path to this example file
CURRENT_DIR = Path(__file__).parent


def main():
    """Run the example."""
    print("Creating a presentation with custom reference PPTX...")

    # Option 1: Specify a custom reference PPTX file and auto-detect blank layout
    presentation1 = Presentation(
        reference_pptx="/path/to/your/custom_reference.pptx",
    )

    # Add a slide with auto-detected blank layout
    slide1 = presentation1.add_slide()
    slide1.add_text(
        text="This slide uses the auto-detected blank layout",
        x="10%",
        y="40%",
        width="80%",
        height="20%",
        font_size=32,
        align="center",
    )

    # Save the presentation
    output_path1 = CURRENT_DIR / "output" / "custom_reference_auto.pptx"
    output_path1.parent.mkdir(exist_ok=True, parents=True)
    presentation1.save(output_path1)
    print(f"Saved: {output_path1}")

    # Option 2: Specify a custom reference PPTX file and specify blank layout index
    presentation2 = Presentation(
        reference_pptx="/path/to/your/custom_reference.pptx",
        blank_layout_index=2,  # Specify the layout index you want to use as blank
    )

    # Add a slide with specified blank layout
    slide2 = presentation2.add_slide()
    slide2.add_text(
        text="This slide uses the specified blank layout (index 2)",
        x="10%",
        y="40%",
        width="80%",
        height="20%",
        font_size=32,
        align="center",
    )

    # Save the presentation
    output_path2 = CURRENT_DIR / "output" / "custom_reference_specific.pptx"
    output_path2.parent.mkdir(exist_ok=True, parents=True)
    presentation2.save(output_path2)
    print(f"Saved: {output_path2}")

    # Option 3: Open an existing PPTX and specify a custom blank layout
    existing_pptx = "/path/to/existing/presentation.pptx"
    # Skip this part if you don't have an existing presentation
    if Path(existing_pptx).exists():
        presentation3 = Presentation.open(
            existing_pptx,
            blank_layout_index=3,  # Use the 4th layout as blank
        )

        # Add a slide with specified blank layout
        slide3 = presentation3.add_slide()
        slide3.add_text(
            text="This slide was added to an existing presentation",
            x="10%",
            y="40%",
            width="80%",
            height="20%",
            font_size=32,
            align="center",
        )

        # Save the presentation
        output_path3 = CURRENT_DIR / "output" / "existing_with_custom_blank.pptx"
        output_path3.parent.mkdir(exist_ok=True, parents=True)
        presentation3.save(output_path3)
        print(f"Saved: {output_path3}")


if __name__ == "__main__":
    main()
