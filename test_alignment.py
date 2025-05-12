from easypptx import Presentation
from easypptx.text import Text

# Create a new presentation
presentation = Presentation()


# Debug function to print slide dimensions
def print_slide_dimensions(slide):
    """Print the slide dimensions for debugging purposes"""
    print("Slide dimensions:")
    print(f"  _slide_width: {slide._slide_width} EMUs ({slide._slide_width / 914400:.2f} inches)")
    print(f"  _slide_height: {slide._slide_height} EMUs ({slide._slide_height / 914400:.2f} inches)")

    # Try to access the actual presentation dimensions for comparison using multiple methods
    print("Trying to access actual presentation dimensions:")

    # Method 1
    try:
        if hasattr(slide.pptx_slide.part.package, "presentation"):
            pres = slide.pptx_slide.part.package.presentation
            print(f"  [Method 1] slide_width: {pres.slide_width} EMUs ({pres.slide_width / 914400:.2f} inches)")
            print(f"  [Method 1] slide_height: {pres.slide_height} EMUs ({pres.slide_height / 914400:.2f} inches)")
    except (AttributeError, TypeError) as e:
        print(f"  [Method 1] Error: {e}")

    # Method 2
    try:
        package = slide.pptx_slide.part.package
        if hasattr(package, "presentation_part") and hasattr(package.presentation_part, "presentation"):
            pres = package.presentation_part.presentation
            print(f"  [Method 2] slide_width: {pres.slide_width} EMUs ({pres.slide_width / 914400:.2f} inches)")
            print(f"  [Method 2] slide_height: {pres.slide_height} EMUs ({pres.slide_height / 914400:.2f} inches)")
    except (AttributeError, TypeError) as e:
        print(f"  [Method 2] Error: {e}")

    # Method 3
    try:
        parent = slide.pptx_slide.part.parent
        if parent and hasattr(parent, "presentation"):
            pres = parent.presentation
            print(f"  [Method 3] slide_width: {pres.slide_width} EMUs ({pres.slide_width / 914400:.2f} inches)")
            print(f"  [Method 3] slide_height: {pres.slide_height} EMUs ({pres.slide_height / 914400:.2f} inches)")
    except (AttributeError, TypeError) as e:
        print(f"  [Method 3] Error: {e}")


# Create a slide first and print its dimensions for debugging
test_slide = presentation.add_slide()
print_slide_dimensions(test_slide)

# Test Case 1: Basic text alignment using slide's add_text method
slide1 = presentation.add_slide()
slide1.add_text(
    text="Slide 1: This text should be center-aligned",
    x="0%",
    y="10%",
    width="100%",
    height="15%",
    font_size=32,
    font_bold=True,
    align="center",
)

slide1.add_text(
    text="This text should be left-aligned",
    x="0%",
    y="30%",
    width="100%",
    height="15%",
    font_size=32,
    font_bold=True,
    align="left",
)

slide1.add_text(
    text="This text should be right-aligned",
    x="0%",
    y="50%",
    width="100%",
    height="15%",
    font_size=32,
    font_bold=True,
    align="right",
)

# Test Case 2: Using the Text class static method
slide2 = presentation.add_slide()
Text.add(
    slide=slide2,
    text="Slide 2: Text.add method - center aligned",
    position={"x": "0%", "y": "10%", "width": "100%", "height": "15%"},
    font_size=32,
    font_bold=True,
    align="center",
)

Text.add(
    slide=slide2,
    text="Text.add method - left aligned",
    position={"x": "0%", "y": "30%", "width": "100%", "height": "15%"},
    font_size=32,
    font_bold=True,
    align="left",
)

Text.add(
    slide=slide2,
    text="Text.add method - right aligned",
    position={"x": "0%", "y": "50%", "width": "100%", "height": "15%"},
    font_size=32,
    font_bold=True,
    align="right",
)

# Test Case 3: Using the Presentation's add_text method
slide3 = presentation.add_slide()
presentation.add_text(
    slide=slide3,
    text="Slide 3: presentation.add_text - center aligned",
    x="0%",
    y="10%",
    width="100%",
    height="15%",
    font_size=32,
    font_bold=True,
    align="center",
)

presentation.add_text(
    slide=slide3,
    text="presentation.add_text - left aligned",
    x="0%",
    y="30%",
    width="100%",
    height="15%",
    font_size=32,
    font_bold=True,
    align="left",
)

presentation.add_text(
    slide=slide3,
    text="presentation.add_text - right aligned",
    x="0%",
    y="50%",
    width="100%",
    height="15%",
    font_size=32,
    font_bold=True,
    align="right",
)

# Test Case 4: Using Grid module
slide4 = presentation.add_slide()
grid = presentation.add_grid(
    slide=slide4,
    rows=3,
    cols=1,
    x="0%",
    y="10%",
    width="100%",
    height="80%",
)

grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide4.add_text,
    text="Slide 4: Grid text - center aligned",
    font_size=32,
    font_bold=True,
    align="center",
)

grid.add_to_cell(
    row=1,
    col=0,
    content_func=slide4.add_text,
    text="Grid text - left aligned",
    font_size=32,
    font_bold=True,
    align="left",
)

grid.add_to_cell(
    row=2,
    col=0,
    content_func=slide4.add_text,
    text="Grid text - right aligned",
    font_size=32,
    font_bold=True,
    align="right",
)

# Save the presentation
presentation.save("output/test_alignment.pptx")
print("Presentation saved as output/test_alignment.pptx")
