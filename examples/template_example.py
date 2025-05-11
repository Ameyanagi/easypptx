"""
Example demonstrating how to use templates in EasyPPTX.

This example shows:
1. Using built-in template presets (title_slide, content_slide, etc.)
2. Creating slides with the template methods
3. Customizing template presets
4. Creating a presentation with consistent styling
"""

from pathlib import Path

import pandas as pd

from easypptx import Presentation

# Create a folder for outputs if it doesn't exist
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# 1. Add a title slide
title_slide = pres.add_title_slide(
    title="EasyPPTX Template System", subtitle="Creating professional presentations with ease"
)

# 2. Add a content slide with a horizontal bar
content_slide = pres.add_content_slide(title="Template Features", use_bar=True)

# Add bullet points to the content slide
content_slide.add_text(
    text=(
        "• Pre-defined slide templates for common layouts\n"
        "• Consistent styling across slides\n"
        "• Automatic positioning with percentage-based coordinates\n"
        "• Smart image placement with auto-centering\n"
        "• Easy table creation with DataFrame support"
    ),
    x="10%",
    y="15%",
    width="80%",
    height="70%",
    font_size=18,
)

# 3. Add a section slide with blue background
section_slide = pres.add_section_slide(title="Image Examples", bg_color="blue")

# 4. Skip the image slide as we don't have sample images available
# Comment out the image slide code for now
# image_slide = pres.add_image_slide(
#     title="Image with Auto-Centering",
#     image_path="examples/assets/sample_image.jpg",  # Replace with your image path
#     label="Sample image with caption",
# )

# 5. Add a comparison slide with two columns
comparison_slide = pres.add_comparison_slide(
    title="Feature Comparison",
    content_texts=[
        "Traditional PowerPoint:\n• Manual positioning\n• Inconsistent styling\n• Tedious repetitive work",
        "EasyPPTX Templates:\n• Percentage-based positioning\n• Consistent preset styles\n• Quick slide creation",
    ],
)

# 6. Add a table slide with DataFrame data
sample_data = {
    "Feature": ["Templates", "Positioning", "Images", "Tables"],
    "Description": [
        "Pre-defined slide layouts",
        "Percentage-based coordinates",
        "Auto-centering and labels",
        "Styled with headers and borders",
    ],
    "Status": ["✓", "✓", "✓", "✓"],
}

df = pd.DataFrame(sample_data)

table_slide = pres.add_table_slide(
    title="Template Features",
    data=df,
    has_header=True,
    custom_style={
        "first_row": {"bold": True, "bg_color": "blue", "text_color": "white"},
        "banded_rows": True,
        "band_color": "lightgray",
    },
)

# 7. Create a custom template and use it
custom_template = {
    "bg_color": "orange",
    "title": {
        "text": "Custom Template",
        "position": {"x": "5%", "y": "5%", "width": "90%", "height": "15%"},
        "font": {"name": "Meiryo", "size": 36, "bold": True},
        "align": "center",
        "vertical": "middle",
        "color": "white",
    },
    "content_area": {"position": {"x": "10%", "y": "30%", "width": "80%", "height": "60%"}},
}

custom_slide = pres.add_slide_from_template(custom_template)
custom_slide.add_text(
    text="This slide uses a custom template with orange background",
    x="10%",
    y="30%",
    width="80%",
    height="60%",
    font_size=24,
    align="center",
    vertical="middle",
    color="white",
)

# 8. Add a "thank you" slide using section_slide template
thank_you_slide = pres.add_section_slide(title="Thank You!", bg_color="green")

# Save the presentation
pres.save(output_dir / "template_example.pptx")
print(f"Presentation saved to {output_dir / 'template_example.pptx'}")
