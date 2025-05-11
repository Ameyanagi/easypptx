"""
Example demonstrating how to use the enhanced template management features.

This example shows:
1. Creating and registering custom templates
2. Saving templates to and loading templates from files
3. Using custom templates in presentations
4. Listing available templates
"""

from pathlib import Path

import pandas as pd

from easypptx import Presentation, TemplateManager
from easypptx.image import Image
from easypptx.table import Table
from easypptx.text import Text

# Create a folder for outputs if it doesn't exist
output_dir = Path("outputs")
output_dir.mkdir(exist_ok=True)

# Create a template directory for this example
template_dir = output_dir / "templates"
template_dir.mkdir(exist_ok=True)

# 1. Create a TemplateManager with a custom template directory
tm = TemplateManager(template_dir=str(template_dir))

# 2. Create a custom template for a product slide
product_template = {
    "bg_color": "white",
    "title": {
        "text": "Product Overview",
        "position": {"x": "5%", "y": "5%", "width": "90%", "height": "8%"},
        "font": {"name": "Meiryo", "size": 32, "bold": True},
        "align": "left",
        "vertical": "middle",
        "color": "darkblue",
    },
    "image_area": {"position": {"x": "5%", "y": "15%", "width": "45%", "height": "60%"}},
    "image_style": {
        "border": True,
        "border_color": "darkblue",
        "border_width": 2,
        "shadow": True,
        "maintain_aspect_ratio": True,
        "center": True,
    },
    "content_area": {"position": {"x": "55%", "y": "15%", "width": "40%", "height": "60%"}},
    "bar": {
        "position": {"x": "5%", "y": "80%", "width": "90%", "height": "1%"},
        "gradient": {
            "start_color": (0x00, 0x70, 0xC0),  # Blue
            "end_color": (0x00, 0xB0, 0x50),  # Green
            "angle": 0,
        },
    },
}

# 3. Register the product template
tm.register("product_slide", product_template)

# 4. Save the template to a file
template_file = tm.save("product_slide")
print(f"Saved template to: {template_file}")

# 5. Create another template for a feature comparison slide
comparison_template = {
    "bg_color": "white",
    "title": {
        "text": "Feature Comparison",
        "position": {"x": "5%", "y": "5%", "width": "90%", "height": "8%"},
        "font": {"name": "Meiryo", "size": 32, "bold": True},
        "align": "center",
        "vertical": "middle",
        "color": "darkblue",
    },
    "table_area": {"position": {"x": "5%", "y": "15%", "width": "90%", "height": "75%"}},
    "table_style": {
        "first_row": {"bold": True, "bg_color": "darkblue", "text_color": "white"},
        "first_column": {"bold": True, "bg_color": "lightgray"},
        "banded_rows": True,
        "band_color": "lightgray",
        "border_color": "darkblue",
        "border_width": 1,
        "header_border_width": 2,
    },
}

# 6. Register and save the comparison template
tm.register("comparison_slide", comparison_template)
tm.save("comparison_slide")

# 7. List all available templates
print("Available templates:")
for template_name in tm.list_templates():
    print(f"- {template_name}")

# 8. Create a presentation using the custom templates
pres = Presentation()

# 8a. Use the template manager from the presentation
pres.template_manager = tm

# 8b. Add a title slide (using a built-in template)
title_slide = pres.add_title_slide(title="Template Manager Example", subtitle="Creating and using custom templates")

# 8c. Add a product slide using our custom template
product_slide = pres.add_slide_from_template("product_slide")

# Add content to the product slide

# Add a product image
Image.add(
    slide=product_slide,
    image_path="examples/assets/sample_image.jpg",  # Replace with your image path
    position=product_template["image_area"]["position"],
    maintain_aspect_ratio=True,
    center=True,
)

# Add product details
Text.add(
    slide=product_slide,
    text=(
        "Product Highlights:\n\n"
        "• Feature 1: Description of feature 1\n"
        "• Feature 2: Description of feature 2\n"
        "• Feature 3: Description of feature 3\n\n"
        "Available in multiple configurations"
    ),
    position=product_template["content_area"]["position"],
    font_name="Meiryo",
    font_size=16,
    align="left",
    vertical_align="top",
)

# 8d. Add a comparison slide using our custom template
comparison_slide = pres.add_slide_from_template("comparison_slide")

# Add a comparison table

# Create sample comparison data
data = pd.DataFrame({
    "Feature": ["Feature 1", "Feature 2", "Feature 3", "Feature 4"],
    "Product A": ["Yes", "Basic", "No", "Premium"],
    "Product B": ["No", "Advanced", "Yes", "Basic"],
    "Product C": ["Yes", "Premium", "Yes", "Premium"],
})

# Add the table
Table.add(
    slide=comparison_slide,
    data=data,
    position=comparison_template["table_area"]["position"],
    has_header=True,
    style=comparison_template["table_style"],
)

# 9. Create another presentation directly using the TemplateManager
pres2 = Presentation()
pres2.template_manager = tm

# 9a. Add a section slide using a built-in template
section_slide = pres2.add_section_slide(title="Direct Template Manager Integration", bg_color="darkblue")

# 9b. Add a slide using a custom template
product_slide2 = pres2.add_slide_from_template("product_slide")
Text.add(
    slide=product_slide2,
    text="This presentation directly uses the custom template from the TemplateManager",
    position={"x": "10%", "y": "40%", "width": "80%", "height": "20%"},
    font_size=20,
    align="center",
    vertical_align="middle",
)

# Save the presentations
pres.save(output_dir / "template_manager_example1.pptx")
pres2.save(output_dir / "template_manager_example2.pptx")

print(f"Presentations saved to: {output_dir}")

# 10. Demonstrate loading a template from a file
new_tm = TemplateManager()
loaded_template_name = new_tm.load(template_file)
print(f"Loaded template '{loaded_template_name}' from file")

# Create a presentation using the loaded template
pres3 = Presentation()
pres3.template_manager = new_tm
pres3.add_title_slide("Loaded Template Example", "Using a template loaded from a file")
pres3.add_slide_from_template(loaded_template_name)
pres3.save(output_dir / "template_loaded_example.pptx")
print(f"Created presentation with loaded template: {output_dir / 'template_loaded_example.pptx'}")
