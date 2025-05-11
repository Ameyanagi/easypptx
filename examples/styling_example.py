"""
Example demonstrating how to use styling options in EasyPPTX templates.

This example shows:
1. Custom image styling (borders, shadows, centering)
2. Custom table styling (headers, borders, alternating rows)
3. Custom chart styling (legends, colors, data labels)
4. Creating a complete presentation with consistent styling
"""

from pathlib import Path

import pandas as pd

from easypptx import Presentation

# Create a folder for outputs if it doesn't exist
output_dir = Path("outputs")
output_dir.mkdir(exist_ok=True)

# Sample data for tables and charts
sales_data = {
    "Product": ["Product A", "Product B", "Product C", "Product D"],
    "Q1 Sales": [120, 75, 90, 110],
    "Q2 Sales": [140, 85, 95, 120],
    "Q3 Sales": [135, 90, 105, 115],
    "Q4 Sales": [150, 100, 110, 130],
}

sales_df = pd.DataFrame(sales_data)

# Create a new presentation
pres = Presentation()

# 1. Add a title slide
title_slide = pres.add_title_slide(
    title="Styling Options in EasyPPTX", subtitle="Customizing images, tables, and charts"
)

# 2. Add a section slide for image styling
section_slide1 = pres.add_section_slide(title="Image Styling", bg_color="blue")

# 3. Add an image slide with custom styling
image_slide = pres.add_image_slide(
    title="Styled Image Example",
    image_path="examples/assets/sample_image.jpg",  # Replace with your image path
    label="Image with blue border and shadow effect",
    custom_style={
        "border": True,
        "border_color": "blue",
        "border_width": 2,
        "shadow": True,
        "maintain_aspect_ratio": True,
        "center": True,
    },
)

# 4. Add a section slide for table styling
section_slide2 = pres.add_section_slide(title="Table Styling", bg_color="green")

# 5. Add a table slide with custom styling
table_slide = pres.add_table_slide(
    title="Quarterly Sales Data",
    data=sales_df,
    has_header=True,
    custom_style={
        "first_row": {"bold": True, "bg_color": "blue", "text_color": "white"},
        "banded_rows": True,
        "band_color": "lightgray",
        "border_color": "black",
        "border_width": 1,
        "header_border_width": 2,
        "text_align": "center",
        "header_align": "center",
        "font_name": "Meiryo",
        "font_size": 12,
        "header_font_size": 14,
    },
)

# 6. Add a section slide for chart styling
section_slide3 = pres.add_section_slide(title="Chart Styling", bg_color="red")

# 7. Add a chart slide with custom styling
chart_slide = pres.add_chart_slide(
    title="Quarterly Sales Comparison",
    data=sales_df,
    chart_type="column",
    category_column="Product",
    value_columns=["Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales"],
    custom_style={
        "has_legend": True,
        "legend_position": "bottom",
        "has_title": True,
        "has_data_labels": True,
        "gridlines": True,
        "has_border": True,
        "border_color": "black",
        "palette": [
            (0x5B, 0x9B, 0xD5),  # Blue
            (0xED, 0x7D, 0x31),  # Orange
            (0xA5, 0xA5, 0xA5),  # Gray
            (0xFF, 0xC0, 0x00),  # Yellow
        ],
    },
)

# 8. Create a custom template with styling for all elements
custom_template = {
    "bg_color": "lightgray",
    "title": {
        "text": "Custom Template with Styling",
        "position": {"x": "5%", "y": "5%", "width": "90%", "height": "10%"},
        "font": {"name": "Meiryo", "size": 24, "bold": True},
        "align": "center",
        "vertical": "middle",
        "color": "darkblue",
    },
    "image_style": {
        "border": True,
        "border_color": "darkblue",
        "shadow": True,
        "maintain_aspect_ratio": True,
        "center": True,
    },
    "table_style": {
        "first_row": {"bold": True, "bg_color": "darkblue", "text_color": "white"},
        "banded_rows": True,
        "band_color": "lightblue",
    },
    "chart_style": {"chart_type": "pie", "has_legend": True, "legend_position": "right", "has_data_labels": True},
}

# 9. Add a slide using the custom template
custom_slide = pres.add_slide_from_template(custom_template)
custom_slide.add_text(
    text="This slide uses a comprehensive custom template with styling options for images, tables, and charts.",
    position={"x": "10%", "y": "20%", "width": "80%", "height": "70%"},
    font_size=16,
    align="center",
    vertical="middle",
)

# 10. Add a "thank you" slide
thank_you_slide = pres.add_slide_from_template("thank_you_slide")

# Save the presentation
pres.save(output_dir / "styling_example.pptx")
print(f"Presentation saved to {output_dir / 'styling_example.pptx'}")
