"""
Quick start example for EasyPPTX.

This example demonstrates the basic usage of EasyPPTX to create a simple
presentation with various elements, using responsive positioning for centered content.
"""

from pathlib import Path

import pandas as pd

from easypptx import Chart, Presentation, Table

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# Slide 1: Title slide with responsive positioning
slide1 = pres.add_slide()
# Use percentage-based positioning with h_align for responsive centering
slide1.add_text(
    text="Getting Started with EasyPPTX",
    x="50%",
    y="30%",
    width="80%",
    height="15%",
    font_size=44,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)
slide1.add_text(
    text="A simple PowerPoint creation library for Python",
    x="50%",
    y="50%",
    width="60%",
    height="10%",
    font_size=24,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Slide 2: Text formatting examples with responsive centered title
slide2 = pres.add_slide()
# Centered title with responsive positioning
slide2.add_text(
    text="Text Formatting",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add different formatted text examples with percentage-based positioning
slide2.add_text(text="Regular text", x="10%", y="25%", width="80%", height="8%")
slide2.add_text(text="Bold text", x="10%", y="35%", width="80%", height="8%", font_bold=True)
slide2.add_text(text="Italic text", x="10%", y="45%", width="80%", height="8%", font_italic=True)
slide2.add_text(
    text="Bold and italic text", x="10%", y="55%", width="80%", height="8%", font_bold=True, font_italic=True
)
slide2.add_text(text="Colored text (Red)", x="10%", y="65%", width="80%", height="8%", color=(255, 0, 0))
slide2.add_text(text="Large text", x="10%", y="75%", width="80%", height="10%", font_size=28)

# Slide 3: Tables with responsive centered title
slide3 = pres.add_slide()
# Centered title with responsive positioning
slide3.add_text(
    text="Tables",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Create a simple table with percentage-based positioning
table_data = [
    ["Product", "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales"],
    ["Widgets", "$10,000", "$12,000", "$15,000", "$18,000"],
    ["Gadgets", "$8,000", "$9,500", "$11,000", "$12,500"],
    ["Doohickeys", "$6,500", "$7,000", "$7,500", "$8,000"],
]
table = Table(slide3)
table.add(table_data, x="10%", y="25%", width="80%", height="25%", first_row_header=True)

# Create a DataFrame and add it as a table with percentage-based positioning
df = pd.DataFrame({
    "Item": ["Product A", "Product B", "Product C"],
    "Revenue": [5000, 7500, 10000],
    "Cost": [2500, 4000, 5500],
    "Profit": [2500, 3500, 4500],
})
table.from_dataframe(df, x="10%", y="60%", width="80%", height="25%", first_row_header=True)

# Slide 4: Charts with responsive centered title
slide4 = pres.add_slide()
# Centered title with responsive positioning
slide4.add_text(
    text="Charts",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add charts with percentage-based positioning
chart = Chart(slide4)
categories = ["Jan", "Feb", "Mar", "Apr", "May"]
values = [10, 15, 13, 18, 20]

# Add a column chart on the left
chart.add_column(
    categories=categories, values=values, x="15%", y="25%", width="30%", height="30%", title="Monthly Sales"
)

# Add a pie chart on the right
categories = ["Product A", "Product B", "Product C", "Product D"]
values = [25, 30, 20, 25]
chart.add_pie(categories=categories, values=values, x="55%", y="25%", width="30%", height="30%", title="Market Share")

# Add a bar chart from DataFrame centered at the bottom
sales_data = pd.DataFrame({
    "Product": ["A", "B", "C", "D"],
    "Sales": [150, 200, 125, 175],
})
chart.from_dataframe(
    df=sales_data,
    chart_type="bar",
    category_column="Product",
    value_column="Sales",
    x="35%",
    y="60%",
    width="30%",
    height="30%",
    title="Sales by Product",
)

# Slide 5: Responsive positioning demonstration
slide5 = pres.add_slide()
# Centered title with responsive positioning
slide5.add_text(
    text="Responsive Positioning",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add explanation with responsive positioning
slide5.add_text(
    text=(
        "EasyPPTX supports responsive positioning that adapts to different aspect ratios. "
        'Elements with h_align="center" automatically adjust their position based on the '
        "aspect ratio, ensuring your content looks great in any format (16:9, 4:3, etc)."
    ),
    x="50%",
    y="25%",
    width="80%",
    height="15%",
    font_size=16,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add a centered shape
slide5.add_shape(
    shape_type=1,  # Rectangle
    x="50%",
    y="45%",
    width="50%",
    height="20%",
    fill_color="blue",
    h_align="center",  # Enable responsive positioning
)

# Add caption with responsive positioning
slide5.add_text(
    text="This shape stays centered in any aspect ratio",
    x="50%",
    y="70%",
    width="60%",
    height="10%",
    font_size=16,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Save the presentation
pres.save(output_dir / "quick_start.pptx")
print(f"Presentation saved to {output_dir / 'quick_start.pptx'}")
