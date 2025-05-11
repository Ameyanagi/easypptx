"""
Basic demonstration of EasyPPTX functionality.

This example creates a simple presentation with text, an image, a table, and a chart.
"""

from pathlib import Path

import pandas as pd

from easypptx import Chart, Presentation, Table

# Create output directory if it doesn't exist
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# Add title slide
slide1 = pres.add_slide()
# Using responsive positioning for centered title and subtitle
slide1.add_text(
    text="EasyPPTX Demo",
    x="50%",
    y="30%",
    width="80%",
    height="15%",
    font_size=44,
    font_bold=True,
    align="center",  # Center text alignment
)
slide1.add_text(
    text="A simple demo of EasyPPTX capabilities",
    x="50%",
    y="50%",
    width="60%",
    height="10%",
    font_size=24,
    align="center",  # Center text alignment
)

# Add content slide with text formatting
slide2 = pres.add_slide()
# Centered title with responsive positioning
slide2.add_text(
    text="Text;lksdja;sd",
    x="0%",
    y="10%",
    width="100%",
    height="15%",
    font_size=32,
    font_bold=True,
    align="center",  # Center text alignment
)
# Regular text examples (left-aligned)
slide2.add_text(text="Regular text", x="10%", y="30%", width="80%", height="10%")
slide2.add_text(text="Bold text", x="10%", y="40%", width="80%", height="10%", font_bold=True)
slide2.add_text(text="Italic text", x="10%", y="50%", width="80%", height="10%", font_italic=True)
slide2.add_text(text="Colored text", x="10%", y="60%", width="80%", height="10%", color=(255, 0, 0))  # Red
slide2.add_text(text="Large text", x="10%", y="70%", width="80%", height="10%", font_size=32)

# Add slide with a table
slide3 = pres.add_slide()
# Centered title with responsive positioning
slide3.add_text(
    text="Table Example",
    x="50%",
    y="10%",
    width="80%",
    height="15%",
    font_size=32,
    font_bold=True,
    align="center",  # Center text alignment
)

# Add a centered table
data = [
    ["Product", "Q1", "Q2", "Q3", "Q4"],
    ["Widgets", 100, 150, 200, 180],
    ["Gadgets", 120, 130, 210, 240],
    ["Gizmos", 90, 110, 130, 150],
]
# Using percentage-based positioning for the table
table = Table(slide3)
table.add(data, x="10%", y="30%", width="80%", height="50%", first_row_header=True)

# Add slide with a chart
slide4 = pres.add_slide()
# Centered title with responsive positioning
slide4.add_text(
    text="Chart Example",
    x="50%",
    y="10%",
    width="80%",
    height="15%",
    font_size=32,
    font_bold=True,
    align="center",  # Center text alignment
)

chart = Chart(slide4)

# Create a pie chart and column chart with percentage-based positioning
categories = ["Category A", "Category B", "Category C", "Category D"]
values = [25, 35, 20, 20]

# Pie chart on the left
chart.add_pie(
    categories=categories,
    values=values,
    x="15%",
    y="30%",
    width="30%",
    height="50%",
    title="Sample Pie Chart",
)

# Column chart on the right
chart.add_column(
    categories=categories,
    values=values,
    x="55%",
    y="30%",
    width="30%",
    height="50%",
    title="Sample Column Chart",
)

# Save the presentation
pres.save(output_dir / "basic_demo.pptx")
print(f"Presentation saved to {output_dir / 'basic_demo.pptx'}")

# Example with DataFrame
slide5 = pres.add_slide()
# Centered title with responsive positioning
slide5.add_text(
    text="DataFrame Example",
    x="50%",
    y="10%",
    width="80%",
    height="15%",
    font_size=32,
    font_bold=True,
    align="center",  # Center text alignment
)

# Create a DataFrame
df = pd.DataFrame({
    "Category": ["Product A", "Product B", "Product C", "Product D"],
    "Sales": [1200, 1800, 1500, 2100],
    "Costs": [800, 1100, 950, 1400],
    "Profit": [400, 700, 550, 700],
})

# Add a table from DataFrame with percentage-based positioning
table = Table(slide5)
table.from_dataframe(df, x="10%", y="30%", width="80%", height="30%", first_row_header=True)

# Add a chart from DataFrame with percentage-based positioning
chart = Chart(slide5)
chart.from_dataframe(
    df,
    chart_type="bar",
    category_column="Category",
    value_column="Profit",
    x="10%",
    y="65%",
    width="80%",
    height="30%",
    title="Profit by Product",
)

# Save the updated presentation
pres.save(output_dir / "basic_demo.pptx")
print(f"Updated presentation saved to {output_dir / 'basic_demo.pptx'}")
