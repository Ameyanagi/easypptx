"""
007_comprehensive_example.py - Comprehensive Example Using Enhanced Grid APIs

This example demonstrates a complete presentation creation workflow using the enhanced
Grid APIs introduced in v0.0.3. It showcases:
- Creating a presentation with multiple slides
- Using autogrid to create empty grids
- Using enhanced grid[row, col].add_xxx() syntax for specific cell access
- Using enhanced grid[row].add_xxx() syntax for sequential row operations
- Creating nested grids
- Adding various content types (text, images, tables, matplotlib plots)
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from easypptx import Presentation

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Ensure images directory exists
images_dir = output_dir / "images"
images_dir.mkdir(exist_ok=True)

# Create example logo image if it doesn't exist
logo_path = images_dir / "company_logo.png"
if not logo_path.exists():
    # Create a simple company logo using matplotlib
    fig, ax = plt.subplots(figsize=(3, 3))
    ax.set_aspect("equal")
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)

    # Draw a simple logo (blue circle with "ABC" text)
    circle = plt.Circle((5, 5), 4, fill=True, color="royalblue")
    ax.add_patch(circle)
    ax.text(
        5,
        5,
        "ABC",
        fontsize=24,
        color="white",
        horizontalalignment="center",
        verticalalignment="center",
        fontweight="bold",
    )

    ax.set_axis_off()
    plt.savefig(logo_path, bbox_inches="tight", transparent=True)
    plt.close()

# Create a new presentation
pres = Presentation()

# -------------------------------------------------------------------------
# Slide 1: Title Slide
# -------------------------------------------------------------------------
slide1 = pres.add_slide()

# Create a 2x1 grid for the title slide
title_grid = pres.add_autogrid(
    slide=slide1,
    content_funcs=None,  # Empty grid
    rows=2,
    cols=1,
    x="5%",
    y="10%",
    width="90%",
    height="80%",
    padding=5.0,
)

# Add a title and subtitle using the enhanced row-based API
title_grid[0].add_text(
    text="Enhanced Grid API Demo",
    font_size=44,
    font_bold=True,
    align="center",
    vertical="middle",
)

title_grid[1].add_text(
    text="A comprehensive example of the new grid access features in EasyPPTX v0.0.3",
    font_size=24,
    align="center",
    vertical="top",
)

# -------------------------------------------------------------------------
# Slide 2: Feature Overview
# -------------------------------------------------------------------------
slide2 = pres.add_slide()

# Add a title to the slide
slide2.add_text(
    text="Enhanced Grid Access Features",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Create a 2x2 grid to showcase features
feature_grid = pres.add_autogrid(
    slide=slide2,
    content_funcs=None,  # Empty grid
    rows=2,
    cols=2,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    padding=8.0,
)

# Use direct cell access with grid[row, col].add_xxx()
feature_grid[0, 0].add_text(
    text="grid[row, col].add_xxx() Syntax",
    font_size=20,
    font_bold=True,
    align="center",
    vertical="top",
)

feature_grid[0, 0].add_text(
    text="• Directly access specific cells\n• Intuitive and concise syntax\n• Similar to numpy/matplotlib indexing",
    font_size=16,
    align="left",
    vertical="middle",
)

# Use direct cell access to add an image
feature_grid[0, 1].add_text(
    text="Easy Content Addition",
    font_size=20,
    font_bold=True,
    align="center",
    vertical="top",
)

feature_grid[0, 1].add_image(
    image_path=str(logo_path),
)

# Use row-based access with grid[row].add_xxx()
feature_grid[1, 0].add_text(
    text="grid[row].add_xxx() Syntax",
    font_size=20,
    font_bold=True,
    align="center",
    vertical="top",
)

feature_grid[1, 0].add_text(
    text="• Automatically adds to next available cell\n• Perfect for sequential content\n• Simplifies row-based layouts",
    font_size=16,
    align="left",
    vertical="middle",
)

# Add data to demonstrate tables
data = [
    ["Feature", "Benefit"],
    ["Cell Proxy", "Direct access"],
    ["Row Proxy", "Sequential access"],
    ["Unified API", "Consistent experience"],
]

feature_grid[1, 1].add_text(
    text="Support for All Content Types",
    font_size=20,
    font_bold=True,
    align="center",
    vertical="top",
)

feature_grid[1, 1].add_table(
    data=data,
    has_header=True,
)

# -------------------------------------------------------------------------
# Slide 3: Data Dashboard Example
# -------------------------------------------------------------------------
slide3 = pres.add_slide()

slide3.add_text(
    text="Data Dashboard Example",
    x="50%",
    y="5%",
    width="90%",
    height="8%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Create a 3x3 grid for the dashboard
dashboard = pres.add_autogrid(
    slide=slide3,
    content_funcs=None,  # Empty grid
    rows=3,
    cols=3,
    x="5%",
    y="15%",
    width="90%",
    height="80%",
    padding=5.0,
)

# Add a header spanning the first row
dashboard.merge_cells(0, 0, 0, 2)
dashboard[0, 0].add_text(
    text="Quarterly Performance Summary (Q2 2025)",
    font_size=22,
    font_bold=True,
    align="center",
    vertical="middle",
)

# Create some sample data
sales_data = pd.DataFrame({
    "Product": ["Widget A", "Widget B", "Widget C", "Widget D"],
    "Q1 Sales": [120, 95, 140, 65],
    "Q2 Sales": [150, 110, 125, 90],
})

# Create visualization for the sales data
fig1 = plt.figure(figsize=(6, 4))
x = np.arange(len(sales_data["Product"]))
width = 0.35
plt.bar(x - width / 2, sales_data["Q1 Sales"], width, label="Q1")
plt.bar(x + width / 2, sales_data["Q2 Sales"], width, label="Q2")
plt.xticks(x, sales_data["Product"])
plt.ylabel("Sales")
plt.title("Quarterly Sales Comparison")
plt.legend()
plt.grid(True, alpha=0.3)
plt.tight_layout()

# Add the chart to the dashboard grid
dashboard[1, 0].add_pyplot(
    figure=fig1,
    dpi=150,
)

# Add the logo to the dashboard
dashboard[1, 1].add_image(
    image_path=str(logo_path),
)

# Create a pie chart for market share
fig2 = plt.figure(figsize=(4, 3))
market_share = [35, 25, 20, 15, 5]
labels = ["Company A", "Company B", "Company C", "Company D", "Others"]
plt.pie(market_share, labels=labels, autopct="%1.1f%%", startangle=90)
plt.title("Market Share")
plt.tight_layout()

# Add the pie chart
dashboard[1, 2].add_pyplot(
    figure=fig2,
    dpi=150,
)

# Add a data table to the dashboard
dashboard[2, 0].add_table(
    data=sales_data,
    has_header=True,
)

# Add a text summary
dashboard[2, 1].add_text(
    text="Q2 Summary:\n\n• Overall sales increased by 15% compared to Q1\n• Product A showed strongest growth at 25%\n• Market share increased from 32% to 35%",
    font_size=14,
)

# Create a nested grid in the bottom-right cell
nested_grid = dashboard[2, 2].add_grid(rows=2, cols=1)

# Add content to the nested grid using the enhanced API
nested_grid[0, 0].add_text(
    text="Key Metrics",
    font_size=16,
    font_bold=True,
    align="center",
)

metrics_data = [
    ["Revenue", "$1.2M", "+15%"],
    ["Profit", "$450K", "+8%"],
    ["Customers", "8,500", "+22%"],
]

nested_grid[1, 0].add_table(
    data=metrics_data,
    has_header=False,
)

# Save the presentation
output_path = output_dir / "007_comprehensive_example.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
