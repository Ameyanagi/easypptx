"""
005_enhanced_grid.py - Enhanced Grid Features Example

This example demonstrates the enhanced Grid features:
1. Using add_autogrid with None for content_funcs to create an empty grid
2. Using the Grid convenience methods (add_textbox, add_image, add_pyplot, add_table)
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

# Create images for example if they don't exist
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

# Create a sample graph
graph_path = images_dir / "graph_image.png"
if not graph_path.exists():
    # Create a simple graph using matplotlib
    x = np.linspace(0, 10, 100)
    y = np.sin(x)

    plt.figure(figsize=(6, 4))
    plt.plot(x, y, "b-", linewidth=2)
    plt.grid(True, alpha=0.3)
    plt.title("Sample Sine Wave")
    plt.xlabel("X Axis")
    plt.ylabel("Y Axis")

    plt.savefig(graph_path)
    plt.close()

# Create a new presentation
pres = Presentation()

# --------------------------------------------------------------------------
# Example 1: Empty Grid with add_autogrid(None)
# --------------------------------------------------------------------------
slide1 = pres.add_slide()
slide1.add_text(
    text="005 - Enhanced Grid Features",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Add a subtitle
slide1.add_text(
    text="Empty Grid with Convenience Methods",
    x="50%",
    y="15%",
    width="90%",
    height="5%",
    font_size=20,
    align="center",
)

# Create an empty 2x2 grid
grid = pres.add_autogrid(
    slide=slide1,
    content_funcs=None,  # Empty grid
    rows=2,
    cols=2,
    x="10%",
    y="25%",
    width="80%",
    height="65%",
    padding=8.0,
)

# Add content to specific cells using convenience methods
grid.add_textbox(
    row=0,
    col=0,
    text="This is a textbox added with grid.add_textbox()",
    font_size=14,
    font_bold=True,
    align="center",
    vertical="middle",
)

# Add an image to a cell
grid.add_image(
    row=0,
    col=1,
    image_path=str(logo_path),
)

# Create a matplotlib figure for the pyplot example
fig = plt.figure(figsize=(4, 3))
x = np.arange(0, 10, 0.1)
y = np.cos(x)
plt.plot(x, y, "r-")
plt.title("Cosine Wave")
plt.grid(True)

# Add the matplotlib figure to a cell
grid.add_pyplot(
    row=1,
    col=0,
    figure=fig,
    dpi=150,
)

# Create sample data for a table
data = [
    ["Product", "Price", "Quantity"],
    ["Widget A", "$10.99", 25],
    ["Widget B", "$15.50", 10],
    ["Widget C", "$5.99", 50],
]

# Add a table to a cell
grid.add_table(
    row=1,
    col=1,
    data=data,
    has_header=True,
)

# --------------------------------------------------------------------------
# Example 2: Multi-purpose dashboard with empty grid
# --------------------------------------------------------------------------
slide2 = pres.add_slide()
slide2.add_text(
    text="Data Dashboard Example",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=28,
    font_bold=True,
    align="center",
)

# Create a 3x3 grid with title
dashboard_grid = pres.add_autogrid(
    slide=slide2,
    content_funcs=None,  # Empty grid
    rows=3,
    cols=3,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    padding=5.0,
    title="Quarterly Performance Dashboard",
    title_height="8%",
)

# Add a large header that spans first row
header_cell = dashboard_grid.merge_cells(0, 0, 0, 2)  # Merge cells in the first row
dashboard_grid.add_textbox(
    row=0,
    col=0,
    text="Company Performance Overview - Q2 2025",
    font_size=18,
    font_bold=True,
    align="center",
)

# Create a pandas DataFrame for sales data
sales_data = pd.DataFrame({
    "Category": ["Product A", "Product B", "Product C", "Product D"],
    "Q1 Sales": [120, 95, 140, 65],
    "Q2 Sales": [150, 110, 125, 90],
})

# Create a bar chart for the sales comparison
fig2 = plt.figure(figsize=(5, 3))
x = np.arange(len(sales_data["Category"]))
width = 0.35
plt.bar(x - width / 2, sales_data["Q1 Sales"], width, label="Q1")
plt.bar(x + width / 2, sales_data["Q2 Sales"], width, label="Q2")
plt.xticks(x, sales_data["Category"])
plt.ylabel("Sales")
plt.title("Quarterly Sales Comparison")
plt.legend()
plt.grid(True, alpha=0.3)
plt.tight_layout()

# Add the sales chart
dashboard_grid.add_pyplot(
    row=1,
    col=0,
    figure=fig2,
    dpi=150,
)

# Add the raw sales data table
dashboard_grid.add_table(
    row=1,
    col=1,
    data=sales_data,
    has_header=True,
)

# Add the company logo
dashboard_grid.add_image(
    row=1,
    col=2,
    image_path=str(logo_path),
)

# Create a pie chart for market share
fig3 = plt.figure(figsize=(4, 3))
market_share = [35, 25, 20, 15, 5]
labels = ["Company A", "Company B", "Company C", "Company D", "Others"]
plt.pie(market_share, labels=labels, autopct="%1.1f%%", startangle=90)
plt.title("Market Share")

# Add the market share pie chart
dashboard_grid.add_pyplot(
    row=2,
    col=0,
    figure=fig3,
    dpi=150,
)

# Add a summary text box
dashboard_grid.add_textbox(
    row=2,
    col=1,
    text=(
        "Q2 2025 Summary:\n\n"
        "• Overall sales increased by 15% compared to Q1\n"
        "• Product A showed the strongest growth at 25%\n"
        "• Market share increased from 32% to 35%\n"
        "• New marketing campaign resulted in 20% more leads"
    ),
    font_size=12,
)

# Add the graph image
dashboard_grid.add_image(
    row=2,
    col=2,
    image_path=str(graph_path),
)

# Save the presentation
output_path = output_dir / "005_enhanced_grid.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
