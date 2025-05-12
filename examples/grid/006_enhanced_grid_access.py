"""
006_enhanced_grid_access.py - Enhanced Grid Access API Example

This example demonstrates the enhanced Grid access API:
1. Using grid[row].add_xxx methods for row-based operations
2. Using grid[row, col].add_xxx methods for direct cell operations
3. Using add_autogrid_slide to create a slide with an empty grid and title
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

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

# Create a new presentation
pres = Presentation()

# --------------------------------------------------------------------------
# Example 1: Enhanced Grid Access API
# --------------------------------------------------------------------------
slide = pres.add_slide()
slide.add_text(
    text="006 - Enhanced Grid Access API",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
)

# Add a subtitle
slide.add_text(
    text="Using grid[row] and grid[row, col] accessors",
    x="50%",
    y="15%",
    width="90%",
    height="5%",
    font_size=20,
    align="center",
)

# Create a 2x2 grid
grid = pres.add_autogrid(
    slide=slide,
    content_funcs=None,  # Empty grid
    rows=2,
    cols=2,
    x="10%",
    y="25%",
    width="80%",
    height="65%",
    padding=8.0,
)

# Method 1: Using grid[row, col].add_text
grid[0, 0].add_text(
    text="This text was added using grid[0, 0].add_text()",
    font_size=14,
    font_bold=True,
    align="center",
    vertical="middle",
)

# Method 2: Using grid[row, col].add_image
grid[0, 1].add_image(
    image_path=str(logo_path),
)

# Method 3: Using grid[row].add_text for the entire row
# (in this case, it adds to the first available cell in row 1)
grid[1].add_text(
    text="This text was added using grid[1].add_text()",
    font_size=14,
    align="center",
    vertical="middle",
)

# Create sample data for a table
data = [
    ["Product", "Price", "Quantity"],
    ["Widget A", "$10.99", 25],
    ["Widget B", "$15.50", 10],
]

# Method 4: Using grid[row].add_table at index 1
# (in this case, it adds to the second cell in row 1)
grid[1].add_table(
    data=data,
    has_header=True,
)

# --------------------------------------------------------------------------
# Example 2: Creating a dashboard using enhanced access API
# --------------------------------------------------------------------------
dashboard_slide = pres.add_slide()
dashboard_slide.add_text(
    text="Dashboard Using Enhanced Grid Access API",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=28,
    font_bold=True,
    align="center",
)

# Create a 3x2 grid
dashboard = pres.add_autogrid(
    slide=dashboard_slide,
    content_funcs=None,  # Empty grid
    rows=3,
    cols=2,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    padding=5.0,
)

# Use grid[row] to add header that spans multiple cells
# Merge first row cells
dashboard.merge_cells(0, 0, 0, 1)
dashboard[0, 0].add_text(
    text="Q3 2025 Performance Summary",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)

# Add content to individual cells using [row, col] syntax
dashboard[1, 0].add_text(
    text="Key Metrics:",
    font_size=14,
    font_bold=True,
)

# Create metrics data for a table
metrics_data = [
    ["Metric", "Value", "Change"],
    ["Revenue", "$1.2M", "+15%"],
    ["Profit", "$450K", "+8%"],
    ["Customers", "8,500", "+22%"],
]

# Add table using [row, col] syntax
dashboard[1, 1].add_table(
    data=metrics_data,
    has_header=True,
)

# Create a simple plot
fig = plt.figure(figsize=(4, 3))
x = np.arange(0, 10, 0.1)
y = np.exp(-x / 2) * np.sin(2 * x)
plt.plot(x, y, "r-")
plt.title("Performance Trend")
plt.grid(True)

# Add plot to bottom-left cell
dashboard[2, 0].add_pyplot(
    figure=fig,
    dpi=150,
)

# Add summary to bottom-right using [row, col] syntax
dashboard[2, 1].add_text(
    text=(
        "Summary:\n\n"
        "• Growth exceeded projections in all key areas\n"
        "• Customer acquisition rate increased by 22%\n"
        "• New product line contributed 35% of revenue\n"
        "• International expansion on track for Q4"
    ),
    font_size=12,
)

# --------------------------------------------------------------------------
# Example 3: Using add_autogrid_slide to create a slide with an empty grid
# --------------------------------------------------------------------------
feature_slide, features_grid = pres.add_autogrid_slide(
    content_funcs=None,  # Empty grid
    rows=4,
    cols=2,
    title="Feature Comparison",
    title_height="15%",
    padding=5.0,
)

# Add feature items to each row without manually tracking columns
# For row 0, this will add to (0,0) then (0,1)
features_grid[0].add_textbox(
    text="Simplified Syntax",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

features_grid[0].add_textbox(
    text="Intuitive and concise API",
    font_size=18,
    align="center",
    vertical="middle",
)

# For row 1, this will add to (1,0) then (1,1)
features_grid[1].add_textbox(
    text="Automatic Placement",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

features_grid[1].add_textbox(
    text="Content flows naturally across rows",
    font_size=18,
    align="center",
    vertical="middle",
)

# For row 2, this will add to (2,0) then (2,1)
features_grid[2].add_textbox(
    text="Reduced Code",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

features_grid[2].add_textbox(
    text="Less typing, fewer parameters",
    font_size=18,
    align="center",
    vertical="middle",
)

# For row 3, this will add to (3,0) then (3,1)
features_grid[3].add_textbox(
    text="All Content Types",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
)

features_grid[3].add_textbox(
    text="Works with text, images, tables, charts",
    font_size=18,
    align="center",
    vertical="middle",
)

# Save the presentation
output_path = output_dir / "006_enhanced_grid_access.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
