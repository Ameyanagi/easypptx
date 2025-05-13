"""
enhanced_slide_methods_example.py - Demonstrates the enhanced slide creation methods

This example demonstrates the use of the enhanced slide creation methods:
- add_grid_slide
- add_pyplot_slide
- add_image_gen_slide
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

from easypptx import Presentation

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)


# Create a sample image for the example
def create_sample_image(filename):
    # Create a simple image using matplotlib
    plt.figure(figsize=(4, 3))
    plt.text(
        0.5,
        0.5,
        "Sample Image",
        fontsize=24,
        ha="center",
        va="center",
        bbox={"boxstyle": "round", "facecolor": "lightblue", "alpha": 0.8},
    )
    plt.axis("off")
    plt.tight_layout()
    plt.savefig(filename)
    plt.close()
    return filename


# Create sample matplotlib figures
def create_bar_chart():
    # Create a bar chart
    categories = ["Category A", "Category B", "Category C", "Category D"]
    values = [15, 30, 45, 10]

    fig, ax = plt.subplots(figsize=(8, 6))
    bars = ax.bar(categories, values, color=["skyblue", "lightgreen", "salmon", "lightgray"])

    # Add value labels on top of each bar
    for bar in bars:
        height = bar.get_height()
        ax.annotate(
            f"{height}",
            xy=(bar.get_x() + bar.get_width() / 2, height),
            xytext=(0, 3),  # 3 points vertical offset
            textcoords="offset points",
            ha="center",
            va="bottom",
        )

    ax.set_title("Sample Bar Chart")
    ax.set_ylabel("Values")
    ax.grid(axis="y", linestyle="--", alpha=0.7)

    return fig


def create_line_chart():
    # Create a line chart
    x = np.linspace(0, 10, 100)
    y1 = np.sin(x)
    y2 = np.cos(x)

    fig, ax = plt.subplots(figsize=(8, 6))
    ax.plot(x, y1, "b-", label="Sine")
    ax.plot(x, y2, "r-", label="Cosine")

    ax.set_title("Sine and Cosine Waves")
    ax.set_xlabel("X")
    ax.set_ylabel("Y")
    ax.legend()
    ax.grid(True)

    return fig


# Create a new presentation
pres = Presentation()

# Example 1: Create a slide with add_grid_slide
slide1, grid1 = pres.add_grid_slide(
    rows=2,
    cols=2,
    title="Enhanced Slide Methods",
    subtitle="Example of add_grid_slide",
    title_height="10%",
    subtitle_height="5%",
    padding=5.0,
)

# Add content to grid cells
grid1[0, 0].add_text(
    "Enhanced slide creation methods provide a consistent API for different content types.",
    font_size=14,
    align="center",
    vertical="middle",
)

grid1[0, 1].add_text(
    "Grid cells can be accessed using [row, col] syntax.",
    font_size=14,
    align="center",
    vertical="middle",
)

# Add content using row-level access
grid1[1].add_text(
    "Rows can be accessed using grid[row] syntax.",
    font_size=14,
    align="center",
    vertical="middle",
)

grid1[1].add_text(
    "Content flows automatically to the next cell in the row.",
    font_size=14,
    align="center",
    vertical="middle",
)

# Example 2: Create a slide with add_pyplot_slide
bar_chart = create_bar_chart()
slide2, plot2 = pres.add_pyplot_slide(
    figure=bar_chart,
    title="Data Visualization",
    subtitle="Using add_pyplot_slide method",
    label="Figure 1: Sample Bar Chart",
    dpi=150,
    border=True,
    shadow=True,
)

# Example 3: Create another slide with add_pyplot_slide (different chart)
line_chart = create_line_chart()
slide3, plot3 = pres.add_pyplot_slide(
    figure=line_chart,
    title="Line Chart Example",
    subtitle="Sine and Cosine Waves",
    dpi=150,
)

# Example 4: Create a slide with add_image_gen_slide
image_path = create_sample_image(output_dir / "sample_image.png")
slide4, image4 = pres.add_image_gen_slide(
    image_path=str(image_path),
    title="Image Example",
    subtitle="Using add_image_gen_slide method",
    label="Figure 3: Sample Image",
    border=True,
    shadow=True,
    maintain_aspect_ratio=True,
)

# Save the presentation
pres.save(output_dir / "enhanced_slide_methods_example.pptx")
print(f"Presentation saved to {output_dir / 'enhanced_slide_methods_example.pptx'}")
