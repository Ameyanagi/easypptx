"""
Example demonstrating the title padding features in EasyPPTX.

This example shows how to control the positioning of titles and content
using padding parameters in different slide types.
"""

import matplotlib.pyplot as plt
import numpy as np

from easypptx import Presentation

# Create a new presentation
pres = Presentation()

# Add a title slide with standard 5% padding (default)
slide1 = pres.add_slide(title="Default Title Positioning", bg_color="lightgray")

# Add slide with custom title padding
slide2 = pres.add_slide(
    title="Custom Title Padding",
    title_padding="15%",  # Apply to both x and y
    bg_color="lightgray",
)

# Add slide with different x and y padding
slide3 = pres.add_slide(
    title="Different X and Y Padding", title_x_padding="20%", title_y_padding="10%", bg_color="lightgray"
)

# Add a grid slide with custom title and content padding
slide4, grid = pres.add_grid_slide(
    rows=2,
    cols=2,
    title="Grid With Custom Padding",
    title_padding="15%",
    content_padding="10%",  # Applies to both x and y
    bg_color="lightgray",
)

# Add content to the grid cells
grid[0, 0].add_text("Top Left", font_bold=True)
grid[0, 1].add_text("Top Right", font_bold=True)
grid[1, 0].add_text("Bottom Left", font_bold=True)
grid[1, 1].add_text("Bottom Right", font_bold=True)

# Create a matplotlib figure for pyplot example
plt.figure(figsize=(8, 6))
x = np.linspace(0, 10, 100)
plt.plot(x, np.sin(x))
plt.title("Sine Wave")
plt.grid(True)

# Add a pyplot slide with custom padding
slide5, pyplot = pres.add_pyplot_slide(
    figure=plt.gcf(),
    title="PyPlot With Custom Padding",
    title_padding="15%",
    content_padding="5%",
    label="Figure 1: Sine Wave",
    label_padding="3%",
    bg_color="lightgray",
)

# Add an image slide with custom padding
slide6, image = pres.add_image_gen_slide(
    image_path="../output/images/company_logo.png",  # Adjust path as needed
    title="Image With Custom Padding",
    title_padding="15%",
    content_padding="8%",
    label="Company Logo",
    label_padding="2%",
    bg_color="lightgray",
)

# Save the presentation
pres.save("title_padding_example.pptx")
print("Presentation saved as 'title_padding_example.pptx'")
