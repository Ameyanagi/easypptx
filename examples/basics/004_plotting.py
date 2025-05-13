"""
Example demonstrating how to add matplotlib and seaborn plots to PowerPoint presentations.

This example shows:
1. Adding a matplotlib figure directly to a slide
2. Adding a seaborn plot to a slide
3. Using the unified add_plot method for different plot types
4. Styling the plots with borders, shadows, etc.
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

from easypptx import Presentation, Pyplot
from easypptx.chart import Chart
from easypptx.text import Text

# Create a folder for outputs if it doesn't exist
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# Add a title slide
title_slide = pres.add_title_slide(
    title="Matplotlib and Seaborn Integration", subtitle="Adding data visualizations to PowerPoint presentations"
)

# 1. Create a basic matplotlib figure
plt.figure(figsize=(10, 6))
x = np.linspace(0, 10, 100)
plt.plot(x, np.sin(x), label="sin(x)")
plt.plot(x, np.cos(x), label="cos(x)")
plt.plot(x, np.exp(-x / 5) * np.sin(x), label="damped sin(x)")
plt.title("Basic Trigonometric Functions")
plt.xlabel("x")
plt.ylabel("y")
plt.grid(True)
plt.legend()

# Add the matplotlib figure to a slide
mpl_slide = pres.add_matplotlib_slide(
    title="Matplotlib Example",
    figure=plt.gcf(),
    label="Figure 1: Basic Trigonometric Functions",
    custom_style={"border": True, "border_color": "blue", "shadow": True},
)

# 2. Create a seaborn plot
plt.figure(figsize=(10, 6))
# Generate some sample data
np.random.seed(42)
data = np.random.randn(100, 2)
df = pd.DataFrame(data, columns=["A", "B"])
df["C"] = np.abs(data[:, 0]) * 10
df["D"] = ["Group 1" if i < 50 else "Group 2" for i in range(100)]

# Create a seaborn scatter plot
sns_plot = sns.scatterplot(data=df, x="A", y="B", hue="D", size="C", sizes=(20, 200))
plt.title("Seaborn Scatter Plot")

# Add the seaborn plot to a slide
sns_slide = pres.add_seaborn_slide(
    title="Seaborn Example",
    seaborn_plot=sns_plot,
    label="Figure 2: Scatter Plot with Groups and Sizes",
    custom_style={"border": True, "border_color": "green", "shadow": True},
)

# 3. Create a more complex seaborn visualization
plt.figure(figsize=(10, 8))
# Generate sample data
tips = pd.DataFrame({
    "total_bill": np.random.uniform(10, 50, 200),
    "tip": np.random.uniform(1, 10, 200),
    "sex": np.random.choice(["Male", "Female"], 200),
    "smoker": np.random.choice(["Yes", "No"], 200),
    "day": np.random.choice(["Sun", "Mon", "Tue", "Wed", "Thur", "Fri", "Sat"], 200),
    "time": np.random.choice(["Lunch", "Dinner"], 200),
    "size": np.random.choice([1, 2, 3, 4, 5, 6], 200),
})

# Create a facetgrid with multiple plots
g = sns.FacetGrid(tips, col="time", row="sex", margin_titles=True, height=3)
g.map_dataframe(sns.scatterplot, x="total_bill", y="tip", hue="day")
g.add_legend()
g.fig.suptitle("Tips by Gender and Time", y=1.05)

# Use the unified add_plot method
facet_slide = pres.add_plot(
    title="FacetGrid Example",
    plot=g,
    plot_type="matplotlib",
    label="Figure 3: Tips by Gender and Time of Day",
    dpi=300,
    custom_style={
        "border": True,
        "border_color": "red",
        "border_width": 2,
        "shadow": True,
        "maintain_aspect_ratio": True,
    },
)

# 4. Create a pandas DataFrame and use a pptx chart
data = pd.DataFrame({
    "Quarter": ["Q1", "Q2", "Q3", "Q4"],
    "Revenue": [100, 120, 135, 150],
    "Expenses": [85, 90, 100, 110],
    "Profit": [15, 30, 35, 40],
})

# Add a PowerPoint chart using the unified add_plot method
chart_slide = pres.add_plot(
    title="PowerPoint Native Chart",
    data=data,
    plot_type="pptx_chart",
    chart_type="column",
    category_column="Quarter",
    value_columns=["Revenue", "Expenses", "Profit"],
    custom_style={
        "has_legend": True,
        "legend_position": "bottom",
        "has_data_labels": True,
        "gridlines": True,
        # No custom palette for now to avoid RGBColor issues
        # "palette": [
        #     (0x5B, 0x9B, 0xD5),  # Blue
        #     (0xED, 0x7D, 0x31),  # Orange
        #     (0x70, 0xAD, 0x47),  # Green
        # ],
    },
)

# 5. Create a combined slide with both matplotlib and pptx chart
# First create a matplotlib pie chart
plt.figure(figsize=(6, 6))
sizes = [35, 25, 20, 20]
labels = ["Product A", "Product B", "Product C", "Product D"]
colors = ["#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000"]
plt.pie(sizes, labels=labels, colors=colors, autopct="%1.1f%%", startangle=90)
plt.axis("equal")
plt.title("Market Share")

# Create a slide with two content areas
comparison_slide = pres.add_comparison_slide(
    title="Data Visualization Comparison",
    content_texts=["", ""],  # Empty placeholders for now
)

# Add matplotlib pie chart to the left side
Pyplot.add(
    slide=comparison_slide,
    figure=plt.gcf(),
    position={"x": "5%", "y": "20%", "width": "42%", "height": "70%"},
    dpi=300,
    style={"border": True, "border_color": "blue", "shadow": True},
)

# Add a label for the pie chart
Text.add(
    slide=comparison_slide,
    text="Matplotlib Pie Chart",
    position={"x": "5%", "y": "15%", "width": "42%", "height": "5%"},
    font_name="Meiryo",
    font_size=14,
    font_bold=True,
    align="center",
)

# Create a small dataset for a PowerPoint chart on the right side
quarterly_data = pd.DataFrame({"Quarter": ["Q1", "Q2", "Q3", "Q4"], "Sales": [120, 150, 135, 180]})

# Add a PowerPoint chart to the right side
# First create a Chart instance
chart_obj = Chart(comparison_slide)

# Get categories and values from DataFrame
categories = quarterly_data["Quarter"].tolist()
values = quarterly_data["Sales"].tolist()

# Now call the instance method
chart_obj.add(
    chart_type="line",
    categories=categories,
    values=values,
    x="53%",
    y="20%",
    width="42%",
    height="70%",
    has_legend=False,
    chart_title="Quarterly Sales",
)

# Add a label for the PowerPoint chart
Text.add(
    slide=comparison_slide,
    text="PowerPoint Native Chart",
    position={"x": "53%", "y": "15%", "width": "42%", "height": "5%"},
    font_name="Meiryo",
    font_size=14,
    font_bold=True,
    align="center",
)

# Save the presentation
pres.save(output_dir / "plot_example.pptx")
print(f"Presentation saved to {output_dir / 'plot_example.pptx'}")
