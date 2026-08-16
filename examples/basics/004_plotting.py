"""
Example demonstrating how to add matplotlib (and seaborn) plots to PowerPoint presentations.

This example shows:
1. Adding a matplotlib figure to a slide with add_pyplot_slide
2. Adding a seaborn plot the same way (seaborn is optional)
3. Adding a native PowerPoint chart with add_chart_slide
4. Styling the plots with borders, shadows, etc.

Requires the plotting extra: pip install "easypptx[plot]"
Seaborn sections run only if seaborn is installed.
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from easypptx import Presentation, Pyplot
from easypptx.chart import Chart
from easypptx.text import Text

try:
    import seaborn as sns
except ImportError:
    sns = None

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
mpl_slide, _ = pres.add_pyplot_slide(
    figure=plt.gcf(),
    title="Matplotlib Example",
    label="Figure 1: Basic Trigonometric Functions",
    border=True,
    border_color="blue",
    shadow=True,
)

# 2. Create a seaborn plot (skipped if seaborn is not installed)
if sns is not None:
    plt.figure(figsize=(10, 6))
    np.random.seed(42)
    data = np.random.randn(100, 2)
    df = pd.DataFrame(data, columns=["A", "B"])
    df["C"] = np.abs(data[:, 0]) * 10
    df["D"] = ["Group 1" if i < 50 else "Group 2" for i in range(100)]

    sns_plot = sns.scatterplot(data=df, x="A", y="B", hue="D", size="C", sizes=(20, 200))
    plt.title("Seaborn Scatter Plot")

    # A seaborn plot is added through its figure, like any matplotlib figure
    sns_slide, _ = pres.add_pyplot_slide(
        figure=sns_plot.figure,
        title="Seaborn Example",
        label="Figure 2: Scatter Plot with Groups and Sizes",
        border=True,
        border_color="green",
        shadow=True,
    )

# 3. Add a native PowerPoint chart from a DataFrame
data = pd.DataFrame({
    "Quarter": ["Q1", "Q2", "Q3", "Q4"],
    "Revenue": [100, 120, 135, 150],
    "Expenses": [85, 90, 100, 110],
    "Profit": [15, 30, 35, 40],
})

chart_slide = pres.add_chart_slide(
    title="PowerPoint Native Chart",
    data=data,
    chart_type="column",
    category_column="Quarter",
    value_columns=["Revenue", "Expenses", "Profit"],
)

# 4. Create a combined slide with both a matplotlib chart and a pptx chart
plt.figure(figsize=(6, 6))
sizes = [35, 25, 20, 20]
labels = ["Product A", "Product B", "Product C", "Product D"]
colors = ["#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000"]
plt.pie(sizes, labels=labels, colors=colors, autopct="%1.1f%%", startangle=90)
plt.axis("equal")
plt.title("Market Share")

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
    font_size=14,
    font_bold=True,
    align="center",
)

# Add a PowerPoint chart to the right side
quarterly_data = pd.DataFrame({"Quarter": ["Q1", "Q2", "Q3", "Q4"], "Sales": [120, 150, 135, 180]})
chart_obj = Chart(comparison_slide)
chart_obj.add(
    chart_type="line",
    categories=quarterly_data["Quarter"].tolist(),
    values=quarterly_data["Sales"].tolist(),
    x="53%",
    y="20%",
    width="42%",
    height="70%",
    has_legend=False,
    title="Quarterly Sales",
)

# Add a label for the PowerPoint chart
Text.add(
    slide=comparison_slide,
    text="PowerPoint Native Chart",
    position={"x": "53%", "y": "15%", "width": "42%", "height": "5%"},
    font_size=14,
    font_bold=True,
    align="center",
)

# Save the presentation
pres.save(output_dir / "plot_example.pptx")
print(f"Presentation saved to {output_dir / 'plot_example.pptx'}")
