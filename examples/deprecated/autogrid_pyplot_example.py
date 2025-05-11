"""
Example demonstrating the autogrid feature with matplotlib plots.

This example shows how to use the Grid.autogrid method to automatically organize
multiple matplotlib plots in a PowerPoint slide.
"""

import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np

from easypptx import Presentation
from easypptx.grid import Grid

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# -----------------------------------------------------
# Simple example: 2x2 grid of plots
# -----------------------------------------------------
slide1 = pres.add_slide()

# Create 4 different matplotlib plots as content functions
# Create matplotlib figures for the grid
figures = []

# Figure 1: Sine Function
fig1 = plt.figure(figsize=(5, 4))
x = np.linspace(0, 10, 100)
y = np.sin(x)
plt.plot(x, y)
plt.title("Sine Function")
plt.grid(True)
figures.append(fig1)

# Figure 2: Cosine Function
fig2 = plt.figure(figsize=(5, 4))
x = np.linspace(0, 10, 100)
y = np.cos(x)
plt.plot(x, y, "r-")
plt.title("Cosine Function")
plt.grid(True)
figures.append(fig2)

# Figure 3: Quadratic Function
fig3 = plt.figure(figsize=(5, 4))
x = np.linspace(0, 10, 100)
y = x**2
plt.plot(x, y, "g-")
plt.title("Quadratic Function")
plt.grid(True)
figures.append(fig3)

# Figure 4: Exponential Function
fig4 = plt.figure(figsize=(5, 4))
x = np.linspace(0, 10, 100)
y = np.exp(x / 10)
plt.plot(x, y, "m-")
plt.title("Exponential Function")
plt.grid(True)
figures.append(fig4)

# Create the grid using the new autogrid_pyplot method
Grid.autogrid_pyplot(
    parent=slide1,
    figures=figures,
    rows=2,
    cols=2,
    title="Different Mathematical Functions",
)

# Save the presentation with just the first slide
output_path = output_dir / "autogrid_pyplot_example.pptx"
print(f"Trying to save to: {output_path.absolute()}")
pres.save(output_path)
print(f"Presentation saved to {output_path}")
# Exit after first slide for testing
sys.exit(0)

# -----------------------------------------------------
# Advanced example: Grid with 7 plots of different types
# -----------------------------------------------------
slide2 = pres.add_slide()


# Create 7 different types of matplotlib plots
def create_bar_chart():
    plt.figure(figsize=(4, 3))
    categories = ["A", "B", "C", "D", "E"]
    values = [23, 45, 56, 78, 32]
    plt.bar(categories, values)
    plt.title("Bar Chart")
    return pres.add_pyplot(slide2, plt.gcf())


def create_scatter_plot():
    plt.figure(figsize=(4, 3))
    x = np.random.rand(50)
    y = np.random.rand(50)
    plt.scatter(x, y)
    plt.title("Scatter Plot")
    plt.grid(True)
    return pres.add_pyplot(slide2, plt.gcf())


def create_histogram():
    plt.figure(figsize=(4, 3))
    data = np.random.randn(1000)
    plt.hist(data, bins=30)
    plt.title("Histogram")
    return pres.add_pyplot(slide2, plt.gcf())


def create_pie_chart():
    plt.figure(figsize=(4, 3))
    labels = ["A", "B", "C", "D"]
    sizes = [15, 30, 45, 10]
    plt.pie(sizes, labels=labels, autopct="%1.1f%%")
    plt.title("Pie Chart")
    return pres.add_pyplot(slide2, plt.gcf())


def create_line_plot():
    plt.figure(figsize=(4, 3))
    x = np.linspace(0, 10, 100)
    y1 = np.sin(x)
    y2 = np.cos(x)
    plt.plot(x, y1, "b-", label="Sin")
    plt.plot(x, y2, "r-", label="Cos")
    plt.legend()
    plt.title("Multiple Lines")
    plt.grid(True)
    return pres.add_pyplot(slide2, plt.gcf())


def create_area_plot():
    plt.figure(figsize=(4, 3))
    x = np.linspace(0, 10, 100)
    y1 = np.sin(x) + 1
    y2 = np.cos(x) + 1
    plt.fill_between(x, 0, y1, alpha=0.5, label="Sin")
    plt.fill_between(x, 0, y2, alpha=0.5, label="Cos")
    plt.legend()
    plt.title("Area Plot")
    return pres.add_pyplot(slide2, plt.gcf())


def create_heatmap():
    plt.figure(figsize=(4, 3))
    data = np.random.rand(10, 10)
    plt.imshow(data, cmap="viridis")
    plt.colorbar()
    plt.title("Heatmap")
    return pres.add_pyplot(slide2, plt.gcf())


# Create the grid and automatically place the plots
# Let autogrid determine the optimal rows and columns
content_funcs = [
    create_bar_chart,
    create_scatter_plot,
    create_histogram,
    create_pie_chart,
    create_line_plot,
    create_area_plot,
    create_heatmap,
]
Grid.autogrid(
    parent=slide2,
    content_funcs=content_funcs,
    title="Various Chart Types",
    padding=3.0,
)

# -----------------------------------------------------
# Example 3: Multiple grids on one slide
# -----------------------------------------------------
slide3 = pres.add_slide()

# Add a title to the slide
slide3.add_text(
    text="Dashboard with Multiple Chart Grids",
    x="50%",
    y="5%",
    width="90%",
    height="10%",
    font_size=28,
    font_bold=True,
    align="center",
    h_align="center",
)


# Create first grid (2 plots in top row)
def create_overview_chart1():
    plt.figure(figsize=(5, 2))
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
    sales = [20, 25, 30, 35, 40, 45]
    expenses = [18, 20, 25, 30, 35, 39]
    plt.plot(months, sales, "b-", label="Sales")
    plt.plot(months, expenses, "r-", label="Expenses")
    plt.legend()
    plt.title("Sales vs Expenses")
    plt.grid(True)
    return pres.add_pyplot(slide3, plt.gcf())


def create_overview_chart2():
    plt.figure(figsize=(5, 2))
    departments = ["HR", "Sales", "Eng", "Marketing", "Ops"]
    budget = [150, 300, 500, 250, 200]
    plt.bar(departments, budget)
    plt.title("Department Budgets ($k)")
    return pres.add_pyplot(slide3, plt.gcf())


overview_funcs = [create_overview_chart1, create_overview_chart2]
Grid.autogrid(
    parent=slide3,
    content_funcs=overview_funcs,
    rows=1,
    cols=2,
    x="5%",
    y="15%",
    width="90%",
    height="30%",
    title=None,
)


# Create second grid (3 plots in middle row)
def create_metric_chart1():
    plt.figure(figsize=(3, 2))
    labels = ["North", "South", "East", "West"]
    sizes = [35, 20, 25, 20]
    plt.pie(sizes, labels=labels, autopct="%1.1f%%")
    plt.title("Sales by Region")
    return pres.add_pyplot(slide3, plt.gcf())


def create_metric_chart2():
    plt.figure(figsize=(3, 2))
    x = np.linspace(0, 10, 50)
    y = 100 + 20 * np.sin(x) + np.random.randn(50) * 5
    plt.plot(x, y, "g-o")
    plt.title("Customer Satisfaction")
    plt.ylim(60, 140)
    return pres.add_pyplot(slide3, plt.gcf())


def create_metric_chart3():
    plt.figure(figsize=(3, 2))
    categories = ["Q1", "Q2", "Q3", "Q4"]
    values = [120, 150, 135, 180]
    plt.bar(categories, values)
    plt.title("Quarterly Revenue ($k)")
    return pres.add_pyplot(slide3, plt.gcf())


metric_funcs = [create_metric_chart1, create_metric_chart2, create_metric_chart3]
Grid.autogrid(
    parent=slide3,
    content_funcs=metric_funcs,
    rows=1,
    cols=3,
    x="5%",
    y="50%",
    width="90%",
    height="20%",
    title=None,
)


# Create third grid (4 small KPIs in bottom row)
def create_kpi1():
    plt.figure(figsize=(2, 2))
    plt.text(0.5, 0.5, "$1.2M", fontsize=24, ha="center")
    plt.axis("off")
    plt.title("Revenue", fontsize=14)
    return pres.add_pyplot(slide3, plt.gcf())


def create_kpi2():
    plt.figure(figsize=(2, 2))
    plt.text(0.5, 0.5, "18.5%", fontsize=24, ha="center", color="green")
    plt.axis("off")
    plt.title("Growth", fontsize=14)
    return pres.add_pyplot(slide3, plt.gcf())


def create_kpi3():
    plt.figure(figsize=(2, 2))
    plt.text(0.5, 0.5, "24.8%", fontsize=24, ha="center")
    plt.axis("off")
    plt.title("Margin", fontsize=14)
    return pres.add_pyplot(slide3, plt.gcf())


def create_kpi4():
    plt.figure(figsize=(2, 2))
    plt.text(0.5, 0.5, "87", fontsize=24, ha="center", color="blue")
    plt.axis("off")
    plt.title("NPS", fontsize=14)
    return pres.add_pyplot(slide3, plt.gcf())


kpi_funcs = [create_kpi1, create_kpi2, create_kpi3, create_kpi4]
Grid.autogrid(
    parent=slide3,
    content_funcs=kpi_funcs,
    rows=1,
    cols=4,
    x="5%",
    y="75%",
    width="90%",
    height="20%",
    title=None,
)

# -----------------------------------------------------
# Example 4: Specific row and column layout for 5 plots
# -----------------------------------------------------
slide4 = pres.add_slide()


# Create 5 different plots for a custom layout
def plot_a():
    plt.figure(figsize=(4, 4))
    x = np.linspace(-5, 5, 100)
    y = x**2
    plt.plot(x, y, "r-")
    plt.title("y = x²")
    plt.grid(True)
    plt.axis("equal")
    return pres.add_pyplot(slide4, plt.gcf())


def plot_b():
    plt.figure(figsize=(4, 4))
    x = np.linspace(-5, 5, 100)
    y = x**3
    plt.plot(x, y, "g-")
    plt.title("y = x³")
    plt.grid(True)
    return pres.add_pyplot(slide4, plt.gcf())


def plot_c():
    plt.figure(figsize=(4, 4))
    x = np.linspace(-5, 5, 100)
    y = np.sin(x)
    plt.plot(x, y, "b-")
    plt.title("y = sin(x)")
    plt.grid(True)
    return pres.add_pyplot(slide4, plt.gcf())


def plot_d():
    plt.figure(figsize=(4, 4))
    x = np.linspace(-5, 5, 100)
    y = np.exp(x)
    plt.plot(x, y, "m-")
    plt.title("y = eˣ")
    plt.grid(True)
    return pres.add_pyplot(slide4, plt.gcf())


def plot_e():
    plt.figure(figsize=(4, 4))
    x = np.linspace(-5, 5, 100)
    y = np.log(np.abs(x) + 0.1)
    plt.plot(x, y, "c-")
    plt.title("y = ln|x|")
    plt.grid(True)
    return pres.add_pyplot(slide4, plt.gcf())


# Create a 3x2 grid for 5 plots
# The last cell will be empty
content_funcs = [plot_a, plot_b, plot_c, plot_d, plot_e]
Grid.autogrid(
    parent=slide4,
    content_funcs=content_funcs,
    rows=3,
    cols=2,
    title="Common Mathematical Functions",
)

# Create just the first slide for testing
# Save the presentation
output_path = output_dir / "autogrid_pyplot_example.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")
print("Open the presentation to see the autogrid examples.")
