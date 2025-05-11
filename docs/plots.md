# Working with Plots and Charts

EasyPPTX provides several ways to add data visualizations to your presentations:

1. **Built-in PowerPoint charts** - Using the `Chart` class with tabular data
2. **Matplotlib figures** - Embedding matplotlib plots directly
3. **Seaborn plots** - Adding seaborn visualizations seamlessly

## PowerPoint Native Charts

For simple charts using PowerPoint's built-in charting capabilities:

```python
from easypptx import Presentation, Chart
import pandas as pd

# Create a presentation
pres = Presentation()
slide = pres.add_slide()

# Create sample data
data = pd.DataFrame({
    "Category": ["A", "B", "C", "D"],
    "Values": [10, 25, 15, 30]
})

# Add a chart
chart = Chart.add(
    slide=slide,
    data=data,
    chart_type="column",
    position={"x": "10%", "y": "20%", "width": "80%", "height": "70%"},
    category_column="Category",
    value_columns="Values",
    has_legend=True,
    has_title=True,
    chart_title="Sample Chart"
)

# You can also use the add_chart_slide convenience method
chart_slide = pres.add_chart_slide(
    title="Chart Example",
    data=data,
    chart_type="pie",
    category_column="Category",
    value_columns="Values",
    custom_style={
        "has_legend": True,
        "legend_position": "right",
        "has_data_labels": True
    }
)
```

## Matplotlib Integration

You can embed matplotlib figures directly in your presentations:

```python
import matplotlib.pyplot as plt
from easypptx import Presentation, Pyplot

# Create a matplotlib figure
plt.figure(figsize=(10, 6))
plt.plot([1, 2, 3, 4], [1, 4, 9, 16])
plt.title('Sample Plot')
plt.grid(True)
plt.xlabel('x')
plt.ylabel('y')

# Create a presentation
pres = Presentation()

# Method 1: Using the Pyplot class directly
slide = pres.add_slide()
Pyplot.add(
    slide=slide,
    figure=plt.gcf(),
    position={"x": "10%", "y": "20%", "width": "80%", "height": "70%"},
    dpi=300,
    style={
        "border": True,
        "border_color": "blue",
        "shadow": True
    }
)

# Method 2: Using the add_matplotlib_slide convenience method
slide = pres.add_matplotlib_slide(
    title="Matplotlib Example",
    figure=plt.gcf(),
    label="Figure 1: Sample Plot",
    dpi=300,
    custom_style={
        "border": True,
        "border_color": "blue",
        "shadow": True
    }
)
```

## Seaborn Integration

You can also add seaborn plots directly:

```python
import seaborn as sns
import pandas as pd
from easypptx import Presentation

# Create a seaborn plot
tips = sns.load_dataset("tips")
sns_plot = sns.barplot(x="day", y="total_bill", data=tips)
plt.title('Tips by Day')

# Create a presentation
pres = Presentation()

# Add a seaborn plot
slide = pres.add_seaborn_slide(
    title="Seaborn Example",
    seaborn_plot=sns_plot,
    label="Figure 1: Average Tips by Day",
    custom_style={
        "border": True,
        "border_color": "green",
        "shadow": True
    }
)
```

## Unified API for Plots

For convenience, EasyPPTX provides a unified `add_plot` method that works with different visualization types:

```python
from easypptx import Presentation
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd

# Create a presentation
pres = Presentation()

# Create a matplotlib plot
plt.figure()
plt.plot([1, 2, 3, 4], [1, 4, 9, 16])
plt.title('Matplotlib Plot')

# Add with the unified method
slide1 = pres.add_plot(
    title="Matplotlib Plot",
    plot=plt.gcf(),
    plot_type="matplotlib",
    label="Figure 1: Sample Plot"
)

# Create a seaborn plot
tips = sns.load_dataset("tips")
sns_plot = sns.barplot(x="day", y="total_bill", data=tips)

# Add with the unified method
slide2 = pres.add_plot(
    title="Seaborn Plot",
    plot=sns_plot,
    plot_type="seaborn",
    label="Figure 2: Tips by Day"
)

# Create data for a PowerPoint chart
data = pd.DataFrame({"Category": ["A", "B", "C"], "Value": [1, 4, 2]})

# Add with the unified method
slide3 = pres.add_plot(
    title="PowerPoint Chart",
    data=data,
    plot_type="pptx_chart",
    chart_type="column",
    category_column="Category",
    value_columns="Value"
)
```

## Customizing Plot Appearance

You can customize how plots appear in your presentations:

```python
# Set custom styling for a matplotlib plot
custom_style = {
    "border": True,
    "border_color": "blue",
    "border_width": 2,
    "shadow": True,
    "maintain_aspect_ratio": True,
    "center": True
}

slide = pres.add_matplotlib_slide(
    title="Styled Plot",
    figure=plt.gcf(),
    custom_style=custom_style
)
```

## Combining Multiple Plots

You can add multiple plots to a single slide:

```python
# Create a comparison slide
slide = pres.add_comparison_slide(
    title="Visualization Comparison",
    content_texts=["", ""]  # Empty placeholders
)

# Add matplotlib plot to the left side
Pyplot.add(
    slide=slide,
    figure=plt.figure1,
    position={"x": "5%", "y": "20%", "width": "42%", "height": "70%"}
)

# Add PowerPoint chart to the right side
Chart.add(
    slide=slide,
    data=data,
    chart_type="line",
    position={"x": "53%", "y": "20%", "width": "42%", "height": "70%"}
)
```

## Examples

Check out these example files for more details:
- `examples/plot_example.py` - Demonstrates matplotlib and seaborn integration
- `examples/chart_example.py` - Shows how to create PowerPoint native charts
