# Working with Plots and Charts

EasyPPTX provides several ways to add data visualizations to your presentations:

1. **Built-in PowerPoint charts** - Using `slide.add_chart` with tabular data
2. **Matplotlib figures** - Embedding matplotlib plots directly
3. **Seaborn plots** - Adding seaborn visualizations by passing their figure

> **Optional dependencies:** pandas support requires `pip install "easypptx[dataframe]"` and matplotlib support requires `pip install "easypptx[plot]"` (or install both with `"easypptx[all]"`). seaborn is not a dependency of EasyPPTX; seaborn plots work by passing their figure to `slide.add_pyplot` or `pres.add_pyplot_slide`.

Matplotlib figures are embedded via in-memory streams, so no temporary files are written.

## PowerPoint Native Charts

For simple charts using PowerPoint's built-in charting capabilities:

```python
from easypptx import Presentation
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
slide.add_chart(
    data=data,
    chart_type="column",
    x="10%", y="20%", width="80%", height="70%",
    category_column="Category",
    value_columns="Values",
    has_legend=True,
    title="Sample Chart"
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

`slide.add_chart` also accepts explicit `categories` and `values` lists instead of `data`. The aliases `value_column` (singular) and `chart_title` are accepted for `value_columns` and `title`.

### Multi-Series Charts

*New in 0.7.0.* Pass several column names to `value_columns` and each column is plotted as a named series:

```python
df = pd.DataFrame({
    "Quarter": ["Q1", "Q2", "Q3"],
    "Revenue": [100, 120, 140],
    "Cost": [80, 85, 90],
    "Profit": [20, 35, 50],
})

slide.add_chart(
    data=df,
    chart_type="line",
    category_column="Quarter",
    value_columns=["Revenue", "Cost", "Profit"],
)
```

The lower-level `Chart.add` accepts an explicit `series` mapping:

```python
from easypptx import Chart

Chart(slide).add(
    chart_type="column",
    categories=["Q1", "Q2", "Q3"],
    series={"Rev": [100, 120, 140], "Cost": [80, 85, 90]},
)
```

## Matplotlib Integration

You can embed matplotlib figures directly in your presentations:

```python
import matplotlib.pyplot as plt
from easypptx import Presentation

# Create a matplotlib figure
plt.figure(figsize=(10, 6))
plt.plot([1, 2, 3, 4], [1, 4, 9, 16])
plt.title('Sample Plot')
plt.grid(True)
plt.xlabel('x')
plt.ylabel('y')

# Create a presentation
pres = Presentation()

# Method 1: slide.add_pyplot (new in 0.7.0)
slide = pres.add_slide()
slide.add_pyplot(
    plt.gcf(),
    x=10,
    y=20,
    width=80,
    height=70,
    dpi=300,
    border=True,
    border_color="blue",
    shadow=True,
)

# Method 2: the add_pyplot_slide convenience method
slide, picture = pres.add_pyplot_slide(
    figure=plt.gcf(),
    title="Matplotlib Example",
    label="Figure 1: Sample Plot",
    dpi=300,
    border=True,
    border_color="blue",
    shadow=True
)
```

> **Removed in 0.7.0:** `pres.add_matplotlib_slide(...)` and `pres.add_plot(...)` have been removed. Use `slide.add_pyplot(...)` or `pres.add_pyplot_slide(...)` instead — see [Migrating to 0.7.0](migration.md).

## Seaborn Integration

seaborn plots are matplotlib figures under the hood, so pass their figure to `add_pyplot_slide`:

```python
import matplotlib.pyplot as plt
import seaborn as sns
from easypptx import Presentation

# Create a seaborn plot
tips = sns.load_dataset("tips")
ax = sns.barplot(x="day", y="total_bill", data=tips)
plt.title('Tips by Day')

# Create a presentation
pres = Presentation()

# Add the seaborn plot's figure
slide, picture = pres.add_pyplot_slide(
    figure=ax.get_figure(),
    title="Seaborn Example",
    label="Figure 1: Average Tips by Day",
    border=True,
    border_color="green",
    shadow=True
)
```

> **Removed in 0.7.0:** `pres.add_seaborn_slide(...)` has been removed. Use `pres.add_pyplot_slide(...)` as shown above.

## Unified API for Plots

Use `pres.add_pyplot_slide(...)` for matplotlib/seaborn figures, or `pres.add_chart_slide(...)` for PowerPoint charts:

```python
from easypptx import Presentation
import matplotlib.pyplot as plt
import pandas as pd

# Create a presentation
pres = Presentation()

# Matplotlib or seaborn figure
plt.figure()
plt.plot([1, 2, 3, 4], [1, 4, 9, 16])
plt.title('Matplotlib Plot')

slide1, picture = pres.add_pyplot_slide(
    figure=plt.gcf(),
    title="Matplotlib Plot",
    label="Figure 1: Sample Plot"
)

# PowerPoint chart from data
data = pd.DataFrame({"Category": ["A", "B", "C"], "Value": [1, 4, 2]})

slide2 = pres.add_chart_slide(
    title="PowerPoint Chart",
    data=data,
    chart_type="column",
    category_column="Category",
    value_columns="Value"
)
```

## Customizing Plot Appearance

You can customize how plots appear in your presentations:

```python
# Set custom styling for a matplotlib plot
slide, picture = pres.add_pyplot_slide(
    figure=plt.gcf(),
    title="Styled Plot",
    border=True,
    border_color="blue",
    shadow=True,
    maintain_aspect_ratio=True
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
slide.add_pyplot(fig1, x="5%", y="20%", width="42%", height="70%")

# Add PowerPoint chart to the right side
slide.add_chart(
    data=data,
    chart_type="line",
    category_column="Category",
    value_columns="Value",
    x="53%", y="20%", width="42%", height="70%"
)
```

## Examples

Check out these example files for more details:
- `examples/plot_example.py` - Demonstrates matplotlib and seaborn integration
- `examples/chart_example.py` - Shows how to create PowerPoint native charts
