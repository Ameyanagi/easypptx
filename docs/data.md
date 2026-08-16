# Data to Slides

*New in 0.8.0.* `slide.add_table(data=...)` and `slide.add_chart(data=...)` accept the same set of data shapes, so whatever container your numbers already live in usually goes straight onto the slide. This page covers the accepted shapes, how chart rendering is routed between native PowerPoint charts and matplotlib, and the `df.pptx` pandas accessor.

## Accepted Data Shapes

Both `add_table` and `add_chart` normalize their `data` argument through one adapter (`easypptx.data`):

| Shape | Charts | Tables |
| --- | --- | --- |
| pandas `DataFrame` | First column is categories, second is values by default; override with `category_column=` / `value_columns=` | Column names become the header row |
| pandas `Series` | Index becomes the categories, the Series `name` labels the series | Two columns: index and values (header from the `name`) |
| polars `DataFrame` | Same column conventions as pandas | Column names become the header row |
| numpy 1D/2D array | Arrays carry no labels: pass `columns=` for series names (default `"Series N"`) and `categories=` for category labels | Pass `columns=` for header names (default `"Column N"`) |
| dict of sequences | Keys are series names: `{"Rev": [...], "Cost": [...]}`; pass `categories=` for category labels | Keys become the header row |
| list of lists | First row is the header; `category_column=` / `value_columns=` select columns by name or index | Passed through unchanged |

None of pandas, polars, or numpy are required dependencies — the adapter duck-types against modules that are already imported (which is always the case when you hold one of their objects).

For inputs that carry no category labels (numpy arrays, dicts of sequences), pass `categories=` alongside `data=` to label them — otherwise labels are positional (`"0"`, `"1"`, ...):

```python
slide.add_chart(
    data={"North": [4, 6, 5], "South": [3, 4, 6]},
    categories=["Jan", "Feb", "Mar"],
)
```

```python
import numpy as np
import pandas as pd
from easypptx import Presentation

pres = Presentation()
slide = pres.add_slide(title="One adapter, many shapes")

# dict of sequences: keys become the series names
slide.add_chart(
    data={"Revenue": [100, 120, 140], "Cost": [80, 85, 90]},
    chart_type="line",
    x=5, y=20, width=42, height=50,
)

# pandas Series: index -> categories, name -> series label
s = pd.Series([4, 8, 15], index=["East", "West", "North"], name="Stores")
slide.add_chart(data=s, chart_type="column", x=52, y=20, width=42, height=50)

# numpy 2D array with labeled columns as a table
arr = np.array([[1.0, 2.0], [3.0, 4.0]])
slide.add_table(data=arr, columns=["Alpha", "Beta"], x=5, y=75, width=40, height=18)
```

The normalization functions are also usable directly — see [`easypptx.data` in the API reference](api_reference.md#data-adapter-easypptxdata).

## Chart Backends: Native vs Matplotlib

`add_chart` routes each chart to one of two backends by its `chart_type`:

- **Native PowerPoint charts** — `"column"`, `"bar"`, `"line"`, `"pie"`, `"area"`, `"scatter"` — remain the default. They stay editable in PowerPoint and pick up the deck theme's palette automatically.
- **Matplotlib-rendered charts** — `"heatmap"`, `"histogram"`, `"box"`, `"violin"` — have no native PowerPoint equivalent, so they are automatically rendered with matplotlib into the same slide region and placed as a picture. This requires the optional plotting extra (`pip install "easypptx[plot]"`), and the call returns a picture shape instead of a chart object.

| Native backend | Matplotlib (`pyplot`) backend |
| --- | --- |
| Chart stays editable in PowerPoint | Static image |
| Colored by the deck theme's palette | Unlimited chart types (heatmap, histogram, box, violin, ...) |
| Small file footprint | Rendered at configurable DPI |

```python
# Automatically routed to matplotlib (no native heatmap exists)
slide.add_chart(
    data={"Mon": [1, 4, 2], "Tue": [3, 1, 5], "Wed": [2, 2, 2]},
    categories=["Morning", "Noon", "Evening"],
    chart_type="heatmap",
    x=10, y=20, width=80, height=70,
)
```

The `backend=` parameter overrides the routing:

- `backend="pyplot"` forces matplotlib rendering for any supported type, including the native ones (`column`, `bar`, `line`, `pie`, `area`, `scatter`).
- `backend="native"` with a non-native type raises a clear `ValueError` listing the native types, rather than silently rendering an image.

```python
# Force a matplotlib rendering of a line chart
slide.add_chart(data=df, chart_type="line", backend="pyplot")

# ValueError: Chart type 'violin' has no native PowerPoint equivalent. ...
slide.add_chart(data=df, chart_type="violin", backend="native")
```

Styling options split accordingly: `x_title`, `y_title`, `y_min`, `y_max`, `title`, `has_legend`, and `palette` apply on both backends, while `show_values`, `number_format`, and `legend_position` are native-only. See [Working with Plots and Charts](plots.md#native-chart-styling) for the full styling reference.

## The df.pptx Pandas Accessor

*New in 0.8.0.* When pandas is available, DataFrames grow a `.pptx` accessor that sends them straight to a slide:

```python
import pandas as pd
import easypptx

pres = easypptx.Presentation()
slide = pres.add_slide(title="Sales")

df = pd.DataFrame({"Region": ["East", "West"], "Sales": [120, 95]})
df.pptx.table(slide, y=20, height=30)
df.pptx.chart(slide, kind="column", y=55, height=40)
```

- `df.pptx.table(slide, **kwargs)` forwards to `slide.add_table(df, **kwargs)` — all table options (`number_format`, `shade_columns`, positioning, ...) work.
- `df.pptx.chart(slide, kind=..., **kwargs)` forwards to `slide.add_chart(data=df, chart_type=kind, **kwargs)`, so `kind` may be any native or matplotlib-backed chart type.

The accessor registers automatically when pandas is already imported at `import easypptx` time, and again whenever a `Presentation` is created. If you imported pandas later and the accessor is missing, register it manually:

```python
import easypptx
import pandas as pd  # imported after easypptx

easypptx.register_pandas_accessor()
```

`register_pandas_accessor()` returns `True` when the accessor is available (or already was) and `False` when pandas is not installed — it never raises just because pandas is absent.

## Related Pages

- [Working with Plots and Charts](plots.md) — chart styling, matplotlib figure embedding
- [Features Overview](features.md#creating-tables) — table formatting (`number_format`, `shade_columns`, styles)
- [API Reference](api_reference.md) — full parameter lists
