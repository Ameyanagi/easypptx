"""
Data-to-slides example: charts and tables from DataFrames, arrays, and dicts.

Shows the 0.8.0 data workflow:
1. One adapter: pandas / polars / numpy / dict / list inputs everywhere
2. Native, editable charts with theme colors, data labels, and axis options
3. Auto-routing to matplotlib for chart types PowerPoint can't draw
4. Formatted and value-shaded tables
5. The df.pptx pandas accessor
"""

from pathlib import Path

import numpy as np
import pandas as pd

from easypptx import Presentation

output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

pres = Presentation(theme="corporate")

# --- Native chart with styling, straight from a DataFrame -------------------
df = pd.DataFrame({
    "Quarter": ["Q1", "Q2", "Q3", "Q4"],
    "Revenue": [1200, 1350, 1480, 1610],
    "Expenses": [900, 940, 1010, 1080],
})

slide = pres.add_slide(title="Financials (native, editable)")
slide.add_chart(
    data=df,
    category_column="Quarter",
    value_columns=["Revenue", "Expenses"],
    show_values=True,
    number_format="#,##0",
    y_title="USD (k)",
    y_min=0,
    y=18,
    height=75,
)

# --- Auto-routed matplotlib chart for a non-native type ---------------------
slide = pres.add_slide(title="Weekly load (matplotlib heatmap)")
matrix = np.random.default_rng(7).random((4, 7))
slide.add_chart(
    data=matrix,
    chart_type="heatmap",  # not a native PowerPoint type -> rendered image
    columns=["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
    y=18,
    height=75,
)

# --- Formatted, value-shaded table ------------------------------------------
slide = pres.add_slide(title="Region performance")
slide.add_table(
    df,
    number_format={"Revenue": "{:,.0f}", "Expenses": "{:,.0f}"},
    shade_columns=["Revenue"],
    shade_color="blue",
    y=20,
    width=60,
    height=50,
)

# --- The pandas accessor ----------------------------------------------------
slide = pres.add_slide(title="df.pptx accessor")
df.pptx.table(slide, y=18, height=30)
df.pptx.chart(slide, kind="line", y=52, height=42, y_title="USD (k)")

# --- A dict works too (no pandas needed) ------------------------------------
slide = pres.add_slide(title="Charts from plain dicts")
slide.add_chart(
    data={"North": [4, 6, 5], "South": [3, 4, 6]},
    categories=["Jan", "Feb", "Mar"],
    chart_type="column",
    y=18,
    height=75,
)

pres.save(output_dir / "data_to_slides.pptx")
print(f"Presentation saved to {output_dir / 'data_to_slides.pptx'}")
