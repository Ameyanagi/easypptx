# Migrating to 0.7.0

EasyPPTX 0.7.0 contains two breaking changes. Most percentage-based code works unchanged; code that relied on floats meaning inches or on the removed `Presentation` pass-through methods needs the updates below.

## 1. Bare numbers are now percentages

Position and size parameters (`x`, `y`, `width`, `height`) treat bare numbers as percentages of the slide dimension: `x=10` is the same as `x="10%"`. In earlier versions a float meant inches. Absolute inches now require the `in_()` helper:

**Before (0.6.x):**

```python
slide.add_text("Hello", x=1.0, y=2.0, width=8.0, height=1.0)  # inches
```

**After (0.7.0):**

```python
from easypptx import in_

# Percentages (recommended, responsive)
slide.add_text("Hello", x=10, y=25, width=80, height=10)

# Or keep exact inches with in_()
slide.add_text("Hello", x=in_(1.0), y=in_(2.0), width=in_(8.0), height=in_(1.0))
```

Related changes:

- Default positions changed from inches to percentages, e.g. `add_text` now defaults to `x=5, y=5, width=90, height=10` and `add_chart` to `x=10, y=20, width=60, height=60`.
- Percentages outside 0-100 are clamped with a warning; use `in_()` for intentional off-slide placement.

## 2. Removed deprecated Presentation methods

The deprecated pass-through and convenience methods on `Presentation` have been removed:

| Removed | Use instead |
| --- | --- |
| `pres.add_text(slide, ...)` | `slide.add_text(...)` |
| `pres.add_image(slide, ...)` | `slide.add_image(...)` |
| `pres.add_shape(slide, ...)` | `slide.add_shape(...)` |
| `pres.add_table(slide, ...)` | `slide.add_table(...)` |
| `pres.add_chart(slide, ...)` | `slide.add_chart(...)` |
| `pres.add_pyplot(slide, ...)` | `slide.add_pyplot(...)` |
| `pres.add_matplotlib_slide(...)` | `pres.add_pyplot_slide(...)` |
| `pres.add_seaborn_slide(...)` | `pres.add_pyplot_slide(...)` |
| `pres.add_plot(...)` | `pres.add_pyplot_slide(...)` or `pres.add_chart_slide(...)` |
| `pres.add_image_slide(...)` | `pres.add_image_gen_slide(...)` |

**Before (0.6.x):**

```python
pres.add_text(slide, "Hello", x="10%", y="20%")
slide_out, picture = pres.add_matplotlib_slide(figure=fig, title="Plot")
```

**After (0.7.0):**

```python
slide.add_text("Hello", x=10, y=20)
slide_out, picture = pres.add_pyplot_slide(figure=fig, title="Plot")
```

## What's new in 0.7.0

Not breaking, but worth adopting:

- [Markdown to deck](markdown.md): `Presentation.from_markdown("deck.md")`
- [Grid upgrades](grid_layout.md): slice spans (`grid[1, :]`), weighted tracks (`Grid(rows=[2, 1], cols=3)`), auto-flow (`grid.next()`), and cell styling (`grid[0, 0].style(...)`)
- [Styles and themes](styling.md): `TextStyle` / `TableStyle` / `ChartStyle` dataclasses, `Theme`, and built-in presets (`Presentation(theme="dark")`)
- Content additions: `slide.add_bullets(...)`, `slide.notes`, `slide.add_pyplot(...)`, and multi-series charts via `value_columns=[...]` or `Chart.add(..., series={...})`
