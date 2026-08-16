"""Matplotlib chart rendering for chart types PowerPoint cannot draw natively.

Used by ``Slide.add_chart`` when the requested chart type is outside the
native set (or when ``backend="pyplot"`` is forced). Requires the optional
plotting extra: ``pip install "easypptx[plot]"``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easypptx.slide import Slide

# Chart kinds that only the matplotlib backend can draw
PYPLOT_ONLY_TYPES = {"heatmap", "histogram", "box", "violin"}

# Native kinds the pyplot backend can also draw when forced
_NATIVE_EQUIVALENTS = {"column", "bar", "line", "pie", "area", "scatter"}

SUPPORTED_TYPES = PYPLOT_ONLY_TYPES | _NATIVE_EQUIVALENTS


def _require_matplotlib() -> Any:
    try:
        import matplotlib

        matplotlib.use("Agg", force=False)
        import matplotlib.pyplot as plt
    except ImportError as err:
        raise ImportError("This chart type needs matplotlib. Install it with: pip install 'easypptx[plot]'") from err
    return plt


def render_chart(
    slide: Slide,
    chart_type: str,
    categories: list,
    series: dict[str, list],
    x: Any,
    y: Any,
    width: Any,
    height: Any,
    title: str | None = None,
    has_legend: bool = True,
    x_title: str | None = None,
    y_title: str | None = None,
    y_min: float | None = None,
    y_max: float | None = None,
    dpi: int = 200,
    palette: list | None = None,
) -> Any:
    """Render a chart with matplotlib and place it on the slide as an image.

    Args:
        slide: Target slide
        chart_type: One of SUPPORTED_TYPES
        categories: Category labels
        series: Mapping of series name -> values
        x: Region left edge (percent or in_() length)
        y: Region top edge (percent or in_() length)
        width: Region width (percent or in_() length)
        height: Region height (percent or in_() length)
        title: Chart title (default: None)
        has_legend: Whether to draw a legend for multi-series data (default: True)
        x_title: X-axis label (default: None)
        y_title: Y-axis label (default: None)
        y_min: Lower y-axis limit (default: None)
        y_max: Upper y-axis limit (default: None)
        dpi: Render resolution (default: 200)
        palette: Series colors as hex strings or RGB tuples (default: None)

    Returns:
        The picture shape placed on the slide

    Raises:
        ValueError: If chart_type is not supported
        ImportError: If matplotlib is not installed
    """
    if chart_type not in SUPPORTED_TYPES:
        raise ValueError(f"Unsupported chart type: {chart_type!r}. Supported: {', '.join(sorted(SUPPORTED_TYPES))}")
    plt = _require_matplotlib()

    def color(i: int) -> Any:
        if not palette:
            return None
        from easypptx.common import resolve_color

        value = palette[i % len(palette)]
        rgb = resolve_color(tuple(value) if isinstance(value, list) else value)
        if rgb is not None:
            return tuple(c / 255 for c in rgb)
        return value  # let matplotlib interpret hex strings and its own names

    fig, ax = plt.subplots(figsize=(8, 5))
    try:
        names = list(series)
        n = len(categories)

        if chart_type in ("column", "bar"):
            positions = range(n)
            group_width = 0.8 / max(len(names), 1)
            for i, name in enumerate(names):
                offset = [p + i * group_width for p in positions]
                if chart_type == "column":
                    ax.bar(offset, series[name], width=group_width, label=name, color=color(i))
                else:
                    ax.barh(offset, series[name], height=group_width, label=name, color=color(i))
            ticks = [p + 0.4 - group_width / 2 for p in positions]
            if chart_type == "column":
                ax.set_xticks(ticks, [str(c) for c in categories])
            else:
                ax.set_yticks(ticks, [str(c) for c in categories])
        elif chart_type in ("line", "area"):
            for i, name in enumerate(names):
                ax.plot(range(n), series[name], label=name, color=color(i))
                if chart_type == "area":
                    ax.fill_between(range(n), series[name], alpha=0.3, color=color(i))
            ax.set_xticks(range(n), [str(c) for c in categories])
        elif chart_type == "scatter":
            for i, name in enumerate(names):
                ax.scatter(range(n), series[name], label=name, color=color(i))
            ax.set_xticks(range(n), [str(c) for c in categories])
        elif chart_type == "pie":
            if len(names) > 1:
                raise ValueError("Pie charts take a single series; got " + ", ".join(names))
            values = series[names[0]]
            pie_colors = [color(i) for i in range(len(values))] if palette else None
            ax.pie(values, labels=[str(c) for c in categories], autopct="%1.1f%%", colors=pie_colors)
        elif chart_type == "heatmap":
            matrix = [series[name] for name in names]
            image = ax.imshow(matrix, aspect="auto", cmap="viridis")
            ax.set_xticks(range(n), [str(c) for c in categories])
            ax.set_yticks(range(len(names)), names)
            fig.colorbar(image, ax=ax)
        elif chart_type == "histogram":
            for i, name in enumerate(names):
                ax.hist(series[name], bins=min(20, max(5, n // 2 or 10)), alpha=0.7, label=name, color=color(i))
        elif chart_type == "box":
            ax.boxplot([series[name] for name in names], tick_labels=names)
        elif chart_type == "violin":
            ax.violinplot([series[name] for name in names], showmedians=True)
            ax.set_xticks(range(1, len(names) + 1), names)

        if title:
            ax.set_title(title)
        # x_title/y_title/y_min/y_max are category/value-axis semantics, so
        # horizontal bars map them to the physical x axis (matching native charts)
        value_axis_is_x = chart_type == "bar"
        if x_title:
            (ax.set_ylabel if value_axis_is_x else ax.set_xlabel)(x_title)
        if y_title:
            (ax.set_xlabel if value_axis_is_x else ax.set_ylabel)(y_title)
        if y_min is not None or y_max is not None:
            if value_axis_is_x:
                ax.set_xlim(left=y_min, right=y_max)
            else:
                ax.set_ylim(bottom=y_min, top=y_max)
        if has_legend and len(names) > 1 and chart_type not in ("pie", "heatmap", "box", "violin"):
            ax.legend()
        fig.tight_layout()

        return slide.add_pyplot(fig, x=x, y=y, width=width, height=height, dpi=dpi)
    finally:
        plt.close(fig)
