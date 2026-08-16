"""The df.pptx pandas accessor.

Registered automatically when pandas is already imported at easypptx import
time (and again whenever a Presentation is created). If you imported pandas
afterwards and the accessor is missing, call
:func:`easypptx.register_pandas_accessor`.

Example:
    ```python
    import pandas as pd
    import easypptx

    pres = easypptx.Presentation()
    slide = pres.add_slide(title="Sales")

    df = pd.DataFrame({"Region": ["East", "West"], "Sales": [120, 95]})
    df.pptx.table(slide, y=20, height=30)
    df.pptx.chart(slide, kind="column", y=55, height=40)
    ```
"""

from __future__ import annotations

import sys
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from easypptx.slide import Slide


def register() -> bool:
    """Register the ``df.pptx`` accessor on pandas DataFrames.

    Returns:
        True if the accessor is registered (or already was), False when
        pandas is not importable.
    """
    try:
        import pandas as pd
        from pandas.api.extensions import register_dataframe_accessor
    except ImportError:
        return False

    existing = getattr(pd.DataFrame, "pptx", None)
    if existing is not None:
        # Class access yields the accessor class itself; instance-style
        # descriptors keep it under _accessor
        accessor_cls = getattr(existing, "_accessor", existing)
        if getattr(accessor_cls, "_easypptx_accessor", False):
            return True
        import warnings

        warnings.warn(
            "DataFrame.pptx is already registered by another library; easypptx will not override it",
            stacklevel=2,
        )
        return False

    @register_dataframe_accessor("pptx")
    class PptxAccessor:
        """Send this DataFrame to an EasyPPTX slide as a table or chart."""

        _easypptx_accessor = True

        def __init__(self, df: Any) -> None:
            self._df = df

        def table(self, slide: Slide, **kwargs: Any) -> Any:
            """Add this DataFrame to a slide as a table.

            Args:
                slide: Target Slide
                **kwargs: Forwarded to Slide.add_table (x, y, width, height,
                    has_header, number_format, shade_columns, ...)

            Returns:
                The created table
            """
            return slide.add_table(self._df, **kwargs)

        def chart(self, slide: Slide, kind: str | None = None, **kwargs: Any) -> Any:
            """Add this DataFrame to a slide as a chart.

            Args:
                slide: Target Slide
                kind: Chart type (native or matplotlib-backed) (default: "column")
                **kwargs: Forwarded to Slide.add_chart (category_column,
                    value_columns, x, y, width, height, show_values, ...)

            Returns:
                The created chart (native) or picture shape (pyplot backend)
            """
            return slide.add_chart(data=self._df, chart_type=kind, **kwargs)

    return True


def maybe_register() -> None:
    """Register the accessor only if pandas is already imported (no import cost)."""
    if "pandas" in sys.modules:
        register()
