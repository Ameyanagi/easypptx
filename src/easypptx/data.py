"""Tabular-data normalization for EasyPPTX.

One adapter for every data shape the chart and table APIs accept:

- pandas DataFrame / Series
- polars DataFrame
- numpy 1D / 2D arrays (with optional ``columns=`` / ``categories=`` labels)
- dict of sequences: ``{"Rev": [...], "Cost": [...]}``
- list of lists whose first row is a header

Everything normalizes to ``(categories, {series_name: values})`` for charts
and to header-plus-rows for tables. pandas / polars / numpy are detected
through ``sys.modules`` so none of them become dependencies.
"""

from __future__ import annotations

import sys
from typing import Any

from easypptx.common import is_dataframe

__all__ = [
    "is_dataframe",
    "is_ndarray",
    "is_polars",
    "is_series",
    "normalize_chart_data",
    "normalize_table_rows",
]


def _module(name: str) -> Any:
    """Return an already-imported module or None (never imports)."""
    return sys.modules.get(name)


def is_series(obj: Any) -> bool:
    """Return True if obj is a pandas Series."""
    pd = _module("pandas")
    return pd is not None and isinstance(obj, pd.Series)


def is_polars(obj: Any) -> bool:
    """Return True if obj is a polars DataFrame."""
    pl = _module("polars")
    return pl is not None and isinstance(obj, pl.DataFrame)


def is_ndarray(obj: Any) -> bool:
    """Return True if obj is a numpy ndarray."""
    np = _module("numpy")
    return np is not None and isinstance(obj, np.ndarray)


def _dict_of_sequences(obj: Any) -> bool:
    return isinstance(obj, dict) and all(not isinstance(v, str | bytes) and hasattr(v, "__len__") for v in obj.values())


def _column_from_frame(data: Any, column: Any, kind: str) -> list:
    """Fetch one column from a pandas DataFrame by label or position."""
    if column in data.columns:
        return data[column].tolist()
    if isinstance(column, int) and 0 <= column < len(data.columns):
        return data.iloc[:, column].tolist()
    raise ValueError(f"{kind} column '{column}' not found in DataFrame")


def _column_from_polars(data: Any, column: Any, kind: str) -> list:
    """Fetch one column from a polars DataFrame by label or position."""
    if isinstance(column, str) and column in data.columns:
        return data.get_column(column).to_list()
    if isinstance(column, int) and 0 <= column < len(data.columns):
        return data.get_column(data.columns[column]).to_list()
    raise ValueError(f"{kind} column '{column}' not found in DataFrame")


def normalize_chart_data(
    data: Any,
    category_column: Any = None,
    value_columns: Any = None,
    categories: list | None = None,
    columns: list[str] | None = None,
) -> tuple[list, dict[str, list]]:
    """Normalize any supported data shape to (categories, {name: values}).

    Args:
        data: DataFrame (pandas or polars), Series, ndarray, dict of
            sequences, or list of lists with a header row
        category_column: Name or index of the category column, for tabular
            inputs (default: first column)
        value_columns: Name(s)/index(es) of value column(s); a list plots
            one series per entry (default: the second column)
        categories: Explicit category labels, for inputs that carry none
            (ndarray, dict) (default: positional labels for ndarray)
        columns: Column names for unlabeled 2D arrays (default: "Series N")

    Returns:
        (categories, series) where series maps series name -> list of values

    Raises:
        ValueError: If a named column is missing or the data is unusable
    """
    # pandas Series: index -> categories, values -> one named series
    if is_series(data):
        cats = categories if categories is not None else [str(i) for i in data.index.tolist()]
        name = str(data.name) if data.name is not None else "Series 1"
        return cats, {name: data.tolist()}

    # dict of sequences: keys are series names
    if _dict_of_sequences(data) and data:
        series = {str(k): list(v) for k, v in data.items()}
        length = len(next(iter(series.values())))
        cats = categories if categories is not None else [str(i) for i in range(length)]
        return cats, series

    # numpy arrays: no labels, so categories/columns fill them in
    if is_ndarray(data):
        if data.ndim == 1:
            cats = categories if categories is not None else [str(i) for i in range(len(data))]
            name = columns[0] if columns else "Series 1"
            return cats, {str(name): data.tolist()}
        if data.ndim == 2:
            n_cols = data.shape[1]
            names = columns if columns is not None else [f"Series {i + 1}" for i in range(n_cols)]
            if len(names) != n_cols:
                raise ValueError(f"columns has {len(names)} names but the array has {n_cols} columns")
            cats = categories if categories is not None else [str(i) for i in range(data.shape[0])]
            return cats, {str(names[i]): data[:, i].tolist() for i in range(n_cols)}
        raise ValueError("Only 1D or 2D arrays are supported for charts")

    # polars DataFrame
    if is_polars(data):
        cols = list(data.columns)
        if not cols:
            return [], {}
        if category_column is None:
            category_column = cols[0]
        selected = value_columns if isinstance(value_columns, list) else [value_columns]
        cats = categories if categories is not None else _column_from_polars(data, category_column, "Category")
        series = {}
        for column in selected:
            if column is None:
                if len(cols) < 2:
                    raise ValueError("DataFrame must have at least two columns for automatic value extraction")
                column = cols[1]
            label = column if isinstance(column, str) else str(cols[column])
            series[label] = _column_from_polars(data, column, "Value")
        return cats, series

    # pandas DataFrame
    if is_dataframe(data):
        if category_column is None:
            category_column = data.columns[0]
        selected = value_columns if isinstance(value_columns, list) else [value_columns]
        cats = categories if categories is not None else _column_from_frame(data, category_column, "Category")
        series = {}
        for column in selected:
            if column is None:
                if len(data.columns) < 2:
                    raise ValueError("DataFrame must have at least two columns for automatic value extraction")
                column = data.columns[1]
            if isinstance(column, int) and column not in data.columns and 0 <= column < len(data.columns):
                label = str(data.columns[column])
            else:
                label = str(column)
            series[label] = _column_from_frame(data, column, "Value")
        return cats, series

    # list of lists with a header row
    if isinstance(data, list | tuple):
        if not data or len(data) < 2:
            return [], {}
        header = list(data[0])

        def column_index(column: Any, default: int, kind: str) -> int:
            if column is None:
                return default
            if isinstance(column, int):
                return column
            try:
                return header.index(column)
            except ValueError:
                raise ValueError(f"{kind} column '{column}' not found in header") from None

        cat_idx = column_index(category_column, 0, "Category")
        cats = categories if categories is not None else [row[cat_idx] for row in data[1:]]
        selected = value_columns if isinstance(value_columns, list) else [value_columns]
        series = {}
        for column in selected:
            if column is None and len(header) < 2:
                raise ValueError("Data must have at least two columns for automatic value extraction")
            val_idx = column_index(column, 1, "Value")
            series[str(header[val_idx])] = [row[val_idx] for row in data[1:]]
        return cats, series

    raise ValueError(
        f"Unsupported chart data type: {type(data).__name__}. "
        "Use a DataFrame, Series, ndarray, dict of sequences, or list of lists."
    )


def normalize_table_rows(data: Any, columns: list[str] | None = None) -> list[list]:
    """Normalize any supported data shape to table rows with a header row.

    Args:
        data: DataFrame (pandas or polars), Series, ndarray, dict of
            sequences, or list of lists (assumed to already include a header)
        columns: Header names for unlabeled 2D arrays (default: "Column N")

    Returns:
        List of rows; the first row is the header (except for plain
        list-of-lists input, which is passed through unchanged)

    Raises:
        ValueError: If the data shape is unsupported
    """
    if is_series(data):
        name = str(data.name) if data.name is not None else "Value"
        return [["", name], *[[str(i), v] for i, v in zip(data.index.tolist(), data.tolist(), strict=True)]]

    if _dict_of_sequences(data) and data:
        names = [str(k) for k in data]
        cols = [list(v) for v in data.values()]
        return [names, *[list(row) for row in zip(*cols, strict=True)]]

    if is_ndarray(data):
        if data.ndim == 1:
            body = [[v] for v in data.tolist()]
            return [[columns[0] if columns else "Column 1"], *body]
        if data.ndim == 2:
            n_cols = data.shape[1]
            names = columns if columns is not None else [f"Column {i + 1}" for i in range(n_cols)]
            if len(names) != n_cols:
                raise ValueError(f"columns has {len(names)} names but the array has {n_cols} columns")
            return [list(names), *data.tolist()]
        raise ValueError("Only 1D or 2D arrays are supported for tables")

    if is_polars(data):
        return [list(data.columns), *[list(row) for row in data.rows()]]

    if is_dataframe(data):
        return [list(data.columns), *data.values.tolist()]

    if isinstance(data, list | tuple):
        return [list(row) for row in data]

    raise ValueError(
        f"Unsupported table data type: {type(data).__name__}. "
        "Use a DataFrame, Series, ndarray, dict of sequences, or list of lists."
    )
