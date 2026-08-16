"""Grid layout module for EasyPPTX."""

from __future__ import annotations

from collections.abc import Callable
from typing import IO, Any, overload

from easypptx.common import filter_to_signature, is_dataframe, merge_defaults
from easypptx.positioning import PositionType, pct, shift_band, to_percent

# Using forward annotations (PEP 563) to avoid circular references


class GridCell:
    """Class representing a cell in a grid.

    This class stores information about a cell's position and dimensions
    within a grid layout.

    Attributes:
        row: Row index
        col: Column index
        x: X position as percentage
        y: Y position as percentage
        width: Width as percentage
        height: Height as percentage
        content: The content placed in this cell (if any)
    """

    def __init__(self, row: int, col: int, x: str, y: str, width: str, height: str) -> None:
        """Initialize a GridCell.

        Args:
            row: Row index of the cell
            col: Column index of the cell
            x: X position as percentage
            y: Y position as percentage
            width: Width as percentage
            height: Height as percentage
        """
        self.row = row
        self.col = col
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.content: Any = None
        self.span_rows = 1
        self.span_cols = 1
        self.is_spanned = False  # Whether this cell is part of another cell's span
        self.pad = 0.0  # Extra per-cell content padding, percent of slide

    def __repr__(self) -> str:
        """Return string representation of the cell."""
        return (
            f"GridCell(row={self.row}, col={self.col}, x={self.x}, "
            f"y={self.y}, width={self.width}, height={self.height})"
        )


def _accepts_position(func: Callable) -> bool:
    """Return True if func can receive x/y/width/height keyword arguments."""
    import inspect

    try:
        sig = inspect.signature(func)
    except (TypeError, ValueError):
        return True
    params = list(sig.parameters.values())
    if any(p.kind is inspect.Parameter.VAR_KEYWORD for p in params):
        return True
    names = {p.name for p in params}
    return {"x", "y", "width", "height"} <= names


class OutOfBoundsError(IndexError):
    """Exception raised when grid coordinates are out of bounds."""


class CellMergeError(ValueError):
    """Exception raised when there's an issue with cell merging."""


class GridCellProxy:
    """Proxy class for accessing grid cells with enhanced syntax.

    This class provides a more intuitive API for accessing grid cells and provides
    direct access to all the add_xxx methods that the Grid class supports.

    Attributes:
        grid: The parent Grid object
        row: The row index of this cell
        col: The column index of this cell
    """

    def __init__(self, grid: Grid, row: int, col: int):
        """Initialize a GridCellProxy.

        Args:
            grid: The parent Grid object
            row: The row index of this cell
            col: The column index of this cell
        """
        self.grid = grid
        self.row = row
        self.col = col

    def add_text(self, text: str, **kwargs) -> Any:
        """Add text to this cell.

        Args:
            text: The text to add
            **kwargs: Additional keyword arguments for the text

        Returns:
            The created text object
        """
        return self.grid.add_textbox(self.row, self.col, text=text, **kwargs)

    def add_image(self, image_path: str | IO[bytes], **kwargs) -> Any:
        """Add an image to this cell.

        Args:
            image_path: The path to the image
            **kwargs: Additional keyword arguments for the image

        Returns:
            The created image object
        """
        return self.grid.add_image(self.row, self.col, image_path=image_path, **kwargs)

    def add_pyplot(self, figure, **kwargs) -> Any:
        """Add a matplotlib figure to this cell.

        Args:
            figure: The matplotlib figure to add
            **kwargs: Additional keyword arguments for the figure

        Returns:
            The created pyplot object
        """
        return self.grid.add_pyplot(self.row, self.col, figure=figure, **kwargs)

    def add_table(self, data, **kwargs) -> Any:
        """Add a table to this cell.

        Args:
            data: The table data to add
            **kwargs: Additional keyword arguments for the table

        Returns:
            The created table object
        """
        return self.grid.add_table(self.row, self.col, data=data, **kwargs)

    def style(
        self,
        fill: str | tuple[int, int, int] | None = None,
        border_color: str | tuple[int, int, int] | None = None,
        border_width: float = 1.0,
        padding: float | str | None = None,
    ) -> GridCellProxy:
        """Style this cell: background fill, border, and content padding.

        Draws a background rectangle covering the full cell and optionally
        sets padding that shrinks the area used by content added afterwards.
        Returns the proxy so calls can be chained.

        Args:
            fill: Background color name or RGB tuple (default: None, no fill)
            border_color: Border color name or RGB tuple (default: None, no border)
            border_width: Border width in points (default: 1.0)
            padding: Extra content padding as percent of the slide (default: None)

        Examples:
            ```python
            grid[0, 0].style(fill="lightgray", padding=2).add_text("Card")
            ```
        """
        from easypptx.common import resolve_color
        from easypptx.positioning import parse_percent

        cell = self.grid.get_cell(self.row, self.col)
        if padding is not None:
            cell.pad = parse_percent(padding)

        if fill is not None or border_color is not None:
            # Background covers the unpadded cell area
            x, y, w, h = self.grid._cell_area(cell, padded=False)
            shape = self.grid.parent.add_shape(x=x, y=y, width=w, height=h)

            fill_rgb = resolve_color(fill)
            if fill_rgb is not None:
                shape.fill.solid()
                shape.fill.fore_color.rgb = fill_rgb
            else:
                shape.fill.background()

            border_rgb = resolve_color(border_color)
            if border_rgb is not None:
                from pptx.util import Pt

                shape.line.color.rgb = border_rgb
                shape.line.width = Pt(border_width)
            else:
                shape.line.fill.background()

        return self

    def add_grid(self, rows: int = 1, cols: int = 1, padding: float = 5.0, **kwargs) -> Grid:
        """Add a nested grid to this cell.

        Args:
            rows: The number of rows in the nested grid
            cols: The number of columns in the nested grid
            padding: The padding for the nested grid
            **kwargs: Additional parameters to pass to the grid

        Returns:
            The created grid object
        """
        return self.grid.add_grid_to_cell(self.row, self.col, rows=rows, cols=cols, padding=padding, **kwargs)


class Grid:
    """Class for creating grid layouts in PowerPoint slides.

    This class provides methods for creating grid layouts and adding content
    to specific cells within the grid. The Grid is iterable and indexable like
    a numpy ndarray or matplotlib subplot grid.

    Attributes:
        parent: The parent Slide or Grid object
        x: X position of the grid as percentage or absolute value
        y: Y position of the grid as percentage or absolute value
        width: Width of the grid as percentage or absolute value
        height: Height of the grid as percentage or absolute value
        rows: Number of rows in the grid
        cols: Number of columns in the grid
        padding: Padding between cells as percentage of cell size
        cells: 2D array of GridCell objects

    Examples:
        ```python
        # Access a cell with grid[row, col]
        cell = grid[0, 1]  # Get cell at row 0, column 1

        # Add content directly to a cell using proxy access
        grid[0, 1].add_text("Cell 0,1", font_size=24)
        grid[1, 0].add_image(image_path="image.png")

        # Access a cell using flat indexing
        grid[0].add_text("First cell (flat index 0)")
        grid[3].add_text("Fourth cell (flat index 3)")

        # Loop through all cells
        for cell in grid:
            print(cell)

        # Loop through cells linearly (flattened)
        for cell in grid.flat:
            print(cell.row, cell.col)

        # Add content using traditional methods
        grid.add_textbox(0, 1, text="Cell 0,1", font_size=24)
        grid.add_image(1, 0, image_path="image.png")
        ```
    """

    def __init__(
        self,
        parent: Any,
        x: PositionType = "0%",
        y: PositionType = "0%",
        width: PositionType = "100%",
        height: PositionType = "100%",
        rows: int | list[float] = 1,
        cols: int | list[float] = 1,
        padding: float = 5.0,
    ) -> None:
        """Initialize a Grid layout.

        Args:
            parent: The parent Slide or Grid object
            x: X position of the grid (default: "0%")
            y: Y position of the grid (default: "0%")
            width: Width of the grid (default: "100%")
            height: Height of the grid (default: "100%")
            rows: Number of rows, or a list of relative row weights,
                e.g. [2, 1] for a top region twice as tall (default: 1)
            cols: Number of columns, or a list of relative column weights (default: 1)
            padding: Padding between cells as percentage of cell size (default: 5.0)
        """
        self.parent = parent
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.row_weights: list[float] = [1.0] * rows if isinstance(rows, int) else [float(w) for w in rows]
        self.col_weights: list[float] = [1.0] * cols if isinstance(cols, int) else [float(w) for w in cols]
        self.rows = len(self.row_weights)
        self.cols = len(self.col_weights)
        self.padding = padding

        # Store template defaults when applied from a template
        self.template_defaults: dict[str, dict[str, Any]] = {
            "text": {},
            "image": {},
            "pyplot": {},
            "table": {},
            "grid": {},
            "global": {},
        }

        # Copy template_defaults from parent if it's a Grid and has template_defaults
        if hasattr(parent, "template_defaults") and isinstance(parent, Grid):
            # Access parent's template_defaults safely
            for key, value in parent.template_defaults.items():
                self.template_defaults[key] = value.copy()

        # Store slide dimensions for percentage calculations
        self._slide_width = self._get_slide_width()
        self._slide_height = self._get_slide_height()

        # Calculate cell dimensions
        self.cells = self._create_cells()

    def _get_slide_width(self) -> int:
        """Get the slide width in EMUs from the parent.

        Returns:
            The slide width in English Metric Units (EMUs)
        """
        # If parent is a Slide object, use its slide width
        if hasattr(self.parent, "_slide_width"):
            return self.parent._slide_width
        # If parent is another Grid, use its slide width
        elif hasattr(self.parent, "_get_slide_width"):
            return self.parent._get_slide_width()
        # Default value if we can't get it: standard 16:9 width
        return 12192768  # 13.33 inches in EMUs

    def _get_slide_height(self) -> int:
        """Get the slide height in EMUs from the parent.

        Returns:
            The slide height in English Metric Units (EMUs)
        """
        # If parent is a Slide object, use its slide height
        if hasattr(self.parent, "_slide_height"):
            return self.parent._slide_height
        # If parent is another Grid, use its slide height
        elif hasattr(self.parent, "_get_slide_height"):
            return self.parent._get_slide_height()
        # Default value if we can't get it (equivalent to 7.5 inches)
        return 6858000  # 7.5 inches in EMUs

    def apply_template_defaults(self, template_data: dict[str, Any]) -> None:
        """Apply template defaults to this grid.

        This method extracts default method arguments from a template and stores them
        for later use by the add_xxx methods.

        Args:
            template_data: Template data dictionary
        """
        # Extract defaults for different element types
        if "defaults" in template_data:
            defaults = template_data["defaults"]

            # Apply defaults for each element type
            for element_type in ["text", "image", "pyplot", "table", "grid", "global"]:
                if element_type in defaults:
                    self.template_defaults[element_type] = defaults[element_type]

    def merge_with_defaults(self, method_type: str, kwargs: dict[str, Any]) -> dict[str, Any]:
        """Merge provided arguments with template defaults.

        Args:
            method_type: The type of method ("text", "image", etc.)
            kwargs: Keyword arguments provided to the method

        Returns:
            Dictionary with merged arguments, where provided args override defaults
        """
        return merge_defaults(self.template_defaults, method_type, kwargs)

    def _create_cells(self) -> list[list[GridCell]]:
        """Create the grid cells based on the layout.

        Returns:
            2D array of GridCell objects
        """
        cells = []

        # Convert percentage values to floats for calculations
        padding_factor = self.padding / 100.0

        # Per-row/column sizes from relative weights (equal splits by default)
        total_row_weight = sum(self.row_weights)
        total_col_weight = sum(self.col_weights)
        col_sizes = [100.0 * w / total_col_weight for w in self.col_weights]
        row_sizes = [100.0 * w / total_row_weight for w in self.row_weights]
        col_offsets = [sum(col_sizes[:i]) for i in range(self.cols)]
        row_offsets = [sum(row_sizes[:i]) for i in range(self.rows)]

        # Create cells
        for row in range(self.rows):
            cell_row = []
            for col in range(self.cols):
                cell_width_percent = col_sizes[col]
                cell_height_percent = row_sizes[row]

                # Effective size excludes the per-cell padding
                effective_cell_width = cell_width_percent * (1 - padding_factor)
                effective_cell_height = cell_height_percent * (1 - padding_factor)
                half_padding_width = (cell_width_percent * padding_factor) / 2
                half_padding_height = (cell_height_percent * padding_factor) / 2
                # Calculate cell position
                x_percent = col_offsets[col] + half_padding_width
                y_percent = row_offsets[row] + half_padding_height

                # Convert to percentage strings
                x_str = f"{x_percent:.2f}%"
                y_str = f"{y_percent:.2f}%"
                width_str = f"{effective_cell_width:.2f}%"
                height_str = f"{effective_cell_height:.2f}%"

                # Create the cell
                cell = GridCell(row, col, x_str, y_str, width_str, height_str)
                cell_row.append(cell)

            cells.append(cell_row)

        return cells

    def get_cell(self, row: int, col: int) -> GridCell:
        """Get a cell at the specified row and column.

        Args:
            row: Row index (0-based)
            col: Column index (0-based)

        Returns:
            The GridCell at the specified position

        Raises:
            OutOfBoundsError: If row or column is out of bounds
        """
        if row < 0 or row >= self.rows or col < 0 or col >= self.cols:
            raise OutOfBoundsError(f"Cell position ({row}, {col}) is out of bounds")

        return self.cells[row][col]

    def merge_cells(self, start_row: int, start_col: int, end_row: int, end_col: int) -> GridCell:
        """Merge cells in the specified range.

        Args:
            start_row: Starting row index (0-based)
            start_col: Starting column index (0-based)
            end_row: Ending row index (0-based, inclusive)
            end_col: Ending column index (0-based, inclusive)

        Returns:
            The merged cell

        Raises:
            OutOfBoundsError: If any row or column is out of bounds
            CellMergeError: If the merged area overlaps with an existing merged cell
        """
        # Validate bounds
        if (
            start_row < 0
            or start_row >= self.rows
            or start_col < 0
            or start_col >= self.cols
            or end_row < 0
            or end_row >= self.rows
            or end_col < 0
            or end_col >= self.cols
        ):
            raise OutOfBoundsError("Merge area is out of bounds")

        # Make sure start coordinates are less than or equal to end coordinates
        if start_row > end_row or start_col > end_col:
            raise CellMergeError("Start coordinates must be less than or equal to end coordinates")

        # Check if any of the cells in the range are already merged
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = self.cells[row][col]
                if cell.is_spanned:
                    raise CellMergeError("Cell is already part of a merged cell")

        # Get the first cell (top-left)
        first_cell = self.cells[start_row][start_col]

        # Calculate the new width and height
        last_cell = self.cells[end_row][end_col]

        # Extract numeric values from percentage strings
        first_x = float(first_cell.x.strip("%"))
        first_y = float(first_cell.y.strip("%"))

        # Calculate the rightmost and bottommost positions
        last_x = float(last_cell.x.strip("%"))
        last_y = float(last_cell.y.strip("%"))
        last_width = float(last_cell.width.strip("%"))
        last_height = float(last_cell.height.strip("%"))

        # Calculate the new width and height
        new_width = (last_x + last_width) - first_x
        new_height = (last_y + last_height) - first_y

        # Update the first cell's dimensions
        first_cell.width = f"{new_width:.2f}%"
        first_cell.height = f"{new_height:.2f}%"
        first_cell.span_rows = end_row - start_row + 1
        first_cell.span_cols = end_col - start_col + 1

        # Mark other cells in the range as spanned
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                if row != start_row or col != start_col:
                    self.cells[row][col].is_spanned = True

        return first_cell

    def _cell_area(self, cell: GridCell, padded: bool = True) -> tuple[str, str, str, str]:
        """Compute a cell's absolute slide position as percent strings.

        The grid's own position may be given in percent or absolute units;
        cell positions are always percentages of the grid area. When padded
        is True, the cell's own content padding (set via style()) is applied.

        Returns:
            (x, y, width, height) as percentage strings
        """
        grid_x = to_percent(self.x, self._slide_width)
        grid_y = to_percent(self.y, self._slide_height)
        grid_w = to_percent(self.width, self._slide_width)
        grid_h = to_percent(self.height, self._slide_height)

        abs_x = grid_x + (to_percent(cell.x, self._slide_width) * grid_w / 100)
        abs_y = grid_y + (to_percent(cell.y, self._slide_height) * grid_h / 100)
        abs_w = to_percent(cell.width, self._slide_width) * grid_w / 100
        abs_h = to_percent(cell.height, self._slide_height) * grid_h / 100

        if padded and cell.pad:
            abs_x += cell.pad
            abs_y += cell.pad
            abs_w -= 2 * cell.pad
            abs_h -= 2 * cell.pad

        return pct(abs_x), pct(abs_y), pct(abs_w), pct(abs_h)

    def add_to_cell(self, row: int, col: int, content_func: Callable, **kwargs) -> Any:
        """Add content to a specific cell in the grid.

        The content function is called with x/y/width/height keyword arguments
        holding the cell's absolute position as percentage strings.

        Args:
            row: Row index (0-based)
            col: Column index (0-based)
            content_func: Function to call to add content (e.g., slide.add_text)
            **kwargs: Additional arguments to pass to the content function

        Returns:
            The object returned by the content function

        Raises:
            OutOfBoundsError: If row or column is out of bounds
            CellMergeError: If the cell is part of a merged cell
        """
        cell = self.get_cell(row, col)

        if cell.is_spanned:
            raise CellMergeError("Cell is part of a merged cell")

        kwargs["x"], kwargs["y"], kwargs["width"], kwargs["height"] = self._cell_area(cell)

        content = content_func(**kwargs)
        cell.content = content
        return content

    def add_grid_to_cell(
        self,
        row: int,
        col: int,
        rows: int = 1,
        cols: int = 1,
        padding: float = 5.0,
        **kwargs,  # Accept any additional parameters
    ) -> Grid:
        """Add a nested grid to a specific cell.

        Args:
            row: Row index (0-based)
            col: Column index (0-based)
            rows: Number of rows in the nested grid (default: 1)
            cols: Number of columns in the nested grid (default: 1)
            padding: Padding between cells as percentage of cell size (default: 5.0)
            **kwargs: Additional parameters (ignored)

        Returns:
            The nested Grid object

        Raises:
            OutOfBoundsError: If row or column is out of bounds
            CellMergeError: If the cell is part of a merged cell
        """
        # Get the cell
        cell = self.get_cell(row, col)

        # Check if the cell is part of a merged cell
        if cell.is_spanned:
            raise CellMergeError("Cell is part of a merged cell")

        # Merge provided kwargs with template defaults
        merged_kwargs = self.merge_with_defaults("grid", {"rows": rows, "cols": cols, "padding": padding})

        # Calculate absolute position for the nested grid
        abs_x, abs_y, abs_width, abs_height = self._cell_area(cell)

        # Create the nested grid
        nested_grid = Grid(
            parent=self.parent,
            x=abs_x,
            y=abs_y,
            width=abs_width,
            height=abs_height,
            rows=merged_kwargs.get("rows", 1),
            cols=merged_kwargs.get("cols", 1),
            padding=merged_kwargs.get("padding", 5.0),
        )

        # Pass template defaults to the nested grid
        for key, value in self.template_defaults.items():
            nested_grid.template_defaults[key] = value.copy()

        # Store the nested grid in the cell
        # Store the nested grid in the cell's content
        if hasattr(cell, "content"):
            cell.content = nested_grid

        return nested_grid

    def __iter__(self):
        """Make Grid iterable to loop through all cells.

        Returns:
            Iterator over all grid cells
        """
        for row in range(self.rows):
            for col in range(self.cols):
                yield self.cells[row][col]

    @overload
    def __getitem__(self, key: int) -> GridCellProxy: ...

    @overload
    def __getitem__(self, key: tuple[int, int]) -> GridCellProxy: ...

    def __getitem__(self, key: int | tuple[int, int]) -> GridCellProxy | GridCell:
        """Access a cell using enhanced indexing with auto-expansion.

        This method supports two types of indexing:
        - grid[idx] for flat access to cells by index (returns GridCellProxy)
        - grid[row, col] for cell-specific operations (returns GridCellProxy)

        For flat indexing, cells are ordered row-wise (0 is top-left, then across the row,
        then to the next row).

        If the requested cell is out of bounds, the grid will raise an OutOfBoundsError.

        Args:
            key: An int for flat access, or a tuple of (row, col) for cell access

        Returns:
            - GridCellProxy if key is a flat index
            - GridCellProxy if key is a tuple of (row, col)

        Raises:
            OutOfBoundsError: If the requested cell is out of bounds
            TypeError: If the key is not in the right format

        Examples:
            ```python
            # Access a specific cell by row, col
            grid[0, 1].add_text("Cell 0,1")

            # Access a cell using flat indexing (0-based)
            grid[0].add_text("First cell (top-left)")
            grid[3].add_text("Fourth cell")
            ```
        """
        if isinstance(key, tuple) and len(key) == 2:
            row, col = key

            # Slice access spans (and merges) a rectangular region:
            # grid[1, :] is the whole second row, grid[0:2, 1] spans two rows
            if isinstance(row, slice) or isinstance(col, slice):
                return self._span(row, col)

            # Access as grid[row, col] -> return a GridCellProxy
            if row < 0 or row >= self.rows or col < 0 or col >= self.cols:
                raise OutOfBoundsError(f"Cell position ({row}, {col}) is out of bounds")
            return GridCellProxy(self, row, col)
        elif isinstance(key, int):
            # Interpret integer keys as flat indices
            if 0 <= key < self.rows * self.cols:
                # This is a valid flat index
                row = key // self.cols
                col = key % self.cols
                return GridCellProxy(self, row, col)
            elif key < 0:
                # Negative index handling for flat access
                total_cells = self.rows * self.cols
                # Convert negative index to positive
                actual_idx = total_cells + key

                if actual_idx < 0:
                    # Still negative after conversion
                    raise OutOfBoundsError(f"Flat index {key} is out of bounds")

                row = actual_idx // self.cols
                col = actual_idx % self.cols
                return GridCellProxy(self, row, col)
            else:
                # This is an out of bounds index
                raise OutOfBoundsError(f"Flat index {key} is out of bounds")
        else:
            raise TypeError(
                "Grid indices must be integers (for flat access) or tuples of the form (row, col) for cell access"
            )

    def _span(self, row: int | slice, col: int | slice) -> GridCellProxy:
        """Merge the region selected by slice indices and return its proxy."""

        def bounds(key: int | slice, size: int, axis: str) -> tuple[int, int]:
            if isinstance(key, slice):
                if key.step not in (None, 1):
                    raise ValueError(f"Grid {axis} slices do not support a step")
                start, stop, _ = key.indices(size)
                if stop <= start:
                    raise OutOfBoundsError(f"Empty {axis} slice {key} for grid of size {size}")
                return start, stop - 1
            if key < 0:
                key += size
            if not 0 <= key < size:
                raise OutOfBoundsError(f"Grid {axis} {key} is out of bounds")
            return key, key

        start_row, end_row = bounds(row, self.rows, "row")
        start_col, end_col = bounds(col, self.cols, "column")

        origin = self.cells[start_row][start_col]
        span_rows = end_row - start_row + 1
        span_cols = end_col - start_col + 1

        # Merge unless this exact span already exists (repeat access is fine)
        if (origin.span_rows, origin.span_cols) != (span_rows, span_cols):
            self.merge_cells(start_row, start_col, end_row, end_col)

        return GridCellProxy(self, start_row, start_col)

    @property
    def flat(self):
        """Flat iterator for this grid, similar to matplotlib's subplot.flat.

        Returns:
            A flat iterator over all cells in the grid
        """
        return GridFlatIterator(self)

    def add_textbox(self, row: int, col: int, text: str, **kwargs) -> Any:
        """Add a text box to a specific cell in the grid.

        This is a convenience method that calls add_to_cell with the parent's
        add_text method.

        Args:
            row: Row index (0-based)
            col: Column index (0-based)
            text: The text content to add
            **kwargs: Additional arguments to pass to the parent's add_text method
                     (font_size, font_bold, align, etc.)

        Returns:
            The created text shape

        Raises:
            OutOfBoundsError: If row or column is out of bounds
            CellMergeError: If the cell is part of a merged cell

        Example:
            ```python
            # Add text to a specific cell
            grid.add_textbox(0, 1, "Hello World", font_size=24, align="center")
            ```
        """
        # Add the text parameter to kwargs
        kwargs["text"] = text

        # Merge provided kwargs with template defaults, dropping default keys
        # that the text method does not accept (e.g. slide-factory defaults)
        merged_kwargs = filter_to_signature(self.parent.add_text, self.merge_with_defaults("text", kwargs), kwargs)

        # Convert list colors to tuples if needed
        if "color" in merged_kwargs and isinstance(merged_kwargs["color"], list) and len(merged_kwargs["color"]) == 3:
            merged_kwargs["color"] = tuple(merged_kwargs["color"])

        # Call add_to_cell with the parent's add_text method
        # The parent's add_text method now accepts **kwargs which will handle any extra parameters
        return self.add_to_cell(row, col, self.parent.add_text, **merged_kwargs)

    def add_image(self, row: int, col: int, image_path: str | IO[bytes], **kwargs) -> Any:
        """Add an image to a specific cell in the grid.

        This is a convenience method that calls add_to_cell with the parent's
        add_image method.

        Args:
            row: Row index (0-based)
            col: Column index (0-based)
            image_path: Path to the image file
            **kwargs: Additional arguments to pass to the parent's add_image method
                     (maintain_aspect_ratio, border, shadow, etc.)

        Returns:
            The created image shape

        Raises:
            OutOfBoundsError: If row or column is out of bounds
            CellMergeError: If the cell is part of a merged cell

        Example:
            ```python
            # Add image to a specific cell
            grid.add_image(1, 0, "path/to/image.jpg", maintain_aspect_ratio=True)
            ```
        """
        # Add the image_path parameter to kwargs
        kwargs["image_path"] = image_path

        # Merge provided kwargs with template defaults, dropping default keys
        # that the image method does not accept
        merged_kwargs = filter_to_signature(self.parent.add_image, self.merge_with_defaults("image", kwargs), kwargs)

        # Call add_to_cell with the parent's add_image method
        return self.add_to_cell(row, col, self.parent.add_image, **merged_kwargs)

    def add_pyplot(self, row: int, col: int, figure, **kwargs) -> Any:
        """Add a matplotlib figure to a specific cell in the grid.

        This is a convenience method that creates a temporary file for the figure
        and then adds it as an image to the cell.

        Args:
            row: Row index (0-based)
            col: Column index (0-based)
            figure: Matplotlib figure object (plt.figure() or plt.gcf())
            **kwargs: Additional arguments for the figure
                     (dpi, file_format, etc.)

        Returns:
            The created image shape

        Raises:
            OutOfBoundsError: If row or column is out of bounds
            CellMergeError: If the cell is part of a merged cell

        Example:
            ```python
            import matplotlib.pyplot as plt

            # Create a matplotlib figure
            plt.figure()
            plt.plot([1, 2, 3, 4], [1, 4, 9, 16])
            plt.title('Sample Plot')

            # Add plot to a specific cell
            grid.add_pyplot(0, 1, plt.gcf(), dpi=300)
            ```
        """
        from easypptx.pyplot import figure_to_stream

        # Merge provided kwargs with template defaults
        merged_kwargs = self.merge_with_defaults("pyplot", kwargs)

        # Set default values
        dpi = merged_kwargs.pop("dpi", 300)
        file_format = merged_kwargs.pop("file_format", "png")

        # Drop template-default keys the image method does not accept, so they
        # don't get forwarded as if the caller passed them explicitly
        merged_kwargs = filter_to_signature(self.parent.add_image, merged_kwargs, kwargs)

        # Render the figure in memory and add it to the cell
        stream = figure_to_stream(figure, dpi=dpi, file_format=file_format)
        return self.add_image(row, col, image_path=stream, **merged_kwargs)

    def add_table(self, row: int, col: int, data, **kwargs) -> Any:
        """Add a table to a specific cell in the grid.

        This is a convenience method that creates a Table object and adds it to the cell.

        Args:
            row: Row index (0-based)
            col: Column index (0-based)
            data: Table data as a list of lists or pandas DataFrame
            **kwargs: Additional arguments for the table
                     (has_header, style, etc.)

        Returns:
            The created table shape

        Raises:
            OutOfBoundsError: If row or column is out of bounds
            CellMergeError: If the cell is part of a merged cell

        Example:
            ```python
            # Add table to a specific cell
            data = [["Name", "Value"], ["Item 1", 100], ["Item 2", 200]]
            grid.add_table(0, 0, data, has_header=True)

            # With pandas DataFrame
            import pandas as pd
            df = pd.DataFrame({"Name": ["Item 1", "Item 2"], "Value": [100, 200]})
            grid.add_table(1, 1, df)
            ```
        """
        from easypptx.table import Table
        from easypptx.table import Table as _Table

        # Get the cell to determine the position and dimensions
        cell = self.get_cell(row, col)

        # Merge provided kwargs with template defaults, dropping default keys
        # that Table.add does not accept
        merged_kwargs = filter_to_signature(_Table.add, self.merge_with_defaults("table", kwargs), kwargs)

        # Create a Table object
        table_obj = Table(self.parent)

        # Use the cell's absolute position on the slide
        merged_kwargs["x"], merged_kwargs["y"], merged_kwargs["width"], merged_kwargs["height"] = self._cell_area(cell)

        # Remove data from kwargs to handle separately
        merged_kwargs.pop("data", None)

        # Handle has_header parameter if provided
        if "has_header" in merged_kwargs:
            merged_kwargs["first_row_header"] = merged_kwargs.pop("has_header")

        # Convert list colors to tuples if needed
        for style_key in ["header_style", "row_style"]:
            if style_key in merged_kwargs:
                style_dict = merged_kwargs[style_key]
                for color_key in ["bg_color", "text_color"]:
                    if (
                        color_key in style_dict
                        and isinstance(style_dict[color_key], list)
                        and len(style_dict[color_key]) == 3
                    ):
                        style_dict[color_key] = tuple(style_dict[color_key])

        # Convert pandas DataFrame to list if needed
        table_data = [list(data.columns), *data.values.tolist()] if is_dataframe(data) else data

        # Create the table with the processed data
        table_shape = table_obj.add(data=table_data, **merged_kwargs)

        # Store the table in the cell's content
        cell.content = table_shape

        return table_shape

    def next(self) -> GridCellProxy:
        """Return a proxy for the next free cell, growing the grid if needed.

        Cells fill in row-major order; when every cell has content, a new
        row is appended automatically.

        Examples:
            ```python
            grid.next().add_text("First free cell")
            grid.next().add_image(image_path="a.png")
            ```
        """
        for row in range(self.rows):
            for col in range(self.cols):
                cell = self.cells[row][col]
                if cell.content is None and not cell.is_spanned:
                    return GridCellProxy(self, row, col)
        self._expand_grid(add_rows=1)
        return GridCellProxy(self, self.rows - 1, 0)

    def append(self, content_func: Callable) -> None:
        """Append content to the grid and automatically update the layout.

        This method adds a new content function to the grid and automatically
        recalculates the grid layout to accommodate the new content. If needed,
        it will expand the grid by adding rows.

        Args:
            content_func: A function that adds content (like add_text, add_image, etc.)

        Returns:
            None

        Examples:
            ```python
            # Create a dynamic grid
            grid = Grid(slide, rows=2, cols=3)

            # Append content, grid will expand as needed
            grid.append(lambda **kwargs: slide.add_text("Item 1", **kwargs))
            grid.append(lambda **kwargs: slide.add_text("Item 2", **kwargs))
            grid.append(lambda **kwargs: slide.add_image("image.png", **kwargs))
            ```
        """
        # Calculate current grid capacity and total items
        capacity = self.rows * self.cols
        cells_used = 0

        # Count used cells
        for row in range(self.rows):
            for col in range(self.cols):
                if self.cells[row][col].content is not None:
                    cells_used += 1

        # If grid is full, add a new row
        if cells_used >= capacity:
            self._expand_grid(add_rows=1, add_cols=0)

        # Find the next available cell
        target_cell = None
        for row in range(self.rows):
            for col in range(self.cols):
                if self.cells[row][col].content is None:
                    target_cell = self.cells[row][col]
                    break
            if target_cell:
                break

        # Add content to the target cell
        if target_cell:
            row, col = target_cell.row, target_cell.col

            # Calculate position and dimensions for content
            x = target_cell.x
            y = target_cell.y
            width = target_cell.width
            height = target_cell.height

            # Call the content function with the cell's position and dimensions
            content = content_func(x=x, y=y, width=width, height=height)

            # Store the content in the cell
            target_cell.content = content

    def _expand_grid(self, add_rows: int = 0, add_cols: int = 0) -> None:
        """Expand the grid by adding rows and/or columns.

        This method increases the size of the grid by adding the specified number
        of rows and/or columns while maintaining the existing content.

        Args:
            add_rows: Number of rows to add (default: 0)
            add_cols: Number of columns to add (default: 0)

        Returns:
            None
        """
        if add_rows <= 0 and add_cols <= 0:
            return  # Nothing to do

        # Save original dimensions
        original_rows = self.rows
        original_cols = self.cols

        # Update dimensions (new rows/columns get weight 1)
        self.row_weights.extend([1.0] * add_rows)
        self.col_weights.extend([1.0] * add_cols)
        self.rows += add_rows
        self.cols += add_cols

        # Recalculate cell dimensions
        new_cells = self._create_cells()

        # Copy content from old cells to new cells where applicable
        for row in range(original_rows):
            for col in range(original_cols):
                if row < self.rows and col < self.cols:
                    new_cells[row][col].content = self.cells[row][col].content

        # Update cells
        self.cells = new_cells

    @classmethod
    def autogrid(
        cls,
        parent: Any,
        content_funcs: list,
        rows: int | None = None,
        cols: int | None = None,
        x: PositionType = "5%",
        y: PositionType = "5%",
        width: PositionType = "90%",
        height: PositionType = "90%",
        padding: float = 5.0,
        title: str | None = None,
        title_height: PositionType = "10%",
        title_align: str = "center",
        column_major: bool = True,  # Use column-major order by default
        **kwargs,  # Accept any additional parameters
    ) -> Grid:
        """Create a grid and automatically place content into cells.

        This method automatically determines the appropriate grid dimensions
        and places the provided content functions into the grid cells.

        Args:
            parent: The parent Slide object
            content_funcs: List of content functions to place in grid cells
            rows: Number of rows (if None, calculated automatically)
            cols: Number of columns (if None, calculated automatically)
            x: X position of the grid (default: "5%")
            y: Y position of the grid (default: "5%")
            width: Width of the grid (default: "90%")
            height: Height of the grid (default: "90%")
            padding: Padding between cells (default: 5.0)
            title: Optional title for the grid (default: None)
            title_height: Height of the title area (default: "10%")
            title_align: Text alignment for the title, one of "left", "center", "right" (default: "center")
            column_major: Whether to fill cells in column-major order (default: True)
                         When True, fills cells down columns first, resulting in a visual layout
                         that matches the specified rows and columns when content is added sequentially.
                         When False, fills cells across rows first.
            **kwargs: Additional parameters (ignored)

        Returns:
            The created Grid object
        """
        # Calculate grid dimensions if not specified
        num_items = len(content_funcs)
        if num_items == 0:
            return cls(parent, x=x, y=y, width=width, height=height)

        if rows is None and cols is None:
            # Determine optimal grid dimensions
            n_cols = max(1, int(num_items**0.5))
            n_rows = (num_items + n_cols - 1) // n_cols
        elif rows is None and cols is not None:
            n_cols = cols
            n_rows = (num_items + n_cols - 1) // n_cols
        elif cols is None and rows is not None:
            n_rows = rows
            n_cols = (num_items + n_rows - 1) // n_rows
        else:
            n_rows, n_cols = rows or 1, cols or 1

        # Adjust grid position and dimensions if a title is provided
        adjusted_y: PositionType = y
        adjusted_height: PositionType = height
        title_y: PositionType = y

        if title:
            slide_height_emu = getattr(parent, "_slide_height", 6858000)
            adjusted_y, adjusted_height = shift_band(y, height, title_height, slide_height_emu)

        # Create the grid
        grid = cls(
            parent=parent,
            x=x,
            y=adjusted_y,
            width=width,
            height=adjusted_height,
            rows=n_rows,
            cols=n_cols,
            padding=padding,
        )

        # Add title if provided
        if title:
            parent.add_text(
                text=title,
                x=x,
                y=title_y,
                width=width,
                height=title_height,
                font_size=24,
                font_bold=True,
                align=title_align,
            )

        # Place content into grid cells
        # Using column-major order (filling columns first) to better match visual expectations
        # when users specify (rows, cols)
        col_idx = 0
        row_idx = 0

        for func in content_funcs:
            # Content functions that can receive x/y/width/height get the
            # cell position; zero-argument functions are called as-is.
            wrapper = func if _accepts_position(func) else (lambda _f=func, **kwargs: _f())

            # Add content to the current cell using the wrapper
            grid.add_to_cell(
                row=row_idx,
                col=col_idx,
                content_func=wrapper,
            )

            # Move to next cell (column-major order: increment row first, then column)
            row_idx += 1
            if row_idx >= n_rows:
                row_idx = 0
                col_idx += 1

            # Stop if we've filled the grid
            if col_idx >= n_cols:
                break

        return grid

    @classmethod
    def autogrid_pyplot(
        cls,
        parent: Any,
        figures: list,
        rows: int | None = None,
        cols: int | None = None,
        x: PositionType = "5%",
        y: PositionType = "5%",
        width: PositionType = "90%",
        height: PositionType = "90%",
        padding: float = 5.0,
        title: str | None = None,
        title_height: PositionType = "10%",
        title_align: str = "center",
        dpi: int = 300,
        file_format: str = "png",
        column_major: bool = True,  # Use column-major order by default
        **kwargs,  # Accept any additional parameters
    ) -> Grid:
        """Create a grid and automatically place matplotlib figures into cells.

        This method automatically determines the appropriate grid dimensions
        and places the provided matplotlib figures into the grid cells.

        Args:
            parent: The parent Slide object
            figures: List of matplotlib figures to place in grid cells
            rows: Number of rows (if None, calculated automatically)
            cols: Number of columns (if None, calculated automatically)
            x: X position of the grid (default: "5%")
            y: Y position of the grid (default: "5%")
            width: Width of the grid (default: "90%")
            height: Height of the grid (default: "90%")
            padding: Padding between cells (default: 5.0)
            title: Optional title for the grid (default: None)
            title_height: Height of the title area (default: "10%")
            title_align: Text alignment for the title, one of "left", "center", "right" (default: "center")
            dpi: Resolution for saved figures (default: 300)
            file_format: Image format for saved figures (default: "png")
            column_major: Whether to fill cells in column-major order (default: True)
                         When True, fills cells down columns first, resulting in a visual layout
                         that matches the specified rows and columns when content is added sequentially.
                         When False, fills cells across rows first.
            **kwargs: Additional parameters (ignored)

        Returns:
            The created Grid object
        """
        from easypptx.pyplot import figure_to_stream

        # Render each figure in memory and create content functions
        content_funcs = []
        for fig in figures:
            stream = figure_to_stream(fig, dpi=dpi, file_format=file_format)

            # Bind the stream via a default argument to avoid loop-variable capture
            def add_image_func(image_stream=stream, **kwargs):
                return parent.add_image(
                    image_path=image_stream,
                    x=kwargs.get("x", "10%"),
                    y=kwargs.get("y", "10%"),
                    width=kwargs.get("width", "80%"),
                    height=kwargs.get("height", "80%"),
                )

            content_funcs.append(add_image_func)

        return cls.autogrid(
            parent=parent,
            content_funcs=content_funcs,
            rows=rows,
            cols=cols,
            x=x,
            y=y,
            width=width,
            height=height,
            padding=padding,
            title=title,
            title_height=title_height,
            title_align=title_align,
            column_major=column_major,
        )


class GridFlatIterator:
    """Flat iterator for a Grid, like matplotlib's subplot.flat.

    This iterator provides a way to loop through all cells in a grid in a flattened manner,
    regardless of their row and column positions.

    Attributes:
        grid: The Grid object to iterate over
        current_index: The current index in the flattened grid
        total_cells: The total number of cells in the grid
    """

    def __init__(self, grid: Grid):
        """Initialize a flat iterator for the grid.

        Args:
            grid: The Grid object to iterate over
        """
        self.grid = grid
        self.current_index = 0
        self.total_cells = grid.rows * grid.cols

    def __iter__(self):
        """Return the iterator itself."""
        return self

    def __next__(self):
        """Get the next cell in the flattened grid.

        Returns:
            The next GridCell object

        Raises:
            StopIteration: When all cells have been iterated through
        """
        if self.current_index >= self.total_cells:
            raise StopIteration

        row = self.current_index // self.grid.cols
        col = self.current_index % self.grid.cols
        self.current_index += 1

        return self.grid.cells[row][col]
