"""Grid layout module for EasyPPTX."""

from collections.abc import Callable
from typing import Any

from easypptx.slide import PositionType


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

    def __repr__(self) -> str:
        """Return string representation of the cell."""
        return (
            f"GridCell(row={self.row}, col={self.col}, x={self.x}, "
            f"y={self.y}, width={self.width}, height={self.height})"
        )


class OutOfBoundsError(IndexError):
    """Exception raised when grid coordinates are out of bounds."""


class CellMergeError(ValueError):
    """Exception raised when there's an issue with cell merging."""


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

        # Loop through all cells
        for cell in grid:
            print(cell)

        # Loop through cells linearly (flattened)
        for cell in grid.flat:
            print(cell.row, cell.col)
        ```
    """

    def __init__(
        self,
        parent: Any,
        x: PositionType = "0%",
        y: PositionType = "0%",
        width: PositionType = "100%",
        height: PositionType = "100%",
        rows: int = 1,
        cols: int = 1,
        padding: float = 5.0,
    ) -> None:
        """Initialize a Grid layout.

        Args:
            parent: The parent Slide or Grid object
            x: X position of the grid (default: "0%")
            y: Y position of the grid (default: "0%")
            width: Width of the grid (default: "100%")
            height: Height of the grid (default: "100%")
            rows: Number of rows (default: 1)
            cols: Number of columns (default: 1)
            padding: Padding between cells as percentage of cell size (default: 5.0)
        """
        self.parent = parent
        self.x = x
        self.y = y
        self.width = width
        self.height = height
        self.rows = rows
        self.cols = cols
        self.padding = padding

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
        # Default value if we can't get it (equivalent to 10 inches)
        return 9144000  # 10 inches in EMUs

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

    def _create_cells(self) -> list:
        """Create the grid cells based on the layout.

        Returns:
            2D array of GridCell objects
        """
        cells = []

        # Convert percentage values to floats for calculations
        padding_factor = self.padding / 100.0

        # Calculate the width and height of each cell including padding
        cell_width_percent = 100.0 / self.cols
        cell_height_percent = 100.0 / self.rows

        # Calculate the effective width and height of each cell (excluding padding)
        effective_cell_width = cell_width_percent * (1 - padding_factor)
        effective_cell_height = cell_height_percent * (1 - padding_factor)

        # Half of the padding (as percentage of total grid size)
        half_padding_width = (cell_width_percent * padding_factor) / 2
        half_padding_height = (cell_height_percent * padding_factor) / 2

        # Create cells
        for row in range(self.rows):
            cell_row = []
            for col in range(self.cols):
                # Calculate cell position
                x_percent = (col * cell_width_percent) + half_padding_width
                y_percent = (row * cell_height_percent) + half_padding_height

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

    def add_to_cell(self, row: int, col: int, content_func: Callable, **kwargs) -> Any:
        """Add content to a specific cell in the grid.

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
        # Get the cell
        cell = self.get_cell(row, col)

        # Check if the cell is part of a merged cell
        if cell.is_spanned:
            raise CellMergeError("Cell is part of a merged cell")

        # Calculate the absolute position based on the grid's position
        # Convert grid position percentage to float
        if isinstance(self.x, str) and self.x.endswith("%"):
            grid_x_percent = float(self.x.strip("%"))
        else:
            # Convert inches to percentage based on actual slide width
            slide_width_inches = self._slide_width / 914400  # Convert EMUs to inches
            grid_x_percent = (float(self.x) / slide_width_inches) * 100

        if isinstance(self.y, str) and self.y.endswith("%"):
            grid_y_percent = float(self.y.strip("%"))
        else:
            # Convert inches to percentage based on actual slide height
            slide_height_inches = self._slide_height / 914400  # Convert EMUs to inches
            grid_y_percent = (float(self.y) / slide_height_inches) * 100

        # Calculate absolute position
        cell_x_percent = float(cell.x.strip("%"))
        cell_y_percent = float(cell.y.strip("%"))

        # Handle the width/height values to correctly handle both string and float types
        width_value = (
            float(self.width.strip("%")) if isinstance(self.width, str) and "%" in self.width else float(self.width)
        )
        height_value = (
            float(self.height.strip("%")) if isinstance(self.height, str) and "%" in self.height else float(self.height)
        )

        abs_x_percent = grid_x_percent + (cell_x_percent * width_value / 100)
        abs_y_percent = grid_y_percent + (cell_y_percent * height_value / 100)

        # Calculate absolute width and height
        cell_width_percent = float(cell.width.strip("%"))
        cell_height_percent = float(cell.height.strip("%"))

        abs_width_percent = cell_width_percent * width_value / 100
        abs_height_percent = cell_height_percent * height_value / 100

        # Format as percentage strings
        kwargs["x"] = f"{abs_x_percent:.2f}%"
        kwargs["y"] = f"{abs_y_percent:.2f}%"
        kwargs["width"] = f"{abs_width_percent:.2f}%"
        kwargs["height"] = f"{abs_height_percent:.2f}%"

        # We previously were going to set word_wrap, but slide.add_text doesn't accept this parameter
        # The word_wrap flag is now set internally in the add_text method

        # Call the content function with the calculated position
        content = content_func(**kwargs)

        # Store the content in the cell
        cell.content = content

        return content

    def add_grid_to_cell(
        self,
        row: int,
        col: int,
        rows: int = 1,
        cols: int = 1,
        padding: float = 5.0,
    ) -> "Grid":
        """Add a nested grid to a specific cell.

        Args:
            row: Row index (0-based)
            col: Column index (0-based)
            rows: Number of rows in the nested grid (default: 1)
            cols: Number of columns in the nested grid (default: 1)
            padding: Padding between cells as percentage of cell size (default: 5.0)

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

        # Calculate absolute position for the nested grid
        if isinstance(self.x, str) and self.x.endswith("%"):
            grid_x_percent = float(self.x.strip("%"))
        else:
            # Convert inches to percentage based on actual slide width
            slide_width_inches = self._slide_width / 914400  # Convert EMUs to inches
            grid_x_percent = (float(self.x) / slide_width_inches) * 100

        if isinstance(self.y, str) and self.y.endswith("%"):
            grid_y_percent = float(self.y.strip("%"))
        else:
            # Convert inches to percentage based on actual slide height
            slide_height_inches = self._slide_height / 914400  # Convert EMUs to inches
            grid_y_percent = (float(self.y) / slide_height_inches) * 100

        # Calculate absolute position
        cell_x_percent = float(cell.x.strip("%"))
        cell_y_percent = float(cell.y.strip("%"))

        # Handle the width/height values to correctly handle both string and float types
        width_value = (
            float(self.width.strip("%")) if isinstance(self.width, str) and "%" in self.width else float(self.width)
        )
        height_value = (
            float(self.height.strip("%")) if isinstance(self.height, str) and "%" in self.height else float(self.height)
        )

        abs_x_percent = grid_x_percent + (cell_x_percent * width_value / 100)
        abs_y_percent = grid_y_percent + (cell_y_percent * height_value / 100)

        # Calculate absolute width and height
        cell_width_percent = float(cell.width.strip("%"))
        cell_height_percent = float(cell.height.strip("%"))

        abs_width_percent = cell_width_percent * width_value / 100
        abs_height_percent = cell_height_percent * height_value / 100

        # Create the nested grid
        nested_grid = Grid(
            parent=self.parent,
            x=f"{abs_x_percent:.2f}%",
            y=f"{abs_y_percent:.2f}%",
            width=f"{abs_width_percent:.2f}%",
            height=f"{abs_height_percent:.2f}%",
            rows=rows,
            cols=cols,
            padding=padding,
        )

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

    def __getitem__(self, key):
        """Access a cell or range of cells using indexing.

        Args:
            key: A tuple of (row, col) or a single index for flattened access

        Returns:
            The requested GridCell object or a list of cells

        Raises:
            OutOfBoundsError: If the requested cell is out of bounds
            TypeError: If the key is not in the right format
        """
        if isinstance(key, tuple) and len(key) == 2:
            # Access as grid[row, col]
            row, col = key
            return self.get_cell(row, col)
        elif isinstance(key, int):
            # Access as grid[flat_index]
            if key < 0 or key >= self.rows * self.cols:
                raise OutOfBoundsError(f"Flat index {key} is out of bounds")
            row = key // self.cols
            col = key % self.cols
            return self.cells[row][col]
        else:
            raise TypeError("Grid indices must be integers or tuples of the form (row, col)")

    @property
    def flat(self):
        """Flat iterator for this grid, similar to matplotlib's subplot.flat.

        Returns:
            A flat iterator over all cells in the grid
        """
        return GridFlatIterator(self)

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
    ) -> "Grid":
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

        Returns:
            The created Grid object
        """
        # Calculate grid dimensions if not specified
        num_items = len(content_funcs)
        if num_items == 0:
            return cls(parent, x=x, y=y, width=width, height=height)

        if rows is None and cols is None:
            # Determine optimal grid dimensions
            cols = max(1, int(num_items**0.5))
            rows = (num_items + cols - 1) // cols
        elif rows is None:
            rows = (num_items + cols - 1) // cols
        elif cols is None:
            cols = (num_items + rows - 1) // rows

        # Adjust grid position and dimensions if a title is provided
        adjusted_y = y
        adjusted_height = height

        if title and isinstance(y, str) and y.endswith("%"):
            y_percent = float(y.strip("%"))
            title_height_percent = float(str(title_height).strip("%"))
            # Adjust y position and height for the grid
            adjusted_y = f"{y_percent:.2f}%"
            title_y = adjusted_y
            adjusted_y = f"{(y_percent + title_height_percent):.2f}%"

            # Adjust height to account for title
            if isinstance(height, str) and height.endswith("%"):
                height_percent = float(height.strip("%"))
                adjusted_height = f"{(height_percent - title_height_percent):.2f}%"

        # Create the grid
        grid = cls(
            parent=parent,
            x=x,
            y=adjusted_y,
            width=width,
            height=adjusted_height,
            rows=rows,
            cols=cols,
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
                align="center",
            )

        # Place content into grid cells
        row_idx = 0
        col_idx = 0

        for func in content_funcs:
            # Create a wrapper function factory to properly capture the current func
            def create_wrapper(content_func):
                def position_agnostic_wrapper(**kwargs):
                    return content_func()

                return position_agnostic_wrapper

            # Create a wrapper specifically for this function
            wrapper = create_wrapper(func)

            # Add content to the current cell using the wrapper
            grid.add_to_cell(
                row=row_idx,
                col=col_idx,
                content_func=wrapper,
            )

            # Move to next cell
            col_idx += 1
            if col_idx >= cols:
                col_idx = 0
                row_idx += 1

            # Stop if we've filled the grid
            if row_idx >= rows:
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
        dpi: int = 300,
        file_format: str = "png",
    ) -> "Grid":
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
            dpi: Resolution for saved figures (default: 300)
            file_format: Image format for saved figures (default: "png")

        Returns:
            The created Grid object
        """
        import os
        import tempfile

        # Create content functions from matplotlib figures
        content_funcs = []

        # Save each figure to a temporary file and create content functions
        temp_files = []
        for fig in figures:
            # Create a temporary file using context manager
            with tempfile.NamedTemporaryFile(suffix=f".{file_format}", delete=False) as temp_file:
                temp_path = temp_file.name
                temp_files.append(temp_path)

            # Save the figure to the temporary file
            fig.savefig(temp_path, dpi=dpi, format=file_format, bbox_inches="tight")

            # Create a closure with an explicit parameter to avoid loop variable capture issues
            def create_content_func(image_path):
                def add_image_func(**kwargs):
                    return parent.add_image(
                        image_path=image_path,
                        x=kwargs.get("x", "10%"),
                        y=kwargs.get("y", "10%"),
                        width=kwargs.get("width", "80%"),
                        height=kwargs.get("height", "80%"),
                    )

                return add_image_func

            # Add the content function to the list with temp_path bound to a parameter
            content_funcs.append(create_content_func(image_path=temp_path))

        # Create the grid
        try:
            grid = cls.autogrid(
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
            )
        finally:
            # Clean up temporary files
            for temp_file_path in temp_files:
                try:
                    if os.path.exists(temp_file_path):
                        os.unlink(temp_file_path)
                except OSError as e:
                    import logging

                    logging.warning(f"Failed to remove temporary file {temp_file_path}: {e}")

        return grid


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
