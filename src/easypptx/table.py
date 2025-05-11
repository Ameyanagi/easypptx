"""Table handling module for EasyPPTX."""

from typing import Any, List, Optional, Union

import pandas as pd
from pptx.table import Table as PPTXTable
from pptx.util import Inches, Pt

# Type for position parameters - accepts either percentage or absolute values
PositionType = Union[float, str]


class Table:
    """Class for handling table operations in PowerPoint slides.

    This class provides methods for creating and manipulating tables on slides.

    Examples:
        ```python
        # Create a table object
        table = Table(slide)

        # Add a simple table
        table.add([["Header 1", "Header 2"], ["Value 1", "Value 2"]])

        # Add a table from pandas DataFrame
        import pandas as pd
        df = pd.DataFrame({'A': [1, 2], 'B': [3, 4]})
        table.from_dataframe(df, x=2, y=2)
        ```
    """

    def __init__(self, slide_obj: "Slide") -> None:  # noqa: F821
        """Initialize a Table object.

        Args:
            slide_obj: The Slide object to add tables to
        """
        self.slide = slide_obj

    def add(
        self,
        data: List[List[Any]],
        x: PositionType = 1.0,
        y: PositionType = 1.0,
        width: Optional[PositionType] = None,
        height: Optional[PositionType] = None,
        first_row_header: bool = True,
        style: Optional[int] = None,
    ) -> PPTXTable:
        """Add a table to the slide.

        Args:
            data: 2D list of table data
            x: X position in inches or percentage (default: 1.0)
            y: Y position in inches or percentage (default: 1.0)
            width: Total width in inches or percentage (default: None, auto-sized)
            height: Total height in inches or percentage (default: None, auto-sized)
            first_row_header: Whether to format the first row as a header (default: True)
            style: Table style ID (default: None)

        Returns:
            The created table object
        """
        if not data:
            raise ValueError("Table data cannot be empty")

        rows = len(data)
        cols = len(data[0])

        # Ensure all rows have the same number of columns
        for row in data:
            if len(row) != cols:
                raise ValueError("All rows must have the same number of columns")

        # Get slide dimensions for percentage conversion
        slide_width = self.slide._get_slide_width()
        slide_height = self.slide._get_slide_height()

        # Convert position values to inches
        x_inches = self.slide._convert_position(x, slide_width)
        y_inches = self.slide._convert_position(y, slide_height)

        # Create table shape
        if width is None:
            # Default width based on columns
            width_inches = cols * 2.0
        else:
            width_inches = self.slide._convert_position(width, slide_width)

        if height is None:
            # Default height based on rows
            height_inches = rows * 0.5
        else:
            height_inches = self.slide._convert_position(height, slide_height)

        table_shape = self.slide.pptx_slide.shapes.add_table(
            rows, cols, Inches(x_inches), Inches(y_inches), Inches(width_inches), Inches(height_inches)
        )
        table = table_shape.table

        # Fill table data
        for i, row_data in enumerate(data):
            for j, cell_data in enumerate(row_data):
                cell = table.cell(i, j)
                cell.text = str(cell_data)

                # Format header row
                if first_row_header and i == 0:
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.bold = True
                        paragraph.font.size = Pt(14)

        # Apply table style if specified
        if style is not None:
            table.style = style

        return table

    def from_dataframe(
        self,
        df: "pd.DataFrame",
        x: PositionType = 1.0,
        y: PositionType = 1.0,
        width: Optional[PositionType] = None,
        height: Optional[PositionType] = None,
        include_index: bool = False,
        first_row_header: bool = True,
        style: Optional[int] = None,
    ) -> PPTXTable:
        """Add a table from a pandas DataFrame.

        Args:
            df: Pandas DataFrame
            x: X position in inches (default: 1.0)
            y: Y position in inches (default: 1.0)
            width: Total width in inches (default: None, auto-sized)
            height: Total height in inches (default: None, auto-sized)
            include_index: Whether to include DataFrame index (default: False)
            first_row_header: Whether to format column names as headers (default: True)
            style: Table style ID (default: None)

        Returns:
            The created table object
        """
        # Convert DataFrame to list format
        if include_index:
            data = [list(df.columns)]
            for idx, row in df.iterrows():
                data.append([idx] + list(row))
        else:
            data = [list(df.columns)] + df.values.tolist()

        return self.add(
            data=data,
            x=x,
            y=y,
            width=width,
            height=height,
            first_row_header=first_row_header,
            style=style,
        )
