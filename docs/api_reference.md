# API Reference

This API reference provides detailed information about the classes and methods in EasyPPTX.

## Presentation Class

The `Presentation` class is the main entry point for creating and manipulating PowerPoint presentations.

```python
class Presentation:
    def __init__(self,
                 aspect_ratio: Optional[str] = "16:9",
                 width_inches: Optional[float] = None,
                 height_inches: Optional[float] = None,
                 template_path: Optional[str] = None,
                 default_bg_color: Optional[Union[str, Tuple[int, int, int]]] = None) -> None:
        """Initialize a new empty presentation.

        Args:
            aspect_ratio: Predefined aspect ratio, one of "16:9" (default), "4:3", "16:10", "A4", "LETTER"
            width_inches: Custom width in inches (overrides aspect_ratio if specified)
            height_inches: Custom height in inches (overrides aspect_ratio if specified)
            template_path: Path to a reference PowerPoint template to use for styles (default: None)
            default_bg_color: Default background color for slides as string name or RGB tuple (default: None)

        Raises:
            ValueError: If an invalid aspect ratio is specified
            FileNotFoundError: If the template file doesn't exist
        """

    @classmethod
    def open(cls, file_path: Union[str, Path]) -> "Presentation":
        """Open an existing PowerPoint presentation.

        Args:
            file_path: Path to the PowerPoint file to open

        Returns:
            A new Presentation object with the loaded presentation

        Raises:
            FileNotFoundError: If the specified file doesn't exist
            ValueError: If the file is not a valid PowerPoint file
        """

    def add_slide(self, layout_index: int = None, bg_color: Optional[Union[str, Tuple[int, int, int]]] = None) -> Slide:
        """Add a new slide to the presentation.

        Args:
            layout_index: Index of the slide layout to use (default: None uses blank layout)
            bg_color: Background color for this slide, overrides default (default: None)

        Returns:
            A new Slide object
        """

    @property
    def slides(self) -> List[Slide]:
        """Get a list of all slides in the presentation.

        Returns:
            List of Slide objects
        """

    def save(self, file_path: Union[str, Path]) -> None:
        """Save the presentation to a file.

        Args:
            file_path: Path where the presentation should be saved
        """
```

## Slide Class

The `Slide` class represents a slide in a PowerPoint presentation and provides methods for adding content.

```python
class Slide:
    def __init__(self, pptx_slide: PPTXSlide) -> None:
        """Initialize a Slide object.

        Args:
            pptx_slide: The python-pptx Slide object
        """

    def add_text(
        self,
        text: str,
        x: PositionType = 1.0,
        y: PositionType = 1.0,
        width: PositionType = 8.0,
        height: PositionType = 1.0,
        font_size: int = 18,
        font_bold: bool = False,
        font_italic: bool = False,
        font_name: str = "Meiryo",
        align: str = "left",
        vertical: str = "top",
        color: Optional[Union[str, Tuple[int, int, int]]] = None,
    ) -> PPTXShape:
        """Add a text box to the slide.

        Args:
            text: The text content
            x: X position in inches or percentage (default: 1.0)
            y: Y position in inches or percentage (default: 1.0)
            width: Width in inches or percentage (default: 8.0)
            height: Height in inches or percentage (default: 1.0)
            font_size: Font size in points (default: 18)
            font_bold: Whether text should be bold (default: False)
            font_italic: Whether text should be italic (default: False)
            font_name: Font name (default: "Meiryo")
            align: Text alignment, one of "left", "center", "right" (default: "left")
            vertical: Vertical alignment, one of "top", "middle", "bottom" (default: "top")
            color: Text color as string name from COLORS dict or RGB tuple (default: None)

        Returns:
            The created shape object
        """

    def add_image(
        self,
        image_path: str,
        x: PositionType = 1.0,
        y: PositionType = 1.0,
        width: Optional[PositionType] = None,
        height: Optional[PositionType] = None,
    ) -> PPTXShape:
        """Add an image to the slide.

        Args:
            image_path: Path to the image file
            x: X position in inches or percentage (default: 1.0)
            y: Y position in inches or percentage (default: 1.0)
            width: Width in inches or percentage (default: None, maintains aspect ratio)
            height: Height in inches or percentage (default: None, maintains aspect ratio)

        Returns:
            The created picture shape

        Raises:
            FileNotFoundError: If the image file doesn't exist
        """

    def add_shape(
        self,
        shape_type: MSO_SHAPE = MSO_SHAPE.RECTANGLE,
        x: PositionType = 1.0,
        y: PositionType = 1.0,
        width: PositionType = 5.0,
        height: PositionType = 1.0,
        fill_color: Optional[Union[str, Tuple[int, int, int]]] = None,
    ) -> PPTXShape:
        """Add a shape to the slide.

        Args:
            shape_type: The shape type (default: MSO_SHAPE.RECTANGLE)
            x: X position in inches or percentage (default: 1.0)
            y: Y position in inches or percentage (default: 1.0)
            width: Width in inches or percentage (default: 5.0)
            height: Height in inches or percentage (default: 1.0)
            fill_color: Fill color as string name from COLORS dict or RGB tuple (default: None)

        Returns:
            The created shape object
        """

    def add_multiple_objects(
        self,
        objects_data: List[dict],
        layout: str = "grid",
        padding_percent: float = 5.0,
        start_x: PositionType = "5%",
        start_y: PositionType = "5%",
        width: PositionType = "90%",
        height: PositionType = "90%",
    ) -> List[PPTXShape]:
        """Add multiple objects to the slide with automatic alignment.

        Args:
            objects_data: List of dictionaries containing object data
                Each dict should have 'type' ('text', 'image', or 'shape') and type-specific parameters
            layout: Layout type ('grid', 'horizontal', 'vertical')
            padding_percent: Padding between objects as percentage of container
            start_x: Starting X position of container in inches or percentage
            start_y: Starting Y position of container in inches or percentage
            width: Width of container in inches or percentage
            height: Height of container in inches or percentage

        Returns:
            List of created shape objects
        """

    def clear(self) -> None:
        """Remove all shapes from the slide."""

    @property
    def title(self) -> Optional[str]:
        """Get the slide title.

        Returns:
            The slide title if it exists, None otherwise
        """

    @title.setter
    def title(self, value: str) -> None:
        """Set the slide title.

        Args:
            value: The title text
        """

    def set_background_color(self, color: Union[str, Tuple[int, int, int]]) -> None:
        """Set the background color of the slide.

        Args:
            color: Background color as string name from COLORS dict or RGB tuple
        """
```

## Text Class

The `Text` class provides methods for adding and formatting text on slides.

```python
class Text:
    def __init__(self, slide_obj: "Slide") -> None:
        """Initialize a Text object.

        Args:
            slide_obj: The Slide object to add text to
        """

    def add_title(
        self,
        text: str,
        font_size: int = 44,
        font_name: str = "Meiryo",
        color: Optional[Union[str, Tuple[int, int, int]]] = "black",
        align: str = "center",
        x: PositionType = "10%",
        y: PositionType = "5%",
        width: PositionType = "80%",
        height: PositionType = "15%",
    ) -> PPTXShape:
        """Add a title to the slide.

        Args:
            text: The title text
            font_size: Font size in points (default: 44)
            font_name: Font name (default: "Meiryo")
            color: Text color as string name from COLORS dict or RGB tuple (default: "black")
            align: Text alignment, one of "left", "center", "right" (default: "center")
            x: X position in inches or percentage (default: "10%")
            y: Y position in inches or percentage (default: "5%")
            width: Width in inches or percentage (default: "80%")
            height: Height in inches or percentage (default: "15%")

        Returns:
            The created shape object
        """

    def add_paragraph(
        self,
        text: str,
        x: PositionType = 1.0,
        y: PositionType = 2.0,
        width: PositionType = 8.0,
        height: PositionType = 1.0,
        font_size: int = 18,
        font_bold: bool = False,
        font_italic: bool = False,
        font_name: str = "Meiryo",
        align: str = "left",
        vertical: str = "top",
        color: Optional[Union[str, Tuple[int, int, int]]] = "black",
    ) -> PPTXShape:
        """Add a paragraph of text to the slide.

        Args:
            text: The paragraph text
            x: X position in inches or percentage (default: 1.0)
            y: Y position in inches or percentage (default: 2.0)
            width: Width in inches or percentage (default: 8.0)
            height: Height in inches or percentage (default: 1.0)
            font_size: Font size in points (default: 18)
            font_bold: Whether text should be bold (default: False)
            font_italic: Whether text should be italic (default: False)
            font_name: Font name (default: "Meiryo")
            align: Text alignment, one of "left", "center", "right" (default: "left")
            vertical: Vertical alignment, one of "top", "middle", "bottom" (default: "top")
            color: Text color as string name from COLORS dict or RGB tuple (default: "black")

        Returns:
            The created shape object
        """

    @staticmethod
    def format_text_frame(
        text_frame: TextFrame,
        font_size: Optional[int] = None,
        font_bold: Optional[bool] = None,
        font_italic: Optional[bool] = None,
        font_name: Optional[str] = None,
        color: Optional[Union[str, Tuple[int, int, int]]] = None,
        align: Optional[str] = None,
        vertical: Optional[str] = None,
    ) -> None:
        """Format an existing text frame.

        Args:
            text_frame: The text frame to format
            font_size: Font size in points (default: None)
            font_bold: Whether text should be bold (default: None)
            font_italic: Whether text should be italic (default: None)
            font_name: Font name (default: None)
            color: Text color as string name from COLORS dict or RGB tuple (default: None)
            align: Text alignment, one of "left", "center", "right" (default: None)
            vertical: Vertical alignment, one of "top", "middle", "bottom" (default: None)
        """
```

## Position Type

EasyPPTX uses a special type for position parameters that supports both absolute and percentage-based positioning:

```python
# Type for position parameters
PositionType = Union[float, str]

# Examples:
x = 1.0       # 1 inch (absolute)
x = "50%"     # 50% of slide width (percentage)
```

## Constants

### Aspect Ratios

The `Presentation` class defines standard aspect ratios:

```python
ASPECT_RATIOS = {
    "16:9": (13.33, 7.5),    # Widescreen (default)
    "4:3": (10, 7.5),        # Standard
    "16:10": (13.33, 8.33),  # Widescreen alternative
    "A4": (11.69, 8.27),     # A4 paper size
    "LETTER": (11, 8.5),     # US Letter paper size
}
```

### Colors

The `Presentation` class defines an expanded color palette:

```python
COLORS = {
    "black": RGBColor(0x10, 0x10, 0x10),
    "darkgray": RGBColor(0x40, 0x40, 0x40),
    "gray": RGBColor(0x80, 0x80, 0x80),
    "lightgray": RGBColor(0xD0, 0xD0, 0xD0),
    "red": RGBColor(0xFF, 0x40, 0x40),
    "green": RGBColor(0x40, 0xFF, 0x40),
    "blue": RGBColor(0x40, 0x40, 0xFF),
    "white": RGBColor(0xFF, 0xFF, 0xFF),
    "yellow": RGBColor(0xFF, 0xD7, 0x00),
    "cyan": RGBColor(0x00, 0xE5, 0xFF),
    "magenta": RGBColor(0xFF, 0x00, 0xFF),
    "orange": RGBColor(0xFF, 0xA5, 0x00)
}
```

### Alignments

The `Presentation` class defines alignment options:

```python
# Text alignment
ALIGN = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT
}

# Vertical alignment
VERTICAL = {
    "top": MSO_ANCHOR.TOP,
    "middle": MSO_ANCHOR.MIDDLE,
    "bottom": MSO_ANCHOR.BOTTOM
}
```

## Shape Types

EasyPPTX uses the `MSO_SHAPE` enum from python-pptx for shape types:

```python
from pptx.enum.shapes import MSO_SHAPE

# Examples:
MSO_SHAPE.RECTANGLE
MSO_SHAPE.OVAL
MSO_SHAPE.ROUNDED_RECTANGLE
MSO_SHAPE.ACTION_BUTTON_HOME
```

For a complete list of shape types, refer to the [python-pptx documentation](https://python-pptx.readthedocs.io/en/latest/api/enum/MsoAutoShapeType.html).

## Grid Class

The `Grid` class provides a powerful layout system for organizing content on slides.

```python
class Grid:
    def __init__(self,
                 parent: Any,
                 x: PositionType = "0%",
                 y: PositionType = "0%",
                 width: PositionType = "100%",
                 height: PositionType = "100%",
                 rows: int = 1,
                 cols: int = 1,
                 padding: float = 5.0) -> None:
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

    def add_grid_to_cell(self, row: int, col: int, rows: int = 1, cols: int = 1, padding: float = 5.0) -> "Grid":
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

    def __iter__(self):
        """Make Grid iterable to loop through all cells.

        Returns:
            Iterator over all grid cells
        """

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

    @property
    def flat(self):
        """Flat iterator for this grid, similar to matplotlib's subplot.flat.

        Returns:
            A flat iterator over all cells in the grid
        """

    @classmethod
    def autogrid(cls, parent: Any, content_funcs: list, rows: int | None = None, cols: int | None = None,
                x: PositionType = "5%", y: PositionType = "5%", width: PositionType = "90%",
                height: PositionType = "90%", padding: float = 5.0, title: str | None = None,
                title_height: PositionType = "10%") -> "Grid":
        """Create a grid and automatically place content into cells.

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

    @classmethod
    def autogrid_pyplot(cls, parent: Any, figures: list, rows: int | None = None, cols: int | None = None,
                       x: PositionType = "5%", y: PositionType = "5%", width: PositionType = "90%",
                       height: PositionType = "90%", padding: float = 5.0, title: str | None = None,
                       title_height: PositionType = "10%", dpi: int = 300, file_format: str = "png") -> "Grid":
        """Create a grid and automatically place matplotlib figures into cells.

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
```

## GridCell Class

The `GridCell` class represents a cell in a grid layout.

```python
class GridCell:
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
```

## GridFlatIterator Class

The `GridFlatIterator` class provides a way to iterate through grid cells in a flattened manner, similar to matplotlib's subplot.flat.

```python
class GridFlatIterator:
    def __init__(self, grid: Grid):
        """Initialize a flat iterator for the grid.

        Args:
            grid: The Grid object to iterate over
        """

    def __iter__(self):
        """Return the iterator itself."""

    def __next__(self):
        """Get the next cell in the flattened grid.

        Returns:
            The next GridCell object

        Raises:
            StopIteration: When all cells have been iterated through
        """
```
