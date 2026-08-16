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
                 reference_pptx: Optional[str] = None,
                 blank_layout_index: Optional[int] = None,
                 default_bg_color: Optional[Union[str, Tuple[int, int, int]]] = None,
                 template_toml: Optional[str] = None,
                 theme: Union[str, Theme, None] = None) -> None:
        """Initialize a new empty presentation.

        Args:
            aspect_ratio: Predefined aspect ratio, one of "16:9" (default), "4:3", "16:10", "A4", "LETTER"
            width_inches: Custom width in inches (overrides aspect_ratio if specified)
            height_inches: Custom height in inches (overrides aspect_ratio if specified)
            template_path: Path to a reference PowerPoint template to use for styles (default: None)
            reference_pptx: Path to a custom reference PPTX file to use (default: None)
            blank_layout_index: Index of blank layout in the slide_layouts (default: None, auto-detected)
            default_bg_color: Default background color for slides as string name or RGB tuple (default: None)
            template_toml: Path to a TOML template used as the default for all new slides (default: None)
            theme: Built-in theme name ("light", "dark", "corporate") or a Theme instance (default: None)

        Raises:
            ValueError: If an invalid aspect ratio is specified
            FileNotFoundError: If the template file doesn't exist
        """

    @classmethod
    def from_markdown(cls,
                      source: Union[str, Path],
                      theme: Union[str, Theme, None] = None,
                      template_toml: Optional[str] = None,
                      base_dir: Union[str, Path, None] = None) -> "Presentation":
        """Build a presentation from markdown text or a markdown file.

        Args:
            source: Markdown content, or a path to a .md file
            theme: Theme name or Theme instance; overrides the frontmatter theme (default: None)
            template_toml: TOML template path; overrides the frontmatter template (default: None)
            base_dir: Directory for resolving relative image paths (default: the file's directory)

        Returns:
            A Presentation with one slide per markdown section
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

    def add_slide(
        self,
        layout_index: Optional[int] = None,
        bg_color: Optional[Union[str, Tuple[int, int, int]]] = None,
        title: Optional[str] = None,
        template_toml: Union[str, Literal[False], None] = None,
        title_padding: Optional[Union[str, float]] = None,
        title_x_padding: Optional[Union[str, float]] = "5%",
        title_y_padding: Optional[Union[str, float]] = "5%",
        title_width: Optional[Union[str, float]] = "90%",
        title_height: Optional[Union[str, float]] = "10%",
        title_font_size: Optional[int] = None,
        title_align: Optional[str] = None,
        title_color: Optional[Union[str, Tuple[int, int, int]]] = None,
    ) -> Slide:
        """Add a new slide to the presentation.

        Args:
            layout_index: Index of the slide layout to use (default: None uses blank layout)
            bg_color: Background color for this slide, overrides default (default: None)
            title: Optional title text for the slide (default: None)
            template_toml: Path to a TOML template file for this slide.
                Pass False to opt out of the presentation's default template;
                None means "use the default template if one is set" (default: None)
            title_padding: Padding around the title, applies to both x and y (default: None)
            title_x_padding: Horizontal padding for the title (default: "5%")
            title_y_padding: Vertical padding for the title (default: "5%")
            title_width: Width of the title area (default: "90%")
            title_height: Height of the title area (default: "10%")
            title_font_size: Font size for the title (default: theme title size or 32)
            title_align: Text alignment for the title (default: theme title align or "center")
            title_color: Title color as a name from COLORS or an RGB tuple (default: theme title color or None)

        Returns:
            A new Slide object
        """

    @property
    def slides(self) -> List[Slide]:
        """Get a list of all slides in the presentation.

        Returns:
            List of Slide objects. Repeated access returns the same cached
            wrapper objects, so state like user_data persists.
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
        x: Optional[PositionType] = None,
        y: Optional[PositionType] = None,
        width: Optional[PositionType] = None,
        height: Optional[PositionType] = None,
        font_size: Optional[int] = None,
        font_bold: Optional[bool] = None,
        font_italic: Optional[bool] = None,
        font_name: Optional[str] = None,
        align: Optional[str] = None,
        vertical: Optional[str] = None,
        color: Optional[Union[str, Tuple[int, int, int]]] = None,
        style: Optional[TextStyle] = None,
    ) -> PPTXShape:
        """Add a text box to the slide.

        Args:
            text: The text content
            x: X position as percent (number or "N%") or in_() length (default: 5)
            y: Y position as percent or in_() length (default: 5)
            width: Width as percent or in_() length (default: 90)
            height: Height as percent or in_() length (default: 10)
            font_size: Font size in points (default: 18)
            font_bold: Whether text should be bold (default: False)
            font_italic: Whether text should be italic (default: False)
            font_name: Font name (default: "Meiryo")
            align: Text alignment, one of "left", "center", "right" (default: "left")
            vertical: Vertical alignment, one of "top", "middle", "bottom" (default: "top")
            color: Text color as string name from COLORS dict or RGB tuple (default: None)
            style: A TextStyle filling any formatting left unset (default: None)

        Returns:
            The created shape object
        """

    def add_bullets(
        self,
        items: List[Union[str, Tuple[str, int]]],
        x: PositionType = 5,
        y: PositionType = 20,
        width: PositionType = 90,
        height: PositionType = 70,
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        color: Optional[Union[str, Tuple[int, int, int]]] = None,
        align: Optional[str] = None,
        bullet: bool = True,
        style: Optional[TextStyle] = None,
    ) -> PPTXShape:
        """Add a bulleted (or plain stacked) list of paragraphs to the slide.

        New in 0.7.0.

        Args:
            items: List entries; each item is a string (level 0) or a
                (text, level) tuple for nested bullets (level 0-4)
            x: X position as percent or in_() length (default: 5)
            y: Y position as percent or in_() length (default: 20)
            width: Width as percent or in_() length (default: 90)
            height: Height as percent or in_() length (default: 70)
            font_size: Font size in points (default: 18)
            font_name: Font name (default: "Meiryo")
            color: Text color name or RGB tuple (default: None)
            align: Text alignment "left"/"center"/"right" (default: "left")
            bullet: Draw bullet characters; False gives plain stacked paragraphs (default: True)
            style: A TextStyle filling any formatting left unset (default: None)

        Returns:
            The created text box shape
        """

    def add_image(
        self,
        image_path: str,
        x: PositionType = 5,
        y: PositionType = 5,
        width: Optional[PositionType] = None,
        height: Optional[PositionType] = None,
    ) -> PPTXShape:
        """Add an image to the slide.

        Args:
            image_path: Path to the image file, or a binary file-like object
            x: X position as percent or in_() length (default: 5)
            y: Y position as percent or in_() length (default: 5)
            width: Width as percent or in_() length (default: None, maintains aspect ratio)
            height: Height as percent or in_() length (default: None, maintains aspect ratio)

        Returns:
            The created picture shape

        Raises:
            FileNotFoundError: If the image file doesn't exist
        """

    def add_pyplot(
        self,
        figure: Any,
        x: PositionType = 5,
        y: PositionType = 15,
        width: PositionType = 90,
        height: PositionType = 75,
        dpi: int = 300,
        file_format: str = "png",
        border: bool = False,
        border_color: str = "black",
        border_width: int = 1,
        shadow: bool = False,
        maintain_aspect_ratio: bool = True,
    ) -> PPTXShape:
        """Add a matplotlib (or seaborn) figure to the slide.

        New in 0.7.0.

        Args:
            figure: Matplotlib figure object (plt.figure(), plt.gcf(), or a
                seaborn plot's .figure)
            x: X position as percent or in_() length (default: 5)
            y: Y position as percent or in_() length (default: 15)
            width: Width as percent or in_() length (default: 90)
            height: Height as percent or in_() length (default: 75)
            dpi: Resolution for the figure (default: 300)
            file_format: Image format ("png" or "jpg") (default: "png")
            border: Whether to draw a border around the image (default: False)
            border_color: Border color name or RGB tuple (default: "black")
            border_width: Border width in points (default: 1)
            shadow: Whether to apply a drop shadow (default: False)
            maintain_aspect_ratio: Whether to keep the figure's aspect ratio (default: True)

        Returns:
            The created picture shape
        """

    def add_shape(
        self,
        shape_type: Union[MSO_SHAPE, str] = MSO_SHAPE.RECTANGLE,
        x: PositionType = 5,
        y: PositionType = 5,
        width: PositionType = 40,
        height: PositionType = 10,
        fill_color: Optional[Union[str, Tuple[int, int, int]]] = None,
    ) -> PPTXShape:
        """Add a shape to the slide.

        Args:
            shape_type: The shape type, as an MSO_SHAPE value or its name,
                e.g. "ROUNDED_RECTANGLE" (default: MSO_SHAPE.RECTANGLE)
            x: X position as percent or in_() length (default: 5)
            y: Y position as percent or in_() length (default: 5)
            width: Width as percent or in_() length (default: 40)
            height: Height as percent or in_() length (default: 10)
            fill_color: Fill color as string name from COLORS dict or RGB tuple (default: None)

        Returns:
            The created shape object

        Raises:
            ValueError: If shape_type is a string that is not an MSO_SHAPE name
        """

    def add_table(
        self,
        data: Any,
        x: PositionType = 5,
        y: PositionType = 20,
        width: Optional[PositionType] = None,
        height: Optional[PositionType] = None,
        has_header: Optional[bool] = None,
        style: Optional[Union[int, dict, TableStyle]] = None,
    ) -> PPTXTable:
        """Add a table to the slide.

        Args:
            data: Table data as a list of lists or pandas DataFrame
            x: X position as percent or in_() length (default: 5)
            y: Y position as percent or in_() length (default: 20)
            width: Total width as percent or in_() length (default: None, auto-sized)
            height: Total height as percent or in_() length (default: None, auto-sized)
            has_header: Whether the first row is a header (default: True)
            style: Table style ID, a style dict with a "style_id" key,
                or a TableStyle (default: None)

        Returns:
            The created table object
        """

    def add_chart(
        self,
        data: Any = None,
        chart_type: Optional[str] = None,
        x: PositionType = 10,
        y: PositionType = 20,
        width: PositionType = 60,
        height: PositionType = 60,
        categories: Optional[list] = None,
        values: Optional[list] = None,
        category_column: Optional[Union[str, int]] = None,
        value_columns: Optional[Union[str, List[str], int, List[int]]] = None,
        title: Optional[str] = None,
        has_legend: Optional[bool] = None,
        legend_position: Optional[str] = None,
        style: Optional[ChartStyle] = None,
    ) -> PPTXChart:
        """Add a chart to the slide.

        Either pass explicit categories and values lists, or pass data
        (a list of lists with a header row, or a pandas DataFrame)
        together with optional category_column / value_columns. Passing
        several value_columns plots each column as a named series.

        Args:
            data: Chart data as a list of lists or pandas DataFrame (default: None)
            chart_type: Type of chart ('column', 'bar', 'line', 'pie', 'area', 'scatter')
                (default: "column")
            x: X position as percent or in_() length (default: 10)
            y: Y position as percent or in_() length (default: 20)
            width: Width as percent or in_() length (default: 60)
            height: Height as percent or in_() length (default: 60)
            categories: Explicit list of category labels (default: None)
            values: Explicit list of data values (default: None)
            category_column: Name or index of the column to use as categories (default: None)
            value_columns: Name(s) or index(es) of column(s) to use as values;
                a list creates one series per column (default: None)
            title: Chart title (default: None)
            has_legend: Whether to show legend (default: True)
            legend_position: Legend position (default: "right")
            style: A ChartStyle filling any chart options left unset (default: None)

        Returns:
            The created chart object

        Raises:
            ValueError: If neither data nor categories/values are provided
        """

    @property
    def notes(self) -> str:
        """The slide's speaker notes (empty string if none exist).

        New in 0.7.0. Assign a string to set the notes:
        slide.notes = "Talking points for this slide"
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
            start_x: Starting X position of container as percent or in_() length
            start_y: Starting Y position of container as percent or in_() length
            width: Width of container as percent or in_() length
            height: Height of container as percent or in_() length

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

Notes:

- `slide.add_chart` accepts `value_column` (singular) as an alias for `value_columns`, and `chart_title` as an alias for `title`. `slide.add_table` accepts `first_row_header` as an alias for `has_header`.
- Unknown parameters passed to Slide content methods trigger a warning instead of being silently ignored. Out-of-range percentages (e.g. `150` or `"150%"`) are clamped with a warning.
- **Removed in 0.7.0:** The pass-through variants on `Presentation` that took a slide as the first argument (`pres.add_text(slide, ...)`, `pres.add_image(slide, ...)`, `pres.add_shape(slide, ...)`, `pres.add_table(slide, ...)`, `pres.add_chart(slide, ...)`, `pres.add_pyplot(slide, ...)`) have been removed; use the `Slide` methods instead. `Presentation.add_matplotlib_slide`, `add_seaborn_slide`, and `add_plot` have been removed in favor of `add_pyplot_slide` (or `slide.add_pyplot`), and `add_image_slide` in favor of `add_image_gen_slide`. See [Migrating to 0.7.0](migration.md).

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
            x: X position as percent or in_() length (default: "10%")
            y: Y position as percent or in_() length (default: "5%")
            width: Width as percent or in_() length (default: "80%")
            height: Height as percent or in_() length (default: "15%")

        Returns:
            The created shape object
        """

    def add_paragraph(
        self,
        text: str,
        x: PositionType = 10,
        y: PositionType = 20,
        width: PositionType = 80,
        height: PositionType = 10,
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
            x: X position as percent or in_() length (default: 10)
            y: Y position as percent or in_() length (default: 20)
            width: Width as percent or in_() length (default: 80)
            height: Height as percent or in_() length (default: 10)
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

EasyPPTX uses a special type for position parameters. **Changed in 0.7.0:** bare numbers are percentages of the slide dimension; absolute inches use the `in_()` helper (floats no longer mean inches):

```python
# Type for position parameters
PositionType = Union[float, str, Length]

# Examples:
from easypptx import in_

x = 50        # 50% of slide width (percentage)
x = "50%"     # 50% of slide width (percentage string)
x = in_(1.5)  # 1.5 inches (absolute Length)
```

Percentages outside 0-100 are clamped with a warning; `in_()` lengths are not clamped.

### Positioning Helpers

The `easypptx.positioning` module provides the `in_()` helper and percent/inch conversion functions used throughout the library:

```python
from easypptx import in_         # in_(1.5) -> absolute length of 1.5 inches

from easypptx.positioning import (
    is_percent,             # True if a value is a percentage string like "50%"
    parse_percent,          # "50%" -> 50.0 (bare numbers pass through)
    pct,                    # 50.0 -> "50.00%"
    to_percent,             # Convert a position value to a percentage of a dimension
    to_inches,              # Convert a position value to inches
    shift_band,             # Shift a position band (e.g. below a title area)
    resolve_padding,        # Resolve combined/x/y padding values
    apply_content_padding,  # Apply padding to a content area
)
```

## Styles and Themes

New in 0.7.0. The `easypptx.styles` module (re-exported from `easypptx`) provides reusable style dataclasses and presentation-wide themes. See [Styling and Formatting](styling.md) for usage.

```python
@dataclass
class TextStyle:
    """Formatting for text content (add_text, add_bullets, titles)."""
    font_name: Optional[str] = None
    font_size: Optional[int] = None
    font_bold: Optional[bool] = None
    font_italic: Optional[bool] = None
    color: Optional[Union[str, Tuple[int, int, int]]] = None
    align: Optional[str] = None
    vertical: Optional[str] = None

@dataclass
class TableStyle:
    """Formatting for tables."""
    has_header: Optional[bool] = None
    style_id: Optional[int] = None

@dataclass
class ChartStyle:
    """Formatting for charts."""
    chart_type: Optional[str] = None
    has_legend: Optional[bool] = None
    legend_position: Optional[str] = None

@dataclass
class Theme:
    """A presentation-wide look: background, title style, and body style.

    Built-in presets are available by name: Presentation(theme="dark").
    Presets: "light", "dark", "corporate".
    """
    name: str = "custom"
    bg_color: Optional[Union[str, Tuple[int, int, int]]] = None
    title: TextStyle = field(default_factory=TextStyle)
    body: TextStyle = field(default_factory=TextStyle)
    accent_color: Optional[Union[str, Tuple[int, int, int]]] = None
```

Style objects are passed via the `style=` parameter of `add_text`, `add_bullets`, `add_table`, and `add_chart`; explicit keyword arguments always beat the style object. Themes set the default slide background, cascade body text styling via template defaults, and style `add_slide` titles.

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

Shared constants (`COLORS`, `ALIGN`, `VERTICAL`, `DEFAULT_FONT`) live in the `easypptx.common` module. They were previously documented as attributes of `Presentation`; those attributes (e.g. `Presentation.COLORS`) still work as aliases.

The color palette:

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

The `easypptx.common` module defines alignment options:

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

EasyPPTX uses the `MSO_SHAPE` enum from python-pptx for shape types. `Slide.add_shape` and `Presentation.add_shape` also accept the shape's name as a string:

```python
from pptx.enum.shapes import MSO_SHAPE

# Enum values:
MSO_SHAPE.RECTANGLE
MSO_SHAPE.OVAL
MSO_SHAPE.ROUNDED_RECTANGLE
MSO_SHAPE.ACTION_BUTTON_HOME

# Equivalent string names:
slide.add_shape(shape_type="ROUNDED_RECTANGLE")
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
                 rows: Union[int, List[float]] = 1,
                 cols: Union[int, List[float]] = 1,
                 padding: float = 5.0) -> None:
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
        """Access a cell (or span of cells) using indexing.

        Supports grid[row, col], flat access grid[index], and — new in
        0.7.0 — slice spans such as grid[1, :] (the whole second row) or
        grid[0:2, 1] (two rows of one column). A slice span merges the
        region and returns its GridCellProxy; step slices are rejected
        with a ValueError.

        Args:
            key: A tuple of (row, col) — either may be a slice — or a
                single int for flattened access

        Returns:
            A GridCellProxy for the requested cell or merged span

        Raises:
            OutOfBoundsError: If the requested cell is out of bounds
            TypeError: If the key is not in the right format
        """

    def next(self) -> "GridCellProxy":
        """Return a proxy for the next free cell, growing the grid if needed.

        New in 0.7.0. Cells fill in row-major order; when every cell has
        content, a new row is appended automatically.

        Example:
            grid.next().add_text("First free cell")
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

## GridCellProxy Class

`grid[row, col]` (and `grid.next()`) return a `GridCellProxy` with content methods (`add_text`, `add_image`, `add_pyplot`, `add_table`, `add_grid`) that size and position content to the cell. New in 0.7.0, the proxy also supports cell styling:

```python
class GridCellProxy:
    def style(self,
              fill: Optional[Union[str, Tuple[int, int, int]]] = None,
              border_color: Optional[Union[str, Tuple[int, int, int]]] = None,
              border_width: float = 1.0,
              padding: Optional[Union[float, str]] = None) -> "GridCellProxy":
        """Style this cell: background fill, border, and content padding.

        Draws a background rectangle covering the full cell and optionally
        sets padding that shrinks the area used by content added afterwards.
        Returns the proxy so calls can be chained.

        Args:
            fill: Background color name or RGB tuple (default: None, no fill)
            border_color: Border color name or RGB tuple (default: None, no border)
            border_width: Border width in points (default: 1.0)
            padding: Extra content padding as percent of the slide (default: None)

        Example:
            grid[0, 0].style(fill="lightgray", border_color="gray",
                             border_width=1, padding=2).add_text("Card")
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
