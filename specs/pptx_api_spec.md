# EasyPPTX API Specification

## Overview

This document outlines the API specification for EasyPPTX, a Python library designed to provide simple yet powerful interfaces for creating and manipulating PowerPoint presentations programmatically. The library aims to be particularly easy to use for both human developers and AI assistants.

## Core Design Principles

1. **Simplicity**: Provide straightforward, intuitive interfaces
2. **Consistency**: Maintain consistent parameter naming and behavior
3. **Flexibility**: Allow for both simple and advanced usage
4. **Error Handling**: Provide clear error messages and graceful fallbacks
5. **Documentation**: Comprehensive docstrings and examples

## API Structure

EasyPPTX is organized around the following main components:

1. **Presentation**: The container for all slides and presentation-level properties
2. **Slide**: Individual slides within a presentation
3. **Text**: Text elements and formatting
4. **Image**: Image handling and placement
5. **Table**: Tables creation and formatting
6. **Chart**: Chart generation and customization

## Detailed API Specification

### 1. Presentation

```python
class Presentation:
    def __init__(self) -> None:
        """Initialize a new empty presentation."""

    @classmethod
    def open(cls, file_path: Union[str, Path]) -> "Presentation":
        """Open an existing PowerPoint presentation."""

    def add_slide(self, layout_index: int = 0) -> Slide:
        """Add a new slide to the presentation."""

    @property
    def slides(self) -> List[Slide]:
        """Get a list of all slides in the presentation."""

    def save(self, file_path: Union[str, Path]) -> None:
        """Save the presentation to a file."""
```

### 2. Slide

```python
class Slide:
    def __init__(self, pptx_slide: PPTXSlide) -> None:
        """Initialize a Slide object."""

    def add_text(
        self,
        text: str,
        x: float = 1.0,
        y: float = 1.0,
        width: float = 8.0,
        height: float = 1.0,
        font_size: int = 18,
        font_bold: bool = False,
        font_italic: bool = False,
    ) -> PPTXShape:
        """Add a text box to the slide."""

    def add_image(
        self,
        image_path: str,
        x: float = 1.0,
        y: float = 1.0,
        width: Optional[float] = None,
        height: Optional[float] = None,
    ) -> PPTXShape:
        """Add an image to the slide."""

    @property
    def shapes(self) -> List[PPTXShape]:
        """Get all shapes on the slide."""

    def clear(self) -> None:
        """Remove all shapes from the slide."""

    @property
    def title(self) -> Optional[str]:
        """Get the slide title."""

    @title.setter
    def title(self, value: str) -> None:
        """Set the slide title."""
```

### 3. Text

```python
class Text:
    def __init__(self, slide_obj: Slide) -> None:
        """Initialize a Text object."""

    def add_title(self, text: str, font_size: int = 44) -> PPTXShape:
        """Add a title to the slide."""

    def add_paragraph(
        self,
        text: str,
        x: float = 1.0,
        y: float = 2.0,
        width: float = 8.0,
        height: float = 1.0,
        font_size: int = 18,
        font_bold: bool = False,
        font_italic: bool = False,
        color: Optional[Tuple[int, int, int]] = None,
    ) -> PPTXShape:
        """Add a paragraph of text to the slide."""

    @staticmethod
    def format_text_frame(
        text_frame: TextFrame,
        font_size: Optional[int] = None,
        font_bold: Optional[bool] = None,
        font_italic: Optional[bool] = None,
        color: Optional[Tuple[int, int, int]] = None,
    ) -> None:
        """Format an existing text frame."""
```

### 4. Image

```python
class Image:
    def __init__(self, slide_obj: Slide) -> None:
        """Initialize an Image object."""

    def add(
        self,
        image_path: Union[str, Path],
        x: float = 1.0,
        y: float = 1.0,
        width: Optional[float] = None,
        height: Optional[float] = None,
        maintain_aspect_ratio: bool = True,
    ) -> PPTXShape:
        """Add an image to the slide."""

    @staticmethod
    def get_image_dimensions(image_path: Union[str, Path]) -> Tuple[int, int]:
        """Get the dimensions of an image file."""
```

### 5. Table

```python
class Table:
    def __init__(self, slide_obj: Slide) -> None:
        """Initialize a Table object."""

    def add(
        self,
        data: List[List[Any]],
        x: float = 1.0,
        y: float = 1.0,
        width: Optional[float] = None,
        height: Optional[float] = None,
        first_row_header: bool = True,
        style: Optional[int] = None,
    ) -> PPTXTable:
        """Add a table to the slide."""

    def from_dataframe(
        self,
        df: pd.DataFrame,
        x: float = 1.0,
        y: float = 1.0,
        width: Optional[float] = None,
        height: Optional[float] = None,
        include_index: bool = False,
        first_row_header: bool = True,
        style: Optional[int] = None,
    ) -> PPTXTable:
        """Add a table from a pandas DataFrame."""
```

### 6. Chart

```python
class Chart:
    def __init__(self, slide_obj: Slide) -> None:
        """Initialize a Chart object."""

    def add(
        self,
        chart_type: str,
        categories: List[str],
        values: List[Union[int, float]],
        x: float = 1.0,
        y: float = 1.0,
        width: float = 6.0,
        height: float = 4.5,
        title: Optional[str] = None,
        has_legend: bool = True,
        **kwargs: Any,
    ) -> PPTXChart:
        """Add a chart to the slide."""

    def add_bar(
        self,
        categories: List[str],
        values: List[Union[int, float]],
        x: float = 1.0,
        y: float = 1.0,
        width: float = 6.0,
        height: float = 4.5,
        title: Optional[str] = None,
        **kwargs: Any,
    ) -> PPTXChart:
        """Add a bar chart to the slide."""

    def add_column(
        self,
        categories: List[str],
        values: List[Union[int, float]],
        x: float = 1.0,
        y: float = 1.0,
        width: float = 6.0,
        height: float = 4.5,
        title: Optional[str] = None,
        **kwargs: Any,
    ) -> PPTXChart:
        """Add a column chart to the slide."""

    def add_pie(
        self,
        categories: List[str],
        values: List[Union[int, float]],
        x: float = 1.0,
        y: float = 1.0,
        width: float = 6.0,
        height: float = 4.5,
        title: Optional[str] = None,
        **kwargs: Any,
    ) -> PPTXChart:
        """Add a pie chart to the slide."""

    def from_dataframe(
        self,
        df: pd.DataFrame,
        chart_type: str,
        category_column: str,
        value_column: str,
        x: float = 1.0,
        y: float = 1.0,
        width: float = 6.0,
        height: float = 4.5,
        title: Optional[str] = None,
        has_legend: bool = True,
        **kwargs: Any,
    ) -> PPTXChart:
        """Create a chart from a pandas DataFrame."""
```

## Future Enhancements

1. **Templates**: Support for using and creating templates
2. **Themes**: Apply consistent styling across presentations
3. **Master Slides**: Manipulation of master slides
4. **Advanced Charts**: More chart types and customization options
5. **SmartArt**: Creation and manipulation of SmartArt graphics
6. **Animations**: Add and control animations
7. **Slide Transitions**: Control slide transitions
8. **Comments**: Add and manage comments
9. **Export**: Export to other formats (PDF, images, etc.)
10. **Merge**: Merge multiple presentations
