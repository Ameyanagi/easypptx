# Auto-Alignment of Multiple Objects

EasyPPTX provides an easy way to arrange multiple objects on a slide with automatic alignment. This feature allows you to quickly create grid layouts, horizontal rows, or vertical columns of elements with consistent spacing.

## Overview

The `add_multiple_objects` method in the `Slide` class automatically arranges objects in one of three layout patterns:

1. **Grid layout**: Arranges objects in a grid pattern (default)
2. **Horizontal layout**: Arranges objects in a single row
3. **Vertical layout**: Arranges objects in a single column

## Basic Usage

```python
from easypptx import Presentation
from pptx.enum.shapes import MSO_SHAPE

# Create a presentation
pres = Presentation()
slide = pres.add_slide()

# Define objects to be added
objects = [
    {"type": "text", "text": "Item 1", "color": "black"},
    {"type": "text", "text": "Item 2", "color": "red"},
    {"type": "text", "text": "Item 3", "color": "blue"},
    {"type": "shape", "shape_type": MSO_SHAPE.RECTANGLE, "fill_color": "green"}
]

# Add objects in a grid layout
slide.add_multiple_objects(
    objects_data=objects,
    layout="grid",
    padding_percent=5.0,
    start_x="10%",
    start_y="20%",
    width="80%",
    height="60%"
)

pres.save("auto_aligned_grid.pptx")
```

## Layout Types

### Grid Layout

The grid layout arranges objects in a grid pattern, automatically determining the optimal number of rows and columns based on the number of objects:

```python
# Add objects in a grid layout
slide.add_multiple_objects(
    objects_data=objects,
    layout="grid",
    padding_percent=5.0
)
```

### Horizontal Layout

The horizontal layout arranges all objects in a single row:

```python
# Add objects in a horizontal layout
slide.add_multiple_objects(
    objects_data=objects,
    layout="horizontal",
    padding_percent=5.0
)
```

### Vertical Layout

The vertical layout arranges all objects in a single column:

```python
# Add objects in a vertical layout
slide.add_multiple_objects(
    objects_data=objects,
    layout="vertical",
    padding_percent=5.0
)
```

## Supported Object Types

You can include different types of objects in your layout:

### Text Objects

```python
{
    "type": "text",
    "text": "My text item",
    "font_size": 24,
    "font_bold": True,
    "font_italic": False,
    "font_name": "Meiryo",
    "align": "center",
    "vertical": "middle",
    "color": "black"
}
```

### Image Objects

```python
{
    "type": "image",
    "image_path": "path/to/image.png"
}
```

### Shape Objects

```python
{
    "type": "shape",
    "shape_type": MSO_SHAPE.RECTANGLE,  # Or any other MSO_SHAPE type
    "fill_color": "blue"  # Or RGB tuple (r, g, b)
}
```

## Container Positioning and Padding

You can control the position and size of the container that holds all objects:

```python
slide.add_multiple_objects(
    objects_data=objects,
    layout="grid",
    padding_percent=5.0,        # Padding between objects
    start_x="5%",               # Starting X position of the container
    start_y="5%",               # Starting Y position of the container
    width="90%",                # Width of the container
    height="90%"                # Height of the container
)
```

- `padding_percent`: Specifies the amount of space between objects as a percentage of the cell size
- All position and size parameters support both percentage-based and absolute positioning

## Implementation Details

The method calculates cell dimensions based on:
1. The number of objects
2. The layout type
3. The container dimensions
4. The specified padding

Each object is then positioned within its calculated cell, maintaining consistent spacing between all elements.
