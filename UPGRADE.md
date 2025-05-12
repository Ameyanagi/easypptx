# Upgrade Guide for EasyPPTX

This document provides information about upgrading to newer versions of EasyPPTX and highlights important changes.

## Upgrading to 0.0.3

### Enhanced Grid Access API

Version 0.0.3 introduces a more intuitive API for accessing and manipulating Grid cells. This new API allows for:

1. **Direct Cell Access**: `grid[row, col].add_xxx()` syntax for directly adding content to specific cells
2. **Sequential Row Access**: `grid[row].add_xxx()` syntax for adding content to the next available cell in a row

#### Before (0.0.2 and earlier)

```python
# Adding content to cells using add_to_cell
grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,
    text="Cell Content",
    font_size=24,
)

# Or using convenience methods with row, col parameters
grid.add_textbox(
    row=0,
    col=1,
    text="Cell Content",
    font_size=24,
)
```

#### After (0.0.3 and later)

```python
# Direct cell access
grid[0, 0].add_text(
    text="Cell Content",
    font_size=24,
)

# Row-based sequential access
grid[1].add_text(
    text="First available cell in row 1",
    font_size=24,
)
grid[1].add_text(
    text="Second available cell in row 1",
    font_size=24,
)
```

### Benefits of the New API

1. **More Intuitive**: Familiar indexing syntax similar to numpy and matplotlib
2. **More Concise**: Requires less code and fewer parameters
3. **Sequential Access**: Automatic placement in rows without tracking column indices
4. **Consistent Interface**: Works with all content types (text, images, tables, charts)

### Example Usage

```python
from easypptx import Presentation

# Create a presentation
pres = Presentation()

# Add a slide with an empty grid
slide = pres.add_slide()
grid = pres.add_autogrid(
    slide=slide,
    content_funcs=None,  # Empty grid
    rows=2,
    cols=2,
)

# Add content directly to specific cells
grid[0, 0].add_text("Top Left", font_size=24)
grid[0, 1].add_image("path/to/image.png")

# Add content sequentially to a row
grid[1].add_text("Bottom Left", font_size=20)
grid[1].add_text("Bottom Right", font_size=20)

# Save the presentation
pres.save("example.pptx")
```

### Compatibility

The original methods (`add_to_cell`, `add_textbox`, etc.) are still fully supported and will remain available for backward compatibility. You can use the new API alongside the existing methods.

### Further Examples

For complete examples of the new API, see:
- `examples/grid/006_enhanced_grid_access.py` - Basic usage of the enhanced APIs
- `examples/grid/007_comprehensive_example.py` - Comprehensive example with multiple slides
- `examples/grid/008_quick_slide_deck.py` - Quick slide deck creation with row-based API
