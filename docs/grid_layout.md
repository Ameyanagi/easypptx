# Grid Layout

EasyPPTX provides a powerful Grid layout system that makes it easy to create well-organized, responsive slides with complex layouts. The Grid system is perfect for creating dashboards, comparison slides, and any content that needs to be arranged in a structured way.

## Basic Grid Concepts

A Grid divides a slide (or a portion of a slide) into rows and columns, creating cells that can contain content. Key features include:

- Percentage-based positioning for responsive layouts
- Cell merging (like in spreadsheets)
- Nested grids for complex layouts
- Automatic padding between cells
- Responsive positioning that adapts to different aspect ratios

## Creating a Grid

To create a Grid, you need a parent (usually a Slide), and you can specify the grid's position, dimensions, and layout:

```python
from easypptx import Presentation, Grid

# Create a presentation
pres = Presentation()
slide = pres.add_slide()

# Create a 2x2 grid that takes up most of the slide
grid = Grid(
    parent=slide,
    x="5%",       # Position from left edge
    y="15%",      # Position from top edge
    width="90%",  # Width of the grid
    height="80%", # Height of the grid
    rows=2,       # Number of rows
    cols=2,       # Number of columns
    padding=5.0,  # Padding between cells (percentage)
    h_align="center",  # Responsive alignment (default: "center")
)
```

## Adding Content to Grid Cells

You can add any content (text, images, shapes, etc.) to a specific cell in the grid:

```python
# Add text to the top-left cell (row 0, column 0)
grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,  # Function to call to add content
    text="Cell Content",          # Parameters for the content function
    font_size=24,
    align="center",
    vertical="middle",
)

# Add a shape to the top-right cell (row 0, column 1)
grid.add_to_cell(
    row=0,
    col=1,
    content_func=slide.add_shape,
    shape_type=1,  # Rectangle
    fill_color="blue",
)
```

## Merging Cells

You can merge cells to create more complex layouts, similar to merging cells in a spreadsheet:

```python
# Merge cells from (0,0) to (1,1) - creating a 2x2 merged cell
merged_cell = grid.merge_cells(0, 0, 1, 1)

# Add content to the merged cell (use the top-left coordinates)
grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,
    text="Merged Cell Content",
    font_size=24,
    align="center",
    vertical="middle",
)
```

## Nested Grids

You can create nested grids for even more complex layouts:

```python
# Add a nested 3x3 grid to a cell in the main grid
nested_grid = grid.add_grid_to_cell(
    row=1,
    col=0,
    rows=3,
    cols=3,
    padding=5.0,
)

# Add content to a cell in the nested grid
nested_grid.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,
    text="Nested Content",
    font_size=16,
    align="center",
    vertical="middle",
)
```

## Dashboard Layout Example

Here's an example of creating a dashboard layout with the Grid system:

```python
# Create a dashboard layout grid
dashboard = Grid(
    parent=slide,
    x="5%",
    y="15%",
    width="90%",
    height="80%",
    rows=3,
    cols=4,
    padding=2.0,
)

# Create header area (spans the entire width)
dashboard.merge_cells(0, 0, 0, 3)
dashboard.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_shape,
    shape_type=1,  # Rectangle
    fill_color="blue",
)
dashboard.add_to_cell(
    row=0,
    col=0,
    content_func=slide.add_text,
    text="Sales Dashboard - FY 2023",
    font_size=24,
    font_bold=True,
    align="center",
    vertical="middle",
    color="white",
)

# Create sidebar (spans two rows)
dashboard.merge_cells(1, 0, 2, 0)
dashboard.add_to_cell(
    row=1,
    col=0,
    content_func=slide.add_shape,
    shape_type=1,  # Rectangle
    fill_color="gray",
)
dashboard.add_to_cell(
    row=1,
    col=0,
    content_func=slide.add_text,
    text="Navigation\n\n• Overview\n• Products\n• Regions",
    font_size=14,
    align="left",
    vertical="top",
    color="white",
)

# Create KPI area (spans 2 columns)
dashboard.merge_cells(1, 1, 1, 2)
dashboard.add_to_cell(
    row=1,
    col=1,
    content_func=slide.add_shape,
    shape_type=1,  # Rectangle
    fill_color="green",
)
dashboard.add_to_cell(
    row=1,
    col=1,
    content_func=slide.add_text,
    text="Revenue: $4.2M\nUp 15% from last year",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
)
```

## Advantages of Grid Layout

1. **Consistent Spacing**: Ensures consistent spacing between elements
2. **Responsive Design**: Adapts to different slide sizes and aspect ratios
3. **Simplified Positioning**: Eliminates the need for precise coordinate calculations
4. **Easy Reorganization**: Makes it easy to reorganize content without recalculating positions
5. **Complex Layouts**: Enables creation of complex layouts with minimal code

## Grid Properties and Methods

### Grid Class

- `parent`: The parent Slide or Grid object
- `x`, `y`: Position of the grid (percentages or absolute values)
- `width`, `height`: Dimensions of the grid
- `rows`, `cols`: Number of rows and columns
- `padding`: Padding between cells (percentage)
- `h_align`: Horizontal alignment for responsive positioning
- `cells`: 2D array of GridCell objects

### Methods

- `get_cell(row, col)`: Get a cell at the specified position
- `merge_cells(start_row, start_col, end_row, end_col)`: Merge cells in the specified range
- `add_to_cell(row, col, content_func, **kwargs)`: Add content to a specific cell
- `add_grid_to_cell(row, col, rows, cols, padding, h_align)`: Add a nested grid to a cell

## Complete Example

See [grid_layout_example.py](../examples/grid_layout_example.py) for a complete example showing basic grids, merged cells, nested grids, and a dashboard layout.