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

## Accessing Grid Cells

EasyPPTX's Grid class provides multiple ways to access and manipulate cells, similar to numpy arrays and matplotlib subplots:

### 1. Using the `[row, col]` Indexing Syntax

```python
# Access a cell directly using grid[row, col]
cell = grid[0, 1]  # Get the cell at row 0, column 1

# Add content to the cell
cell.content = slide.add_text(
    text="Cell at [0, 1]",
    x=cell.x,
    y=cell.y,
    width=cell.width,
    height=cell.height,
    font_size=24,
    align="center",
    vertical="middle",
)
```

### 2. Using Flat Indexing

```python
# Access a cell using a flat index (row-major order)
cell = grid[3]  # Get the cell at flat index 3 (row 1, col 1 in a 2x2 grid)

# Add content to the cell
cell.content = slide.add_text(
    text="Cell at flat index 3",
    x=cell.x,
    y=cell.y,
    width=cell.width,
    height=cell.height,
    font_size=24,
    align="center",
    vertical="middle",
)
```

### 3. Using Iteration

```python
# Iterate through all cells in the grid
for cell in grid:
    # Add content to each cell
    cell.content = slide.add_text(
        text=f"Cell at [{cell.row}, {cell.col}]",
        x=cell.x,
        y=cell.y,
        width=cell.width,
        height=cell.height,
        font_size=24,
        align="center",
        vertical="middle",
    )
```

### 4. Using Flat Iteration (like matplotlib)

```python
# Iterate through cells in a flattened manner
for cell in grid.flat:
    # Check cell properties and add content conditionally
    if cell.row == 1:  # Only add to cells in the second row
        cell.content = slide.add_text(
            text=f"Flat cell at [{cell.row}, {cell.col}]",
            x=cell.x,
            y=cell.y,
            width=cell.width,
            height=cell.height,
            font_size=24,
            align="center",
            vertical="middle",
        )
```

### 5. Using the Traditional Method

```python
# Get a cell using the get_cell method
cell = grid.get_cell(row=1, col=2)
```

## Grid Properties and Methods

### Grid Class

- `parent`: The parent Slide or Grid object
- `x`, `y`: Position of the grid (percentages or absolute values)
- `width`, `height`: Dimensions of the grid
- `rows`, `cols`: Number of rows and columns
- `padding`: Padding between cells (percentage)
- `h_align`: Horizontal alignment for responsive positioning
- `cells`: 2D array of GridCell objects
- `flat`: Property that returns a flat iterator for the grid (like matplotlib's subplot.flat)

### Methods

- `get_cell(row, col)`: Get a cell at the specified position
- `merge_cells(start_row, start_col, end_row, end_col)`: Merge cells in the specified range
- `add_to_cell(row, col, content_func, **kwargs)`: Add content to a specific cell
- `add_grid_to_cell(row, col, rows, cols, padding, h_align)`: Add a nested grid to a cell
- `__iter__()`: Makes Grid objects iterable
- `__getitem__(key)`: Enables accessing cells via grid[row, col] or grid[index]

### Convenience Methods

The Grid class provides convenient methods for adding common elements directly:

- `add_textbox(row, col, text, **kwargs)`: Add a text box to a specific cell
- `add_image(row, col, image_path, **kwargs)`: Add an image to a specific cell
- `add_pyplot(row, col, figure, **kwargs)`: Add a matplotlib figure to a specific cell
- `add_table(row, col, data, **kwargs)`: Add a table to a specific cell

Example usage:

```python
# Add text directly to a cell
grid.add_textbox(0, 0, "Hello World", font_size=24, align="center")

# Add an image to a cell
grid.add_image(0, 1, "path/to/image.jpg", maintain_aspect_ratio=True)

# Add a matplotlib figure to a cell
import matplotlib.pyplot as plt
fig = plt.figure()
plt.plot([1, 2, 3, 4], [1, 4, 9, 16])
grid.add_pyplot(1, 0, fig, dpi=300)

# Add a table to a cell
data = [["Name", "Value"], ["Item 1", 100], ["Item 2", 200]]
grid.add_table(1, 1, data, has_header=True)
```

## Automatic Grid Layout

EasyPPTX can automatically arrange content in a grid:

```python
# Create content functions
def create_text1():
    return slide.add_text(
        text="Content 1",
        font_size=24,
        align="center",
        vertical="middle",
    )

def create_text2():
    return slide.add_text(
        text="Content 2",
        font_size=24,
        align="center",
        vertical="middle",
    )

# More content functions...

# Use add_autogrid to arrange content automatically
content_funcs = [create_text1, create_text2, ...]
grid = pres.add_autogrid(
    slide=slide,
    content_funcs=content_funcs,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    padding=5.0,
    title="Auto Grid Example",
)
```

### Creating Empty Grids with add_autogrid

You can also create an empty grid using add_autogrid by passing None for content_funcs, and then populate it later using the convenience methods:

```python
# Create an empty 2x2 grid
grid = pres.add_autogrid(
    slide=slide,
    content_funcs=None,  # This creates an empty grid
    rows=2,              # Must specify rows and cols when content_funcs is None
    cols=2,
    x="10%",
    y="20%",
    width="80%",
    height="70%",
    padding=5.0,
)

# Add content to specific cells using convenience methods
grid.add_textbox(0, 0, "Top Left Cell", font_size=18, align="center")
grid.add_image(0, 1, "path/to/image.jpg")
grid.add_pyplot(1, 0, matplotlib_figure, dpi=150)
grid.add_table(1, 1, data=[["A", "B"], [1, 2]])
```

This approach gives you the ability to create an empty grid layout first and then selectively add different types of content to specific cells. It's particularly useful for creating dashboards and complex layouts.

## Arranging Matplotlib Plots

You can automatically arrange matplotlib plots in a grid:

```python
import matplotlib.pyplot as plt
import numpy as np

# Create matplotlib figures
fig1 = plt.figure(figsize=(4, 3))
categories = ["A", "B", "C", "D", "E"]
values = [23, 45, 56, 78, 32]
plt.bar(categories, values)
plt.title("Bar Chart")

fig2 = plt.figure(figsize=(4, 3))
x = np.random.rand(50)
y = np.random.rand(50)
plt.scatter(x, y)
plt.title("Scatter Plot")

# More figures...

# Arrange plots in a grid
figures = [fig1, fig2, ...]
from easypptx.grid import Grid
grid = Grid.autogrid_pyplot(
    parent=slide,
    figures=figures,
    x="5%",
    y="20%",
    width="90%",
    height="75%",
    rows=2,
    cols=2,
    padding=5.0,
    title="Matplotlib Plots in Grid",
    title_height="5%",
    dpi=300,
)
```

## One-Step Grid Slide Creation

For a quicker workflow, create a slide with a grid in one step:

```python
# Create a slide with a grid
slide, grid = pres.add_grid_slide(
    rows=3,
    cols=3,
    title="Grid Slide Example",
    title_height="15%",
    padding=3.0,
)

# Now add content to the grid cells
for row in range(3):
    for col in range(3):
        grid.add_to_cell(
            row=row,
            col=col,
            content_func=slide.add_text,
            text=f"Cell [{row},{col}]",
            font_size=18,
            align="center",
            vertical="middle",
        )
```

## One-Step AutoGrid Slide Creation

Similarly, create a slide with an auto-arranged grid in one step:

```python
# Content functions
content_funcs = [create_text1, create_text2, ...]

# Create a slide with auto-arranged content
slide, grid = pres.add_autogrid_slide(
    content_funcs=content_funcs,
    rows=2,
    cols=2,
    title="Auto Grid Slide",
    title_height="15%",
)
```

## Complete Examples

See the following examples for detailed demonstrations:

- [001_basic_grid.py](../examples/grid/001_basic_grid.py): Basic grid creation and usage
- [002_grid_indexing.py](../examples/grid/002_grid_indexing.py): Grid indexing and iteration features
- [003_nested_grid.py](../examples/grid/003_nested_grid.py): Nested grids and cell merging
- [004_autogrid.py](../examples/grid/004_autogrid.py): Automatic grid layout features
- [005_enhanced_grid.py](../examples/grid/005_enhanced_grid.py): Empty grids and convenience methods
