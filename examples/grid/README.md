# Grid Layout Examples

This directory contains examples demonstrating the grid layout features of EasyPPTX.

## Examples

Each example has a numerical prefix for easy reference, and creates a corresponding output file with the same prefix.

### 001_basic_grid.py
- Basic grid creation and usage
- Shows how to create a 2x2 grid and add content to each cell
- Output: `output/001_basic_grid.pptx`

### 002_grid_indexing.py
- Demonstrates different ways to access and manipulate grid cells
- Includes examples of using:
  - Tuple indexing: `grid[row, col]`
  - Flat indexing: `grid[index]`
  - Iteration: `for cell in grid`
  - Flat iteration: `for cell in grid.flat`
- Output: `output/002_grid_indexing.pptx`

### 003_nested_grid.py
- Shows how to create nested grids (grids within grid cells)
- Demonstrates cell merging in both main and nested grids
- Output: `output/003_nested_grid.pptx`

### 004_autogrid.py
- Demonstrates automatic grid layout features:
  - Using `autogrid` to automatically arrange content
  - Using `autogrid_pyplot` to arrange matplotlib plots
- Output: `output/004_autogrid.pptx`

### 005_enhanced_grid.py
- Shows enhanced Grid convenience methods:
  - `add_textbox`, `add_image`, `add_pyplot`, and `add_table`
- Demonstrates creating empty grids with `add_autogrid(content_funcs=None)`
- Output: `output/005_enhanced_grid.pptx`

### 006_enhanced_grid_access.py
- Demonstrates enhanced Grid access API with more intuitive syntax:
  - Using `grid[row, col].add_xxx()` for direct cell access and manipulation
  - Using `grid[row].add_xxx()` for automatically adding content to the next available cell in a row
- Shows how these methods make creating complex layouts more intuitive
- Output: `output/006_enhanced_grid_access.pptx`

### 007_comprehensive_example.py
- Provides a comprehensive example showcasing the enhanced Grid APIs:
  - Creating a complete presentation with multiple slides
  - Using the enhanced syntax for various content types (text, images, tables, charts)
  - Creating nested grids with enhanced access pattern
  - Building a data dashboard using enhanced Grid APIs
- Output: `output/007_comprehensive_example.pptx`

### 008_quick_slide_deck.py
- Demonstrates using the row-based API (`grid[row].add_xxx()`) to quickly create a slide deck
- Shows how content can be added sequentially to rows without tracking column indices
- Showcases automatic content placement and flow across rows
- Output: `output/008_quick_slide_deck.pptx`

### Additional Files
- `001_basic_grid_updated.py`: Updated version of the basic grid example using the enhanced API

## Running Examples

To run an example, use the following command from the project root:

```bash
cd /path/to/easypptx
python examples/grid/001_basic_grid.py
```

This will create the output file in the `output/` directory.

## Recommended Usage Pattern

The enhanced Grid API enables a more concise and intuitive syntax:

```python
from easypptx import Presentation

# Create a new presentation
pres = Presentation()

# Add a slide with an empty 2x2 grid
slide = pres.add_slide()
grid = pres.add_autogrid(
    slide=slide,
    content_funcs=None,  # Empty grid
    rows=2,
    cols=2
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
