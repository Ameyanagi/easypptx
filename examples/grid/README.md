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

## Running Examples

To run an example, use the following command from the project root:

```bash
cd /path/to/easypptx
python examples/grid/001_basic_grid.py
```

This will create the output file in the `output/` directory.
