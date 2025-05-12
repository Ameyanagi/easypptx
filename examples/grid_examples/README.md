# Grid Examples

This directory contains examples demonstrating various grid layout features and indexing methods in EasyPPTX.

## Examples

### dynamic_grid_example.py
- Demonstrates creating a grid-based slide with dynamic sizing
- Shows how to use tuple indexing syntax with (row, col) format
- Output: `output/dynamic_grid_example.pptx`

### flat_grid_example.py
- Shows how flat indexing works in grid layouts
- Includes examples of row access API
- Output: `output/flat_grid_example.pptx`

### flat_indexing_fix.py
- Demonstrates the fixed flat indexing implementation
- Shows how to access cells with a sequential numbering scheme
- Output: `output/flat_indexing_fix.pptx`

### updated_flat_grid_example.py
- Shows improved row access methods using the row() method
- Compares different grid access APIs
- Output: `output/updated_flat_grid_example.pptx`

## Running Examples

To run an example, use the following command from the project root:

```bash
cd /path/to/easypptx
python examples/grid_examples/dynamic_grid_example.py
```

This will create the output file in the `output/` directory.
