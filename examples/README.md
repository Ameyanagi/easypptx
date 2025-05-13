# EasyPPTX Examples

This directory contains example scripts that demonstrate how to use the EasyPPTX library to create professional PowerPoint presentations programmatically.

## Getting Started

To run these examples, make sure you have installed EasyPPTX and its dependencies:

```bash
pip install easypptx
# or
pip install -e .  # If you're in the project root
```

## Example Categories

The examples are organized into the following categories:

### Basics
Simple examples to get started with EasyPPTX:
- **001_quickstart.py**: A minimal example showing how to create a simple presentation
- **002_basic_demo.py**: Introduction to basic features
- **003_object_api.py**: Object-oriented API demonstration
- **004_plotting.py**: Including matplotlib charts and plots

### Styling
Examples related to visual styling of presentations:
- **001_basic_styling.py**: Basic styling options
- **002_dark_theme.py**: Creating presentations with dark backgrounds
- **003_responsive_positioning.py**: Percentage-based positioning for responsive layouts

### Templates
Examples showing how to use templates for consistent presentation design:
- **001_template_basic.py**: Basic TOML template usage
- **002_template_presets.py**: Built-in template presets
- **003_template_toml_manager.py**: TOML template management
- **004_template_manager.py**: Template Manager API
- **005_template_toml_reference.py**: TOML templates with reference PPTX

### Layouts
Examples showing different layout options:
- **001_alignment.py**: Element alignment strategies

### Grid
Examples demonstrating the Grid layout system:
- **001_basic_grid.py**: Basic grid creation and usage
- **002_grid_indexing.py**: Different ways to access grid cells
- **003_nested_grid.py**: Creating nested grids and merging cells
- **004_autogrid.py**: Automatic grid layout
- **005_enhanced_grid.py**: Enhanced Grid with convenience methods

### Advanced
More complex examples:
- **001_aspect_ratio.py**: Working with different aspect ratios
- **002_blank_layout.py**: Custom slide layouts and blank layouts
- **003_comprehensive.py**: A full-featured business presentation
- **004_enhanced_slides.py**: Advanced slide manipulation methods
- **005_extended_features.py**: Various extended features

## Running the Examples

You can run any example directly:

```bash
python examples/basics/001_quickstart.py
python examples/templates/001_template_basic.py
python examples/grid/001_basic_grid.py
```

## Example Structure

Each example follows a similar pattern:

1. **Setup**: Imports, data preparation, and directory creation
2. **Content Creation**: Building slides with text, images, tables, and charts
3. **Saving**: Exporting the final presentation to a PPTX file

## Output Directory

By default, all examples save their presentations to an `output` directory at the root of the project. Check this directory for the generated PPTX files.

## Next Steps

After exploring these examples, you can:

1. Review the [documentation](https://ameyanagi.github.io/EasyPPTX/) for detailed API reference
2. Examine the [test suite](/tests) for additional usage patterns
3. Create your own presentations based on these templates
