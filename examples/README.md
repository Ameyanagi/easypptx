# EasyPPTX Examples

This directory contains example scripts that demonstrate how to use the EasyPPTX library to create professional PowerPoint presentations programmatically.

## Getting Started

To run these examples, make sure you have installed EasyPPTX and its dependencies:

```bash
pip install easypptx
# or
pip install -e .  # If you're in the project root
```

## Available Examples

### Grid Layout Examples

The `grid/` directory contains examples that demonstrate the Grid layout capabilities:

- **001_basic_grid.py**: Basic grid creation and usage
- **002_grid_indexing.py**: Different ways to access grid cells (tuple indexing, flat indexing, iteration)
- **003_nested_grid.py**: Creating nested grids and merging cells
- **004_autogrid.py**: Automatic grid layout and arrangement of matplotlib plots

```bash
python examples/grid/001_basic_grid.py
python examples/grid/002_grid_indexing.py
python examples/grid/003_nested_grid.py
python examples/grid/004_autogrid.py
```

See [grid/README.md](grid/README.md) for more details.

### 1. quick_start.py

A simple introduction to EasyPPTX basics. This example demonstrates:
- Creating a presentation with multiple slides
- Basic text formatting
- Simple tables and charts

```bash
python examples/quick_start.py
```

### 2. basic_demo.py

An introductory demo that shows how to work with:
- Text formatting options
- Creating tables from data
- Basic chart creation
- Using pandas DataFrames as data sources

```bash
python examples/basic_demo.py
```

### 3. comprehensive_example.py

A full-featured business presentation that demonstrates the complete range of EasyPPTX capabilities:

- **Title Slide**: Company branding with logo and formatted text
- **Agenda Slide**: Structured content outline
- **Company Overview**: Image integration with a metrics table
- **Annual Sales Performance**: Line chart with summary statistics
- **Product Analysis (Table)**: pandas DataFrame conversion to formatted table
- **Product Analysis (Charts)**: Multiple chart types with insight callouts
- **Regional Sales Breakdown**: Pie chart with corresponding data table
- **Customer Satisfaction**: Column chart with conditional formatting
- **Future Outlook**: Text bullets with growth projection chart
- **Thank You / Q&A**: Contact information and branding elements

```bash
python examples/comprehensive_example.py
```

### 4. aspect_ratio_example.py

This example demonstrates how to create presentations with different aspect ratios:

- **16:9 Presentation**: The default widescreen format
- **4:3 Presentation**: Standard/legacy format
- **16:10 Presentation**: Alternative widescreen format
- **A4 Presentation**: Common paper size format
- **Custom Dimensions**: Setting arbitrary width and height in inches

The example creates multiple presentations with different aspect ratios and provides a summary table comparing their dimensions.

```bash
python examples/aspect_ratio_example.py
```

### 5. extended_features_example.py

This example demonstrates the extended features of EasyPPTX:

- **Percentage-Based Positioning**: Position and size elements using percentages instead of absolute inches
- **Default Fonts and Colors**: Use Meiryo as the default font and predefined color scheme
- **Auto-Alignment**: Automatically arrange multiple objects in grid, horizontal, or vertical layouts
- **Reference PowerPoint Templates**: Use existing PowerPoint files as templates

The example creates several presentations showing how to use these features.

```bash
python examples/extended_features_example.py
```

### 6. dark_theme_example.py

This example demonstrates how to create modern presentations with dark backgrounds:

- **Dark Backgrounds**: Set black or dark-colored backgrounds for slides
- **High-Contrast Colors**: Use vibrant text colors for better visibility
- **Layered Design**: Create depth with overlapping shapes and gradients
- **Modern Layout**: Combine auto-alignment with strategic object placement

The example creates two presentations: a basic dark theme presentation and a gradient effect presentation.

```bash
python examples/dark_theme_example.py
```

### 7. blank_layout_example.py

This example demonstrates using blank slide layouts and templates:

- **Blank Layouts**: Create slides using the default blank layout (layout 6)
- **Different Layouts**: Use different slide layout templates
- **Custom Backgrounds**: Apply background colors to blank slides
- **Template Usage**: Use existing presentations as templates for new slides

The example creates a presentation with slides using different layouts and another presentation based on an existing template.

```bash
python examples/blank_layout_example.py
```

## Example Structure

Each example follows a similar pattern:

1. **Setup**: Imports, data preparation, and directory creation
2. **Content Creation**: Building slides with text, images, tables, and charts
3. **Saving**: Exporting the final presentation to a PPTX file

## Comprehensive Example Details

The `comprehensive_example.py` file creates a complete business presentation with 10 slides. It showcases:

### Slide Types
- Title slide
- Content slides with various layouts
- Data visualization slides
- Summary slide

### Text Features
- Title and paragraph text
- Font formatting (size, bold, italic)
- Text coloring
- Positioned text elements

### Image Handling
- Logo placement
- Full-width images
- Images with specified dimensions
- Automatic sample image generation (using PIL)

### Tables
- Basic tables from lists of data
- Tables from pandas DataFrames
- Header formatting
- Data tables accompanying charts

### Charts
- Line charts for time series data
- Bar and column charts for comparisons
- Pie charts for distributions
- Charts with titles and legends
- Charts from pandas DataFrames

### Data Handling
- Static data arrays
- Calculated metrics
- pandas DataFrames for data manipulation
- Conditional formatting based on data analysis

## Customizing the Examples

Feel free to modify these examples to suit your specific needs:

- Change colors, fonts, and positioning
- Use your own data sources
- Add additional slides or elements
- Experiment with different chart types and layouts

## Output Directory

By default, all examples save their presentations to an `output` directory that is created automatically. Check this directory for the generated PPTX files.

## Next Steps

After exploring these examples, you can:

1. Review the [documentation](https://ameyanagi.github.io/EasyPPTX/) for detailed API reference
2. Examine the [test suite](/tests) for additional usage patterns
3. Create your own presentations based on these templates
