# EasyPPTX Template System

EasyPPTX provides several ways to work with templates:

1. **Reference Templates**: Automatic built-in templates based on aspect ratio
2. **File-based templates**: Using existing PowerPoint files as templates
3. **Template Presets**: Using the built-in template system with predefined slide layouts
4. **TOML/JSON Template Files**: Saving and loading templates as TOML or JSON files

## Part 1: Reference Templates

EasyPPTX automatically uses built-in reference templates for standard aspect ratios, ensuring your presentations have a consistent, professional look without requiring any additional configuration.

### Built-in Reference Templates

The library includes the following reference templates:

- `reference_16x9.pptx`: Used automatically for 16:9 (widescreen) presentations
- `reference_4x3.pptx`: Used automatically for 4:3 (standard) presentations

These templates provide well-designed slide masters, layouts, and theme elements appropriate for each aspect ratio.

### Automatic Template Selection

The reference templates are automatically applied based on the aspect ratio:

```python
from easypptx import Presentation

# Creates a presentation with 16:9 aspect ratio
# Automatically uses reference_16x9.pptx as the template
presentation = Presentation()  # Default is 16:9

# Creates a presentation with 4:3 aspect ratio
# Automatically uses reference_4x3.pptx as the template
presentation = Presentation(aspect_ratio="4:3")
```

### When Reference Templates Are Not Used

In certain cases, the reference templates will not be applied:

1. When a custom template is specified:
   ```python
   presentation = Presentation(template_path="custom_template.pptx")
   ```

2. When custom dimensions are provided:
   ```python
   presentation = Presentation(width_inches=13.33, height_inches=7.5)
   ```

3. For aspect ratios without reference templates:
   ```python
   presentation = Presentation(aspect_ratio="16:10")
   presentation = Presentation(aspect_ratio="A4")
   ```

In these cases, EasyPPTX will create a blank presentation and apply the specified dimensions or use the custom template.

## Part 2: File-based Templates

EasyPPTX supports using existing PowerPoint files as templates for your presentations. This allows you to leverage pre-designed slides, themes, master layouts, and styles.

### Basic Usage

To create a presentation based on an existing template:

```python
from easypptx import Presentation, Text

# Create a presentation using an existing template
pres = Presentation(template_path="template.pptx")

# Add a new slide (uses blank layout by default)
slide = pres.add_slide()

# Add content to the slide
text = Text(slide)
text.add_title("Slide Using Template")
text.add_paragraph("This presentation is based on a template file")

# Save the presentation with a new name
pres.save("template_based_presentation.pptx")
```

### Slide Layouts

By default, EasyPPTX uses the blank slide layout (typically index 6) for new slides. You can specify a different layout by using the `layout_index` parameter:

```python
# Add a slide with title layout (usually index 0)
title_slide = pres.add_slide(layout_index=0)

# Add a slide with title and content layout (usually index 1)
content_slide = pres.add_slide(layout_index=1)

# Add a blank slide explicitly
blank_slide = pres.add_slide(layout_index=6)

# Add a blank slide (default behavior when layout_index is not specified)
default_slide = pres.add_slide()
```

### Benefits of Using File-based Templates

1. **Consistent branding**: Use corporate templates with logos, color schemes, and layouts
2. **Professional design**: Leverage professionally designed templates
3. **Time-saving**: Avoid recreating complex designs from scratch
4. **Theme consistency**: Maintain consistent fonts, colors, and styles

### Accessing Slide Layouts

When using a template, you can access different slide layouts by index:

```python
# Create a presentation from a template
pres = Presentation(template_path="template.pptx")

# Add a slide using the first layout (usually the title slide)
title_slide = pres.add_slide(layout_index=0)

# Add a slide using the second layout (usually a content slide)
content_slide = pres.add_slide(layout_index=1)

# Add a slide using the third layout
third_layout_slide = pres.add_slide(layout_index=2)
```

The available layouts depend on the template file. Most PowerPoint templates include:

- Title slide layout (index 0)
- Title and content layout (index 1)
- Section header layout (index 2)
- Two-content layout (index 3)
- Comparison layout (index 4)
- Title only layout (index 5)
- Blank layout (index 6)

## Part 3: Template Presets

The EasyPPTX template preset system allows you to create consistent and professional-looking presentations with minimal code. It provides pre-defined templates for common slide types and layouts that can be customized to your needs.

### Overview

Template presets in EasyPPTX are defined as dictionaries that specify:
- Slide background colors
- Text elements (title, subtitle, content) with their positioning, fonts, and colors
- Decorative elements like horizontal bars with gradient colors
- Areas for content, images, tables, and charts

### Built-in Templates

EasyPPTX includes the following built-in templates:

### Basic Templates

- `title_slide`: For presentation title slides with title and subtitle
- `content_slide`: For general content with title and horizontal gradient bar
- `section_slide`: For section dividers with full-screen title on colored background
- `comparison_slide`: For comparing two items side by side
- `image_slide`: For displaying images with title and optional caption
- `table_slide`: For displaying data tables with title
- `chart_slide`: For charts and graphs with title
- `thank_you_slide`: For ending the presentation

### Advanced Templates

- `quote_slide`: For displaying quotes with author attribution and stylized quotation marks
- `bullets_slide`: For displaying bullet points with a title
- `agenda_slide`: For showing a meeting or presentation agenda with optional image
- `team_slide`: For introducing team members with photos and titles
- `statement_slide`: For emphasizing important statements with a colored background
- `dashboard_slide`: For creating dashboards with multiple charts and visualizations
- `timeline_slide`: For displaying project timelines or roadmaps with connected steps

### Styling Options

EasyPPTX templates provide extensive styling options for different types of slide elements:

#### Image Styling

You can customize how images appear in your presentations:

```python
# Add an image slide with custom styling
slide = pres.add_image_slide(
    title="Styled Image",
    image_path="example.jpg",
    label="Beautiful scenery",
    custom_style={
        "border": True,
        "border_color": "blue",
        "border_width": 2,
        "shadow": True,
        "maintain_aspect_ratio": True,
        "center": True
    }
)
```

Available image style options:
- `border`: Whether to show a border around the image (Boolean)
- `border_color`: Color name or RGB tuple for the border
- `border_width`: Width of the border in points
- `shadow`: Whether to apply a shadow effect (Boolean)
- `rounded_corners`: Whether to apply rounded corners (Boolean)
- `maintain_aspect_ratio`: Whether to maintain the image's aspect ratio (Boolean)
- `center`: Whether to center the image in its container (Boolean)
- `brightness`: Brightness adjustment from -1.0 to 1.0
- `contrast`: Contrast adjustment from -1.0 to 1.0

#### Table Styling

Tables can be styled with headers, borders, and alternating rows:

```python
# Add a table slide with custom styling
slide = pres.add_table_slide(
    title="Styled Table",
    data=my_data,
    has_header=True,
    custom_style={
        "first_row": {
            "bold": True,
            "bg_color": "darkblue",
            "text_color": "white"
        },
        "banded_rows": True,
        "band_color": "lightgray",
        "border_width": 1,
        "header_border_width": 2
    }
)
```

Available table style options:
- `first_row`: Dictionary of styling for the header row
  - `bold`: Whether header text is bold (Boolean)
  - `bg_color`: Background color for header row
  - `text_color`: Text color for header row
- `banded_rows`: Whether to use alternating row colors (Boolean)
- `band_color`: Color for alternating rows
- `border_color`: Color for table borders
- `border_width`: Width of regular cell borders
- `header_border_width`: Width of header row borders
- `text_align`: Text alignment in cells ("left", "center", "right")
- `header_align`: Text alignment in header cells
- `font_name`: Font name for table text
- `font_size`: Font size for regular cells
- `header_font_size`: Font size for header cells

#### Chart Styling

Charts can be styled with custom colors, legends, and labels:

```python
# Add a chart slide with custom styling
slide = pres.add_chart_slide(
    title="Styled Chart",
    data=my_data,
    chart_type="pie",
    custom_style={
        "has_legend": True,
        "legend_position": "right",
        "has_data_labels": True,
        "palette": [
            (0x5B, 0x9B, 0xD5),  # Blue
            (0xED, 0x7D, 0x31),  # Orange
            (0xA5, 0xA5, 0xA5),  # Gray
        ]
    }
)
```

Available chart style options:
- `chart_type`: Type of chart ("column", "bar", "line", "pie", "scatter", "area")
- `has_legend`: Whether to show a legend (Boolean)
- `legend_position`: Where to place the legend ("top", "bottom", "left", "right")
- `has_title`: Whether to show a chart title (Boolean)
- `title_font_size`: Font size for the chart title
- `palette`: List of RGB tuples or color names for data series
- `has_data_labels`: Whether to show data labels (Boolean)
- `gridlines`: Whether to show gridlines (Boolean)
- `has_border`: Whether to show a border around the chart (Boolean)
- `border_color`: Color for the chart border

### Using Template Presets

There are two ways to use template presets:

#### 1. Convenience Methods

EasyPPTX provides convenience methods for creating slides from templates:

```python
from easypptx import Presentation

pres = Presentation()

# Add a title slide
title_slide = pres.add_title_slide(
    title="My Presentation",
    subtitle="Created with EasyPPTX"
)

# Add a content slide with bullet points
content_slide = pres.add_content_slide(title="Key Points")
content_slide.add_text(
    text="• Point 1\n• Point 2\n• Point 3",
    position={"x": "10%", "y": "15%", "width": "80%", "height": "70%"}
)

# Add a section slide
section_slide = pres.add_section_slide(
    title="New Section",
    bg_color="blue"
)

# Add an image slide with caption
image_slide = pres.add_image_slide(
    title="Image Example",
    image_path="path/to/image.jpg",
    label="This is a caption for the image"
)

# Add a comparison slide
comparison_slide = pres.add_comparison_slide(
    title="Comparison",
    content_texts=["Left side content", "Right side content"]
)

# Add a table slide
table_slide = pres.add_table_slide(
    title="Data Table",
    data=my_dataframe,  # Or list of lists
    has_header=True
)
```

#### 2. Using Template Directly

You can also use the template system directly:

```python
from easypptx import Presentation

pres = Presentation()

# Create slide using a built-in template
slide = pres.add_slide_from_template("title_slide")

# Update text in the slide
title_shapes = [shape for shape in slide.shapes if shape.has_text_frame]
if title_shapes:
    title_shapes[0].text_frame.text = "My Custom Title"
```

### Custom Templates

You can create custom templates by defining a dictionary with the template properties:

```python
custom_template = {
    "bg_color": "orange",
    "title": {
        "text": "Custom Template",
        "position": {"x": "5%", "y": "5%", "width": "90%", "height": "15%"},
        "font": {"name": "Meiryo", "size": 36, "bold": True},
        "align": "center",
        "vertical": "middle",
        "color": "white"
    },
    "content_area": {
        "position": {"x": "10%", "y": "30%", "width": "80%", "height": "60%"}
    }
}

custom_slide = pres.add_slide_from_template(custom_template)
```

### Template Structure

A template is a dictionary with the following structure:

```python
{
    "bg_color": "color_name_or_rgb_tuple",  # Optional background color

    "title": {  # Optional title element
        "text": "Default title text",
        "position": {"x": "10%", "y": "5%", "width": "80%", "height": "10%"},
        "font": {"name": "Font name", "size": 32, "bold": True},
        "align": "center",  # "left", "center", or "right"
        "vertical": "middle",  # "top", "middle", or "bottom"
        "color": "black"
    },

    "subtitle": {  # Optional subtitle element (similar to title)
        # Same properties as title
    },

    "bar": {  # Optional decorative bar
        "position": {"x": "0%", "y": "10%", "width": "100%", "height": "2%"},
        "gradient": {  # Optional gradient fill
            "start_color": RGBColor(0xE0, 0xE5, 0xF7),
            "end_color": RGBColor(0x95, 0xAB, 0xEA),
            "angle": 0  # 0 for horizontal, 90 for vertical
        }
    },

    "content_area": {  # Optional content area
        "position": {"x": "5%", "y": "15%", "width": "90%", "height": "80%"}
    },

    # Styling for images
    "image_style": {
        "border": True,  # Whether the image should have a border
        "border_color": "gray",  # Color of the border
        "border_width": 1,  # Width of the border
        "shadow": True,  # Whether the image should have a shadow
        "rounded_corners": False,  # Whether the image should have rounded corners
        "maintain_aspect_ratio": True,  # Whether to maintain the image aspect ratio
        "center": True,  # Whether to center the image in its container
        "brightness": 0,  # Brightness adjustment (-1.0 to 1.0)
        "contrast": 0  # Contrast adjustment (-1.0 to 1.0)
    },

    # Styling for tables
    "table_style": {
        "first_row": {  # Styling for the header row
            "bold": True,
            "bg_color": "blue",
            "text_color": "white"
        },
        "banded_rows": True,  # Whether to use alternating row colors
        "band_color": "lightgray",  # Color for alternate rows
        "border_color": "black",  # Color for table borders
        "border_width": 1,  # Width of table borders
        "header_border_width": 2,  # Width of header row border
        "text_align": "center",  # Alignment for cell text
        "header_align": "center",  # Alignment for header text
        "font_name": "Meiryo",  # Font for table text
        "font_size": 12,  # Font size for table text
        "header_font_size": 14  # Font size for header text
    },

    # Styling for charts
    "chart_style": {
        "chart_type": "column",  # Type of chart (column, bar, line, pie, scatter, area)
        "has_legend": True,  # Whether to show a legend
        "legend_position": "bottom",  # Position of the legend (top, bottom, left, right)
        "has_title": True,  # Whether to show a chart title
        "title_font_size": 14,  # Font size for the chart title
        "palette": [  # Custom colors for chart series
            RGBColor(0x5B, 0x9B, 0xD5),  # Blue
            RGBColor(0xED, 0x7D, 0x31),  # Orange
            RGBColor(0xA5, 0xA5, 0xA5),  # Gray
            RGBColor(0xFF, 0xC0, 0x00),  # Yellow
            RGBColor(0x4C, 0xAF, 0x50),  # Green
            RGBColor(0x9C, 0x27, 0xB0)   # Purple
        ],
        "has_data_labels": False,  # Whether to show data labels
        "gridlines": True,  # Whether to show gridlines
        "has_border": True,  # Whether to show chart border
        "border_color": "black"  # Color of chart border
    },

    # Other optional elements like image_area, table_area, chart_area,
    # left_content, right_content, etc.
}
```

### Percentage-based Positioning

All positioning in templates uses percentages of the slide dimensions, making the templates responsive to different slide sizes and aspect ratios:

```python
position = {"x": "10%", "y": "20%", "width": "80%", "height": "60%"}
```

## Combining Multiple Template Systems

You can combine multiple template systems for maximum flexibility:

1. Starting with a reference template for your chosen aspect ratio
2. Using a custom PowerPoint template file for specific design needs
3. Adding slides using template presets for consistent layout and content
4. Saving your custom templates as TOML/JSON files for reuse

```python
from easypptx import Presentation

# Create presentation from a file-based template
pres = Presentation(template_path="corporate_template.pptx")

# Add slides using template presets
title_slide = pres.add_title_slide(
    title="Quarterly Report",
    subtitle="Q1 2023"
)

content_slide = pres.add_content_slide("Key Highlights")
content_slide.add_text(
    text="• Revenue increased by 15%\n• New product launch successful\n• Expansion to 3 new markets",
    position={"x": "10%", "y": "15%", "width": "80%", "height": "70%"}
)

pres.save("combined_template_approach.pptx")
```

## Advanced Template Examples

Let's look at some examples of the advanced templates:

### Quote Slide

```python
from easypptx import Presentation

pres = Presentation()

# Add a quote slide
quote_slide = pres.add_slide_from_template("quote_slide")

# Update the quote elements (find text frames and update text)
quote_shapes = [shape for shape in quote_slide.shapes if shape.has_text_frame]
for shape in quote_shapes:
    if "Your quote goes here" in shape.text_frame.text:
        shape.text_frame.text = "The best way to predict the future is to create it."
    elif "Author Name" in shape.text_frame.text:
        shape.text_frame.text = "Abraham Lincoln"
```

### Bullets Slide

```python
# Add a bullets slide
bullets_slide = pres.add_slide_from_template("bullets_slide")

# Update the title and bullet points
bullet_shapes = [shape for shape in bullets_slide.shapes if shape.has_text_frame]
for shape in bullet_shapes:
    if "Key Points" in shape.text_frame.text:
        shape.text_frame.text = "Key Benefits"
    elif "Point" in shape.text_frame.text:
        shape.text_frame.text = """• Professional-looking presentations
• Percentage-based positioning
• Customizable templates
• Support for matplotlib visualizations"""
```

### Dashboard Slide

```python
from easypptx import Presentation, Chart, Pyplot
import pandas as pd
import matplotlib.pyplot as plt

pres = Presentation()

# Add a dashboard slide
dashboard_slide = pres.add_slide_from_template("dashboard_slide")

# Get the chart positions from the template
template = pres.template.get_preset("dashboard_slide")
chart_positions = template["charts"]

# Create sample data for a chart
data = pd.DataFrame({
    'Quarter': ['Q1', 'Q2', 'Q3', 'Q4'],
    'Revenue': [100, 120, 135, 150]
})

# Add a chart to the top left position
Chart.add(
    slide=dashboard_slide,
    data=data,
    chart_type="column",
    position=chart_positions["top_left"]["position"],
    category_column="Quarter",
    value_columns="Revenue",
    has_title=True,
    chart_title="Quarterly Revenue"
)

# Add a matplotlib plot to another position
plt.figure()
plt.plot([1, 2, 3, 4], [10, 20, 15, 25])
plt.title('Sample Trend')

Pyplot.add(
    slide=dashboard_slide,
    figure=plt.gcf(),
    position=chart_positions["top_right"]["position"]
)
```

### Timeline Slide

```python
from easypptx import Presentation, Text

pres = Presentation()

# Add a timeline slide
timeline_slide = pres.add_slide_from_template("timeline_slide")

# Update the title
timeline_title = [shape for shape in timeline_slide.shapes if shape.has_text_frame][0]
timeline_title.text_frame.text = "Project Timeline"

# Create timeline steps (simplified example)
step_box1 = timeline_slide.add_shape(
    position={"x": "5%", "y": "20%", "width": "15%", "height": "15%"},
    fill_color="blue"
)

Text.add(
    slide=timeline_slide,
    text="Step 1: Planning",
    position={"x": "5%", "y": "20%", "width": "15%", "height": "15%"},
    font_bold=True,
    align="center",
    vertical_align="middle",
    color="white"
)

# Add more steps and connectors as needed
```

## Advanced Templates Details

Each advanced template comes with specific elements designed for its purpose:

### Quote Slide
- `quote`: The main quote text
- `author`: Attribution for the quote
- `quotation_marks`: Large decorative quotation mark

### Bullets Slide
- `title`: The slide title
- `bullet_points`: Text area for bullet points

### Agenda Slide
- `title`: The agenda title
- `agenda_items`: Text area for agenda items
- `image_area`: Optional area for an image

### Team Slide
- `title`: The team section title
- `members_area`: Area for team member profiles
- `member_image_style`: Styling for member photos

### Statement Slide
- `bg_color`: Background color (default: blue)
- `statement`: Large prominent text

### Dashboard Slide
- `title`: Dashboard title
- `charts`: Positions for four charts (top_left, top_right, bottom_left, bottom_right)
- `chart_style`: Default styling for charts

### Timeline Slide
- `title`: Timeline title
- `timeline_area`: Area for the timeline
- `step_colors`: List of colors for timeline steps
- `connector_color`: Color for connecting lines

## Part 4: TOML/JSON Template Files

EasyPPTX allows you to save and load templates as TOML or JSON files. TOML (Tom's Obvious Minimal Language) is a user-friendly configuration file format that is more readable than JSON, while maintaining full compatibility with the template system.

### TemplateManager

The `TemplateManager` class handles saving and loading templates:

```python
from easypptx import Presentation, TemplateManager

# Create a template manager (default saves in ~/.easypptx/templates)
tm = TemplateManager()

# Or specify a custom directory
tm = TemplateManager(template_dir="/path/to/templates")
```

### Saving Templates

You can save any template (built-in or custom) to TOML or JSON format:

```python
# Create a presentation and template manager
pres = Presentation()
tm = TemplateManager()

# Save a built-in template to TOML
toml_path = tm.save("title_slide", format="toml")
print(f"Template saved to: {toml_path}")

# Save a built-in template to JSON
json_path = tm.save("content_slide", format="json")
print(f"Template saved to: {json_path}")

# Create and save a custom template
custom_template = {
    "bg_color": "green",
    "title": {
        "text": "My Custom Template",
        "position": {"x": "10%", "y": "5%", "width": "80%", "height": "10%"},
        "font": {"name": "Arial", "size": 36, "bold": True},
        "align": "center",
        "vertical": "middle",
        "color": "white"
    },
    "content_area": {
        "position": {"x": "10%", "y": "20%", "width": "80%", "height": "70%"}
    }
}

# Register the custom template
tm.register("my_custom", custom_template)

# Save to TOML with a specific path
custom_path = tm.save("my_custom", file_path="templates/my_custom.toml")
```

### Loading Templates

You can load templates from TOML or JSON files:

```python
# Load a template from a TOML file
template_name = tm.load("templates/my_custom.toml")

# The template is now registered and can be used
slide = pres.add_slide_from_template(template_name)

# Load with a custom name
custom_name = tm.load("templates/my_custom.toml", template_name="renamed_template")
slide = pres.add_slide_from_template(custom_name)
```

### Benefits of TOML/JSON Templates

1. **Sharing**: Templates can be easily shared between projects and users
2. **Version Control**: Templates can be stored in version control systems
3. **Readability**: TOML format is more readable than JSON and easier to edit by hand
4. **Portability**: Templates are independent of the PowerPoint format
5. **Flexibility**: Templates can be modified with text editors or programmatically

### TOML vs JSON

EasyPPTX supports both formats, each with its advantages:

- **TOML**: More readable, better for manual editing
- **JSON**: More widely used, better for programmatic generation

### RGB Color Handling

EasyPPTX automatically handles the serialization and deserialization of `RGBColor` objects in templates:

```python
# Templates with RGB colors serialize/deserialize correctly
template_with_colors = {
    "bg_color": RGBColor(240, 240, 240),
    "title": {
        "color": RGBColor(10, 20, 30)
    }
}

# Register and save to TOML
tm.register("with_colors", template_with_colors)
tm.save("with_colors", format="toml")

# Load back and the RGB colors are preserved
loaded_name = tm.load("~/.easypptx/templates/with_colors.toml")
loaded_template = tm.get(loaded_name)

# RGB colors are reconstructed correctly
assert isinstance(loaded_template["bg_color"], RGBColor)
assert loaded_template["bg_color"][0] == 240
```

## See Also

Check out these example files for demonstrations of the template system:
- `examples/template_example.py` - Basic templates
- `examples/advanced_templates_example.py` - Advanced templates
- `examples/template_manager_example.py` - Template management
- `examples/template_toml_example.py` - TOML/JSON template export and import
