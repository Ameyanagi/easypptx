# Dark Theme Support

EasyPPTX supports creating modern, visually appealing presentations with dark backgrounds and vibrant, high-contrast text colors.

## Overview

Dark-themed presentations have several advantages:

1. **Reduced eye strain**: Dark backgrounds with light text can be easier on the eyes, especially in low-light environments
2. **Modern appearance**: Dark themes are popular in contemporary design and can make presentations look more professional
3. **Color emphasis**: Bright colors stand out more dramatically against dark backgrounds
4. **Stylistic flexibility**: Dark backgrounds provide a good base for creative design elements

## Creating Dark-Themed Presentations

### Setting Default Background Color

You can set a default background color for all slides in a presentation:

```python
from easypptx import Presentation, Text

# Create a presentation with black background by default
pres = Presentation(default_bg_color="black")

# Create a presentation with dark blue background
pres = Presentation(default_bg_color=(0, 20, 40))  # RGB values
```

### Setting Individual Slide Backgrounds

You can also set different background colors for individual slides:

```python
# Add a slide with the default background color
slide1 = pres.add_slide()

# Add a slide with a specific background color
slide2 = pres.add_slide(bg_color="darkgray")

# Add a slide with a custom RGB color
slide3 = pres.add_slide(bg_color=(30, 30, 50))  # Dark purple-gray

# Update an existing slide's background
slide1.set_background_color("blue")
```

## High-Contrast Text and Elements

When using dark backgrounds, it's important to use high-contrast text and elements:

```python
text = Text(slide)

# Bright title text
text.add_title(
    "Dark Theme Presentation",
    color="cyan",
    align="center"
)

# White body text for readability
text.add_paragraph(
    "This text stands out against the dark background",
    color="white",
    font_size=24
)

# Use vibrant accent colors for emphasis
text.add_paragraph(
    "Important information",
    color="yellow",
    font_bold=True
)
```

## Expanded Color Palette

EasyPPTX includes an expanded color palette suitable for dark-themed presentations:

```python
# Basic colors
"black"        # RGBColor(0x10, 0x10, 0x10)
"white"        # RGBColor(0xFF, 0xFF, 0xFF)

# Grayscale shades
"darkgray"     # RGBColor(0x40, 0x40, 0x40)
"gray"         # RGBColor(0x80, 0x80, 0x80)
"lightgray"    # RGBColor(0xD0, 0xD0, 0xD0)

# Vibrant colors
"red"          # RGBColor(0xFF, 0x40, 0x40)
"green"        # RGBColor(0x40, 0xFF, 0x40)
"blue"         # RGBColor(0x40, 0x40, 0xFF)
"yellow"       # RGBColor(0xFF, 0xD7, 0x00)
"cyan"         # RGBColor(0x00, 0xE5, 0xFF)
"magenta"      # RGBColor(0xFF, 0x00, 0xFF)
"orange"       # RGBColor(0xFF, 0xA5, 0x00)
```

## Creating Visual Depth

To create visual depth in dark presentations, you can layer shapes and use gradient-like effects:

```python
# Add layered shapes for visual interest
slide.add_shape(
    shape_type=MSO_SHAPE.ROUNDED_RECTANGLE,
    x="10%",
    y="30%",
    width="80%",
    height="15%",
    fill_color=(0, 60, 100)  # Dark blue
)

slide.add_shape(
    shape_type=MSO_SHAPE.ROUNDED_RECTANGLE,
    x="15%",
    y="40%",
    width="70%",
    height="15%",
    fill_color=(0, 80, 100)  # Lighter blue
)

# Add text on top of the shapes
text.add_paragraph(
    "Layered design elements",
    x="20%",
    y="42%",
    width="60%",
    height="10%",
    color="white",
    align="center",
    vertical="middle"
)
```

## Best Practices for Dark Themes

1. **Maintain contrast**: Ensure there's sufficient contrast between text and background
2. **Use color sparingly**: Don't overuse bright colors; save them for emphasis
3. **Test readability**: Ensure all text is easily readable against the dark background
4. **Consider projection**: Test how your presentation looks when projected, as colors may appear differently
5. **Be consistent**: Maintain a consistent color scheme throughout the presentation

## Complete Example

See the [dark_theme_example.py](https://github.com/Ameyanagi/EasyPPTX/tree/main/examples/dark_theme_example.py) file for a complete example of creating dark-themed presentations.
