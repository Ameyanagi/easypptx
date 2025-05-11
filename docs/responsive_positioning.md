# Responsive Positioning

EasyPPTX supports responsive positioning that adapts to different aspect ratios, ensuring your presentations look great across different display formats.

## Problem: Aspect Ratio Challenges

When presentations are viewed or exported with different aspect ratios (e.g., switching from 16:9 to 4:3), elements that were centered in one ratio may appear off-center in another. This is because:

1. The ratio between width and height changes
2. Standard percentage-based positioning doesn't account for this change
3. Content tends to get pushed to the left in wider aspect ratios

## Solution: Responsive Positioning

EasyPPTX implements responsive positioning that automatically adjusts element positions based on the aspect ratio, ensuring consistent layout across different formats.

### Using Horizontal Alignment

All element positioning methods in EasyPPTX (add_text, add_image, add_shape) support an optional `h_align` parameter:

```python
# Add a centered title that adjusts for different aspect ratios
slide.add_text(
    text="Centered Title",
    x="50%",
    y="5%",
    width="80%",
    height="10%",
    align="center",
    h_align="center"  # Enable responsive horizontal alignment
)

# Add a wide image that stays centered
slide.add_image(
    image_path="path/to/image.jpg",
    x="10%",
    y="20%",
    width="80%",
    height="60%",
    h_align="center"  # Enable responsive horizontal alignment
)
```

### Automatic Centering

Several methods apply responsive centering automatically when appropriate:

1. Text elements with `align="center"` automatically use `h_align="center"`
2. Images with width > 50% automatically use `h_align="center"`
3. Shapes with width > 50% automatically use `h_align="center"`

This means most centered content will automatically adapt to different aspect ratios without explicit configuration.

## How It Works

When responsive positioning is enabled, EasyPPTX:

1. Detects the current aspect ratio of the presentation
2. Compares it to the standard 16:9 ratio
3. Calculates an adjustment factor based on the ratio difference
4. Applies this adjustment to horizontal positions
5. Ensures elements remain visually balanced regardless of dimensions

## Example

The `responsive_positioning_example.py` file demonstrates this feature by creating presentations with different aspect ratios (16:9, 4:3, 16:10, A4) and showing how elements maintain their visual balance across formats.

```python
from easypptx import Presentation

# Create presentations with different aspect ratios
pres = Presentation(aspect_ratio="4:3")

# Add a slide with responsive positioning
slide = pres.add_slide()
pres.add_text(
    slide=slide,
    text="Centered Content",
    x="50%",
    y="5%",
    width="80%",
    height="10%",
    font_size=32,
    align="center",
    h_align="center"  # Enable responsive positioning
)

# Save the presentation
pres.save("responsive_example_4_3.pptx")
```

## Best Practices

1. **Use responsive positioning for key elements**: Apply it to titles, important images, and central content
2. **Use percentage-based positioning**: Always use percentages (e.g., "50%") for responsive layouts
3. **Center align important content**: Text with `align="center"` automatically gets responsive positioning
4. **Explicitly set `h_align="center"` for large elements**: For maximum control, explicitly set h_align

## Compatibility

Responsive positioning is compatible with all other EasyPPTX features, including:
- Template presets
- Styling options
- Color and font customization
- Image and shape manipulation

## Default Behavior

To ensure backward compatibility, responsive positioning is applied selectively:

1. **Explicitly centered elements**: When `h_align="center"` is specified
2. **Centered text**: When `align="center"` is specified
3. **Wide images/shapes**: When width > 50% and centered positioning is likely desired

For other elements, standard positioning is used to maintain precise control.
