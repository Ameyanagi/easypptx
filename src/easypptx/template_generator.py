"""Module for generating default templates."""

import os
from pathlib import Path

import tomli_w


def generate_default_template(output_path: str | Path | None = None) -> dict:
    """Generate a basic default template with global defaults.

    This function creates a template with global defaults for all methods.
    The template follows a cascading defaults system:
    1. Method-specific defaults override global defaults
    2. Global defaults apply to all methods unless overridden
    3. If not specified, hardcoded defaults apply

    Args:
        output_path: Optional path to save the template TOML file. If None, only returns the template dict.

    Returns:
        The template as a Python dictionary
    """
    template = {
        # Background color (RGB tuple as list)
        "bg_color": [255, 255, 255],
        # Global defaults apply to all methods
        "defaults": {
            # Global defaults for all methods
            "global": {
                # Positions and dimensions
                "x": "5%",
                "y": "5%",
                "width": "90%",
                "height": "10%",
                # Text formatting
                "font_name": "Arial",
                "font_size": 16,
                "font_bold": False,
                "font_italic": False,
                "align": "left",
                "vertical": "top",
                "color": [50, 50, 50],
                # Padding
                "padding": 5.0,
                # Content spacing
                "title_padding": "5%",
                "subtitle_padding": "5%",
                "content_padding": "5%",
                # Grid defaults
                "rows": 2,
                "cols": 2,
                "title_height": "12%",
                "subtitle_height": "8%",
                "title_font_size": 32,
                "subtitle_font_size": 22,
                "title_align": "center",
                "subtitle_align": "center",
                "content_y_padding": "22%",
                # Image defaults
                "maintain_aspect_ratio": True,
                "border": False,
                "shadow": False,
                # Shape defaults
                "fill_color": [230, 230, 240],
                "line_color": [200, 200, 210],
                "line_width": 1.0,
                # Table defaults
                "first_row_header": True,
                "banded_rows": True,
                # Chart defaults
                "chart_type": "column",
                "has_legend": True,
                "legend_position": "bottom",
                "has_title": True,
                "has_data_labels": False,
                "gridlines": True,
            }
        },
    }

    # Save template to TOML file if output_path is provided
    if output_path:
        output_path = Path(output_path)

        # Ensure parent directory exists
        os.makedirs(output_path.parent, exist_ok=True)

        # Write template to TOML file
        with open(output_path, "wb") as f:
            tomli_w.dump(template, f)

        print(f"Default template saved to: {output_path}")
        print("Edit the template file directly to add comments and customize settings.")

    return template


def generate_template_with_comments(output_path: str | Path) -> None:
    """Generate a comprehensive default template with comments for all available settings.

    This function first generates a basic template with global defaults,
    then adds commented sections for all other settings.

    Args:
        output_path: Path to save the template TOML file.
    """
    # First generate the basic template
    generate_default_template(output_path)

    # Now append commented sections
    with open(output_path, "a") as f:
        f.write("\n\n# -----------------------------------------------------\n")
        f.write("# COMMENTED SECTIONS - UNCOMMENT AND MODIFY AS NEEDED\n")
        f.write("# -----------------------------------------------------\n\n")

        # Reference PPTX
        f.write("# Reference PowerPoint file (optional)\n")
        f.write('# reference_pptx = "path/to/your/reference.pptx"\n\n')

        # Method-specific defaults for grid slides
        f.write("# Method-specific defaults for grid slides\n")
        f.write("# [defaults.grid_slide]\n")
        f.write("# rows = 2\n")
        f.write("# cols = 2\n")
        f.write('# title_height = "12%"\n')
        f.write('# subtitle_height = "8%"\n')
        f.write("# title_font_size = 32\n")
        f.write("# subtitle_font_size = 22\n")
        f.write('# title_align = "center"\n')
        f.write('# subtitle_align = "center"\n')
        f.write("# padding = 8.0\n")
        f.write('# content_y_padding = "22%"  # Allow space for title + subtitle\n\n')

        # Grid-specific defaults
        f.write("# Method-specific defaults for grid objects and cells\n")
        f.write("# [defaults.grid]\n")
        f.write("# rows = 2\n")
        f.write("# cols = 2\n")
        f.write("# padding = 8.0\n")
        f.write('# x = "5%"\n')
        f.write('# y = "15%"  # Allow space for title\n')
        f.write('# width = "90%"\n')
        f.write('# height = "80%"\n\n')

        # Text defaults
        f.write("# Method-specific defaults for text elements\n")
        f.write("# [defaults.text]\n")
        f.write("# font_size = 16\n")
        f.write('# font_name = "Arial"\n')
        f.write('# align = "left"\n')
        f.write('# vertical = "top"\n')
        f.write("# color = [40, 40, 45]  # Dark gray\n\n")

        # Image defaults
        f.write("# Method-specific defaults for image elements\n")
        f.write("# [defaults.image]\n")
        f.write('# x = "10%"\n')
        f.write('# y = "10%"\n')
        f.write('# width = "80%"\n')
        f.write('# height = "70%"\n')
        f.write("# maintain_aspect_ratio = true\n")
        f.write("# border = false\n")
        f.write("# shadow = false\n\n")

        # Shape defaults
        f.write("# Method-specific defaults for shape elements\n")
        f.write("# [defaults.shape]\n")
        f.write("# fill_color = [230, 230, 240]\n")
        f.write("# line_color = [200, 200, 210]\n")
        f.write("# line_width = 1.0\n")
        f.write("# shadow = false\n\n")

        # Table defaults
        f.write("# Method-specific defaults for table elements\n")
        f.write("# [defaults.table]\n")
        f.write("# font_size = 14\n")
        f.write("# border = true\n")
        f.write("# first_row_header = true\n")
        f.write("# banded_rows = true\n\n")
        f.write("# [defaults.table.header_style]\n")
        f.write("# bg_color = [30, 50, 100]\n")
        f.write("# text_color = [255, 255, 255]\n")
        f.write("# font_bold = true\n\n")
        f.write("# [defaults.table.row_style]\n")
        f.write("# bg_color = [240, 240, 245]\n")
        f.write("# text_color = [40, 40, 45]\n\n")

        # Chart defaults
        f.write("# Method-specific defaults for chart elements\n")
        f.write("# [defaults.chart]\n")
        f.write('# chart_type = "column"\n')
        f.write("# has_legend = true\n")
        f.write('# legend_position = "bottom"\n')
        f.write("# has_title = true\n")
        f.write("# title_font_size = 14\n")
        f.write("# has_data_labels = false\n")
        f.write("# gridlines = true\n\n")

        # Theme elements
        f.write("# Title element configuration (optional theme element)\n")
        f.write("# [title]\n")
        f.write('# align = "center"\n')
        f.write('# vertical = "middle"\n')
        f.write("# color = [30, 50, 100]  # Dark blue\n\n")
        f.write("# [title.position]\n")
        f.write('# x = "5%"\n')
        f.write('# y = "5%"\n')
        f.write('# width = "90%"\n')
        f.write('# height = "12%"\n\n')
        f.write("# [title.font]\n")
        f.write('# name = "Arial"\n')
        f.write("# size = 32\n")
        f.write("# bold = true\n\n")
        f.write("# [title.padding]\n")
        f.write('# top = "2%"\n')
        f.write('# left = "5%"\n\n')

        # Subtitle element
        f.write("# Subtitle element configuration (optional theme element)\n")
        f.write("# [subtitle]\n")
        f.write('# align = "center"\n')
        f.write('# vertical = "middle"\n')
        f.write("# color = [60, 80, 120]  # Medium blue\n\n")
        f.write("# [subtitle.position]\n")
        f.write('# x = "5%"\n')
        f.write('# y = "18%"\n')
        f.write('# width = "90%"\n')
        f.write('# height = "8%"\n\n')
        f.write("# [subtitle.font]\n")
        f.write('# name = "Arial"\n')
        f.write("# size = 22\n")
        f.write("# bold = false\n\n")

        # Footer element
        f.write("# Footer element configuration (optional theme element)\n")
        f.write("# [footer]\n")
        f.write('# align = "right"\n')
        f.write('# vertical = "middle"\n')
        f.write("# color = [100, 100, 120]  # Gray\n\n")
        f.write("# [footer.position]\n")
        f.write('# x = "5%"\n')
        f.write('# y = "90%"\n')
        f.write('# width = "90%"\n')
        f.write('# height = "5%"\n\n')
        f.write("# [footer.font]\n")
        f.write('# name = "Arial"\n')
        f.write("# size = 12\n")
        f.write("# bold = false\n\n")

    print(f"Added commented sections to template: {output_path}")


# Example usage
if __name__ == "__main__":
    # Generate template with comments
    template_path = Path(__file__).parent.parent.parent / "templates" / "default_template.toml"
    generate_template_with_comments(template_path)
