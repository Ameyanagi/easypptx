"""Template module for EasyPPTX."""

import json
import os
from pathlib import Path
from typing import Any, cast

import tomli
import tomli_w
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN


class Template:
    """Class for managing presentation templates and presets."""

    def __init__(self) -> None:
        """Initialize a Template object with predefined presets."""
        # Define text alignment mappings
        self.align_dict = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}

        # Define vertical alignment mappings
        self.vertical_dict = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}

        # Store reference PPTX paths for presets
        self.preset_references: dict[str, str] = {}

        # Store blank layout indices for presets
        self.preset_blank_layouts: dict[str, int] = {}

        # Define standard colors
        self.color_dict = {
            "black": RGBColor(0x40, 0x40, 0x40),
            "darkgray": RGBColor(0x40, 0x40, 0x40),
            "gray": RGBColor(0x80, 0x80, 0x80),
            "lightgray": RGBColor(0xD0, 0xD0, 0xD0),
            "red": RGBColor(0xFF, 0x40, 0x40),
            "green": RGBColor(0x40, 0xFF, 0x40),
            "blue": RGBColor(0x40, 0x40, 0xFF),
            "white": RGBColor(0xFF, 0xFF, 0xFF),
            "yellow": RGBColor(0xFF, 0xD7, 0x00),
            "cyan": RGBColor(0x00, 0xE5, 0xFF),
            "magenta": RGBColor(0xFF, 0x00, 0xFF),
            "orange": RGBColor(0xFF, 0xA5, 0x00),
        }

        # Define default image styling
        self.default_image_style = {
            "border": False,
            "border_color": "black",
            "border_width": 1,
            "shadow": False,
            "rounded_corners": False,
            "maintain_aspect_ratio": True,
            "center": True,
            "brightness": 0,
            "contrast": 0,
        }

        # Define default table styling
        self.default_table_style = {
            "first_row": {"bold": True, "bg_color": "blue", "text_color": "white"},
            "banded_rows": True,
            "band_color": "lightgray",
            "border_color": "black",
            "border_width": 1,
            "header_border_width": 2,
            "text_align": "center",
            "header_align": "center",
            "font_name": "Meiryo",
            "font_size": 12,
            "header_font_size": 14,
        }

        # Define default chart styling
        self.default_chart_style = {
            "chart_type": "column",  # column, bar, line, pie, scatter, area
            "has_legend": True,
            "legend_position": "bottom",  # top, bottom, left, right
            "has_title": True,
            "title_font_size": 14,
            "palette": [
                RGBColor(0x5B, 0x9B, 0xD5),  # Blue
                RGBColor(0xED, 0x7D, 0x31),  # Orange
                RGBColor(0xA5, 0xA5, 0xA5),  # Gray
                RGBColor(0xFF, 0xC0, 0x00),  # Yellow
                RGBColor(0x4C, 0xAF, 0x50),  # Green
                RGBColor(0x9C, 0x27, 0xB0),  # Purple
            ],
            "has_data_labels": False,
            "gridlines": True,
            "has_border": True,
            "border_color": "black",
        }

        # Define predefined template presets
        self.presets = {
            # Title slide with title and subtitle
            "title_slide": {
                "title": {
                    "text": "Presentation Title",
                    "position": {"x": "10%", "y": "25%", "width": "80%", "height": "30%"},
                    "font": {"name": "Meiryo", "size": 44, "bold": True},
                    "align": "center",
                    "vertical": "middle",
                    "color": "black",
                },
                "subtitle": {
                    "text": "Subtitle or Author",
                    "position": {"x": "20%", "y": "60%", "width": "60%", "height": "20%"},
                    "font": {"name": "Meiryo", "size": 24, "bold": False},
                    "align": "center",
                    "vertical": "middle",
                    "color": "black",
                },
            },
            # Content slide with title and horizontal bar
            "content_slide": {
                "title": {
                    "text": "Content Slide",
                    "position": {"x": "1%", "y": "2%", "width": "95%", "height": "5%"},
                    "font": {"name": "Meiryo", "size": 30, "bold": True},
                    "align": "left",
                    "vertical": "top",
                    "color": "black",
                },
                "bar": {
                    "position": {"x": "0%", "y": "10%", "width": "100%", "height": "2%"},
                    "gradient": {
                        "start_color": RGBColor(0xE0, 0xE5, 0xF7),
                        "end_color": RGBColor(0x95, 0xAB, 0xEA),
                        "angle": 0,
                    },
                },
                "content_area": {"position": {"x": "5%", "y": "15%", "width": "90%", "height": "80%"}},
            },
            # Section slide with prominent title on colored background
            "section_slide": {
                "bg_color": "blue",
                "title": {
                    "text": "Section Title",
                    "position": {"x": "10%", "y": "40%", "width": "80%", "height": "20%"},
                    "font": {"name": "Meiryo", "size": 44, "bold": True},
                    "align": "center",
                    "vertical": "middle",
                    "color": "white",
                },
            },
            # Image slide with title and centered image
            "image_slide": {
                "title": {
                    "text": "Image Title",
                    "position": {"x": "10%", "y": "5%", "width": "80%", "height": "10%"},
                    "font": {"name": "Meiryo", "size": 30, "bold": True},
                    "align": "center",
                    "vertical": "middle",
                    "color": "black",
                },
                "image_area": {
                    "position": {"x": "10%", "y": "20%", "width": "80%", "height": "70%"},
                },
                "image_style": {
                    "border": True,
                    "border_color": "gray",
                    "border_width": 1,
                    "shadow": True,
                    "maintain_aspect_ratio": True,
                    "center": True,
                },
            },
            # Table slide with title and data table
            "table_slide": {
                "title": {
                    "text": "Table Title",
                    "position": {"x": "10%", "y": "5%", "width": "80%", "height": "10%"},
                    "font": {"name": "Meiryo", "size": 30, "bold": True},
                    "align": "center",
                    "vertical": "middle",
                    "color": "black",
                },
                "table_area": {
                    "position": {"x": "10%", "y": "20%", "width": "80%", "height": "70%"},
                },
                "table_style": {
                    "first_row": {"bold": True, "bg_color": "blue", "text_color": "white"},
                    "banded_rows": True,
                    "band_color": "lightgray",
                },
            },
            # Chart slide with title and chart area
            "chart_slide": {
                "title": {
                    "text": "Chart Title",
                    "position": {"x": "10%", "y": "5%", "width": "80%", "height": "10%"},
                    "font": {"name": "Meiryo", "size": 30, "bold": True},
                    "align": "center",
                    "vertical": "middle",
                    "color": "black",
                },
                "chart_area": {
                    "position": {"x": "10%", "y": "20%", "width": "80%", "height": "70%"},
                },
                "chart_style": {
                    "chart_type": "column",
                    "has_legend": True,
                    "legend_position": "bottom",
                },
            },
            # Comparison slide with title and two content areas
            "comparison_slide": {
                "title": {
                    "text": "Comparison",
                    "position": {"x": "10%", "y": "5%", "width": "80%", "height": "10%"},
                    "font": {"name": "Meiryo", "size": 30, "bold": True},
                    "align": "center",
                    "vertical": "middle",
                    "color": "black",
                },
                "left_content": {
                    "position": {"x": "5%", "y": "20%", "width": "42%", "height": "75%"},
                },
                "right_content": {
                    "position": {"x": "53%", "y": "20%", "width": "42%", "height": "75%"},
                },
            },
        }

    def get_preset(self, preset_name: str) -> dict:
        """Get a preset template by name.

        Args:
            preset_name: Name of the preset to retrieve

        Returns:
            The preset template dictionary

        Raises:
            ValueError: If the preset name doesn't exist
        """
        if preset_name not in self.presets:
            valid_presets = ", ".join(self.presets.keys())
            raise ValueError(f"Unknown preset: {preset_name}. Valid presets are: {valid_presets}")

        return cast(dict, self.presets[preset_name])

    def create_custom_preset(self, **kwargs) -> dict:
        """Create a custom preset with specified parameters."""
        # Implementation here
        return {}

    def get_image_style(self, preset: dict) -> dict:
        """Get image styling from a preset.

        Args:
            preset: Template preset dictionary

        Returns:
            Dictionary of image styling options
        """
        # Start with default styling
        style = self.default_image_style.copy()

        # Update with style from preset if available
        if "image_style" in preset:
            style.update(preset["image_style"])

        return style

    def get_table_style(self, preset: dict) -> dict:
        """Get table styling from a preset.

        Args:
            preset: Template preset dictionary

        Returns:
            Dictionary of table styling options
        """
        # Start with default styling
        style = self.default_table_style.copy()

        # Update with style from preset if available
        if "table_style" in preset:
            # Handle nested dictionaries by updating them separately
            table_style = cast(dict, preset.get("table_style", {}))
            if "first_row" in table_style and "first_row" in style:
                first_row_style = cast(dict, style.get("first_row", {}))
                first_row_style.update(table_style.get("first_row", {}))

            # Update all other keys
            for key, value in table_style.items():
                if key != "first_row":
                    style[key] = value

        return style

    def get_chart_style(self, preset: dict) -> dict:
        """Get chart styling from a preset.

        Args:
            preset: Template preset dictionary

        Returns:
            Dictionary of chart styling options
        """
        # Start with default styling
        style = self.default_chart_style.copy()

        # Update with style from preset if available
        if "chart_style" in preset:
            style.update(preset["chart_style"])

        return style

    def set_reference_pptx(self, preset_name: str, reference_pptx: str, blank_layout_index: int | None = None) -> None:
        """Set a reference PPTX file for a preset.

        Args:
            preset_name: Name of the preset
            reference_pptx: Path to the reference PPTX file
            blank_layout_index: Index of the blank layout (default: None, auto-detect)

        Raises:
            ValueError: If the preset name doesn't exist
        """
        if preset_name not in self.presets:
            valid_presets = ", ".join(self.presets.keys())
            raise ValueError(f"Unknown preset: {preset_name}. Valid presets are: {valid_presets}")

        self.preset_references[preset_name] = reference_pptx

        if blank_layout_index is not None:
            self.preset_blank_layouts[preset_name] = blank_layout_index

    def get_reference_pptx(self, preset_name: str) -> str | None:
        """Get the reference PPTX file path for a preset if specified.

        Args:
            preset_name: Name of the preset

        Returns:
            Path to the reference PPTX file or None if not specified
        """
        return self.preset_references.get(preset_name)

    def get_blank_layout_index(self, preset_name: str) -> int | None:
        """Get the blank layout index for a preset if specified.

        Args:
            preset_name: Name of the preset

        Returns:
            Index of the blank layout or None if not specified
        """
        return self.preset_blank_layouts.get(preset_name)


class TemplateManager:
    """Class for managing templates, providing easy save/load functionality."""

    def __init__(self, template_dir: str | None = None):
        """Initialize a TemplateManager.

        Args:
            template_dir: Optional directory for template files (default: None)
        """
        # Initialize the Template object for built-in templates
        self.template = Template()

        # Dictionary for additional registered templates
        self.registered_templates: dict[str, dict[str, Any]] = {}

        # Dictionary for reference PPTX file paths associated with templates
        self.template_references: dict[str, str] = {}

        # Dictionary for blank layout indices associated with templates
        self.blank_layout_indices: dict[str, int] = {}

        # Set the template directory
        self.template_dir = template_dir
        if self.template_dir is None:
            # Use default directory in user's home directory
            self.template_dir = os.path.join(str(Path.home()), ".easypptx", "templates")

        # Create the template directory if it doesn't exist
        os.makedirs(self.template_dir, exist_ok=True)

    def get(self, template_name: str) -> dict:
        """Get a template by name.

        Args:
            template_name: Name of the template

        Returns:
            Template dictionary

        Raises:
            ValueError: If the template is not found
        """
        # First check registered templates
        if template_name in self.registered_templates:
            return cast(dict, self.registered_templates[template_name])

        # Then check built-in templates
        if template_name in self.template.presets:
            return cast(dict, self.template.presets[template_name])

        # If not found, raise an error
        raise ValueError(
            f"Template '{template_name}' not found. Available templates: " + ", ".join(self.list_templates())
        )

    def register(self, name: str, template: dict) -> None:
        """Register a custom template.

        Args:
            name: Name for the template
            template: Template dictionary
        """
        self.registered_templates[name] = template

    def list_templates(self) -> list:
        """List all available templates.

        Returns:
            List of template names
        """
        # Combine built-in and registered templates
        return list(self.template.presets.keys()) + list(self.registered_templates.keys())

    def get_reference_pptx(self, template_name: str) -> str | None:
        """Get the reference PPTX file path for a template if specified.

        Args:
            template_name: Name of the template

        Returns:
            Path to the reference PPTX file or None if not specified
        """
        return self.template_references.get(template_name)

    def get_blank_layout_index(self, template_name: str) -> int | None:
        """Get the blank layout index for a template if specified.

        Args:
            template_name: Name of the template

        Returns:
            Index of the blank layout or None if not specified
        """
        return self.blank_layout_indices.get(template_name)

    def save(self, template_name: str, file_path: str | None = None, file_format: str = "toml") -> str:
        """Save a template to a file.

        Args:
            template_name: Name of the template to save
            file_path: Path to save the template (default: None, uses name in template_dir)
            format: File format to use, 'toml' or 'json' (default: 'toml')

        Returns:
            Path where the template was saved

        Raises:
            ValueError: If the template is not found or format is invalid
        """
        # Get the template
        template = self.get(template_name)

        # Determine the file path
        if file_path is None:
            extension = ".toml" if file_format.lower() == "toml" else ".json"
            file_path = os.path.join(self.template_dir, f"{template_name}{extension}")

        # Create the directory if it doesn't exist
        os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)

        # Convert RGBColor objects to serializable values
        template_serializable = self._prepare_for_serialization(template)

        # Save the template to the file
        if file_format.lower() == "toml":
            # Convert to TOML and save
            with open(file_path, "wb") as f:
                tomli_w.dump(template_serializable, f)
        elif file_format.lower() == "json":
            # Convert to JSON and save
            with open(file_path, "w") as f:
                json.dump(template_serializable, f, indent=2)
        else:
            raise ValueError(f"Unsupported format: {file_format}. Use 'toml' or 'json'")

        return file_path

    def load(self, file_path: str, template_name: str | None = None) -> str:
        """Load a template from a file.

        Args:
            file_path: Path to the template file
            template_name: Name to register the template as (default: None, uses filename)

        Returns:
            The name the template was registered as

        Raises:
            FileNotFoundError: If the template file is not found
            ValueError: If the file format is invalid
        """
        # Check if the file exists
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Template file not found: {file_path}")

        # Determine the template name
        if template_name is None:
            template_name = os.path.splitext(os.path.basename(file_path))[0]

        # Determine the file format from extension
        file_extension = os.path.splitext(file_path)[1].lower()

        # Get the base directory for relative paths
        base_dir = os.path.dirname(os.path.abspath(file_path))

        # Load the template from the file based on format
        if file_extension == ".toml":
            # Load TOML
            with open(file_path, "rb") as f:
                template_data = tomli.load(f)
        elif file_extension == ".json":
            # Load JSON
            with open(file_path) as f:
                template_data = json.load(f)
        else:
            raise ValueError(f"Unsupported file format: {file_extension}. Supported formats: .toml, .json")

        # Extract reference_pptx and blank_layout_index if present
        reference_pptx = template_data.pop("reference_pptx", None)
        blank_layout_index = template_data.pop("blank_layout_index", None)

        # Process reference_pptx path (resolve relative paths)
        if reference_pptx is not None:
            # If it's a relative path, make it absolute relative to the TOML file location
            if not os.path.isabs(reference_pptx):
                reference_pptx = os.path.normpath(os.path.join(base_dir, reference_pptx))

            # Store the reference path
            self.template_references[template_name] = reference_pptx

        # Store blank layout index if specified
        if blank_layout_index is not None:
            self.blank_layout_indices[template_name] = blank_layout_index

        # Convert serialized color values back to RGBColor objects
        template = self._process_after_deserialization(template_data)

        # Register the template
        self.register(template_name, template)

        return template_name

    def _prepare_for_serialization(self, obj: Any) -> Any:
        """Convert RGBColor objects to serializable values.

        Args:
            obj: Object to prepare for serialization

        Returns:
            Serializable version of the object
        """
        if isinstance(obj, dict):
            result = {}
            for key, value in obj.items():
                result[key] = self._prepare_for_serialization(value)
            return result
        elif isinstance(obj, list):
            return [self._prepare_for_serialization(item) for item in obj]
        elif isinstance(obj, RGBColor):
            # Convert RGBColor to a dictionary representation
            return {"__rgbcolor__": [obj[0], obj[1], obj[2]]}
        else:
            return obj

    def _process_after_deserialization(self, obj: Any) -> Any:
        """Convert serialized color values back to RGBColor objects.

        Args:
            obj: Object after deserialization

        Returns:
            Processed object with RGBColor objects
        """
        if isinstance(obj, dict):
            if "__rgbcolor__" in obj and len(obj) == 1:
                # Convert dictionary representation back to RGBColor
                rgb_values = obj["__rgbcolor__"]
                return RGBColor(rgb_values[0], rgb_values[1], rgb_values[2])
            else:
                # Process nested dictionaries
                result = {}
                for key, value in obj.items():
                    result[key] = self._process_after_deserialization(value)
                return result
        elif isinstance(obj, list):
            return [self._process_after_deserialization(item) for item in obj]
        else:
            return obj
