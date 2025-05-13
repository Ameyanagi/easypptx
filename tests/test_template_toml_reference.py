"""Tests for template TOML reference handling."""

import os
import tempfile

import tomli_w

from easypptx import Presentation
from easypptx.template import TemplateManager


def test_template_toml_color_list_to_tuple_conversion():
    """Test conversion of color list from TOML to tuples for template handling."""
    # Create a temporary TOML template file
    with tempfile.NamedTemporaryFile(suffix=".toml", delete=False) as temp_file:
        # Write test template with list color format
        template_data = {
            "bg_color": [255, 255, 255],  # Color as a list
            "title": {
                "text": "Test Template",
                "position": {"x": "10%", "y": "30%", "width": "80%", "height": "20%"},
                "font": {"name": "Arial", "size": 32, "bold": True},
                "align": "center",
                "vertical": "middle",
                "color": [0, 102, 204],  # Color as a list
            },
        }
        tomli_w.dump(template_data, temp_file)
        temp_file_path = temp_file.name

    try:
        # Initialize the template manager
        template_manager = TemplateManager()

        # Load custom template
        template_name = template_manager.load(temp_file_path)

        # Get template data from manager
        template_data = template_manager.get(template_name)

        # Check that bg_color is available (still as a list)
        assert "bg_color" in template_data
        assert isinstance(template_data["bg_color"], list)

        # Create presentation instance
        presentation = Presentation()

        # Add template to presentation's template manager
        presentation.template_manager.register(template_name, template_data)

        # The following line should not raise an error if bg_color is properly converted
        slide = presentation.add_slide_from_template(template_name)

        # Verify slide was created successfully
        assert slide is not None

    finally:
        # Clean up temporary file
        os.unlink(temp_file_path)


def test_template_toml_reference_path_resolution():
    """Test that reference_pptx paths in TOML templates are properly resolved."""
    # Create a temporary TOML template file
    with tempfile.NamedTemporaryFile(suffix=".toml", delete=False) as temp_file:
        # Path to built-in reference PPTX
        ref_path = "../src/easypptx/reference_16x9.pptx"

        # Write test template with reference PPTX
        template_data = {
            "reference_pptx": ref_path,
            "blank_layout_index": 6,
            "bg_color": [240, 240, 240],  # Light gray background
            "title": {
                "text": "Reference Template Test",
                "position": {"x": "10%", "y": "30%", "width": "80%", "height": "20%"},
                "font": {"name": "Arial", "size": 32, "bold": True},
                "align": "center",
            },
        }
        tomli_w.dump(template_data, temp_file)
        temp_file_path = temp_file.name

    try:
        # Initialize the template manager
        template_manager = TemplateManager()

        # Load custom template
        template_name = template_manager.load(temp_file_path)

        # Check if reference path is registered correctly
        ref_pptx = template_manager.get_reference_pptx(template_name)
        assert ref_pptx is not None
        # Ensure path is stored and resolved correctly
        assert os.path.basename(ref_path) == os.path.basename(ref_pptx)

        # Create presentation instance
        presentation = Presentation()

        # Register template and reference with presentation's template manager
        presentation.template_manager.register(template_name, template_manager.get(template_name))
        presentation.template_manager.template_references[template_name] = ref_pptx

        # The following line should create a slide using the reference PPTX
        slide = presentation.add_slide_from_template(template_name)

        # Verify slide was created successfully
        assert slide is not None

    finally:
        # Clean up temporary file
        os.unlink(temp_file_path)
