"""Tests for template TOML/JSON export and import."""

import os
import tempfile

from pptx.dml.color import RGBColor

from easypptx.template import Template, TemplateManager


class TestTemplateToml:
    """Test the Template and TemplateManager with TOML and JSON."""

    def test_template_init(self):
        """Test that Template initializes with presets."""
        template = Template()
        assert template.presets
        assert "title_slide" in template.presets
        assert "content_slide" in template.presets
        assert "section_slide" in template.presets

    def test_get_preset(self):
        """Test getting a preset from Template."""
        template = Template()
        preset = template.get_preset("title_slide")
        assert preset
        assert "title" in preset
        assert "subtitle" in preset

    def test_template_manager_init(self):
        """Test that TemplateManager initializes correctly."""
        # Test with default directory
        tm = TemplateManager()
        assert tm.template
        assert tm.template_dir
        assert os.path.isdir(tm.template_dir)

        # Test with custom directory
        with tempfile.TemporaryDirectory() as temp_dir:
            tm = TemplateManager(template_dir=temp_dir)
            assert tm.template_dir == temp_dir
            assert os.path.isdir(tm.template_dir)

    def test_register_and_get_template(self):
        """Test registering and getting a template."""
        tm = TemplateManager()

        # Create a custom template
        custom_template = {
            "bg_color": "blue",
            "title": {
                "text": "Custom Title",
                "position": {"x": "10%", "y": "10%", "width": "80%", "height": "20%"},
                "font": {"name": "Arial", "size": 36, "bold": True},
                "align": "center",
                "vertical": "middle",
                "color": "white",
            },
        }

        # Register the template
        tm.register("custom", custom_template)

        # Get the template
        retrieved = tm.get("custom")
        assert retrieved == custom_template

        # Check it's in list_templates
        assert "custom" in tm.list_templates()

    def test_rgb_color_serialization(self):
        """Test the RGB color serialization and deserialization."""
        tm = TemplateManager()

        # Test serialization
        color = RGBColor(10, 20, 30)
        serialized = tm._prepare_for_serialization(color)
        assert isinstance(serialized, dict)
        assert "__rgbcolor__" in serialized
        assert serialized["__rgbcolor__"] == [10, 20, 30]

        # Test deserialization
        deserialized = tm._process_after_deserialization(serialized)
        assert isinstance(deserialized, RGBColor)
        assert deserialized[0] == 10
        assert deserialized[1] == 20
        assert deserialized[2] == 30

        # Test complex structure
        complex_obj = {
            "color1": RGBColor(10, 20, 30),
            "nested": {"color2": RGBColor(40, 50, 60)},
            "list": [RGBColor(70, 80, 90), {"color3": RGBColor(100, 110, 120)}],
        }

        serialized = tm._prepare_for_serialization(complex_obj)
        deserialized = tm._process_after_deserialization(serialized)

        assert isinstance(deserialized["color1"], RGBColor)
        assert isinstance(deserialized["nested"]["color2"], RGBColor)
        assert isinstance(deserialized["list"][0], RGBColor)
        assert isinstance(deserialized["list"][1]["color3"], RGBColor)

    def test_toml_export_import(self):
        """Test exporting and importing templates with TOML."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create a TemplateManager with the temp directory
            tm = TemplateManager(template_dir=temp_dir)

            # Create a test template with RGBColor objects
            test_template = {
                "bg_color": RGBColor(240, 240, 240),
                "title": {
                    "text": "Test Template",
                    "position": {"x": "10%", "y": "10%", "width": "80%", "height": "20%"},
                    "font": {"name": "Arial", "size": 36, "bold": True},
                    "align": "center",
                    "vertical": "middle",
                    "color": RGBColor(10, 20, 30),
                },
            }

            # Register the template
            tm.register("test_toml", test_template)

            # Save to TOML
            toml_path = tm.save("test_toml", file_format="toml")
            assert os.path.exists(toml_path)

            # Clear registered templates
            tm.registered_templates = {}

            # Load from TOML
            loaded_name = tm.load(toml_path)
            assert loaded_name == "test_toml"

            # Get the loaded template
            loaded_template = tm.get("test_toml")

            # Check values
            assert loaded_template["bg_color"][0] == 240
            assert loaded_template["bg_color"][1] == 240
            assert loaded_template["bg_color"][2] == 240
            assert loaded_template["title"]["text"] == "Test Template"
            assert loaded_template["title"]["color"][0] == 10
            assert loaded_template["title"]["color"][1] == 20
            assert loaded_template["title"]["color"][2] == 30

    def test_json_export_import(self):
        """Test exporting and importing templates with JSON."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create a TemplateManager with the temp directory
            tm = TemplateManager(template_dir=temp_dir)

            # Create a test template with RGBColor objects
            test_template = {
                "bg_color": RGBColor(240, 240, 240),
                "title": {
                    "text": "Test Template",
                    "position": {"x": "10%", "y": "10%", "width": "80%", "height": "20%"},
                    "font": {"name": "Arial", "size": 36, "bold": True},
                    "align": "center",
                    "vertical": "middle",
                    "color": RGBColor(10, 20, 30),
                },
            }

            # Register the template
            tm.register("test_json", test_template)

            # Save to JSON
            json_path = tm.save("test_json", file_format="json")
            assert os.path.exists(json_path)

            # Clear registered templates
            tm.registered_templates = {}

            # Load from JSON
            loaded_name = tm.load(json_path)
            assert loaded_name == "test_json"

            # Get the loaded template
            loaded_template = tm.get("test_json")

            # Check values
            assert loaded_template["bg_color"][0] == 240
            assert loaded_template["bg_color"][1] == 240
            assert loaded_template["bg_color"][2] == 240
            assert loaded_template["title"]["text"] == "Test Template"
            assert loaded_template["title"]["color"][0] == 10
            assert loaded_template["title"]["color"][1] == 20
            assert loaded_template["title"]["color"][2] == 30

    def test_builtin_template_export_import(self):
        """Test exporting and importing built-in templates."""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create a TemplateManager with the temp directory
            tm = TemplateManager(template_dir=temp_dir)

            # Export a built-in template
            toml_path = tm.save(template_name="title_slide", file_path=None, file_format="toml")
            assert os.path.exists(toml_path)

            # Load with a custom name
            loaded_name = tm.load(toml_path, "my_title_slide")
            assert loaded_name == "my_title_slide"

            # Get the loaded template
            loaded_template = tm.get("my_title_slide")
            assert "title" in loaded_template
            assert "subtitle" in loaded_template

    def test_get_image_style(self):
        """Test getting image style from a template."""
        template = Template()

        # Test with a preset that has image_style
        preset = {"image_style": {"border": True, "border_color": "red", "shadow": True}}

        style = template.get_image_style(preset)
        assert style["border"] is True
        assert style["border_color"] == "red"
        assert style["shadow"] is True

        # Test with defaults
        default_style = template.get_image_style({})
        assert default_style["border"] is False
        assert default_style["border_color"] == "black"
        assert default_style["shadow"] is False

    def test_get_table_style(self):
        """Test getting table style from a template."""
        template = Template()

        # Test with a preset that has table_style
        preset = {
            "table_style": {
                "first_row": {"bg_color": "red", "text_color": "white"},
                "banded_rows": False,
                "border_color": "blue",
            }
        }

        style = template.get_table_style(preset)
        assert style["first_row"]["bg_color"] == "red"
        assert style["first_row"]["text_color"] == "white"
        assert style["banded_rows"] is False
        assert style["border_color"] == "blue"

        # Test with defaults
        default_style = template.get_table_style({})
        # The default style should have 'blue' bg_color
        assert isinstance(default_style["first_row"], dict)
        assert "bg_color" in default_style["first_row"]
        assert default_style["banded_rows"] is True

    def test_get_chart_style(self):
        """Test getting chart style from a template."""
        template = Template()

        # Test with a preset that has chart_style
        preset = {"chart_style": {"chart_type": "pie", "has_legend": False, "has_title": True, "has_data_labels": True}}

        style = template.get_chart_style(preset)
        assert style["chart_type"] == "pie"
        assert style["has_legend"] is False
        assert style["has_title"] is True
        assert style["has_data_labels"] is True

        # Test with defaults
        default_style = template.get_chart_style({})
        assert default_style["chart_type"] == "column"
        assert default_style["has_legend"] is True
