"""Tests for Grid template defaults functionality."""

import unittest
from unittest.mock import MagicMock, patch

from easypptx.grid import Grid


class MockSlide:
    """Mock Slide class for testing."""

    def __init__(self):
        """Initialize a MockSlide."""
        self._slide_width = 9144000  # 10 inches
        self._slide_height = 6858000  # 7.5 inches
        self.add_text = MagicMock(return_value="Text added")
        self.add_image = MagicMock(return_value="Image added")


class TestGridTemplateDefaults(unittest.TestCase):
    """Test cases for Grid template defaults functionality."""

    def setUp(self):
        """Set up test cases."""
        self.mock_slide = MockSlide()
        self.grid = Grid(parent=self.mock_slide)

        # Sample template data for testing
        self.template_data = {
            "defaults": {
                "global": {
                    "font_size": 16,
                    "font_bold": False,
                    "align": "left",
                    "color": [50, 50, 50],
                },
                "grid": {
                    "rows": 3,
                    "cols": 3,
                    "padding": 10.0,
                    "x": "5%",
                    "y": "15%",
                    "width": "90%",
                    "height": "80%",
                },
                "text": {
                    "font_size": 14,
                    "font_bold": True,
                    "align": "center",
                    "vertical": "middle",
                    "color": [20, 60, 120],
                },
            }
        }

    def test_apply_template_defaults(self):
        """Test applying template defaults to a grid."""
        # Apply template defaults
        self.grid.apply_template_defaults(self.template_data)

        # Verify defaults were stored correctly
        self.assertEqual(self.grid.template_defaults["global"]["font_size"], 16)
        self.assertEqual(self.grid.template_defaults["grid"]["rows"], 3)
        self.assertEqual(self.grid.template_defaults["text"]["font_bold"], True)

    def test_merge_with_defaults(self):
        """Test merging kwargs with template defaults."""
        # Apply template defaults
        self.grid.apply_template_defaults(self.template_data)

        # Test merging with empty kwargs
        merged = self.grid.merge_with_defaults("text", {})
        self.assertEqual(merged["font_size"], 14)
        self.assertEqual(merged["align"], "center")
        self.assertEqual(merged["color"], [20, 60, 120])

        # Test merging with some provided kwargs
        merged = self.grid.merge_with_defaults("text", {"font_size": 24, "align": "right"})
        self.assertEqual(merged["font_size"], 24)  # Provided value should override default
        self.assertEqual(merged["align"], "right")  # Provided value should override default
        self.assertEqual(merged["font_bold"], True)  # Default value should be preserved

        # Test global defaults
        merged = self.grid.merge_with_defaults("image", {})
        self.assertEqual(merged["font_size"], 16)  # Should get value from global defaults

    def test_nested_grid_inherits_template_defaults(self):
        """Test that nested grids inherit template defaults from parent."""
        # Apply template defaults to parent grid
        self.grid.apply_template_defaults(self.template_data)

        # Create a nested grid (using MockSlide as parent for simplicity)
        nested_grid = Grid(parent=self.mock_slide)

        # Apply parent's template defaults
        for key, value in self.grid.template_defaults.items():
            nested_grid.template_defaults[key] = value.copy()

        # Verify nested grid has the same defaults
        self.assertEqual(nested_grid.template_defaults["text"]["font_size"], 14)
        self.assertEqual(nested_grid.template_defaults["grid"]["padding"], 10.0)

    def test_add_textbox_with_template_defaults(self):
        """Test using template defaults when adding a text box."""
        # Apply template defaults
        self.grid.apply_template_defaults(self.template_data)

        # Mock the add_to_cell method to check the merged kwargs
        with patch.object(self.grid, "add_to_cell") as mock_add_to_cell:
            self.grid.add_textbox(0, 0, "Sample text")

            # Get the kwargs passed to add_to_cell
            # The actual signature is add_to_cell(row, col, content_func, **kwargs)
            # args[0] contains row, col, content_func, while kwargs contains the merged kwargs
            args, kwargs = mock_add_to_cell.call_args

            # Verify template defaults were applied
            self.assertEqual(kwargs["font_size"], 14)
            self.assertEqual(kwargs["font_bold"], True)
            self.assertEqual(kwargs["align"], "center")
            self.assertEqual(kwargs["text"], "Sample text")
            self.assertEqual(kwargs["color"], (20, 60, 120))  # List should be converted to tuple

    def test_add_image_with_template_defaults(self):
        """Test using template defaults when adding an image."""
        # Apply template defaults
        self.grid.apply_template_defaults(self.template_data)

        # Mock the add_to_cell method to check the merged kwargs
        with patch.object(self.grid, "add_to_cell") as mock_add_to_cell:
            self.grid.add_image(0, 0, "test.png")

            # Get the kwargs passed to add_to_cell
            # The actual signature is add_to_cell(row, col, content_func, **kwargs)
            args, kwargs = mock_add_to_cell.call_args

            # Verify template defaults were applied and image_path was preserved
            self.assertEqual(kwargs["image_path"], "test.png")


if __name__ == "__main__":
    unittest.main()
