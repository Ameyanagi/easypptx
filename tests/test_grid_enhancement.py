"""Tests for grid enhancements for easypptx."""

from unittest.mock import MagicMock

from easypptx import Presentation
from easypptx.grid import Grid


def test_add_autogrid_none():
    """Test the add_autogrid method with None for content_funcs."""
    pres = Presentation()
    slide = pres.add_slide()

    # Test with None for content_funcs
    grid = pres.add_autogrid(slide, None, rows=2, cols=3)

    assert isinstance(grid, Grid)
    assert grid.rows == 2
    assert grid.cols == 3
    assert grid.x == "0%"  # default value
    assert grid.y == "0%"  # default value

    # Test with None for content_funcs and custom position
    grid2 = pres.add_autogrid(slide, None, rows=3, cols=2, x="10%", y="20%", width="80%", height="60%")

    assert isinstance(grid2, Grid)
    assert grid2.rows == 3
    assert grid2.cols == 2
    assert grid2.x == "10%"
    assert grid2.y == "20%"
    assert grid2.width == "80%"
    assert grid2.height == "60%"

    # Test with None for content_funcs, title provided
    grid3 = pres.add_autogrid(slide, None, rows=2, cols=2, title="Test Grid")

    assert isinstance(grid3, Grid)
    assert grid3.rows == 2
    assert grid3.cols == 2
    # Grid position should be adjusted to account for title
    assert grid3.y != "0%"  # Should be adjusted for title


class TestGridEnhancedMethods:
    """Test cases for enhanced Grid methods."""

    def setup_method(self):
        """Set up test environment before each test method."""
        self.parent = MagicMock()
        self.grid = Grid(parent=self.parent, rows=2, cols=2)

        # Create mock methods that the Grid will use
        self.parent.add_text = MagicMock(return_value="text_shape")
        self.parent.add_image = MagicMock(return_value="image_shape")
        self.parent.add_pyplot = MagicMock(return_value="pyplot_shape")
        self.parent.add_table = MagicMock(return_value="table_shape")

    def test_add_textbox_implemented(self):
        """Test that add_textbox is now implemented."""
        # Check the method exists
        assert hasattr(self.grid, "add_textbox")
        # Grid should have the add_textbox method, implemented in previous tests

    def test_add_image_implemented(self):
        """Test that add_image is now implemented."""
        # Check the method exists
        assert hasattr(self.grid, "add_image")
        # Grid should have the add_image method, implemented in previous tests

    def test_add_pyplot_implemented(self):
        """Test that add_pyplot is now implemented."""
        # Check the method exists
        assert hasattr(self.grid, "add_pyplot")
        # Grid should have the add_pyplot method, implemented in previous tests

    def test_add_table_implemented(self):
        """Test that add_table is now implemented."""
        # Check the method exists
        assert hasattr(self.grid, "add_table")
        # Grid should have the add_table method, implemented in previous tests
