"""Tests for enhanced Grid methods."""

import unittest
from unittest.mock import MagicMock

from easypptx.grid import Grid


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

    def test_add_textbox(self):
        """Test the add_textbox method."""
        result = self.grid.add_textbox(row=0, col=1, text="Test Text", font_size=24, font_bold=True)

        # Verify that the parent's add_text method was called with the correct parameters
        self.parent.add_text.assert_called_once()
        call_kwargs = self.parent.add_text.call_args[1]
        assert "x" in call_kwargs
        assert "y" in call_kwargs
        assert "width" in call_kwargs
        assert "height" in call_kwargs
        assert call_kwargs["text"] == "Test Text"
        assert call_kwargs["font_size"] == 24
        assert call_kwargs["font_bold"] is True

        # Verify that the result is correct
        assert result == "text_shape"
        assert self.grid.cells[0][1].content == "text_shape"

    def test_add_image(self):
        """Test the add_image method."""
        result = self.grid.add_image(row=1, col=0, image_path="test.png", border=True)

        # Verify that the parent's add_image method was called with the correct parameters
        self.parent.add_image.assert_called_once()
        call_kwargs = self.parent.add_image.call_args[1]
        assert "x" in call_kwargs
        assert "y" in call_kwargs
        assert "width" in call_kwargs
        assert "height" in call_kwargs
        assert call_kwargs["image_path"] == "test.png"
        assert call_kwargs["border"] is True

        # Verify that the result is correct
        assert result == "image_shape"
        assert self.grid.cells[1][0].content == "image_shape"

    def test_add_pyplot(self):
        """Test the add_pyplot method renders the figure to an in-memory stream."""
        import io

        # Mock the figure and savefig methods
        mock_figure = MagicMock()
        mock_figure.savefig = MagicMock()

        result = self.grid.add_pyplot(row=1, col=1, figure=mock_figure, dpi=300)

        # Verify that figure.savefig was called with an in-memory stream
        mock_figure.savefig.assert_called_once()
        stream = mock_figure.savefig.call_args[0][0]
        assert isinstance(stream, io.BytesIO)
        assert mock_figure.savefig.call_args[1]["dpi"] == 300

        # Verify that add_image was called with the same stream
        self.parent.add_image.assert_called_once()
        assert self.parent.add_image.call_args[1]["image_path"] is stream

        # Verify correct result and cell content
        assert result == "image_shape"
        assert self.grid.cells[1][1].content == "image_shape"

    def test_add_table(self):
        """Test the add_table method."""
        data = [["A", "B"], [1, 2]]

        # Mock the Table class and its add method
        mock_table = MagicMock()
        mock_table.add.return_value = "table_shape"

        with unittest.mock.patch("easypptx.table.Table", return_value=mock_table):
            result = self.grid.add_table(row=0, col=0, data=data, has_header=True)

            # Verify that the Table was created
            mock_table.add.assert_called_once()

            # Check that the correct data was passed
            call_kwargs = mock_table.add.call_args[1]
            assert call_kwargs["data"] == data
            assert call_kwargs["first_row_header"] is True

            # Verify that the result is correct
            assert result == "table_shape"
            assert self.grid.cells[0][0].content == "table_shape"
