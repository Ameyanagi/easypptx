"""Tests for grid slide related methods."""

from unittest.mock import MagicMock, patch

from easypptx import Presentation


class TestGridSlideMethods:
    """Test the grid slide related methods in the Presentation class."""

    def test_add_grid_slide(self):
        """Test the add_grid_slide method."""
        # Create a presentation
        pres = Presentation()

        # Mock slide.add_text to avoid actual PowerPoint operations
        pres.add_slide = MagicMock()
        slide_mock = MagicMock()
        slide_mock.add_text = MagicMock()
        pres.add_slide.return_value = slide_mock

        # Create a patch for the Grid class
        with patch("easypptx.presentation.Grid") as mock_grid:
            # Setup the mock
            grid_instance = MagicMock()
            mock_grid.return_value = grid_instance

            # Call the method
            result = pres.add_grid_slide(
                rows=3,
                cols=2,
                title="Test Title",
                subtitle="Test Subtitle",
                padding=5.0,
            )

            # Verify the result
            assert isinstance(result, tuple)
            assert len(result) == 2
            assert result[0] == slide_mock  # The slide
            assert result[1] == grid_instance  # The grid

            # Verify add_slide was called
            pres.add_slide.assert_called_once()

            # Verify add_text was called twice (for title and subtitle)
            assert slide_mock.add_text.call_count == 2

            # Verify Grid was instantiated with correct parameters
            mock_grid.assert_called_once()
            grid_args = mock_grid.call_args[1]
            assert grid_args["parent"] == slide_mock
            assert grid_args["rows"] == 3
            assert grid_args["cols"] == 2
            assert grid_args["padding"] == 5.0

    def test_add_grid_slide_without_titles(self):
        """Test the add_grid_slide method without title or subtitle."""
        # Create a presentation
        pres = Presentation()

        # Mock slide.add_text to avoid actual PowerPoint operations
        pres.add_slide = MagicMock()
        slide_mock = MagicMock()
        slide_mock.add_text = MagicMock()
        pres.add_slide.return_value = slide_mock

        # Create a patch for the Grid class
        with patch("easypptx.presentation.Grid") as mock_grid:
            # Setup the mock
            grid_instance = MagicMock()
            mock_grid.return_value = grid_instance

            # Call the method without title or subtitle
            result = pres.add_grid_slide(
                rows=2,
                cols=2,
                x="10%",
                y="10%",
                width="80%",
                height="80%",
                padding=5.0,
            )

            # Verify the result
            assert isinstance(result, tuple)
            assert len(result) == 2
            assert result[0] == slide_mock  # The slide
            assert result[1] == grid_instance  # The grid

            # Verify add_slide was called
            pres.add_slide.assert_called_once()

            # Verify add_text was not called (no title or subtitle)
            slide_mock.add_text.assert_not_called()

            # Verify Grid was instantiated with correct parameters
            mock_grid.assert_called_once()
            grid_args = mock_grid.call_args[1]
            assert grid_args["parent"] == slide_mock
            assert grid_args["rows"] == 2
            assert grid_args["cols"] == 2
            assert grid_args["x"] == "10%"
            assert grid_args["y"] == "10%"
            assert grid_args["width"] == "80%"
            assert grid_args["height"] == "80%"
            assert grid_args["padding"] == 5.0
