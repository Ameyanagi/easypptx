"""Tests for title padding functionality in all slide creation methods."""

from unittest.mock import MagicMock, patch

from easypptx import Presentation


class TestTitlePadding:
    """Test the title padding feature in various slide creation methods."""

    def test_add_slide_with_title_padding(self):
        """Test add_slide with title padding."""
        # Create a presentation with mocked internals
        pres = Presentation()

        # Create a mock Slide instance with a tracked add_text method
        slide_mock = MagicMock()

        # Patch the add_slide method to return our mock slide
        with (
            patch.object(Presentation, "add_slide_from_template"),
            patch.object(pres.pptx_presentation.slides, "add_slide", return_value=MagicMock()),
            patch("easypptx.presentation.Slide", return_value=slide_mock),
        ):
            # Call the add_slide method with title_padding
            pres.add_slide(title="Test Title", title_padding="10%")

            # Verify add_text was called on our mock slide
            slide_mock.add_text.assert_called_once()
            args = slide_mock.add_text.call_args[1]
            assert args["text"] == "Test Title"
            assert args["x"] == "10%"
            assert args["y"] == "10%"

            # Reset the mock for the next test
            slide_mock.add_text.reset_mock()

            # Test with separate x and y padding
            pres.add_slide(title="Test Title", title_x_padding="15%", title_y_padding="5%")

            # Verify add_text was called with correct arguments
            slide_mock.add_text.assert_called_once()
            args = slide_mock.add_text.call_args[1]
            assert args["text"] == "Test Title"
            assert args["x"] == "15%"
            assert args["y"] == "5%"

            # Reset the mock for the next test
            slide_mock.add_text.reset_mock()

            # Test that title_padding overrides individual paddings
            pres.add_slide(title="Test Title", title_padding="20%", title_x_padding="15%", title_y_padding="5%")

            # Verify add_text was called with correct arguments
            slide_mock.add_text.assert_called_once()
            args = slide_mock.add_text.call_args[1]
            assert args["text"] == "Test Title"
            assert args["x"] == "20%"
            assert args["y"] == "20%"

    def test_add_grid_slide_with_title_padding(self):
        """Test add_grid_slide with title padding."""
        # Create a presentation
        pres = Presentation()

        # Mock slide.add_text to avoid actual PowerPoint operations
        pres.add_slide = MagicMock()
        slide_mock = MagicMock()
        slide_mock.add_text = MagicMock()
        pres.add_slide.return_value = slide_mock

        # Create a patch for Grid
        with patch("easypptx.presentation.Grid") as mock_grid:
            # Setup the mock
            grid_instance = MagicMock()
            mock_grid.return_value = grid_instance

            # Call the method with title padding
            pres.add_grid_slide(rows=2, cols=2, title="Test Title", title_padding="10%")

            # Verify add_text was called with the right parameters
            slide_mock.add_text.assert_called_once()
            args = slide_mock.add_text.call_args[1]
            assert args["text"] == "Test Title"
            assert args["x"] == "10%"
            assert args["y"] == "10%"

            # Test with content padding
            slide_mock.add_text.reset_mock()
            mock_grid.reset_mock()

            # Call the method with content padding
            pres.add_grid_slide(rows=2, cols=2, title="Test Title", content_padding="15%")

            # Verify Grid was called with correct parameters
            mock_grid.assert_called_once()
            grid_args = mock_grid.call_args[1]
            assert grid_args["x"] == "15%"

    def test_add_autogrid_slide_with_title_padding(self):
        """Test add_autogrid_slide with title padding."""
        # Create a presentation
        pres = Presentation()

        # Mock add_slide to avoid actual PowerPoint operations
        pres.add_slide = MagicMock()
        slide_mock = MagicMock()
        slide_mock.add_text = MagicMock()
        pres.add_slide.return_value = slide_mock

        # Mock add_autogrid to avoid actual grid operations
        pres.add_autogrid = MagicMock()
        grid_mock = MagicMock()
        pres.add_autogrid.return_value = grid_mock

        # Call the method with title padding
        pres.add_autogrid_slide(rows=2, cols=2, title="Test Title", title_padding="10%", content_padding="15%")

        # Verify add_text was called with the right parameters
        slide_mock.add_text.assert_called_once()
        args = slide_mock.add_text.call_args[1]
        assert args["text"] == "Test Title"
        assert args["x"] == "10%"
        assert args["y"] == "10%"

        # Verify add_autogrid was called with correct parameters
        pres.add_autogrid.assert_called_once()
        grid_args = pres.add_autogrid.call_args[1]
        assert grid_args["x"] == "15%"

    def test_add_pyplot_slide_with_title_padding(self):
        """Test add_pyplot_slide with title padding."""
        # Create a presentation
        pres = Presentation()

        # Mock add_slide to avoid actual PowerPoint operations
        pres.add_slide = MagicMock()
        slide_mock = MagicMock()
        slide_mock.add_text = MagicMock()
        pres.add_slide.return_value = slide_mock

        # Mock the Pyplot.add method
        with patch("easypptx.presentation.Pyplot") as mock_pyplot:
            # Set up the mock
            pyplot_instance = MagicMock()
            mock_pyplot.add.return_value = pyplot_instance

            # Create a mock figure
            figure_mock = MagicMock()

            # Call the method with padding
            pres.add_pyplot_slide(
                figure=figure_mock,
                title="Test Title",
                title_padding="10%",
                content_x_padding="15%",
                content_y_padding="5%",
                label="Test Label",
                label_padding="8%",
            )

            # Verify title text was added with correct padding
            assert slide_mock.add_text.call_count >= 1
            title_args = slide_mock.add_text.call_args_list[0][1]
            assert title_args["text"] == "Test Title"
            assert title_args["x"] == "10%"
            assert title_args["y"] == "10%"

            # Verify Pyplot.add was called with correct parameters
            mock_pyplot.add.assert_called_once()
            pyplot_args = mock_pyplot.add.call_args[1]
            assert pyplot_args["position"]["x"] == "15%"

    def test_add_image_gen_slide_with_title_padding(self):
        """Test add_image_gen_slide with title padding."""
        # Create a presentation
        pres = Presentation()

        # Mock add_slide to avoid actual PowerPoint operations
        pres.add_slide = MagicMock()
        slide_mock = MagicMock()
        slide_mock.add_text = MagicMock()
        pres.add_slide.return_value = slide_mock

        # Mock the Image class
        with patch("easypptx.presentation.Image") as mock_image:
            # Set up the mock
            image_instance = MagicMock()
            mock_image.return_value = image_instance
            image_shape = MagicMock()
            image_instance.add.return_value = image_shape

            # Call the method with padding
            pres.add_image_gen_slide(
                image_path="test.png",
                title="Test Title",
                title_padding="10%",
                content_x_padding="15%",
                content_y_padding="5%",
                label="Test Label",
                label_padding="8%",
            )

            # Verify title text was added with correct padding
            assert slide_mock.add_text.call_count >= 1
            title_args = slide_mock.add_text.call_args_list[0][1]
            assert title_args["text"] == "Test Title"
            assert title_args["x"] == "10%"
            assert title_args["y"] == "10%"

            # Verify Image.add was called with correct parameters
            image_instance.add.assert_called_once()
            image_args = image_instance.add.call_args[1]
            assert image_args["x"] == "15%"
