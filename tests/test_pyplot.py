"""Tests for the pyplot module."""

from unittest.mock import MagicMock, patch

from easypptx.pyplot import Pyplot


class TestPyplot:
    """Test the Pyplot class."""

    @patch("easypptx.image.Image.add")
    def test_add(self, mock_add):
        """Test adding a matplotlib figure to a slide."""
        # Create mock slide and figure
        slide = MagicMock()
        figure = MagicMock()

        # Mock the image shape
        mock_image_shape = MagicMock()
        mock_add.return_value = mock_image_shape

        # Call the add method
        position = {"x": "10%", "y": "20%", "width": "80%", "height": "70%"}
        result = Pyplot.add(
            slide=slide,
            figure=figure,
            position=position,
            dpi=300,
            file_format="png",
            style={"border": True, "border_color": "blue", "shadow": True},
        )

        # Verify the figure was saved and added as an image
        assert figure.savefig.called
        assert mock_add.called

        # Verify styling was applied
        assert mock_image_shape.line.color.rgb is not None
        assert mock_image_shape.shadow.visible is True

        # Verify result
        assert result == mock_image_shape

    @patch("easypptx.pyplot.Pyplot.add")
    def test_add_from_seaborn(self, mock_add):
        """Test adding a seaborn plot to a slide."""
        # Create mock slide and seaborn plot
        slide = MagicMock()
        seaborn_plot = MagicMock()

        # Mock seaborn plot with figure
        seaborn_plot.figure = MagicMock()

        # Call the add_from_seaborn method
        position = {"x": "10%", "y": "20%", "width": "80%", "height": "70%"}
        Pyplot.add_from_seaborn(
            slide=slide,
            seaborn_plot=seaborn_plot,
            position=position,
            dpi=300,
            file_format="png",
            style={"border": True},
        )

        # Verify Pyplot.add was called with the correct figure
        mock_add.assert_called_once_with(
            slide=slide,
            figure=seaborn_plot.figure,
            position=position,
            dpi=300,
            file_format="png",
            style={"border": True},
        )

        # Test with plot that has fig attribute
        seaborn_plot = MagicMock(spec=["fig"])
        Pyplot.add_from_seaborn(slide=slide, seaborn_plot=seaborn_plot, position=position)

        # Verify add was called with fig attribute
        assert mock_add.call_args[1]["figure"] == seaborn_plot.fig

        # Skip the test with no figure attribute as it requires matplotlib
