"""Tests for the Image class."""

import os
import tempfile
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from PIL import Image as PILImage

from easypptx import Image


def create_test_image(width=100, height=50):
    """Create a temporary test image file and return its path."""
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        temp_path = Path(tmp.name)

    # Create a simple test image
    img = PILImage.new("RGB", (width, height), color="red")
    img.save(temp_path)

    return temp_path


class TestImage:
    """Test cases for the Image class."""

    def setup_method(self):
        """Set up test environment before each test method."""
        self.slide = MagicMock()
        self.image = Image(self.slide)
        self.test_image_path = create_test_image()

    def teardown_method(self):
        """Clean up test environment after each test method."""
        if os.path.exists(self.test_image_path):
            os.unlink(self.test_image_path)

    def test_image_init(self):
        """Test Image initialization."""
        assert self.image.slide == self.slide

    def test_add_image_basic(self):
        """Test adding an image with default parameters."""
        # Call add method with default parameters
        result = self.image.add(self.test_image_path)

        # Verify slide.add_image was called with correct parameters
        self.slide.add_image.assert_called_once()
        call_args = self.slide.add_image.call_args[0]

        # First argument should be the image path as string
        assert call_args[0] == str(self.test_image_path)

        # Verify the result is the return value from slide.add_image
        assert result == self.slide.add_image.return_value

    def test_add_image_with_position(self):
        """Test adding an image with custom position."""
        # Call add method with custom position
        result = self.image.add(self.test_image_path, x=2.5, y=3.5)

        # Verify slide.add_image was called with correct parameters
        self.slide.add_image.assert_called_once()
        call_args = self.slide.add_image.call_args[0]

        assert call_args[0] == str(self.test_image_path)
        assert call_args[1] == 2.5  # x position
        assert call_args[2] == 3.5  # y position

        # Verify the result is the return value from slide.add_image
        assert result == self.slide.add_image.return_value

    def test_add_image_with_dimensions(self):
        """Test adding an image with custom dimensions."""
        # Call add method with custom dimensions
        result = self.image.add(self.test_image_path, width=4.0, height=3.0, maintain_aspect_ratio=False)

        # Verify slide.add_image was called with correct parameters
        self.slide.add_image.assert_called_once()
        call_args = self.slide.add_image.call_args[0]

        assert call_args[0] == str(self.test_image_path)
        assert call_args[3] == 4.0  # width
        assert call_args[4] == 3.0  # height

        # Verify the result is the return value from slide.add_image
        assert result == self.slide.add_image.return_value

    def test_add_image_maintain_aspect_ratio_width(self):
        """Test adding an image with width while maintaining aspect ratio."""
        # Call add method with width only and maintain_aspect_ratio=True
        result = self.image.add(self.test_image_path, width=4.0, maintain_aspect_ratio=True)

        # Verify slide.add_image was called with correct parameters
        self.slide.add_image.assert_called_once()
        call_args = self.slide.add_image.call_args[0]

        assert call_args[0] == str(self.test_image_path)
        assert call_args[3] == 4.0  # width
        # Height should be calculated based on aspect ratio (width/2 since our test image is 100x50)
        assert call_args[4] == 2.0  # height

        # Verify the result is the return value from slide.add_image
        assert result == self.slide.add_image.return_value

    def test_add_image_maintain_aspect_ratio_height(self):
        """Test adding an image with height while maintaining aspect ratio."""
        # Call add method with height only and maintain_aspect_ratio=True
        result = self.image.add(self.test_image_path, height=2.0, maintain_aspect_ratio=True)

        # Verify slide.add_image was called with correct parameters
        self.slide.add_image.assert_called_once()
        call_args = self.slide.add_image.call_args[0]

        assert call_args[0] == str(self.test_image_path)
        # Width should be calculated based on aspect ratio (height*2 since our test image is 100x50)
        assert call_args[3] == 4.0  # width
        assert call_args[4] == 2.0  # height

        # Verify the result is the return value from slide.add_image
        assert result == self.slide.add_image.return_value

    def test_add_image_nonexistent_file(self):
        """Test adding a nonexistent image raises FileNotFoundError."""
        with pytest.raises(FileNotFoundError):
            self.image.add("nonexistent_image.png")

    def test_get_image_dimensions(self):
        """Test getting image dimensions."""
        # Get the dimensions of the test image
        width, height = Image.get_image_dimensions(self.test_image_path)

        # Verify dimensions match the test image
        assert width == 100
        assert height == 50

    def test_get_image_dimensions_nonexistent_file(self):
        """Test getting dimensions of a nonexistent image raises FileNotFoundError."""
        with pytest.raises(FileNotFoundError):
            Image.get_image_dimensions("nonexistent_image.png")
