"""Unit tests for responsive positioning feature."""

import unittest

from easypptx import Presentation


class TestResponsivePositioning(unittest.TestCase):
    """Tests for responsive positioning in different aspect ratios."""

    def test_percentage_positioning(self):
        """Test that percentage-based positioning works correctly."""
        # Create a 16:9 presentation
        pres = Presentation(aspect_ratio="16:9")
        slide = pres.add_slide()

        # Add text with percentage positioning
        text_shape = slide.add_text(text="Test Text", x="50%", y="10%", width="50%", height="10%", align="center")

        # Verify text was added
        self.assertIsNotNone(text_shape)

        # Add shape with percentage positioning
        shape = slide.add_shape(x="50%", y="50%", width="50%", height="10%", fill_color="blue")

        # Verify shape was added
        self.assertIsNotNone(shape)

        # Test successful if no errors were raised
        self.assertTrue(True)

    def test_percentage_conversion(self):
        """Test that percentage values are correctly converted to absolute positions."""
        # Create a presentation
        pres = Presentation(aspect_ratio="16:9")
        slide = pres.add_slide()

        # Get slide dimensions
        slide_width = slide._get_slide_width()

        # Convert a percentage position (50%)
        position_inches = slide._convert_position("50%", slide_width)

        # Expected value (50% of slide width)
        expected_inches = (slide_width / 914400) * 0.5

        # Verify conversion is correct
        self.assertAlmostEqual(position_inches, expected_inches, places=5)

        # Test another percentage (25%)
        position_inches = slide._convert_position("25%", slide_width)
        expected_inches = (slide_width / 914400) * 0.25
        self.assertAlmostEqual(position_inches, expected_inches, places=5)


if __name__ == "__main__":
    unittest.main()
