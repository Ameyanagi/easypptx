"""Unit tests for responsive positioning feature."""

import unittest
from easypptx import Presentation


class TestResponsivePositioning(unittest.TestCase):
    """Tests for responsive positioning in different aspect ratios."""

    def test_h_align_parameter_presence(self):
        """Test that the h_align parameter is correctly passed to the underlying methods."""
        # Create a 16:9 presentation
        pres = Presentation(aspect_ratio="16:9")
        slide = pres.add_slide()
        
        # Add text with h_align parameter
        text_box = slide.add_text(
            text="Test Text",
            x="50%",
            y="10%",
            width="50%",
            height="10%",
            align="center",
            h_align="center"
        )
        
        # Add image with h_align parameter
        # This would normally error if h_align wasn't properly implemented
        image_shape = slide.add_image(
            image_path="output/images/company_logo.png",
            x="50%", 
            y="30%",
            width="50%",
            h_align="center"
        )
        
        # Add shape with h_align parameter
        shape = slide.add_shape(
            x="50%",
            y="50%",
            width="50%",
            height="10%",
            fill_color="blue",
            h_align="center"
        )
        
        # Test successful if no errors were raised
        self.assertTrue(True)
    
    def test_aspect_ratio_adjustment(self):
        """Test that different aspect ratios produce different adjustments."""
        # Create presentations with different aspect ratios
        pres_standard = Presentation(aspect_ratio="4:3") 
        slide_standard = pres_standard.add_slide()
        
        # Get internal slide dimensions
        standard_width = slide_standard._get_slide_width()
        
        # Standard positioning (no h_align)
        standard_pos = slide_standard._convert_position("50%", standard_width)
        
        # Responsive positioning (with h_align="center")
        responsive_pos = slide_standard._convert_position("50%", standard_width, h_align="center")
        
        # In non-16:9 ratio, there should be a noticeable difference
        # Centered elements should be adjusted to account for aspect ratio difference
        self.assertNotEqual(round(standard_pos, 3), round(responsive_pos, 3))
        
        # Verify that 4:3 adjustment is different from 16:9 adjustment
        # The exact values aren't as important as the fact that they differ
        standard_ratio = 16/9
        actual_ratio = slide_standard._get_slide_width() / slide_standard._get_slide_height()
        self.assertNotAlmostEqual(standard_ratio, actual_ratio, delta=0.01)


if __name__ == "__main__":
    unittest.main()