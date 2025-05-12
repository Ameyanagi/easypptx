"""Tests for reference templates in presentation."""

import os
import sys
import unittest
from unittest.mock import MagicMock, patch

# Add the src directory to the path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


class TestReferenceTemplates(unittest.TestCase):
    """Test the usage of reference templates based on aspect ratio."""

    @patch("src.easypptx.presentation.PPTXPresentation")
    def test_uses_reference_template_for_16x9(self, mock_pptx):
        """Test that a 16:9 presentation uses the reference_16x9.pptx template by default."""
        # Setup mock
        mock_pptx.return_value.slide_layouts = [MagicMock() for _ in range(10)]

        # Import here to ensure our mock is applied
        from src.easypptx.presentation import Presentation

        # Create a presentation with 16:9 aspect ratio
        _ = Presentation(aspect_ratio="16:9")

        # Check if the mock was called with a path containing reference_16x9.pptx
        args, _ = mock_pptx.call_args
        self.assertTrue(args)
        template_path = args[0]
        self.assertIn("reference_16x9.pptx", template_path)

    @patch("src.easypptx.presentation.PPTXPresentation")
    def test_uses_reference_template_for_4x3(self, mock_pptx):
        """Test that a 4:3 presentation uses the reference_4x3.pptx template by default."""
        # Setup mock
        mock_pptx.return_value.slide_layouts = [MagicMock() for _ in range(10)]

        # Import here to ensure our mock is applied
        from src.easypptx.presentation import Presentation

        # Create a presentation with 4:3 aspect ratio
        _ = Presentation(aspect_ratio="4:3")

        # Check if the mock was called with a path containing reference_4x3.pptx
        args, _ = mock_pptx.call_args
        self.assertTrue(args)
        template_path = args[0]
        self.assertIn("reference_4x3.pptx", template_path)

    @patch("src.easypptx.presentation.Path.exists")
    @patch("src.easypptx.presentation.PPTXPresentation")
    def test_custom_template_overrides_reference(self, mock_pptx, mock_exists):
        """Test that providing a custom template overrides the reference template."""
        # Setup mocks
        mock_pptx.return_value.slide_layouts = [MagicMock() for _ in range(10)]
        mock_exists.return_value = True

        # Import here to ensure our mock is applied
        from src.easypptx.presentation import Presentation

        # Create a presentation with a custom template
        _ = Presentation(aspect_ratio="16:9", template_path="custom_template.pptx")

        # Check if the mock was called with the custom template path
        args, _ = mock_pptx.call_args
        self.assertTrue(args)
        template_path = args[0]
        self.assertEqual(template_path, "custom_template.pptx")

    @patch("src.easypptx.presentation.Path.exists")
    @patch("src.easypptx.presentation.PPTXPresentation")
    def test_custom_dimensions_skip_reference_template(self, mock_pptx, mock_exists):
        """Test that custom dimensions don't use reference templates."""
        # Setup mocks
        mock_pptx.return_value.slide_layouts = [MagicMock() for _ in range(10)]
        mock_exists.return_value = True

        # Import here to ensure our mock is applied
        from src.easypptx.presentation import Presentation

        # Create a presentation with custom dimensions
        _ = Presentation(width_inches=10, height_inches=7.5)

        # The first call will be with no arguments (blank presentation)
        args, _ = mock_pptx.call_args
        self.assertEqual(args, ())

    @patch("src.easypptx.presentation.Path.exists")
    @patch("src.easypptx.presentation.PPTXPresentation")
    def test_other_aspect_ratios_skip_reference_template(self, mock_pptx, mock_exists):
        """Test that other aspect ratios don't use reference templates."""
        # Setup mocks
        mock_pptx.return_value.slide_layouts = [MagicMock() for _ in range(10)]
        mock_exists.return_value = True

        # Import here to ensure our mock is applied
        from src.easypptx.presentation import Presentation

        # Create a presentation with 16:10 aspect ratio
        _ = Presentation(aspect_ratio="16:10")

        # The presentation should be created without any template path
        args, _ = mock_pptx.call_args
        self.assertEqual(args, ())
