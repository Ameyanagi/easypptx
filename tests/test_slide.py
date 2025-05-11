"""Tests for the Slide class."""

from unittest.mock import MagicMock

from easypptx import Presentation


def test_slide_add_text():
    """Test adding text to a slide."""
    pres = Presentation()
    slide = pres.add_slide()

    # Add text
    text_box = slide.add_text("Test text")
    assert text_box is not None

    # Check that the text frame contains the expected text
    assert text_box.text_frame.text == "Test text"


def test_slide_add_text_with_formatting():
    """Test adding formatted text to a slide."""
    pres = Presentation()
    slide = pres.add_slide()

    # Add formatted text
    text_box = slide.add_text("Test text", font_size=24, font_bold=True, font_italic=True)

    # Check formatting
    p = text_box.text_frame.paragraphs[0]
    assert p.font.size.pt == 24
    assert p.font.bold is True
    assert p.font.italic is True


def test_slide_clear():
    """Test clearing a slide."""
    # This test requires mocking since we can't directly test _spTree removal
    slide = MagicMock()
    mock_shape = MagicMock()
    mock_element = MagicMock()
    mock_shape._element = mock_element

    # Setup the slide's shapes collection with _spTree
    slide.pptx_slide.shapes = MagicMock()
    slide.pptx_slide.shapes._spTree = MagicMock()
    slide.pptx_slide.shapes.__iter__.return_value = [mock_shape]

    # Call clear method through the Slide class
    from easypptx.slide import Slide

    slide_obj = Slide(slide.pptx_slide)
    slide_obj.clear()

    # Verify that remove was called on _spTree for the shape element
    slide.pptx_slide.shapes._spTree.remove.assert_called_with(mock_element)


def test_slide_title_property():
    """Test the title property."""
    # Mock the slide with a title
    slide = MagicMock()
    mock_title = MagicMock()
    mock_title.text = "Test Title"
    slide.pptx_slide.shapes.title = mock_title

    # Call the title getter
    from easypptx.slide import Slide

    result = Slide.title.__get__(Slide(slide.pptx_slide))

    # Verify the result
    assert result == "Test Title"

    # Call the title setter
    slide_obj = Slide(slide.pptx_slide)
    slide_obj.title = "New Title"

    # Verify the title was updated
    assert mock_title.text == "New Title"


def test_slide_title_property_no_title_shape():
    """Test the title property when there's no title shape."""
    # Mock the slide without a title
    slide = MagicMock()
    slide.pptx_slide.shapes.title = None

    # Call the title getter
    from easypptx.slide import Slide

    result = Slide.title.__get__(Slide(slide.pptx_slide))

    # Verify the result is None
    assert result is None


def test_slide_convert_position():
    """Test the percentage-based and absolute position conversion."""
    # Setup
    from easypptx.slide import Slide

    slide = MagicMock()
    slide_obj = Slide(slide.pptx_slide)

    # Mock slide dimension (10 inches = 9144000 EMUs)
    slide_dimension = 9144000

    # Test percentage conversion
    percent_50 = slide_obj._convert_position("50%", slide_dimension)
    assert percent_50 == 5.0  # 50% of 10 inches = 5 inches

    percent_25 = slide_obj._convert_position("25%", slide_dimension)
    assert percent_25 == 2.5  # 25% of 10 inches = 2.5 inches

    # Test absolute positioning
    absolute_3 = slide_obj._convert_position(3.0, slide_dimension)
    assert absolute_3 == 3.0  # Should remain as 3.0 inches
