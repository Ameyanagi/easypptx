"""Image handling module for EasyPPTX."""

from __future__ import annotations

from pathlib import Path
from typing import IO, TYPE_CHECKING

from PIL import Image as PILImage
from pptx.shapes.autoshape import Shape as PPTXShape
from pptx.util import Inches

from easypptx.positioning import PositionType, to_inches

if TYPE_CHECKING:
    from easypptx.slide import Slide


class Image:
    """Class for handling image operations in PowerPoint slides.

    This class provides methods for adding and manipulating images on slides.

    Examples:
        ```python
        # Create an image object
        image = Image(slide)

        # Add an image
        image.add("example.png", x=2, y=2)

        # Add an image with specific dimensions
        image.add("example.jpg", x=1, y=1, width=4, height=3)
        ```
    """

    def __init__(self, slide_obj: Slide) -> None:
        """Initialize an Image object.

        Args:
            slide_obj: The Slide object to add images to
        """
        self.slide = slide_obj

    def add(
        self,
        image_path: str | Path | IO[bytes],
        x: PositionType = 5,
        y: PositionType = 5,
        width: PositionType | None = None,
        height: PositionType | None = None,
        maintain_aspect_ratio: bool = True,
    ) -> PPTXShape:
        """Add an image to the slide.

        Args:
            image_path: Path to the image file, or a binary file-like object
            x: X position in inches or percentage (default: 1.0)
            y: Y position in inches or percentage (default: 1.0)
            width: Width in inches or percentage (default: None, uses image's width)
            height: Height in inches or percentage (default: None, uses image's height)
            maintain_aspect_ratio: Whether to maintain aspect ratio when only one
                                  dimension is specified (default: True)

        Returns:
            The created picture shape

        Raises:
            FileNotFoundError: If the image file doesn't exist
        """
        source: str | IO[bytes]
        if isinstance(image_path, str | Path):
            image_path_obj = Path(image_path)
            if not image_path_obj.exists():
                raise FileNotFoundError(f"Image file not found: {image_path}")
            source = str(image_path_obj)
        else:
            source = image_path

        # Preserve the image's aspect ratio via physical dimensions
        if maintain_aspect_ratio and (width is not None or height is not None):
            with PILImage.open(source) as img:
                img_width, img_height = img.size
            aspect_ratio = img_width / img_height
            if not isinstance(source, str):
                source.seek(0)

            slide_w = self.slide._get_slide_width()
            slide_h = self.slide._get_slide_height()

            if width is not None and height is not None:
                # Both given: fit within the box, centered
                box_x = to_inches(x, slide_w)
                box_y = to_inches(y, slide_h)
                box_w = to_inches(width, slide_w)
                box_h = to_inches(height, slide_h)
                fit_w = min(box_w, box_h * aspect_ratio)
                fit_h = fit_w / aspect_ratio
                x = Inches(box_x + (box_w - fit_w) / 2)
                y = Inches(box_y + (box_h - fit_h) / 2)
                width = Inches(fit_w)
                height = Inches(fit_h)
            elif width is not None:
                # Height follows from the physical width
                height = Inches(to_inches(width, slide_w) / aspect_ratio)
            elif height is not None:
                # Width follows from the physical height
                width = Inches(to_inches(height, slide_h) * aspect_ratio)

        return self.slide.add_image(source, x, y, width, height)

    @staticmethod
    def get_image_dimensions(image_path: str | Path) -> tuple[int, int]:
        """Get the dimensions of an image file.

        Args:
            image_path: Path to the image file

        Returns:
            A tuple containing (width, height) in pixels

        Raises:
            FileNotFoundError: If the image file doesn't exist
        """
        image_path_obj = Path(image_path)
        if not image_path_obj.exists():
            raise FileNotFoundError(f"Image file not found: {image_path}")

        with PILImage.open(image_path_obj) as img:
            return img.size
