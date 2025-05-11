"""Image handling module for EasyPPTX."""

from pathlib import Path
from typing import TYPE_CHECKING

from PIL import Image as PILImage
from pptx.shapes.autoshape import Shape as PPTXShape

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

    def __init__(self, slide_obj: "Slide") -> None:
        """Initialize an Image object.

        Args:
            slide_obj: The Slide object to add images to
        """
        self.slide = slide_obj

    def add(
        self,
        image_path: str | Path,
        x: float | str = 1.0,
        y: float | str = 1.0,
        width: float | str | None = None,
        height: float | str | None = None,
        maintain_aspect_ratio: bool = True,
    ) -> PPTXShape:
        """Add an image to the slide.

        Args:
            image_path: Path to the image file
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
        image_path_obj = Path(image_path)
        if not image_path_obj.exists():
            raise FileNotFoundError(f"Image file not found: {image_path}")

        # Get image dimensions if needed
        if maintain_aspect_ratio and (width is not None or height is not None):
            with PILImage.open(image_path_obj) as img:
                img_width, img_height = img.size
                aspect_ratio = img_width / img_height

                if width is not None and height is None:
                    # Calculate height based on width
                    # If width is percentage-based, assume proportional height
                    if isinstance(width, str) and width.endswith("%"):
                        height = f"{float(width.strip('%')) / aspect_ratio}%"
                    else:
                        height = float(width) / aspect_ratio
                elif height is not None and width is None:
                    # Calculate width based on height
                    # If height is percentage-based, assume proportional width
                    if isinstance(height, str) and height.endswith("%"):
                        width = f"{float(height.strip('%')) * aspect_ratio}%"
                    else:
                        width = float(height) * aspect_ratio

        # Pass positional arguments for compatibility with tests
        return self.slide.add_image(str(image_path_obj), x, y, width, height)

    @staticmethod
    def get_image_dimensions(image_path: str | Path) -> tuple:
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
