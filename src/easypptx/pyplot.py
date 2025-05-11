"""Pyplot integration module for EasyPPTX."""

import os
import tempfile
from typing import Any

from easypptx.image import Image


class Pyplot:
    """Class for adding matplotlib/seaborn plots to slides."""

    @staticmethod
    def add(
        slide,
        figure,
        position: dict[str, float | str],
        dpi: int = 300,
        file_format: str = "png",
        style: dict[str, Any] | None = None,
    ):
        """Add a matplotlib or seaborn figure to a slide.

        Args:
            slide: Slide object to add the plot to
            figure: Matplotlib figure object (plt.figure(), sns.FacetGrid, etc.)
            position: Position dictionary with x, y, width, height as percentages
            dpi: Resolution for the figure (default: 300)
            file_format: Image format ("png" or "jpg") (default: "png")
            style: Dictionary of style options for the image (default: None)

        Returns:
            Image shape object

        Example:
            ```python
            import matplotlib.pyplot as plt
            from easypptx import Presentation, Plot

            # Create a matplotlib figure
            plt.figure(figsize=(10, 6))
            plt.plot([1, 2, 3, 4], [1, 4, 9, 16])
            plt.title('Sample Plot')

            # Add it to a slide
            pres = Presentation()
            slide = pres.add_slide()

            Plot.add(
                slide=slide,
                figure=plt.gcf(),
                position={"x": "10%", "y": "20%", "width": "80%", "height": "70%"}
            )
            ```
        """
        # Apply default styling if not provided
        if style is None:
            style = {"maintain_aspect_ratio": True, "center": True, "border": False}

        # Create a temporary file to save the figure
        with tempfile.NamedTemporaryFile(suffix=f".{file_format}", delete=False) as tmp:
            temp_path = tmp.name

        try:
            # Save the figure to the temporary file
            figure.savefig(temp_path, dpi=dpi, format=file_format, bbox_inches="tight")

            # Add the image with styling
            img = Image(slide)
            x = position.get("x", "10%")
            y = position.get("y", "20%")
            width = position.get("width", "80%")
            height = position.get("height", "70%")

            image_shape = img.add(
                image_path=temp_path,
                x=x,
                y=y,
                width=width,
                height=height,
                maintain_aspect_ratio=style.get("maintain_aspect_ratio", True),
            )

            # Apply any additional styling options from the slide's presentation
            if hasattr(slide, "pptx_slide") and hasattr(slide.pptx_slide, "shapes"):
                # Access the presentation to get color definitions
                from easypptx.presentation import Presentation

                # Apply border if specified
                if style.get("border", False):
                    image_shape.line.color.rgb = Presentation.COLORS.get(
                        style.get("border_color", "black"), Presentation.COLORS["black"]
                    )
                    image_shape.line.width = style.get("border_width", 1)

                # Apply shadow if specified
                if style.get("shadow", False):
                    image_shape.shadow.inherit = False
                    image_shape.shadow.visible = True
                    image_shape.shadow.blur_radius = 5
                    image_shape.shadow.distance = 3
                    image_shape.shadow.angle = 45

            return image_shape

        finally:
            # Clean up the temporary file
            if os.path.exists(temp_path):
                os.unlink(temp_path)

    @staticmethod
    def add_from_seaborn(
        slide,
        seaborn_plot,
        position: dict[str, float | str],
        dpi: int = 300,
        file_format: str = "png",
        style: dict[str, Any] | None = None,
    ):
        """Add a seaborn plot to a slide.

        Args:
            slide: Slide object to add the plot to
            seaborn_plot: Seaborn plot object (sns.barplot, sns.heatmap, etc.)
            position: Position dictionary with x, y, width, height as percentages
            dpi: Resolution for the figure (default: 300)
            file_format: Image format ("png" or "jpg") (default: "png")
            style: Dictionary of style options for the image (default: None)

        Returns:
            Image shape object

        Example:
            ```python
            import seaborn as sns
            from easypptx import Presentation, Plot

            # Create a seaborn plot
            tips = sns.load_dataset("tips")
            sns_plot = sns.barplot(x="day", y="total_bill", data=tips)

            # Add it to a slide
            pres = Presentation()
            slide = pres.add_slide()

            Plot.add_from_seaborn(
                slide=slide,
                seaborn_plot=sns_plot,
                position={"x": "10%", "y": "20%", "width": "80%", "height": "70%"}
            )
            ```
        """
        # Get the figure from the seaborn plot
        if hasattr(seaborn_plot, "figure"):
            figure = seaborn_plot.figure
        elif hasattr(seaborn_plot, "fig"):
            figure = seaborn_plot.fig
        else:
            try:
                import matplotlib.pyplot as plt

                figure = plt.gcf()
            except ImportError as err:
                raise ImportError(
                    "Matplotlib is required for plots with no figure attribute. "
                    "Please install matplotlib or use a plot object with a figure attribute."
                ) from err

        return Pyplot.add(slide=slide, figure=figure, position=position, dpi=dpi, file_format=file_format, style=style)
