"""Pyplot integration module for EasyPPTX."""

from __future__ import annotations

import io
from typing import Any

from easypptx.common import COLORS, apply_shadow
from easypptx.image import Image


def figure_to_stream(figure: Any, dpi: int = 300, file_format: str = "png") -> io.BytesIO:
    """Render a matplotlib figure to an in-memory image stream.

    Args:
        figure: Matplotlib figure object
        dpi: Resolution for the figure (default: 300)
        file_format: Image format ("png" or "jpg") (default: "png")

    Returns:
        A BytesIO stream positioned at the start of the image data
    """
    stream = io.BytesIO()
    figure.savefig(stream, dpi=dpi, format=file_format, bbox_inches="tight")
    stream.seek(0)
    return stream


class Pyplot:
    """Class for adding matplotlib/seaborn plots to slides."""

    @staticmethod
    def add(
        slide: Any,
        figure: Any,
        position: dict[str, float | str],
        dpi: int = 300,
        file_format: str = "png",
        style: dict[str, Any] | None = None,
    ) -> Any:
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
            from easypptx import Presentation, Pyplot

            plt.figure(figsize=(10, 6))
            plt.plot([1, 2, 3, 4], [1, 4, 9, 16])

            pres = Presentation()
            slide = pres.add_slide()

            Pyplot.add(
                slide=slide,
                figure=plt.gcf(),
                position={"x": "10%", "y": "20%", "width": "80%", "height": "70%"}
            )
            ```
        """
        if style is None:
            style = {"maintain_aspect_ratio": True, "center": True, "border": False}

        # Render the figure in memory (no temporary files)
        stream = figure_to_stream(figure, dpi=dpi, file_format=file_format)

        img = Image(slide)
        image_shape = img.add(
            image_path=stream,
            x=position.get("x", "10%"),
            y=position.get("y", "20%"),
            width=position.get("width", "80%"),
            height=position.get("height", "70%"),
            maintain_aspect_ratio=style.get("maintain_aspect_ratio", True),
        )

        # Apply border if specified
        if style.get("border", False):
            image_shape.line.color.rgb = COLORS.get(style.get("border_color", "black"), COLORS["black"])
            image_shape.line.width = style.get("border_width", 1)

        # Apply shadow if specified
        if style.get("shadow", False):
            apply_shadow(image_shape)

        return image_shape

    @staticmethod
    def add_from_seaborn(
        slide: Any,
        seaborn_plot: Any,
        position: dict[str, float | str],
        dpi: int = 300,
        file_format: str = "png",
        style: dict[str, Any] | None = None,
    ) -> Any:
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
        """
        # Get the figure from the seaborn plot
        if hasattr(seaborn_plot, "figure"):
            figure = seaborn_plot.figure
        elif hasattr(seaborn_plot, "fig"):
            figure = seaborn_plot.fig
        else:
            try:
                import matplotlib.pyplot as plt
            except ImportError as err:
                raise ImportError(
                    "Matplotlib is required for plots with no figure attribute. "
                    "Install it with: pip install 'easypptx[plot]'"
                ) from err
            figure = plt.gcf()

        return Pyplot.add(slide=slide, figure=figure, position=position, dpi=dpi, file_format=file_format, style=style)
