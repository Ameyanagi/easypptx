"""EasyPPTX - Simple PowerPoint manipulation library."""

from easypptx.chart import Chart
from easypptx.grid import Grid
from easypptx.image import Image
from easypptx.presentation import Presentation
from easypptx.pyplot import Pyplot
from easypptx.slide import Slide
from easypptx.table import Table
from easypptx.template import Template, TemplateManager
from easypptx.text import Text

__version__ = "0.3.0"

__all__ = ["Chart", "Grid", "Image", "Presentation", "Pyplot", "Slide", "Table", "Template", "TemplateManager", "Text"]
