"""EasyPPTX - Simple PowerPoint manipulation library."""

from easypptx.chart import Chart
from easypptx.grid import Grid
from easypptx.image import Image
from easypptx.presentation import Presentation
from easypptx.pyplot import Pyplot
from easypptx.slide import Slide
from easypptx.table import Table
from easypptx.template import Template, TemplateManager
from easypptx.template_generator import generate_default_template, generate_template_with_comments
from easypptx.text import Text

__version__ = "0.5.3"

__all__ = [
    "Chart",
    "Grid",
    "Image",
    "Presentation",
    "Pyplot",
    "Slide",
    "Table",
    "Template",
    "TemplateManager",
    "Text",
    "generate_default_template",
    "generate_template_with_comments",
]
