"""EasyPPTX - Simple PowerPoint manipulation library."""

from importlib.metadata import PackageNotFoundError, version

from easypptx.chart import Chart
from easypptx.grid import Grid
from easypptx.image import Image
from easypptx.positioning import in_
from easypptx.presentation import Presentation
from easypptx.pyplot import Pyplot
from easypptx.slide import Slide
from easypptx.table import Table
from easypptx.template import Template, TemplateManager
from easypptx.template_generator import generate_default_template, generate_template_with_comments
from easypptx.text import Text

try:
    __version__ = version("easypptx")
except PackageNotFoundError:  # pragma: no cover - package not installed
    __version__ = "0.0.0"

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
    "in_",
]
