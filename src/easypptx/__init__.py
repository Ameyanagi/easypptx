"""EasyPPTX - Simple PowerPoint manipulation library."""

from importlib.metadata import PackageNotFoundError, version

from easypptx.chart import Chart
from easypptx.grid import Grid
from easypptx.image import Image
from easypptx.markdown import from_markdown
from easypptx.pandas_accessor import register as register_pandas_accessor
from easypptx.positioning import in_
from easypptx.presentation import Presentation
from easypptx.pyplot import Pyplot
from easypptx.slide import Slide
from easypptx.styles import ChartStyle, TableStyle, TextStyle, Theme
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
    "ChartStyle",
    "Grid",
    "Image",
    "Presentation",
    "Pyplot",
    "Slide",
    "Table",
    "TableStyle",
    "Template",
    "TemplateManager",
    "Text",
    "TextStyle",
    "Theme",
    "from_markdown",
    "generate_default_template",
    "generate_template_with_comments",
    "in_",
    "register_pandas_accessor",
]

# Register df.pptx when pandas is already loaded (free — no import happens)
from easypptx.pandas_accessor import maybe_register as _maybe_register_pandas_accessor

_maybe_register_pandas_accessor()
