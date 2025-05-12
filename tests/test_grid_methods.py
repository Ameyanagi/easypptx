"""Tests for the grid methods in the Presentation class."""

from easypptx import Presentation
from easypptx.grid import Grid


def test_add_grid():
    """Test the add_grid method."""
    pres = Presentation()
    slide = pres.add_slide()

    grid = pres.add_grid(slide, rows=2, cols=3)

    assert isinstance(grid, Grid)
    assert grid.rows == 2
    assert grid.cols == 3
    assert grid.x == "0%"
    assert grid.y == "0%"
    assert grid.width == "100%"
    assert grid.height == "100%"

    # Test with custom dimensions
    grid2 = pres.add_grid(slide, x="10%", y="20%", width="80%", height="60%", rows=3, cols=2)

    assert isinstance(grid2, Grid)
    assert grid2.rows == 3
    assert grid2.cols == 2
    assert grid2.x == "10%"
    assert grid2.y == "20%"
    assert grid2.width == "80%"
    assert grid2.height == "60%"


def test_add_grid_slide():
    """Test the add_grid_slide method."""
    pres = Presentation()

    # Test without title
    slide, grid = pres.add_grid_slide(rows=2, cols=2)

    assert slide is not None
    assert isinstance(grid, Grid)
    assert grid.rows == 2
    assert grid.cols == 2
    assert grid.x == "0%"
    assert grid.y == "0%"
    assert grid.width == "100%"
    assert grid.height == "100%"

    # Test with title
    slide2, grid2 = pres.add_grid_slide(rows=3, cols=3, title="Test Grid", title_height="15%")

    assert slide2 is not None
    assert isinstance(grid2, Grid)
    assert grid2.rows == 3
    assert grid2.cols == 3
    assert grid2.x == "0%"
    assert grid2.y == "15.00%"
    assert grid2.width == "100%"
    assert grid2.height == "85.00%"


def test_add_autogrid():
    """Test the add_autogrid method."""
    pres = Presentation()
    slide = pres.add_slide()

    # Mock content functions
    def func1():
        return None

    def func2():
        return None

    content_funcs = [func1, func2]

    # Test with explicit rows and cols
    grid = pres.add_autogrid(slide, content_funcs, rows=1, cols=2)

    assert isinstance(grid, Grid)
    assert grid.rows == 1
    assert grid.cols == 2

    # Test with auto-calculated rows and cols
    grid2 = pres.add_autogrid(slide, content_funcs)

    assert isinstance(grid2, Grid)
    # For 2 content functions, we should get a 1x2 or 2x1 grid
    assert grid2.rows * grid2.cols >= len(content_funcs)


def test_add_autogrid_slide():
    """Test the add_autogrid_slide method."""
    pres = Presentation()

    # Mock content functions
    def func1():
        return None

    def func2():
        return None

    content_funcs = [func1, func2]

    # Test without title
    slide, grid = pres.add_autogrid_slide(content_funcs, rows=1, cols=2)

    assert slide is not None
    assert isinstance(grid, Grid)
    assert grid.rows == 1
    assert grid.cols == 2

    # Test with title
    slide2, grid2 = pres.add_autogrid_slide(content_funcs, rows=1, cols=2, title="Test AutoGrid", title_height="15%")

    assert slide2 is not None
    assert isinstance(grid2, Grid)
    assert grid2.rows == 1
    assert grid2.cols == 2
    # Expect two decimal places in percentage values
    assert float(grid2.y.strip("%")) == 15.00
    assert float(grid2.height.strip("%")) == 85.00
