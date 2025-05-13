"""Tests for grid flat indexing access API."""

from unittest.mock import MagicMock

import pytest

from easypptx import Presentation
from easypptx.grid import Grid, GridCellProxy


def test_grid_flat_access():
    """Test the flat indexing access API for Grid."""
    # Create a mock parent
    parent = MagicMock()
    parent.add_text = MagicMock(return_value="text_shape")
    parent.add_image = MagicMock(return_value="image_shape")

    # Create a grid
    grid = Grid(parent=parent, rows=2, cols=2)

    # Test accessing a cell using flat indexing
    cell_proxy0 = grid[0]  # Should get the cell at (0,0)
    assert isinstance(cell_proxy0, GridCellProxy)
    assert cell_proxy0.row == 0
    assert cell_proxy0.col == 0

    # Test accessing other cells
    cell_proxy1 = grid[1]  # Should get the cell at (0,1)
    assert isinstance(cell_proxy1, GridCellProxy)
    assert cell_proxy1.row == 0
    assert cell_proxy1.col == 1

    cell_proxy2 = grid[2]  # Should get the cell at (1,0)
    assert isinstance(cell_proxy2, GridCellProxy)
    assert cell_proxy2.row == 1
    assert cell_proxy2.col == 0

    cell_proxy3 = grid[3]  # Should get the cell at (1,1)
    assert isinstance(cell_proxy3, GridCellProxy)
    assert cell_proxy3.row == 1
    assert cell_proxy3.col == 1

    # Test adding content using flat index
    result = cell_proxy0.add_text("Text at 0,0")
    assert result == "text_shape"

    # Test with negative index
    cell_proxy_neg = grid[-1]  # Should get the last cell (1,1)
    assert cell_proxy_neg.row == 1
    assert cell_proxy_neg.col == 1

    # Test out of bounds
    with pytest.raises(IndexError):
        grid[4]  # Out of bounds for a 2x2 grid

    with pytest.raises(IndexError):
        grid[-5]  # Out of bounds negative index


def test_add_autogrid_slide_with_flat_indexing():
    """Test the add_autogrid_slide method with flat indexing."""
    pres = Presentation()

    # Test with empty grid
    slide, grid = pres.add_autogrid_slide(content_funcs=None, rows=3, cols=2, title="Test Title")

    # Verify grid dimensions
    assert grid.rows == 3
    assert grid.cols == 2

    # Test adding content using flat indexing
    grid[0].add_text("First cell (0,0)")
    grid[1].add_text("Second cell (0,1)")
    grid[2].add_text("Third cell (1,0)")
    grid[3].add_text("Fourth cell (1,1)")
    grid[4].add_text("Fifth cell (2,0)")
    grid[5].add_text("Sixth cell (2,1)")

    # Test out of bounds
    with pytest.raises(IndexError):
        grid[6]  # Out of bounds for a 3x2 grid
