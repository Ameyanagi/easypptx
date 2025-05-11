"""Tests for the Grid class."""

from unittest.mock import MagicMock

import pytest

from easypptx.grid import Grid, GridCell


class TestGrid:
    """Test cases for the Grid class."""

    def setup_method(self):
        """Set up test environment before each test method."""
        self.slide = MagicMock()
        self.parent = self.slide

    def test_grid_init(self):
        """Test Grid initialization with default parameters."""
        grid = Grid(parent=self.parent)

        # Verify grid properties
        assert grid.parent == self.parent
        assert grid.x == "0%"
        assert grid.y == "0%"
        assert grid.width == "100%"
        assert grid.height == "100%"
        assert grid.rows == 1
        assert grid.cols == 1
        assert grid.padding == 5.0

        # Verify a cell was created
        assert len(grid.cells) == 1
        assert len(grid.cells[0]) == 1
        assert isinstance(grid.cells[0][0], GridCell)

    def test_grid_init_custom(self):
        """Test Grid initialization with custom parameters."""
        grid = Grid(
            parent=self.parent,
            x="10%",
            y="20%",
            width="80%",
            height="60%",
            rows=3,
            cols=2,
            padding=10.0,
        )

        # Verify grid properties
        assert grid.parent == self.parent
        assert grid.x == "10%"
        assert grid.y == "20%"
        assert grid.width == "80%"
        assert grid.height == "60%"
        assert grid.rows == 3
        assert grid.cols == 2
        assert grid.padding == 10.0

        # Verify cells were created
        assert len(grid.cells) == 3
        assert len(grid.cells[0]) == 2
        assert isinstance(grid.cells[0][0], GridCell)

    def test_create_cells(self):
        """Test cell creation with padding calculation."""
        grid = Grid(parent=self.parent, rows=2, cols=2, padding=10.0)

        # Check if cells are created with correct dimensions
        assert len(grid.cells) == 2
        assert len(grid.cells[0]) == 2

        # Check first cell position and dimensions
        # With 10% padding, each cell should be about 45% width/height
        # and positioned with appropriate padding
        cell = grid.cells[0][0]
        assert float(cell.x.strip("%")) < 10  # Should be positioned with padding
        assert float(cell.y.strip("%")) < 10  # Should be positioned with padding
        assert float(cell.width.strip("%")) > 40  # Width accounting for padding
        assert float(cell.height.strip("%")) > 40  # Height accounting for padding

    def test_get_cell(self):
        """Test retrieving a cell from the grid."""
        grid = Grid(parent=self.parent, rows=3, cols=3)

        # Get a cell and verify it's the correct one
        cell = grid.get_cell(1, 2)
        assert cell.row == 1
        assert cell.col == 2

        # Test out of bounds
        with pytest.raises(IndexError):
            grid.get_cell(3, 3)

    def test_merge_cells(self):
        """Test merging cells."""
        grid = Grid(parent=self.parent, rows=3, cols=3)

        # Merge a 2x2 area
        merged_cell = grid.merge_cells(0, 0, 1, 1)

        # Verify the merged cell dimensions
        assert merged_cell.span_rows == 2
        assert merged_cell.span_cols == 2

        # Verify other cells in the merged area are marked as spanned
        assert grid.cells[0][1].is_spanned
        assert grid.cells[1][0].is_spanned
        assert grid.cells[1][1].is_spanned

        # Verify cells outside the merged area are not spanned
        assert not grid.cells[0][2].is_spanned
        assert not grid.cells[2][0].is_spanned

        # Test merging invalid ranges
        with pytest.raises(IndexError):
            grid.merge_cells(0, 0, 3, 3)  # Out of bounds

        with pytest.raises(ValueError):
            grid.merge_cells(1, 1, 0, 0)  # End before start

        # Test merging cells that are already part of a merged area
        with pytest.raises(ValueError):
            grid.merge_cells(0, 1, 2, 2)  # Includes already merged cell

    def test_add_to_cell(self):
        """Test adding content to a cell."""
        grid = Grid(parent=self.parent, rows=2, cols=2)

        # Create a mock content function
        content_func = MagicMock()
        content_func.return_value = "test_content"

        # Add content to a cell
        result = grid.add_to_cell(0, 1, content_func, text="Test Text")

        # Verify content function was called with correct parameters
        content_func.assert_called_once()
        call_kwargs = content_func.call_args[1]

        # Verify position parameters were calculated
        assert "x" in call_kwargs
        assert "y" in call_kwargs
        assert "width" in call_kwargs
        assert "height" in call_kwargs
        assert "text" in call_kwargs
        assert call_kwargs["text"] == "Test Text"

        # Verify result is what content_func returned
        assert result == "test_content"

        # Verify content was stored in the cell
        assert grid.cells[0][1].content == "test_content"

        # Test adding to a spanned cell
        grid.merge_cells(0, 0, 1, 0)
        with pytest.raises(ValueError):
            grid.add_to_cell(1, 0, content_func)  # Cell is spanned

    def test_add_grid_to_cell(self):
        """Test adding a nested grid to a cell."""
        grid = Grid(parent=self.parent, rows=2, cols=2)

        # Add a nested grid
        nested_grid = grid.add_grid_to_cell(0, 1, rows=3, cols=3, padding=2.0)

        # Verify the nested grid was created with the correct properties
        assert isinstance(nested_grid, Grid)
        assert nested_grid.parent == self.parent
        assert nested_grid.rows == 3
        assert nested_grid.cols == 3
        assert nested_grid.padding == 2.0

        # Verify position inheritance from parent grid
        assert nested_grid.x != grid.x  # Should be adjusted
        assert nested_grid.y != grid.y  # Should be adjusted
        assert nested_grid.width != grid.width  # Should be scaled
        assert nested_grid.height != grid.height  # Should be scaled

        # Verify content was stored in the cell
        assert grid.cells[0][1].content == nested_grid

        # Test adding to a spanned cell
        grid.merge_cells(1, 0, 1, 1)
        with pytest.raises(ValueError):
            grid.add_grid_to_cell(1, 1, rows=2, cols=2)  # Cell is spanned


class TestGridCell:
    """Test cases for the GridCell class."""

    def test_grid_cell_init(self):
        """Test GridCell initialization."""
        cell = GridCell(row=1, col=2, x="10%", y="20%", width="30%", height="40%")

        # Verify cell properties
        assert cell.row == 1
        assert cell.col == 2
        assert cell.x == "10%"
        assert cell.y == "20%"
        assert cell.width == "30%"
        assert cell.height == "40%"
        assert cell.content is None
        assert cell.span_rows == 1
        assert cell.span_cols == 1
        assert not cell.is_spanned

    def test_grid_cell_repr(self):
        """Test GridCell string representation."""
        cell = GridCell(row=1, col=2, x="10%", y="20%", width="30%", height="40%")

        # Verify string representation contains all relevant information
        repr_str = repr(cell)
        assert "GridCell" in repr_str
        assert "row=1" in repr_str
        assert "col=2" in repr_str
        assert "x=10%" in repr_str
        assert "y=20%" in repr_str
        assert "width=30%" in repr_str
        assert "height=40%" in repr_str
