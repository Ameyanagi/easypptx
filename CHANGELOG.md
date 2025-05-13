# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.3.0] - 2025-05-13

### Added
- Custom reference PPTX file support
  - Added `reference_pptx` parameter to Presentation constructor
  - Added support for specifying reference PPTX files in TOML template files
  - Added automatic blank layout detection for reference PPTX files
  - Added `blank_layout_index` parameter for specifying which layout to use as blank
- New methods in TemplateManager:
  - Added `get_reference_pptx` and `get_blank_layout_index` methods
- New examples:
  - Added example showing how to use custom reference PPTX files directly
  - Added example demonstrating reference PPTX specification in TOML files
- Updated documentation with comprehensive examples of the new features

## [0.0.7] - 2025-05-12

### Added
- Dynamic Grid features for easier content management
  - Added `append()` method to Grid for auto-layout updates
  - Implemented auto-expansion for out-of-bounds cell access
  - Grid now automatically expands when accessing cells beyond current dimensions
  - Added example demonstrating dynamic grid features in `examples/dynamic_grid_example.py`
- Added title_align parameter to Grid.autogrid and Grid.autogrid_pyplot
- Maintained backward compatibility for existing code

## [0.0.6] - 2025-05-12

### Added
- Enhanced Grid indexing functionality
  - Added support for flat indexing with `grid[idx]` (row-major ordering)
  - Negative indices now supported for flat indexing (e.g., `grid[-1]` for last cell)
  - Added example showing different grid indexing methods
- Prioritized backward compatibility for existing code

## [0.0.5] - 2025-05-12

### Added
- Title alignment control for all slide creation methods
  - Added `title_align` parameter to `add_grid_slide`
  - Added `title_align` parameter to `add_autogrid_slide`
  - Added `title_align`, `subtitle_align`, and `label_align` parameters to `add_pyplot_slide`
  - All alignment parameters support "left", "center", and "right" values

## [0.0.4] - 2025-05-12

### Added
- New Grid Features
  - Enhanced row-level Grid access API with `grid[row].add_xxx()` methods
  - Added `reset()` method to GridRowProxy to allow reusing rows
  - Implemented convenient `add_textbox()` alias for consistent API
- New Slide Creation Methods with Consistent API
  - Added `add_grid_slide` for creating slides with grid layouts
  - Added `add_pyplot_slide` for creating slides with matplotlib/seaborn figures
  - Added `add_image_gen_slide` for creating slides with images
  - All methods support title, subtitle, and flexible positioning
  - Consistent parameter naming and return value patterns
  - Return both the slide and the content object for easy customization
- Updated add_autogrid_slide method to support empty grids
- Added examples demonstrating all the new features
- Comprehensive tests for the new functionality

## [0.0.3] - 2025-05-12

### Added
- Enhanced Grid access API with intuitive syntax
  - Added `grid[row, col].add_xxx()` for direct cell access
  - Added `grid[row].add_xxx()` for sequential row operations
- Added GridCellProxy and GridRowProxy classes to support new functionality
- Added example file demonstrating enhanced grid access patterns
- Updated documentation to reflect new features

## [0.0.2] - Previous release

### Added
- Initial project structure
- Basic PowerPoint presentation creation functionality
- Grid layout system
- Basic examples

[0.3.0]: https://github.com/Ameyanagi/easypptx/compare/v0.2.0...v0.3.0
[0.0.7]: https://github.com/Ameyanagi/easypptx/compare/v0.0.6...v0.0.7
[0.0.6]: https://github.com/Ameyanagi/easypptx/compare/v0.0.5...v0.0.6
[0.0.5]: https://github.com/Ameyanagi/easypptx/compare/v0.0.4...v0.0.5
[0.0.4]: https://github.com/Ameyanagi/easypptx/compare/v0.0.3...v0.0.4
[0.0.3]: https://github.com/Ameyanagi/easypptx/compare/v0.0.2...v0.0.3
[0.0.2]: https://github.com/Ameyanagi/easypptx/releases/tag/v0.0.2
