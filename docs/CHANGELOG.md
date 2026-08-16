# Changelog

## [0.7.0] - 2026-08-16

### Breaking Changes
- Bare numbers in position parameters are now percentages (`x=10` == `"10%"`); floats no longer mean inches. Absolute inches use the new `in_()` helper.
- Default positions changed from inches to percentages (e.g. `add_text`: x=5, y=5, width=90, height=10; `add_chart`: x=10, y=20, width=60, height=60).
- Removed the deprecated `Presentation` pass-throughs (`add_text`, `add_image`, `add_shape`, `add_table`, `add_chart`, `add_pyplot`) and `add_matplotlib_slide`, `add_seaborn_slide`, `add_plot`, `add_image_slide`. Use the `Slide` methods, `add_pyplot_slide`, and `add_image_gen_slide`.

### Added
- Markdown-to-deck conversion: `Presentation.from_markdown` / `easypptx.from_markdown`
- Grid upgrades: slice spans (`grid[1, :]`), weighted tracks (`rows=[2, 1]`), auto-flow (`grid.next()`), and cell styling (`grid[r, c].style(...)`)
- `slide.add_bullets`, `slide.notes`, and `slide.add_pyplot`
- Multi-series charts via `value_columns=[...]` and `Chart.add(..., series={...})`
- `TextStyle`, `TableStyle`, `ChartStyle` dataclasses with a `style=` parameter, plus `Theme` and built-in presets (`light`, `dark`, `corporate`)

## [Unreleased]

### Added
- Percentage-based positioning for all elements (text, images, shapes)
- Auto-alignment of multiple objects (grid, horizontal, vertical layouts)
- Dark theme support with custom background colors
- Expanded color palette for modern designs
- Default styling with Meiryo font and color scheme
- PowerPoint template support with blank layout as default
- New helper methods for slide dimension calculation
- Comprehensive documentation for all features
- Extended examples demonstrating new features
- Blank layout example demonstrating different slide layouts

### Changed
- Updated text formatting to support named colors
- Improved positioning flexibility with multiple coordinate systems
- Enhanced method parameter documentation
- Standardized API for consistent usage patterns

### Fixed
- Slide dimension calculation for various contexts
- Test suite compatibility with new features

## Documentation

### Added
- User guides for percentage-based positioning
- User guides for auto-alignment of multiple objects
- User guides for styling and formatting
- User guides for PowerPoint templates
- API reference documentation
- Features overview page
- Expanded example code in documentation
- Updated modules documentation
