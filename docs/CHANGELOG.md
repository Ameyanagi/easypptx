# Changelog

## [0.8.0] - 2026-08-16

Non-breaking release.

### Added
- Universal data adapter (`easypptx.data`): `slide.add_chart(data=...)` and `slide.add_table(data=...)` accept pandas DataFrames and Series, polars DataFrames, numpy 1D/2D arrays (label with `columns=`), dicts of sequences, and lists of lists with a header row — all duck-typed, none required as dependencies
- Chart backend routing: `chart_type="heatmap"`/`"histogram"`/`"box"`/`"violin"` render via matplotlib (requires `easypptx[plot]`) into the same slide region and return a picture shape; `backend="pyplot"` forces matplotlib for any type, `backend="native"` raises a clear error for non-native types. Native charts (`column`, `bar`, `line`, `pie`, `area`, `scatter`) remain the default and stay editable
- Native chart styling on `add_chart` and `Chart.add`: `show_values=`, `number_format=` (Excel-style, for data labels), `x_title=`, `y_title=`, `y_min=`, `y_max=`, and `palette=`
- `Theme.palette`: themed presentations color native chart series automatically (the built-in `light`/`dark`/`corporate` themes ship palettes); an explicit `palette=` wins
- Table formatting on `add_table`: `number_format=` (Python format spec or per-column dict), `shade_columns=` + `shade_color=` value-scaled background tints, `columns=` header names for numpy arrays, and `style=` now accepts GUID strings with small ints mapping to built-in PowerPoint table styles
- `df.pptx` pandas accessor: `df.pptx.table(slide, ...)` and `df.pptx.chart(slide, kind=..., ...)`; auto-registers when pandas is imported before easypptx (or when a `Presentation` is created), with `easypptx.register_pandas_accessor()` for manual registration
- Text fitting (`easypptx.textfit`): `fit=` on `add_text` and `add_bullets` — `"shrink"` (default) writes a fitting font size into the file using CJK-aware wrap estimation so decks render correctly outside PowerPoint too, `"resize"` grows the box, `"none"` opts out
- The markdown renderer allocates block heights from estimated line counts

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
