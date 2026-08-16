# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.10.0] - 2026-08-16

### Added
- **Fluent deck builder**: `Deck(theme=...)` / `pres.deck()` — build a whole presentation as one chain (`.title_slide().slide().bullets().chart().save()`). Calls validate eagerly (errors at the call site); rendering is lazy with content-aware layout: fixed blocks (text/bullets/tables) get natural estimated heights, flexible blocks (charts/images/figures) expand into the remaining space, and no coordinates are needed. Explicit x/y/width/height on any call opts that block out of auto-layout
- **Auto-pagination**: overflowing content flows onto "(cont.)" slides and long bullet lists split across pages at full size; `slide(paginate=False)` compresses instead
- `section()` divider slides, `notes()`, and a `tap(fn)` escape hatch to the low-level Slide API; a rendered deck cannot be reused (builder is consumed)

## [0.9.0] - 2026-08-16

Professional visual defaults. Decks built with a theme now look designed out of the box; render-verified on both light and dark themes.

### Changed
- **Themed titles** are left-aligned with a short accent bar underneath (disable with `Theme(title_accent=False)`); built-in theme typography refined (softer body grays, consistent title sizes)
- **Themed tables** are fully styled: theme-colored header row, subtle row banding, per-theme body colors, cell padding, vertical centering, and automatic right-alignment for numeric columns (headers follow their column)
- **Bullets** gained vertical rhythm: paragraph spacing, 1.1 line spacing, and nested levels stepping down in size
- **Chart gridlines** are luminance-aware: they fade toward white on light decks and toward black on dark decks
- **Markdown title slides** get a centered accent rule and refined type scale
- `Theme` gained `title_accent` and `table` fields; custom themes can supply their own table spec

## [0.8.4] - 2026-08-16

Polish pass driven by rendered-PNG inspection (PowerPoint + LibreOffice engines).

### Changed
- Data labels position outside the bar end (columns/bars) or above the point (lines/scatter) instead of overlapping neighbors
- Themed native charts get subtle gridlines derived from the theme text color (previously near-invisible on dark themes)
- `Slide.add_chart` legend defaults to "bottom", giving the plot the full slide width (pass legend_position="right" for the old layout)
- matplotlib-backed heatmaps render the input matrix in its original orientation (rows stay rows, series across the x axis)
- matplotlib-backed charts use the deck font when it is installed for matplotlib

## [0.8.3] - 2026-08-16

### Fixed
- Chart legends reserve their own layout space instead of floating over the plot area (found in PowerPoint-rendered visual QA)

## [0.8.2] - 2026-08-16

Driven by rendered-PNG visual QA of generated decks.

### Fixed
- Native chart text (axis labels, tick values, data labels, legend, axis titles) follows the deck theme's body color — charts on dark themes were nearly unreadable; explicit `font_color=` overrides
- matplotlib-backed charts on themed decks render with a transparent background and theme-colored text/spines instead of a white panel
- Grid-cell tables route through `Slide.add_table`, so DataFrames-and-friends, `number_format`, `shade_columns`, and styles work inside grids (previously a TypeError)

### Added
- `transparent=` on `Slide.add_pyplot` / `figure_to_stream`; `font_color=` on `Slide.add_chart` / `Chart.add`

## [0.8.1] - 2026-08-16

### Fixed
- `Grid.append()` skips spanned (merged-over) cells and places content with absolute slide coordinates; it now returns the created object instead of None
- Grid expansion preserves merged-region geometry (spans survive `next()`-triggered growth)
- Pie charts with axis options keep their palette colors (the axis warning no longer returned early)
- List-of-lists chart data raises `ValueError` for out-of-range integer column selectors instead of leaking `IndexError`

### Security
- Release workflow and its composite action pin all third-party GitHub Actions to full commit SHAs

## [0.8.0] - 2026-08-16

### Added
- **Universal data adapter**: `add_chart` and `add_table` accept pandas DataFrames and Series, polars DataFrames, numpy 1D/2D arrays (with `columns=`/`categories=` labels), dicts of sequences, and list-of-lists — all through one normalization layer (`easypptx.data`), detected via `sys.modules` with zero new dependencies
- **Chart backend routing**: native PowerPoint charts remain the default (editable, theme-aware); `chart_type="heatmap"/"histogram"/"box"/"violin"` renders via matplotlib into the same slide region; `backend="pyplot"` forces image rendering for any type; `backend="native"` raises a clear error for non-native types
- **Native chart styling**: `show_values=`/`number_format=` data labels, `x_title`/`y_title` axis titles, `y_min`/`y_max` limits, and `palette=` series colors — with the deck `Theme` palette applied automatically (built-in themes ship palettes)
- **Table formatting**: `number_format=` (Python format specs, per-column dicts), `shade_columns=`/`shade_color=` value-scaled cell tinting (from raw values, not formatted strings), and real table styles via GUID-backed ids
- **`df.pptx` pandas accessor**: `df.pptx.table(slide, ...)` and `df.pptx.chart(slide, kind=...)`; auto-registered when pandas is loaded, or explicitly via `easypptx.register_pandas_accessor()`
- **Text fitting**: `fit="shrink"` (default) computes a fitting font size client-side — decks render correctly in LibreOffice/previews, not just after PowerPoint recalculates autofit; `fit="resize"` grows the box; `fit="none"` opts out. CJK-aware width estimation (`easypptx.textfit`); the markdown renderer now allocates block heights from estimated line counts

### Changed
- `Slide.add_chart` gained `backend`, `columns`, `show_values`, `number_format`, `x_title`, `y_title`, `y_min`, `y_max`, `palette`
- `Slide.add_table` gained `columns`, `number_format`, `shade_columns`, `shade_color`; `style` accepts GUID strings
- `Theme` gained a `palette` field; built-in light/dark/corporate themes define palettes

## [0.7.0] - 2026-08-16

### Added
- **Markdown → deck**: `Presentation.from_markdown()` converts a markdown document into a presentation — frontmatter (`theme`, `template`, `aspect_ratio`), `#` title slide with subtitle, `##` slide breaks, nested bullet lists, images, GFM tables, fenced code blocks, `<!-- notes: ... -->` speaker notes, `::: columns` side-by-side layout, and `---` forced breaks
- **Grid upgrades**: slice spans (`grid[1, :]`, `grid[0:2, 1]` merge and return the region), weighted tracks (`Grid(rows=[2, 1])`), auto-flow (`grid.next()` fills the next free cell and grows the grid), and per-cell styling (`grid[r, c].style(fill=..., border_color=..., padding=...)`)
- **Content**: `Slide.add_bullets()` with nesting levels (real bullets or plain stacked paragraphs), `Slide.notes` speaker-notes property, `Slide.add_pyplot()`, and multi-series charts (`value_columns=[...]` plots every listed column; `Chart.add(series={...})`)
- **Styling**: `TextStyle`/`TableStyle`/`ChartStyle` objects accepted by the Slide content methods (explicit arguments always win), and `Theme` with built-in presets — `Presentation(theme="dark")` (also "light", "corporate")
- `in_()` helper for absolute inch positions, exported from the package root

### Changed
- **BREAKING — positions**: bare numbers are now percentages (`x=10` means 10%, identical to `"10%"`); absolute inches require `in_(1.5)`. Content-method defaults converted to clean percentages (e.g. `add_text` defaults to x=5, y=5, width=90, height=10)
- **BREAKING — removed the APIs deprecated in 0.6.0**: `Presentation.add_text/add_image/add_shape/add_table/add_chart/add_pyplot` (pass-through variants) and `add_matplotlib_slide`/`add_seaborn_slide`/`add_plot`/`add_image_slide`. Use the `Slide` methods and `add_pyplot_slide`/`add_image_gen_slide`
- `Slide.add_chart` parameters `chart_type`/`has_legend`/`legend_position` accept None (resolved from ChartStyle, then defaults)

### Removed
- AI-workflow scaffolding from the repository (ai_docs/, .states/, specs/, CLAUDE.md, cookiecutter leftovers)

## [0.6.0] - 2026-08-16

### Added
- `Slide.add_table` and `Slide.add_chart` — content methods now live on `Slide`, so `slide.add_table(...)` / `slide.add_chart(...)` work as the README always showed
- `Slide.add_shape` (and `Presentation.add_shape`) accept string shape names such as `"ROUNDED_RECTANGLE"` in addition to `MSO_SHAPE` enums
- `Presentation.add_slide` gained `title_color`; `template_toml=False` opts a single slide out of the presentation's default template
- New shared modules: `easypptx.positioning` (percent/inch conversion and layout arithmetic) and `easypptx.common` (colors, alignment, font constants)
- Common parameter aliases are accepted: `has_header`/`first_row_header`, `value_column`/`value_columns`, `vertical`/`vertical_align`, `chart_title`/`title`
- `py.typed` marker — the library's annotations are now visible to downstream type checkers
- Optional dependency extras: `easypptx[dataframe]` (pandas), `easypptx[plot]` (matplotlib), `easypptx[all]`

### Changed
- Unknown keyword arguments now emit a warning instead of being silently discarded (reverts the silent `**kwargs` behavior from 0.5.6); out-of-range percentages are clamped with a warning
- Invalid or missing template files now raise `FileNotFoundError`/`ValueError` instead of printing a warning and silently falling back to the default template
- `Presentation.slides` returns cached `Slide` wrappers, so `user_data` and object identity are preserved across accesses
- matplotlib figures are embedded via in-memory streams instead of temporary files
- TOML templates are cached by path and mtime, so per-slide `template_toml` no longer re-reads the file for every slide
- pandas, matplotlib, and seaborn are no longer hard dependencies; pandas/matplotlib moved to extras and seaborn was dropped (pass a seaborn plot's figure to `add_pyplot_slide`)
- Title/subtitle/content layout arithmetic now works with inch values as well as percent strings
- Tooling: mypy replaced by `ty`, pre-commit replaced by `lefthook`, tox removed; CI matrix now matches `requires-python` (3.12/3.13)

### Deprecated
- `Presentation.add_text/add_image/add_shape/add_table/add_chart/add_pyplot` (pass-through variants taking a `slide` argument) — use the `Slide` methods directly
- `Presentation.add_matplotlib_slide`, `add_seaborn_slide`, `add_plot` — use `add_pyplot_slide`
- `Presentation.add_image_slide` — use `add_image_gen_slide`

### Fixed
- `content_y_padding` was applied twice in `add_grid_slide`, shifting and shrinking the grid double
- `Presentation.add_image`'s documented `maintain_aspect_ratio` parameter now actually takes effect (routed through `Image.add`); `crop`/`center` warn that they are unimplemented
- `add_slide_from_template` no longer silently replaces the presentation (discarding existing slides) when a template names a different reference PPTX — it now raises if slides exist
- `Table.from_dataframe(include_index=True)` no longer raises a column-count error
- `Grid.add_table` and grid-cell content are now positioned absolutely on the slide (previously cell-relative coordinates were used as slide coordinates)
- `Grid.autogrid` passes the computed cell position to content functions that can accept it (previously all positions were dropped)
- `__version__` is derived from package metadata instead of a hardcoded, out-of-date string
- Fixed latent `None`-arithmetic in `Grid.autogrid` row/column inference
- Color constants unified: "black" is 0x101010 everywhere (the template module's diverged copy was removed)

## [0.5.6] - 2025-05-14

### Fixed
- Modified all content-related methods in Slide and Grid classes to accept additional parameters
- Fixed `TypeError` when methods receive unexpected parameters like 'padding'
- Added `**kwargs` parameter to ensure API flexibility and backward compatibility
- Updated documentation for all affected methods

## [0.5.5] - 2025-05-14

### Fixed
- Version bump for minor changes

## [0.5.4] - 2025-05-14

### Fixed
- Improved template defaults cascade priority to ensure consistent behavior
- Fixed `defaults.global` settings not being respected in alignment and other properties
- Ensured proper inheritance of global defaults into method-specific defaults
- Enhanced consistency between Grid and Slide classes for template handling

### Added
- New example demonstrating how to use global defaults in templates: `examples/templates/011_template_global_defaults.py`

## [0.5.3] - 2025-05-13

### Fixed
- Title alignment settings from TOML templates now properly apply to `add_grid_slide` and `add_autogrid_slide` methods
- Fixed `title_align` parameter to respect template settings when not explicitly specified

## [0.5.2] - 2025-05-13

### Fixed
- Added example showing how to properly use template-based alignment settings: `examples/templates/001_template_align_fix.py`

## [0.5.1] - 2025-05-13

### Fixed
- Updated all template TOML files to use RGB arrays instead of hex color codes for better compatibility
- Prevent color errors when using hex codes in template files
- Ensured backwards compatibility with predefined color names

## [0.5.0] - 2025-05-13

### Added
- Title and content padding parameters for all slide creation methods
  - Added `title_padding`, `title_x_padding`, and `title_y_padding` parameters
  - Added `subtitle_padding`, `subtitle_x_padding`, and `subtitle_y_padding` parameters
  - Added `content_padding`, `content_x_padding`, and `content_y_padding` parameters
  - Added `label_padding`, `label_x_padding`, and `label_y_padding` parameters
- Enhanced positioning control in all slide types (standard, grid, autogrid, pyplot, image)
- New example demonstrating title padding features in `examples/styling/004_title_padding.py`
- Comprehensive test suite for title padding functionality

### Fixed
- Fixed recursive call between `add_slide` and `add_slide_from_template` causing maximum recursion error

## [0.4.0] - 2025-05-13

### Added
- Direct TOML template integration
  - Added `template_toml` parameter to Presentation constructor for default template application
  - Added `template_toml` parameter to add_slide method for per-slide template customization
  - Implemented template priority system (slide-specific template > default template)
  - Reorganized template examples with numbered convention for clarity
  - Added comprehensive example demonstrating the new template_toml feature
- Updated documentation with new template_toml usage examples

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

[0.5.2]: https://github.com/Ameyanagi/easypptx/compare/v0.5.1...v0.5.2
[0.5.1]: https://github.com/Ameyanagi/easypptx/compare/v0.5.0...v0.5.1
[0.5.0]: https://github.com/Ameyanagi/easypptx/compare/v0.4.0...v0.5.0
[0.4.0]: https://github.com/Ameyanagi/easypptx/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/Ameyanagi/easypptx/compare/v0.2.0...v0.3.0
[0.0.7]: https://github.com/Ameyanagi/easypptx/compare/v0.0.6...v0.0.7
[0.0.6]: https://github.com/Ameyanagi/easypptx/compare/v0.0.5...v0.0.6
[0.0.5]: https://github.com/Ameyanagi/easypptx/compare/v0.0.4...v0.0.5
[0.0.4]: https://github.com/Ameyanagi/easypptx/compare/v0.0.3...v0.0.4
[0.0.3]: https://github.com/Ameyanagi/easypptx/compare/v0.0.2...v0.0.3
[0.0.2]: https://github.com/Ameyanagi/easypptx/releases/tag/v0.0.2
