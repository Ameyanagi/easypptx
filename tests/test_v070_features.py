"""Tests for the 0.7.0 features: grid upgrades, bullets, notes, styles, themes, markdown."""

import pandas as pd
import pytest
from PIL import Image as PILImage

from easypptx import ChartStyle, Presentation, TableStyle, TextStyle, Theme, in_
from easypptx.positioning import to_percent

SLIDE_W = 12192768
SLIDE_H = 6858000


class TestGridUpgrades:
    def test_weighted_rows(self):
        pres = Presentation()
        slide = pres.add_slide()
        grid = pres.add_grid(slide=slide, rows=[2, 1], cols=1, padding=0.0)
        assert to_percent(grid.cells[0][0].height, SLIDE_H) == pytest.approx(66.67, abs=0.05)
        assert to_percent(grid.cells[1][0].height, SLIDE_H) == pytest.approx(33.33, abs=0.05)
        assert to_percent(grid.cells[1][0].y, SLIDE_H) == pytest.approx(66.67, abs=0.05)

    def test_weighted_cols(self):
        pres = Presentation()
        slide = pres.add_slide()
        grid = pres.add_grid(slide=slide, rows=1, cols=[1, 3], padding=0.0)
        assert to_percent(grid.cells[0][0].width, SLIDE_W) == pytest.approx(25.0, abs=0.05)
        assert to_percent(grid.cells[0][1].width, SLIDE_W) == pytest.approx(75.0, abs=0.05)

    def test_slice_span_row(self):
        pres = Presentation()
        slide = pres.add_slide()
        grid = pres.add_grid(slide=slide, rows=2, cols=2, padding=0.0)
        shape = grid[1, :].add_text("spans the whole bottom row")
        # The merged origin covers both columns
        origin = grid.cells[1][0]
        assert origin.span_cols == 2
        assert grid.cells[1][1].is_spanned
        assert shape is not None
        # Repeat access to the same span must not raise
        grid[1, :]

    def test_slice_span_column(self):
        pres = Presentation()
        slide = pres.add_slide()
        grid = pres.add_grid(slide=slide, rows=3, cols=2, padding=0.0)
        grid[0:2, 1].add_text("two rows tall")
        assert grid.cells[0][1].span_rows == 2

    def test_slice_step_rejected(self):
        pres = Presentation()
        slide = pres.add_slide()
        grid = pres.add_grid(slide=slide, rows=3, cols=3)
        with pytest.raises(ValueError, match="step"):
            grid[::2, 0]

    def test_next_autoflow_and_growth(self):
        pres = Presentation()
        slide = pres.add_slide()
        grid = pres.add_grid(slide=slide, rows=1, cols=2)
        grid.next().add_text("first")
        grid.next().add_text("second")
        assert grid.rows == 1
        grid.next().add_text("third grows the grid")
        assert grid.rows == 2
        assert grid.cells[1][0].content is not None

    def test_cell_style_padding_and_fill(self):
        pres = Presentation()
        slide = pres.add_slide()
        grid = pres.add_grid(slide=slide, rows=1, cols=1, padding=0.0)
        before = len(slide.shapes)
        grid[0, 0].style(fill="lightgray", padding=5).add_text("card")
        # Background shape plus the text box were added
        assert len(slide.shapes) == before + 2
        # Padding shrinks the content area
        x, y, w, h = grid._cell_area(grid.cells[0][0])
        assert to_percent(x, SLIDE_W) == pytest.approx(5.0, abs=0.05)
        assert to_percent(w, SLIDE_W) == pytest.approx(90.0, abs=0.05)


class TestBulletsAndNotes:
    def test_add_bullets_levels(self):
        pres = Presentation()
        slide = pres.add_slide()
        shape = slide.add_bullets(["top", ("nested", 1), "top again"])
        paragraphs = shape.text_frame.paragraphs
        assert [p.text for p in paragraphs] == ["top", "nested", "top again"]
        assert [p.level for p in paragraphs] == [0, 1, 0]

    def test_add_bullets_plain(self):
        pres = Presentation()
        slide = pres.add_slide()
        shape = slide.add_bullets(["a", "b"], bullet=False)
        assert len(shape.text_frame.paragraphs) == 2

    def test_notes_roundtrip(self):
        pres = Presentation()
        slide = pres.add_slide()
        assert slide.notes == ""
        slide.notes = "remember the demo"
        assert slide.notes == "remember the demo"


class TestMultiSeriesCharts:
    def test_dataframe_multi_series(self):
        pres = Presentation()
        slide = pres.add_slide()
        df = pd.DataFrame({"Q": ["Q1", "Q2"], "Rev": [1, 2], "Cost": [3, 4]})
        chart = slide.add_chart(data=df, category_column="Q", value_columns=["Rev", "Cost"])
        assert [s.name for s in chart.plots[0].series] == ["Rev", "Cost"]

    def test_list_data_multi_series(self):
        pres = Presentation()
        slide = pres.add_slide()
        data = [["Q", "Rev", "Cost"], ["Q1", 1, 3], ["Q2", 2, 4]]
        chart = slide.add_chart(data=data, value_columns=["Rev", "Cost"])
        assert [s.name for s in chart.plots[0].series] == ["Rev", "Cost"]

    def test_series_length_mismatch_raises(self):
        from easypptx.chart import Chart

        pres = Presentation()
        slide = pres.add_slide()
        with pytest.raises(ValueError, match="same length"):
            Chart(slide).add("column", ["a", "b"], series={"s": [1]})


class TestStylesAndThemes:
    def test_text_style_applies(self):
        pres = Presentation()
        slide = pres.add_slide()
        heading = TextStyle(font_size=28, font_bold=True, color="red")
        shape = slide.add_text("styled", style=heading)
        p = shape.text_frame.paragraphs[0]
        assert p.font.size.pt == 28
        assert p.font.bold is True

    def test_explicit_beats_style(self):
        pres = Presentation()
        slide = pres.add_slide()
        shape = slide.add_text("x", font_size=40, style=TextStyle(font_size=12))
        assert shape.text_frame.paragraphs[0].font.size.pt == 40

    def test_table_and_chart_styles(self):
        pres = Presentation()
        slide = pres.add_slide()
        table = slide.add_table([["H", "V"], ["a", 1]], style=TableStyle(has_header=False))
        assert table.cell(0, 0).text_frame.paragraphs[0].font.bold is None
        chart = slide.add_chart(data=[["C", "V"], ["a", 1]], style=ChartStyle(chart_type="pie", has_legend=False))
        assert chart.has_legend is False

    def test_builtin_theme_sets_background_and_text(self):
        pres = Presentation(theme="dark")
        slide = pres.add_slide(title="T")
        # Theme body color cascades into plain text via template defaults
        shape = slide.add_text("body")
        from easypptx.common import COLORS

        assert shape.text_frame.paragraphs[0].font.color.rgb == COLORS["white"]

    def test_unknown_theme_raises(self):
        with pytest.raises(ValueError, match="Unknown theme"):
            Presentation(theme="vaporwave")

    def test_custom_theme(self):
        theme = Theme(bg_color=(1, 2, 3), body=TextStyle(color="yellow"))
        pres = Presentation(theme=theme)
        assert pres.default_bg_color == (1, 2, 3)


class TestMarkdown:
    def test_full_document(self, tmp_path):
        image = tmp_path / "chart.png"
        PILImage.new("RGB", (80, 40), "blue").save(image)
        md = f"""---
theme: dark
---

# Deck Title

The subtitle

## Bullets <!-- notes: my notes -->

- one
  - one point five
- two

## Media

![chart]({image.name})

| A | B |
|---|---|
| 1 | 2 |

## Code

```
print("hi")
```

---

Closing words.
"""
        path = tmp_path / "deck.md"
        path.write_text(md)

        pres = Presentation.from_markdown(path)
        assert len(pres.slides) == 5
        assert pres.slides[1].notes == "my notes"
        # Relative image resolved against the md file's directory
        out = tmp_path / "deck.pptx"
        pres.save(out)
        assert out.exists()

    def test_columns_layout(self):
        md = """## Two Columns

::: columns
- left

- right
:::
"""
        pres = Presentation.from_markdown(md)
        slide = pres.slides[0]
        # Title + two bullet boxes side by side
        boxes = [s for s in slide.shapes if s.has_text_frame]
        assert len(boxes) == 3
        left, right = sorted(boxes[1:], key=lambda s: s.left)
        assert left.left < right.left
        assert left.top == right.top

    def test_theme_override_beats_frontmatter(self):
        md = "---\ntheme: dark\n---\n\n## S\n\ntext\n"
        pres = Presentation.from_markdown(md, theme="light")
        assert pres.theme is not None
        assert pres.theme.name == "light"

    def test_missing_file_raises(self):
        with pytest.raises(FileNotFoundError):
            Presentation.from_markdown("no_such_deck.md")


class TestNumericPercent:
    def test_bare_numbers_are_percent(self):
        pres = Presentation()
        slide = pres.add_slide()
        shape = slide.add_text("q", x=50, y=0, width=50, height=10)
        assert shape.left == pytest.approx(pres.pptx_presentation.slide_width / 2, abs=10)

    def test_in_helper_is_inches(self):
        pres = Presentation()
        slide = pres.add_slide()
        shape = slide.add_text("q", x=in_(1.0), y=in_(0.5), width=in_(2.0), height=in_(1.0))
        assert shape.left == 914400
        assert shape.top == 457200
