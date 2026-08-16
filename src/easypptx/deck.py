"""Fluent deck builder for EasyPPTX.

Build a whole presentation as one readable chain — no coordinates needed:

    from easypptx import Deck

    (Deck(theme="dark")
        .title_slide("Q3 Review", subtitle="Finance")
        .slide("Highlights", notes="keep it short")
        .bullets(["Revenue +12%", ("EMEA strongest", 1)])
        .chart(df, kind="column", value_columns=["Rev", "Cost"])
        .slide("Data")
        .table(df, shade_columns=["Sales"])
        .save("q3.pptx"))

Every call validates its arguments immediately (errors point at the call
site), but rendering happens at :meth:`save`/:meth:`build`, where the
layout engine stacks blocks with natural heights, lets charts and images
expand into the remaining space, and paginates overflow onto
"(cont.)" slides instead of shrinking text into illegibility.

Explicit ``x``/``y``/``width``/``height`` on any content call opts that
block out of automatic layout and places it exactly where you say.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Any, Self

from easypptx.textfit import estimate_lines

if TYPE_CHECKING:
    from collections.abc import Callable

    from easypptx.presentation import Presentation
    from easypptx.slide import Slide

# Layout constants (percent of slide height unless noted)
_CONTENT_TOP = 19  # below the title band and accent bar
_CONTENT_BOTTOM = 94
_MARGIN_X = 5  # percent of slide width
_CONTENT_WIDTH = 100 - 2 * _MARGIN_X
_GAP = 2.5
_MIN_FLEX_HEIGHT = 30  # charts/images need at least this much to read well
_BODY_FONT = 18
_LINE_FACTOR = 1.35  # line height + paragraph spacing, relative to font size


@dataclass
class _Block:
    """One content block queued for a slide."""

    kind: str  # "text" | "bullets" | "chart" | "table" | "image" | "pyplot" | "tap"
    payload: dict[str, Any] = field(default_factory=dict)
    kwargs: dict[str, Any] = field(default_factory=dict)

    @property
    def positioned(self) -> bool:
        """Blocks with explicit geometry are excluded from automatic layout."""
        return any(k in self.kwargs for k in ("x", "y", "width", "height"))

    @property
    def flexible(self) -> bool:
        """Flexible blocks expand to fill the space fixed blocks leave over."""
        return self.kind in ("chart", "image", "pyplot")


@dataclass
class _SlideSpec:
    """One planned slide."""

    kind: str = "content"  # "content" | "title" | "section"
    title: str | None = None
    subtitle: str | None = None
    notes: str | None = None
    paginate: bool = True
    blocks: list[_Block] = field(default_factory=list)


def _content_width_inches(pres: Presentation) -> float:
    return pres._slide_width_emu / 914400 * _CONTENT_WIDTH / 100


def _estimate_fixed_height(block: _Block, content_width_in: float) -> float:
    """Natural height of a fixed block, in percent of slide height."""
    if block.kind == "bullets":
        total = 0.0
        for item in block.payload["items"]:
            text, level = item if isinstance(item, tuple) else (item, 0)
            size = max(10, _BODY_FONT - 2 * int(level))
            lines = estimate_lines(str(text), size, content_width_in * (1 - 0.04 * int(level)))
            total += lines * size * _LINE_FACTOR / 72 / 7.5 * 100
        return max(6.0, total)
    if block.kind == "text":
        lines = estimate_lines(block.payload["text"], _BODY_FONT, content_width_in)
        return max(6.0, lines * _BODY_FONT * _LINE_FACTOR / 72 / 7.5 * 100)
    if block.kind == "table":
        return max(12.0, len(block.payload["rows"]) * 6.0)
    return 0.0  # tap and flexible blocks contribute no fixed height


def _split_bullets(block: _Block, available: float, content_width_in: float) -> tuple[_Block, _Block | None]:
    """Split a bullets block so the first part fits in the available height."""
    items = block.payload["items"]
    kept: list[Any] = []
    used = 0.0
    for i, item in enumerate(items):
        text, level = item if isinstance(item, tuple) else (item, 0)
        size = max(10, _BODY_FONT - 2 * int(level))
        h = estimate_lines(str(text), size, content_width_in) * size * _LINE_FACTOR / 72 / 7.5 * 100
        if kept and used + h > available:
            rest = _Block("bullets", {"items": items[i:]}, dict(block.kwargs))
            return _Block("bullets", {"items": kept}, dict(block.kwargs)), rest
        kept.append(item)
        used += h
    return block, None


class Deck:
    """A lazily-rendered, chainable presentation builder.

    Create standalone (``Deck(theme="dark")``) or from an existing
    presentation (``pres.deck()``). Every method returns the deck, so a
    whole presentation reads as one expression ending in :meth:`save`.
    """

    def __init__(
        self,
        theme: Any = None,
        template_toml: str | None = None,
        aspect_ratio: str = "16:9",
        presentation: Presentation | None = None,
    ) -> None:
        """Initialize a deck builder.

        Args:
            theme: Theme name or Theme instance for a new presentation (default: None)
            template_toml: TOML template path for a new presentation (default: None)
            aspect_ratio: Aspect ratio for a new presentation (default: "16:9")
            presentation: Build onto an existing Presentation instead of
                creating one; theme/template arguments are then ignored (default: None)
        """
        if presentation is not None:
            self._pres = presentation
        else:
            from easypptx.presentation import Presentation

            self._pres = Presentation(aspect_ratio=aspect_ratio, theme=theme, template_toml=template_toml)
        self._specs: list[_SlideSpec] = []
        self._finalized = False

    # ----- slide starters -------------------------------------------------

    def title_slide(self, title: str, subtitle: str | None = None, notes: str | None = None) -> Self:
        """Start the deck with a title slide.

        Args:
            title: The presentation title
            subtitle: Optional subtitle line (default: None)
            notes: Speaker notes for the slide (default: None)
        """
        self._check_open()
        self._specs.append(_SlideSpec(kind="title", title=title, subtitle=subtitle, notes=notes))
        return self

    def slide(self, title: str | None = None, *, notes: str | None = None, paginate: bool = True) -> Self:
        """Start a new content slide; following content calls land on it.

        Args:
            title: Slide title (default: None)
            notes: Speaker notes (default: None)
            paginate: Spill overflowing content onto "(cont.)" slides;
                False shrinks text to fit instead (default: True)
        """
        self._check_open()
        self._specs.append(_SlideSpec(title=title, notes=notes, paginate=paginate))
        return self

    def section(self, title: str, notes: str | None = None) -> Self:
        """Add a section-divider slide with a large centered title.

        Args:
            title: The section name
            notes: Speaker notes (default: None)
        """
        self._check_open()
        self._specs.append(_SlideSpec(kind="section", title=title, notes=notes))
        return self

    # ----- content --------------------------------------------------------

    def text(self, text: str, **kwargs: Any) -> Self:
        """Add a paragraph to the current slide.

        Args:
            text: The paragraph text
            **kwargs: Forwarded to Slide.add_text; explicit x/y/width/height
                opt out of automatic layout
        """
        if not isinstance(text, str):
            raise TypeError(f"text must be a string, got {type(text).__name__}")
        self._current().blocks.append(_Block("text", {"text": text}, kwargs))
        return self

    def bullets(self, items: list[str | tuple[str, int]], **kwargs: Any) -> Self:
        """Add a bulleted list to the current slide.

        Args:
            items: Strings or (text, level) tuples, as in Slide.add_bullets
            **kwargs: Forwarded to Slide.add_bullets
        """
        if not isinstance(items, list | tuple) or not items:
            raise TypeError("bullets() needs a non-empty list of items")
        for item in items:
            if isinstance(item, tuple):
                if len(item) != 2 or not isinstance(item[1], int):
                    raise TypeError(f"Bullet tuples must be (text, level), got {item!r}")
            elif not isinstance(item, str):
                raise TypeError(f"Bullet items must be strings or (text, level) tuples, got {type(item).__name__}")
        self._current().blocks.append(_Block("bullets", {"items": list(items)}, kwargs))
        return self

    def chart(self, data: Any = None, kind: str | None = None, **kwargs: Any) -> Self:
        """Add a chart to the current slide.

        Args:
            data: Any shape Slide.add_chart accepts (default: None with
                explicit categories=/values= in kwargs)
            kind: Chart type, native or matplotlib-backed (default: "column")
            **kwargs: Forwarded to Slide.add_chart
        """
        from easypptx.chart import Chart
        from easypptx.data import normalize_chart_data
        from easypptx.plot_backend import SUPPORTED_TYPES

        chart_type = kind if kind is not None else kwargs.get("chart_type", "column")
        kwargs.pop("chart_type", None)
        if chart_type not in Chart.CHART_TYPES and chart_type not in SUPPORTED_TYPES:
            raise ValueError(
                f"Unknown chart type: {chart_type!r}. Choose from "
                f"{', '.join(sorted(set(Chart.CHART_TYPES) | SUPPORTED_TYPES))}"
            )
        if data is None and not ("categories" in kwargs and "values" in kwargs):
            raise ValueError("chart() needs data= or both categories= and values=")
        if data is not None:
            # Validate the data shape now so errors point at this call
            normalize_chart_data(
                data,
                kwargs.get("category_column"),
                kwargs.get("value_columns", kwargs.get("value_column")),
                categories=kwargs.get("categories"),
                columns=kwargs.get("columns"),
            )
        self._current().blocks.append(_Block("chart", {"data": data, "chart_type": chart_type}, kwargs))
        return self

    def table(self, data: Any, **kwargs: Any) -> Self:
        """Add a table to the current slide.

        Args:
            data: Any shape Slide.add_table accepts
            **kwargs: Forwarded to Slide.add_table
        """
        from easypptx.data import normalize_table_rows

        rows = normalize_table_rows(data, columns=kwargs.get("columns"))
        self._current().blocks.append(_Block("table", {"data": data, "rows": rows}, kwargs))
        return self

    def image(self, path: str | Path, **kwargs: Any) -> Self:
        """Add an image to the current slide.

        Args:
            path: Path to the image file (checked immediately)
            **kwargs: Forwarded to Image.add
        """
        if not Path(path).exists():
            raise FileNotFoundError(f"Image file not found: {path}")
        self._current().blocks.append(_Block("image", {"path": str(path)}, kwargs))
        return self

    def pyplot(self, figure: Any, **kwargs: Any) -> Self:
        """Add a matplotlib figure to the current slide.

        Args:
            figure: A matplotlib figure (rendered at save time)
            **kwargs: Forwarded to Slide.add_pyplot
        """
        if not hasattr(figure, "savefig"):
            raise TypeError("pyplot() needs a matplotlib figure (an object with savefig)")
        self._current().blocks.append(_Block("pyplot", {"figure": figure}, kwargs))
        return self

    def notes(self, text: str) -> Self:
        """Set (or extend) the current slide's speaker notes.

        Args:
            text: The notes text; appended on a new line if notes exist
        """
        spec = self._current()
        spec.notes = f"{spec.notes}\n{text}" if spec.notes else text
        return self

    def tap(self, fn: Callable[[Slide], None]) -> Self:
        """Escape hatch: run a callback with the rendered Slide at save time.

        Args:
            fn: Called with the Slide after the slide's other content renders
        """
        if not callable(fn):
            raise TypeError("tap() needs a callable")
        self._current().blocks.append(_Block("tap", {"fn": fn}))
        return self

    # ----- terminal -------------------------------------------------------

    def build(self) -> Presentation:
        """Render every queued slide and return the Presentation."""
        self._check_open()
        self._finalized = True
        for spec in self._paginated():
            self._render(spec)
        return self._pres

    def save(self, path: str | Path) -> Presentation:
        """Render the deck and save it.

        Args:
            path: Output .pptx path

        Returns:
            The rendered Presentation
        """
        pres = self.build()
        pres.save(path)
        return pres

    @property
    def presentation(self) -> Presentation:
        """The underlying Presentation (unrendered until build/save)."""
        return self._pres

    # ----- internals ------------------------------------------------------

    def _check_open(self) -> None:
        if self._finalized:
            raise RuntimeError("This deck was already rendered; create a new Deck to build another")

    def _current(self) -> _SlideSpec:
        self._check_open()
        if not self._specs or self._specs[-1].kind != "content":
            raise RuntimeError("Start a slide first: call .slide(...) before adding content")
        return self._specs[-1]

    def _paginated(self) -> list[_SlideSpec]:
        """Split content slides whose fixed blocks overflow the content area."""
        content_width_in = _content_width_inches(self._pres)
        available_total = _CONTENT_BOTTOM - _CONTENT_TOP
        out: list[_SlideSpec] = []
        for spec in self._specs:
            if spec.kind != "content" or not spec.paginate:
                out.append(spec)
                continue

            pending = list(spec.blocks)
            page = 0
            while True:
                page_spec = _SlideSpec(
                    title=spec.title if page == 0 else f"{spec.title} (cont.)" if spec.title else None,
                    notes=spec.notes if page == 0 else None,
                    paginate=spec.paginate,
                )
                used = 0.0
                flex_count = 0
                placed_auto = False
                while pending:
                    block = pending[0]
                    if block.positioned or block.kind == "tap":
                        page_spec.blocks.append(pending.pop(0))
                        continue
                    if block.flexible:
                        need = (flex_count + 1) * _MIN_FLEX_HEIGHT + flex_count * _GAP
                        if placed_auto and used + need > available_total:
                            break
                        flex_count += 1
                        placed_auto = True
                        page_spec.blocks.append(pending.pop(0))
                        continue
                    height = _estimate_fixed_height(block, content_width_in)
                    gap = _GAP if placed_auto else 0.0
                    remaining = available_total - used - flex_count * (_MIN_FLEX_HEIGHT + _GAP) - gap
                    if height <= remaining:
                        page_spec.blocks.append(pending.pop(0))
                        used += height + gap
                        placed_auto = True
                        continue
                    # Overflow: split bullets to fill the rest of the page
                    if block.kind == "bullets" and remaining > 10:
                        head, rest = _split_bullets(block, remaining, content_width_in)
                        if rest is not None:
                            page_spec.blocks.append(head)
                            pending[0] = rest
                            placed_auto = True
                            break
                    if not placed_auto:
                        # A lone unsplittable oversized block: place it and
                        # let rendering compress it
                        page_spec.blocks.append(pending.pop(0))
                        used += height + gap
                        placed_auto = True
                    break
                out.append(page_spec)
                if not pending:
                    break
                page += 1
        return out

    def _render(self, spec: _SlideSpec) -> None:
        pres = self._pres
        theme = pres.theme

        if spec.kind == "title":
            slide = pres.add_slide()
            slide.add_text(
                spec.title or "",
                x=8,
                y=32,
                width=84,
                height=16,
                font_size=40,
                font_bold=True,
                align="center",
                vertical="middle",
            )
            if theme is not None and theme.title_accent and theme.accent_color is not None:
                rule = slide.add_shape(x=44, y=50, width=12, height=0.9, fill_color=theme.accent_color)
                rule.line.fill.background()
            if spec.subtitle:
                slide.add_text(spec.subtitle, x=10, y=54, width=80, height=8, font_size=18, align="center")
        elif spec.kind == "section":
            slide = pres.add_slide()
            slide.add_text(
                spec.title or "",
                x=8,
                y=38,
                width=84,
                height=14,
                font_size=34,
                font_bold=True,
                align="center",
                vertical="middle",
            )
            if theme is not None and theme.title_accent and theme.accent_color is not None:
                rule = slide.add_shape(x=46, y=53, width=8, height=0.9, fill_color=theme.accent_color)
                rule.line.fill.background()
        else:
            slide = pres.add_slide(title=spec.title) if spec.title else pres.add_slide()
            self._render_blocks(slide, spec)

        if spec.notes:
            slide.notes = spec.notes

    def _render_blocks(self, slide: Slide, spec: _SlideSpec) -> None:
        content_width_in = _content_width_inches(self._pres)
        top = _CONTENT_TOP if spec.title else 8
        available = _CONTENT_BOTTOM - top

        auto = [b for b in spec.blocks if not b.positioned and b.kind != "tap"]
        fixed_heights = {id(b): _estimate_fixed_height(b, content_width_in) for b in auto if not b.flexible}
        flex_blocks = [b for b in auto if b.flexible]

        gaps = _GAP * max(0, len(auto) - 1)
        leftover = available - sum(fixed_heights.values()) - gaps
        flex_height = max(_MIN_FLEX_HEIGHT * 0.6, leftover / len(flex_blocks)) if flex_blocks else 0.0

        # Compress fixed blocks proportionally if the page still overflows
        # (always: pagination estimates can drift, and paginate=False relies on it)
        scale = 1.0
        needed = sum(fixed_heights.values()) + len(flex_blocks) * _MIN_FLEX_HEIGHT * 0.6 + gaps
        if needed > available:
            scale = max(
                0.35,
                (available - gaps - len(flex_blocks) * _MIN_FLEX_HEIGHT * 0.6) / max(sum(fixed_heights.values()), 1.0),
            )

        y = float(top)
        for block in spec.blocks:
            if block.kind == "tap":
                block.payload["fn"](slide)
                continue
            if block.positioned:
                self._render_one(slide, block, None, None)
                continue
            height = flex_height if block.flexible else fixed_heights[id(block)] * scale
            self._render_one(slide, block, y, height)
            y += height + _GAP

    def _render_one(self, slide: Slide, block: _Block, y: float | None, height: float | None) -> None:
        kwargs = dict(block.kwargs)
        if y is not None:
            kwargs.setdefault("x", _MARGIN_X)
            kwargs.setdefault("y", y)
            kwargs.setdefault("width", _CONTENT_WIDTH)
            kwargs.setdefault("height", height)

        if block.kind == "text":
            slide.add_text(block.payload["text"], **kwargs)
        elif block.kind == "bullets":
            slide.add_bullets(block.payload["items"], **kwargs)
        elif block.kind == "chart":
            slide.add_chart(data=block.payload["data"], chart_type=block.payload["chart_type"], **kwargs)
        elif block.kind == "table":
            slide.add_table(block.payload["data"], **kwargs)
        elif block.kind == "image":
            from easypptx.image import Image

            Image(slide).add(block.payload["path"], **kwargs)
        elif block.kind == "pyplot":
            slide.add_pyplot(block.payload["figure"], **kwargs)
