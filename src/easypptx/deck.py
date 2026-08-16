"""Fluent deck builder for EasyPPTX.

Build a whole presentation as one readable chain — no coordinates needed:

    from easypptx import Deck

    (Deck(theme="dark", footer="Acme Corp")
        .title_slide("Q3 Review", subtitle="Finance")
        .slide("Highlights", kicker="Q3 FINANCIALS", notes="keep it short")
        .stats([("+12%", "Revenue growth", "+3pt"), ("$1.6M", "Q4 revenue")])
        .bullets(["Revenue +12%", ("EMEA strongest", 1)])
        .chart(df, kind="column", value_columns=["Rev", "Cost"],
               emphasize="Rev", headline="Revenue accelerated every quarter")
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

import copy
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
_UNTITLED_TOP = 8
_CONTENT_BOTTOM = 92  # leaves room for the footer line
_MARGIN_X = 5  # percent of slide width
_CONTENT_WIDTH = 100 - 2 * _MARGIN_X
_GAP = 2.5
_MIN_FLEX_HEIGHT = 30  # charts/images need at least this much to read well
_BODY_FONT = 18
_LINE_FACTOR = 1.35  # line height + paragraph spacing, relative to font size
_STATS_HEIGHT = 22
_HEADLINE_HEIGHT = 7


@dataclass
class _Block:
    """One content block queued for a slide."""

    kind: str  # text | bullets | chart | table | image | pyplot | stats | compare | tap
    payload: dict[str, Any] = field(default_factory=dict)
    kwargs: dict[str, Any] = field(default_factory=dict)

    @property
    def positioned(self) -> bool:
        """Blocks with explicit geometry are excluded from automatic layout."""
        return any(k in self.kwargs for k in ("x", "y", "width", "height"))

    @property
    def flexible(self) -> bool:
        """Flexible blocks expand to fill the space fixed blocks leave over."""
        return self.kind in ("chart", "image", "pyplot", "compare")


@dataclass
class _SlideSpec:
    """One planned slide."""

    kind: str = "content"  # "content" | "title" | "section"
    title: str | None = None
    subtitle: str | None = None
    kicker: str | None = None
    notes: str | None = None
    paginate: bool = True
    blocks: list[_Block] = field(default_factory=list)


def _content_width_inches(pres: Presentation) -> float:
    return pres._slide_width_emu / 914400 * _CONTENT_WIDTH / 100


def _bullet_item_height(item: Any, content_width_in: float) -> float:
    """Height of one bullet item, percent of slide height (shared estimator)."""
    text, level = item if isinstance(item, tuple) else (item, 0)
    level = max(0, min(4, int(level)))
    size = max(10, _BODY_FONT - 2 * level)
    lines = estimate_lines(str(text), size, content_width_in * (1 - 0.04 * level))
    return lines * size * _LINE_FACTOR / 72 / 7.5 * 100


def _estimate_fixed_height(block: _Block, content_width_in: float) -> float:
    """Natural height of a fixed block, in percent of slide height."""
    if block.kind == "bullets":
        return max(6.0, sum(_bullet_item_height(i, content_width_in) for i in block.payload["items"]))
    if block.kind == "text":
        lines = estimate_lines(block.payload["text"], _BODY_FONT, content_width_in)
        return max(6.0, lines * _BODY_FONT * _LINE_FACTOR / 72 / 7.5 * 100)
    if block.kind == "table":
        return max(12.0, len(block.payload["rows"]) * 6.0)
    if block.kind == "stats":
        return _STATS_HEIGHT
    return 0.0  # tap and flexible blocks contribute no fixed height


def _flex_extra(block: _Block) -> float:
    """Extra fixed height a flexible block carries (e.g. a chart headline)."""
    return _HEADLINE_HEIGHT if block.kind == "chart" and block.payload.get("headline") else 0.0


def _split_bullets(block: _Block, available: float, content_width_in: float) -> tuple[_Block | None, _Block | None]:
    """Split a bullets block so the first part fits in the available height.

    Returns (head, rest); head is None when not even the first item fits.
    """
    items = block.payload["items"]
    kept: list[Any] = []
    used = 0.0
    for i, item in enumerate(items):
        h = _bullet_item_height(item, content_width_in)
        if used + h > available:
            if not kept:
                return None, block
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
        footer: str | None = None,
        page_numbers: bool = True,
    ) -> None:
        """Initialize a deck builder.

        Args:
            theme: Theme name or Theme instance for a new presentation (default: None)
            template_toml: TOML template path for a new presentation (default: None)
            aspect_ratio: Aspect ratio for a new presentation (default: "16:9")
            presentation: Build onto an existing Presentation instead of
                creating one; theme/template arguments are then ignored (default: None)
            footer: Footer text on every content slide; falls back to the
                theme's footer field (default: None)
            page_numbers: Show page numbers on themed content slides (default: True)
        """
        if presentation is not None:
            self._pres = presentation
        else:
            from easypptx.presentation import Presentation

            self._pres = Presentation(aspect_ratio=aspect_ratio, theme=theme, template_toml=template_toml)
        if footer is None and self._pres.theme is not None:
            footer = self._pres.theme.footer
        self._footer = footer
        self._page_numbers = page_numbers
        self._specs: list[_SlideSpec] = []
        self._finalized = False

    # ----- slide starters -------------------------------------------------

    def title_slide(self, title: str, subtitle: str | None = None, notes: str | None = None) -> Self:
        """Start the deck with a composed title slide.

        Args:
            title: The presentation title
            subtitle: Optional subtitle line (default: None)
            notes: Speaker notes for the slide (default: None)
        """
        self._check_open()
        self._specs.append(_SlideSpec(kind="title", title=title, subtitle=subtitle, notes=notes))
        return self

    def slide(
        self,
        title: str | None = None,
        *,
        kicker: str | None = None,
        notes: str | None = None,
        paginate: bool = True,
    ) -> Self:
        """Start a new content slide; following content calls land on it.

        Args:
            title: Slide title (default: None)
            kicker: Small label above the title, e.g. a section or topic tag
                (default: None)
            notes: Speaker notes (default: None)
            paginate: Spill overflowing content onto "(cont.)" slides;
                False compresses content to fit instead (default: True)
        """
        self._check_open()
        self._specs.append(_SlideSpec(title=title, kicker=kicker, notes=notes, paginate=paginate))
        return self

    def section(self, title: str, notes: str | None = None) -> Self:
        """Add a section-divider slide with a large title and accent rule.

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
            items: Strings or (text, level) tuples with level 0-4
            **kwargs: Forwarded to Slide.add_bullets
        """
        if not isinstance(items, list | tuple) or not items:
            raise TypeError("bullets() needs a non-empty list of items")
        for item in items:
            if isinstance(item, tuple):
                if len(item) != 2 or not isinstance(item[1], int) or not 0 <= item[1] <= 4:
                    raise TypeError(f"Bullet tuples must be (text, level) with level 0-4, got {item!r}")
            elif not isinstance(item, str):
                raise TypeError(f"Bullet items must be strings or (text, level) tuples, got {type(item).__name__}")
        self._current().blocks.append(_Block("bullets", {"items": list(items)}, kwargs))
        return self

    def chart(self, data: Any = None, kind: str | None = None, headline: str | None = None, **kwargs: Any) -> Self:
        """Add a chart to the current slide.

        Args:
            data: Any shape Slide.add_chart accepts (default: None with
                explicit categories=/values= in kwargs)
            kind: Chart type, native or matplotlib-backed (default: "column")
            headline: A message line rendered above the chart — say the
                takeaway, not the topic (default: None)
            **kwargs: Forwarded to Slide.add_chart (including emphasize=)
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
        # Canonicalize aliases so eager validation matches render behavior
        if "value_column" in kwargs and "value_columns" not in kwargs:
            kwargs["value_columns"] = kwargs.pop("value_column")
        explicit_categories = kwargs.get("categories")
        explicit_values = kwargs.get("values")
        if explicit_categories is not None and explicit_values is not None:
            if len(explicit_categories) != len(explicit_values):
                raise ValueError("categories and values must have the same length")
        elif data is None:
            raise ValueError("chart() needs data= or both categories= and values=")
        if data is not None:
            # Validate now and snapshot so later mutation can't break rendering
            normalize_chart_data(
                data,
                kwargs.get("category_column"),
                kwargs.get("value_columns"),
                categories=explicit_categories,
                columns=kwargs.get("columns"),
            )
            data = copy.deepcopy(data)
        self._current().blocks.append(
            _Block("chart", {"data": data, "chart_type": chart_type, "headline": headline}, kwargs)
        )
        return self

    def table(self, data: Any, **kwargs: Any) -> Self:
        """Add a table to the current slide.

        Args:
            data: Any shape Slide.add_table accepts
            **kwargs: Forwarded to Slide.add_table
        """
        from easypptx.data import normalize_table_rows

        rows = normalize_table_rows(data, columns=kwargs.get("columns"))
        if not rows:
            raise ValueError("table() needs at least one row of data")
        width = len(rows[0])
        if width == 0 or any(len(r) != width for r in rows):
            raise ValueError("table() needs rectangular, non-empty rows")
        kwargs.pop("columns", None)  # already applied during normalization
        self._current().blocks.append(_Block("table", {"rows": copy.deepcopy(rows)}, kwargs))
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

    def stats(self, items: list[tuple | dict], **kwargs: Any) -> Self:
        """Add a row of KPI stat tiles: big numbers with labels.

        Args:
            items: One to five (value, label) or (value, label, delta)
                tuples, or dicts with value/label/delta keys. Deltas
                starting with "-" render in the negative tone
            **kwargs: x/y/width/height for explicit placement

        Examples:
            ```python
            deck.stats([("+12%", "Revenue growth", "+3pt vs Q2"),
                        ("$1.6M", "Q4 revenue"),
                        ("38", "New logos", "-4 vs plan")])
            ```
        """
        if not isinstance(items, list | tuple) or not 1 <= len(items) <= 5:
            raise ValueError("stats() needs between 1 and 5 items")
        normalized: list[dict[str, str | None]] = []
        for item in items:
            if isinstance(item, dict):
                value, label, delta = item.get("value"), item.get("label"), item.get("delta")
            elif isinstance(item, tuple) and len(item) in (2, 3):
                value, label = item[0], item[1]
                delta = item[2] if len(item) == 3 else None
            else:
                raise TypeError(f"stats items must be (value, label[, delta]) or dicts, got {item!r}")
            if value is None or label is None:
                raise ValueError(f"stats items need value and label, got {item!r}")
            normalized.append({
                "value": str(value),
                "label": str(label),
                "delta": None if delta is None else str(delta),
            })
        self._current().blocks.append(_Block("stats", {"items": normalized}, kwargs))
        return self

    def compare(self, left: tuple[str, list], right: tuple[str, list], **kwargs: Any) -> Self:
        """Add two side-by-side cards, each with a heading and bullet list.

        Args:
            left: (heading, bullet items) for the left card
            right: (heading, bullet items) for the right card
            **kwargs: x/y/width/height for explicit placement

        Examples:
            ```python
            deck.compare(("Now", ["Manual exports", "Weekly cadence"]),
                         ("With rollout", ["Automated", "Daily"]))
            ```
        """
        for name, side in (("left", left), ("right", right)):
            if not (isinstance(side, tuple) and len(side) == 2 and isinstance(side[0], str)):
                raise TypeError(f"compare() {name} side must be (heading, items)")
            if not isinstance(side[1], list | tuple) or not side[1]:
                raise TypeError(f"compare() {name} side needs a non-empty items list")
        self._current().blocks.append(_Block("compare", {"left": left, "right": right}, kwargs))
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

        Taps run after the slide's other content has rendered.

        Args:
            fn: Called with the completed Slide
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
        specs = self._paginated()
        total = len(specs)
        for number, spec in enumerate(specs, start=1):
            self._render(spec, number, total)
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

    def _available_height(self, spec: _SlideSpec) -> float:
        return _CONTENT_BOTTOM - (_CONTENT_TOP if spec.title else _UNTITLED_TOP)

    def _paginated(self) -> list[_SlideSpec]:
        """Split content slides whose blocks overflow the content area."""
        content_width_in = _content_width_inches(self._pres)
        out: list[_SlideSpec] = []
        for spec in self._specs:
            if spec.kind != "content" or not spec.paginate:
                out.append(spec)
                continue

            available = self._available_height(spec)
            pending = list(spec.blocks)
            page = 0
            while True:
                page_spec = _SlideSpec(
                    title=spec.title if page == 0 else f"{spec.title} (cont.)" if spec.title else None,
                    kicker=spec.kicker if page == 0 else None,
                    notes=spec.notes if page == 0 else None,
                    paginate=spec.paginate,
                )
                fixed_sum = 0.0
                flex_sum = 0.0
                auto_count = 0

                while pending:
                    block = pending[0]
                    if block.positioned or block.kind == "tap":
                        page_spec.blocks.append(pending.pop(0))
                        continue
                    if block.flexible:
                        need = _MIN_FLEX_HEIGHT + _flex_extra(block)
                        gaps = _GAP * max(0, auto_count)  # one more gap when adding to non-empty page
                        if auto_count and fixed_sum + flex_sum + need + gaps > available:
                            break
                        flex_sum += need
                        auto_count += 1
                        page_spec.blocks.append(pending.pop(0))
                        continue
                    height = _estimate_fixed_height(block, content_width_in)
                    gaps = _GAP * max(0, auto_count)
                    if fixed_sum + flex_sum + height + gaps <= available:
                        fixed_sum += height
                        auto_count += 1
                        page_spec.blocks.append(pending.pop(0))
                        continue
                    # Overflow: try to split bullets to fill the rest of the page
                    if block.kind == "bullets":
                        remaining = available - fixed_sum - flex_sum - gaps
                        head, rest = _split_bullets(block, remaining, content_width_in)
                        if head is not None and rest is not None:
                            page_spec.blocks.append(head)
                            pending[0] = rest
                            auto_count += 1
                            break
                        if head is None and auto_count:
                            break  # nothing fits here; next page gets a fresh try
                        if head is not None and rest is None:
                            # Estimator disagreement at the boundary: place it
                            page_spec.blocks.append(pending.pop(0))
                            fixed_sum += height
                            auto_count += 1
                            continue
                    if not auto_count:
                        # A lone oversized block: place it and let rendering compress
                        fixed_sum += height
                        auto_count += 1
                        page_spec.blocks.append(pending.pop(0))
                    break
                out.append(page_spec)
                if not pending:
                    break
                page += 1
        return out

    # ----- rendering ------------------------------------------------------

    def _theme_colors(self) -> dict[str, Any]:
        from easypptx.styles import card_tone, muted_tone, theme_rgb

        theme = self._pres.theme
        if theme is None:
            return {"accent": None, "card": None, "muted": None, "body": None, "bg": None, "theme": None}
        return {
            "theme": theme,
            "accent": theme_rgb(theme.accent_color),
            "card": card_tone(theme),
            "muted": muted_tone(theme),
            "body": theme_rgb(theme.body.color),
            "bg": theme_rgb(theme.bg_color),
        }

    def _render(self, spec: _SlideSpec, number: int, total: int) -> None:
        pres = self._pres
        colors = self._theme_colors()
        accent = colors["accent"]

        if spec.kind == "title":
            slide = pres.add_slide()
            self._render_title_slide(slide, spec, colors)
        elif spec.kind == "section":
            slide = pres.add_slide()
            if accent is not None:
                rule = slide.add_shape(x=8, y=36, width=7, height=1.1, fill_color=accent)
                rule.line.fill.background()
            slide.add_text(
                spec.title or "",
                x=8,
                y=39,
                width=84,
                height=16,
                font_size=36,
                font_bold=True,
                align="left",
                vertical="middle",
            )
        else:
            if spec.kicker:
                slide = pres.add_slide()
                self._render_kicker_title(slide, spec, colors)
            else:
                slide = pres.add_slide(title=spec.title) if spec.title else pres.add_slide()
            self._render_blocks(slide, spec, colors)
            self._render_footer(slide, colors, number, total)

        if spec.notes:
            slide.notes = spec.notes

    def _render_title_slide(self, slide: Slide, spec: _SlideSpec, colors: dict) -> None:
        accent, card = colors["accent"], colors["card"]
        if card is not None:
            panel = slide.add_shape(x=68, y=0, width=32, height=100, fill_color=card)
            panel.line.fill.background()
        if accent is not None:
            edge = slide.add_shape(x=67.4, y=0, width=0.6, height=100, fill_color=accent)
            edge.line.fill.background()
            bar = slide.add_shape(x=8, y=52, width=9, height=1.1, fill_color=accent)
            bar.line.fill.background()
        slide.add_text(
            spec.title or "",
            x=8,
            y=34,
            width=56,
            height=16,
            font_size=40,
            font_bold=True,
            align="left",
            vertical="middle",
        )
        if spec.subtitle:
            slide.add_text(
                spec.subtitle, x=8, y=56, width=56, height=8, font_size=17, align="left", color=colors["muted"]
            )

    def _render_kicker_title(self, slide: Slide, spec: _SlideSpec, colors: dict) -> None:
        accent = colors["accent"]
        theme = colors["theme"]
        slide.add_text(
            spec.kicker or "",
            x=5,
            y=3.5,
            width=90,
            height=4.5,
            font_size=12,
            font_bold=True,
            color=accent,
            align="left",
            fit="none",
        )
        if spec.title:
            slide.add_text(
                spec.title,
                x=5,
                y=8,
                width=90,
                height=9,
                font_size=30,
                font_bold=True,
                align="left",
                color=theme.title.color if theme else None,
            )
        if accent is not None and theme is not None and theme.title_accent:
            bar = slide.add_shape(x=5, y=17.4, width=7, height=0.9, fill_color=accent)
            bar.line.fill.background()

    def _render_footer(self, slide: Slide, colors: dict, number: int, total: int) -> None:
        muted = colors["muted"]
        if muted is None:
            return
        if self._footer:
            slide.add_text(self._footer, x=5, y=95, width=60, height=4, font_size=9, color=muted, fit="none")
        if self._page_numbers:
            slide.add_text(
                f"{number} / {total}",
                x=85,
                y=95,
                width=10,
                height=4,
                font_size=9,
                color=muted,
                align="right",
                fit="none",
            )

    def _render_blocks(self, slide: Slide, spec: _SlideSpec, colors: dict) -> None:
        content_width_in = _content_width_inches(self._pres)
        top = _CONTENT_TOP if spec.title else _UNTITLED_TOP
        available = _CONTENT_BOTTOM - top

        auto = [b for b in spec.blocks if not b.positioned and b.kind != "tap"]
        fixed_heights = {id(b): _estimate_fixed_height(b, content_width_in) for b in auto if not b.flexible}
        flex_blocks = [b for b in auto if b.flexible]
        flex_extras = sum(_flex_extra(b) for b in flex_blocks)

        gap = _GAP
        gaps = gap * max(0, len(auto) - 1)
        fixed_sum = sum(fixed_heights.values())
        needed = fixed_sum + len(flex_blocks) * _MIN_FLEX_HEIGHT + flex_extras + gaps

        scale = 1.0
        if needed > available:
            # Compress the entire budget (fixed, flex, gaps) proportionally
            scale = available / needed
            gap *= scale
            gaps = gap * max(0, len(auto) - 1)
        flex_height = 0.0
        if flex_blocks:
            leftover = available - fixed_sum * scale - flex_extras * scale - gaps
            flex_height = max(_MIN_FLEX_HEIGHT * scale, leftover / len(flex_blocks))

        taps: list[_Block] = []
        y = float(top)
        for block in spec.blocks:
            if block.kind == "tap":
                taps.append(block)
                continue
            if block.positioned:
                self._render_one(slide, block, None, None, colors)
                continue
            height = flex_height + _flex_extra(block) * scale if block.flexible else fixed_heights[id(block)] * scale
            self._render_one(slide, block, y, height, colors)
            y += height + gap

        # Taps run after the slide's content, in their chained order
        for block in taps:
            block.payload["fn"](slide)

    def _render_one(self, slide: Slide, block: _Block, y: float | None, height: float | None, colors: dict) -> None:
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
            headline = block.payload.get("headline")
            if headline:
                hx = kwargs.get("x", _MARGIN_X)
                hy = kwargs.get("y", _CONTENT_TOP)
                hw = kwargs.get("width", _CONTENT_WIDTH)
                slide.add_text(
                    headline, x=hx, y=hy, width=hw, height=_HEADLINE_HEIGHT - 1, font_size=16, font_bold=True
                )
                if not isinstance(hy, str):
                    kwargs["y"] = float(hy) + _HEADLINE_HEIGHT
                if kwargs.get("height") is not None and not isinstance(kwargs["height"], str):
                    kwargs["height"] = max(10.0, float(kwargs["height"]) - _HEADLINE_HEIGHT)
            slide.add_chart(data=block.payload["data"], chart_type=block.payload["chart_type"], **kwargs)
        elif block.kind == "table":
            slide.add_table(block.payload["rows"], **kwargs)
        elif block.kind == "image":
            from easypptx.image import Image

            Image(slide).add(block.payload["path"], **kwargs)
        elif block.kind == "pyplot":
            slide.add_pyplot(block.payload["figure"], **kwargs)
        elif block.kind == "stats":
            self._render_stats(slide, block, kwargs, colors)
        elif block.kind == "compare":
            self._render_compare(slide, block, kwargs, colors)

    def _render_stats(self, slide: Slide, block: _Block, kwargs: dict, colors: dict) -> None:
        items = block.payload["items"]
        x0 = float(kwargs.get("x", _MARGIN_X))
        y0 = float(kwargs.get("y", _CONTENT_TOP))
        width = float(kwargs.get("width", _CONTENT_WIDTH))
        height = float(kwargs.get("height") or _STATS_HEIGHT)
        tile_gap = 2.0
        tile_w = (width - tile_gap * (len(items) - 1)) / len(items)
        accent, card, muted = colors["accent"], colors["card"], colors["muted"]

        for i, item in enumerate(items):
            tx = x0 + i * (tile_w + tile_gap)
            if card is not None:
                tile = slide.add_shape(
                    shape_type="ROUNDED_RECTANGLE", x=tx, y=y0, width=tile_w, height=height, fill_color=card
                )
                tile.line.fill.background()
            slide.add_text(
                item["value"],
                x=tx + 2,
                y=y0 + 2,
                width=tile_w - 4,
                height=height * 0.45,
                font_size=30,
                font_bold=True,
                color=accent,
                fit="none",
            )
            slide.add_text(
                item["label"].upper(),
                x=tx + 2,
                y=y0 + height * 0.52,
                width=tile_w - 4,
                height=height * 0.2,
                font_size=10,
                font_bold=True,
                color=muted,
                fit="none",
            )
            if item["delta"]:
                positive = not item["delta"].startswith("-")
                delta_color = (0x3E, 0x9B, 0x63) if positive else (0xC7, 0x4A, 0x4A)
                slide.add_text(
                    item["delta"],
                    x=tx + 2,
                    y=y0 + height * 0.73,
                    width=tile_w - 4,
                    height=height * 0.22,
                    font_size=11,
                    color=delta_color,
                    fit="none",
                )

    def _render_compare(self, slide: Slide, block: _Block, kwargs: dict, colors: dict) -> None:
        x0 = float(kwargs.get("x", _MARGIN_X))
        y0 = float(kwargs.get("y", _CONTENT_TOP))
        width = float(kwargs.get("width", _CONTENT_WIDTH))
        height = float(kwargs.get("height") or 60.0)
        card_gap = 2.5
        card_w = (width - card_gap) / 2
        accent, card = colors["accent"], colors["card"]

        for i, (heading, items) in enumerate((block.payload["left"], block.payload["right"])):
            cx = x0 + i * (card_w + card_gap)
            if card is not None:
                panel = slide.add_shape(
                    shape_type="ROUNDED_RECTANGLE", x=cx, y=y0, width=card_w, height=height, fill_color=card
                )
                panel.line.fill.background()
            slide.add_text(
                heading, x=cx + 2.5, y=y0 + 2.5, width=card_w - 5, height=6, font_size=16, font_bold=True, color=accent
            )
            slide.add_bullets(list(items), x=cx + 2.5, y=y0 + 10, width=card_w - 5, height=height - 13)
