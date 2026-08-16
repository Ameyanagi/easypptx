# The Deck Builder

The fluent builder turns a whole presentation into one readable chain —
no coordinates required. It validates every call immediately (errors point
at your code) and renders lazily at `save()`, where the layout engine
stacks blocks with natural heights, lets charts and images expand into the
remaining space, and paginates overflow onto "(cont.)" slides.

```python
from easypptx import Deck

(Deck(theme="dark")
    .title_slide("Q3 Review", subtitle="Finance")
    .slide("Highlights", notes="keep it short")
    .bullets(["Revenue +12%", ("EMEA strongest", 1)])
    .chart(df, kind="column", value_columns=["Rev", "Cost"], show_values=True)
    .section("Details")
    .slide("Data")
    .table(df, shade_columns=["Sales"])
    .save("q3.pptx"))
```

Build onto an existing presentation with `pres.deck()`.

## Slide starters

| Method | Creates |
|---|---|
| `title_slide(title, subtitle=None, notes=None)` | Title slide with an accent rule |
| `slide(title=None, *, notes=None, paginate=True)` | Content slide; following content calls land here |
| `section(title, notes=None)` | Section divider with a large centered title |

## Content methods

`text`, `bullets`, `chart`, `table`, `image`, and `pyplot` mirror the
`Slide` methods and forward their keyword arguments. Three extras:

- `notes(text)` sets or appends the current slide's speaker notes.
- `tap(fn)` runs `fn(slide)` with the rendered `Slide` at save time — the
  escape hatch to the full low-level API.
- Passing explicit `x`/`y`/`width`/`height` to any content call opts that
  block out of automatic layout and places it exactly where you say.

## Layout and pagination

Blocks without explicit geometry are laid out automatically:

- **Fixed blocks** (text, bullets, tables) get their natural estimated
  height, using the same CJK-aware measurement as `fit="shrink"`.
- **Flexible blocks** (charts, images, matplotlib figures) split the
  remaining space between them, so a bullets-plus-chart slide gives the
  chart everything the bullets don't need.
- **Overflow paginates**: content that doesn't fit continues on a
  "Title (cont.)" slide, and long bullet lists split across pages at full
  size. Pass `paginate=False` on `slide()` to compress instead.

## Errors happen at the call site

The chain validates eagerly: unknown chart types, unusable data shapes,
missing image files, and content before any `slide()` raise immediately.
A deck that has been rendered (`build()`/`save()`) cannot be extended —
create a new one.
