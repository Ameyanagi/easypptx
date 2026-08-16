# Markdown to Presentation

EasyPPTX can build a whole deck from a markdown document. Each section of the document becomes a slide, so you can draft a presentation in any text editor and convert it in one call:

```python
from easypptx import Presentation

pres = Presentation.from_markdown("deck.md")
pres.save("deck.pptx")
```

The module-level function `easypptx.from_markdown` is equivalent:

```python
from easypptx import from_markdown

pres = from_markdown("deck.md")
```

`source` can be a path to a `.md` file or the markdown text itself:

```python
pres = Presentation.from_markdown("""
# Quarterly Review

Q3 2026

## Highlights

- Revenue up 12%
  - EMEA strongest region
- Costs down 3%
""")
```

## Supported Syntax

| Markdown | Result |
| --- | --- |
| `--- key: value ---` frontmatter | Deck settings: `theme`, `template`, `aspect_ratio` |
| `# Heading` (first heading) | Title slide; the next paragraph becomes the subtitle |
| `## Heading` | New slide with that title |
| `---` on its own line | Forced slide break |
| `-` / `*` / `1.` lists | Bullets; 2-space indentation nests (levels 0-4) |
| `![alt](path)` | Image; relative paths resolve against the markdown file's directory |
| GFM `\|` tables | PowerPoint tables (first row is the header) |
| Fenced code blocks (```` ``` ````) | Monospaced code text box |
| `<!-- notes: ... -->` | Speaker notes for the slide |
| `::: columns` ... `:::` | Blocks inside are laid out side by side |

Basic inline markdown (`**bold**`, `*italic*`, `` `code` ``, links) is stripped to plain text. Content blocks on a slide share the vertical space automatically, weighted by their size.

## Full Example

`deck.md`:

````markdown
---
theme: dark
aspect_ratio: 16:9
---

# Quarterly Review

Q3 2026 - Sales Team

## Highlights

- Revenue up 12%
  - EMEA strongest region
- Costs down 3%
- New logo customers: 14

<!-- notes: Lead with the EMEA story, then costs. -->

## Revenue by Region

| Region | Revenue | Growth |
| ------ | ------- | ------ |
| EMEA   | $2.1M   | +18%   |
| AMER   | $1.6M   | +9%    |
| APAC   | $0.8M   | +7%    |

## Side by Side

::: columns

![architecture](images/architecture.png)

- Simple deployment
- One binary
- No external services

:::

## How to Install

```python
pip install easypptx
```

---

Questions?
````

Convert it:

```python
from easypptx import Presentation

pres = Presentation.from_markdown("deck.md")
pres.save("review.pptx")
```

This produces a title slide, three content slides (bullets, a table, and a two-column layout), a code slide, and a closing slide from the forced `---` break. The `<!-- notes: ... -->` comment becomes speaker notes on the Highlights slide.

## Frontmatter

An optional frontmatter block at the top of the file sets deck-wide options:

```markdown
---
theme: dark
template: my_template.toml
aspect_ratio: 4:3
---
```

- `theme`: A built-in theme name (`light`, `dark`, `corporate`); see [Styling and Formatting](styling.md)
- `template`: Path to a TOML template; see [PowerPoint Templates](templates.md)
- `aspect_ratio`: One of `16:9` (default), `4:3`, `16:10`, `A4`, `LETTER`

## Function Reference

```python
Presentation.from_markdown(
    source,              # markdown text, or a path to a .md file
    theme=None,          # theme name or Theme instance; overrides frontmatter
    template_toml=None,  # TOML template path; overrides frontmatter
    base_dir=None,       # directory for relative image paths
                         # (default: the markdown file's directory, or the cwd)
)
```

Arguments passed to the function win over the frontmatter, so you can render the same document with different themes:

```python
light = Presentation.from_markdown("deck.md", theme="light")
dark = Presentation.from_markdown("deck.md", theme="dark")
```

The returned object is an ordinary `Presentation`: you can keep adding slides, restyle content, or tweak individual slides before saving.
