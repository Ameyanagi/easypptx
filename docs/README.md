# EasyPPTX Documentation

This directory contains documentation for the EasyPPTX library.

## Documentation Structure

- `index.md` - Landing page with overview and quick start examples
- `features.md` - Comprehensive features overview
- `percentage_positioning.md` - Guide to percentage-based positioning
- `auto_alignment.md` - Guide to automatic object alignment
- `styling.md` - Styling and formatting options
- `templates.md` - Using PowerPoint templates
- `api_reference.md` - Detailed API reference
- `modules.md` - Auto-generated API documentation
- `CHANGELOG.md` - Record of changes and new features

## Building Documentation

The documentation uses MkDocs with the Material theme. To build the documentation, you need to install MkDocs and the required plugins:

```bash
pip install mkdocs mkdocs-material mkdocstrings
```

Then, you can build the documentation with:

```bash
mkdocs build
```

Or serve it locally for testing:

```bash
mkdocs serve
```

## Documentation Guidelines

When contributing to the documentation, please follow these guidelines:

1. Use Markdown for all documentation files
2. Include code examples for all features
3. Provide clear explanations of parameters and return values
4. Use descriptive headings and subheadings
5. Maintain consistent formatting and style
6. Update the CHANGELOG.md when adding or modifying features
