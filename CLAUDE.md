# CLAUDE.md - EasyPPTX Development Guide

This guide provides comprehensive instructions for working with the EasyPPTX codebase.

## Project Overview

EasyPPTX is a Python library for easily creating and manipulating PowerPoint presentations programmatically with simple APIs. It aims to provide intuitive interfaces that are straightforward enough for LLMs to use effectively.

## Build/Test/Run Commands

The project uses `make` commands for most common tasks. These are preferred for consistency:

```bash
# Install dependencies and pre-commit hooks
make install

# Run code quality checks (linting, type checking, etc.)
make check

# Run all tests with coverage
make test

# Build documentation and serve locally
make docs

# Build package
make build
```

For more granular control, you can also use direct `uv` commands:

```bash
# Install dependencies
uv pip install -e ".[dev]"

# Run specific tests
uv run pytest tests/test_specific.py

# Run with coverage
uv run pytest --cov=src/easypptx

# Run linting
uv run ruff check src/ tests/

# Run type checking
uv run mypy src/ tests/

# Build documentation
uv run mkdocs build

# List all available make commands
make help
```

## Development Approach

### Test-Driven Development (TDD)

1. Write tests first that define expected behavior
2. Implement code to satisfy the tests
3. Refactor while maintaining test coverage
4. Repeat for each new feature

### Error Handling Strategy

- Use explicit exception types defined in the codebase
- Provide clear error messages with troubleshooting suggestions
- Ensure graceful degradation when possible
- Log errors appropriately with context information
- For API functions, follow the pattern of raising specific exceptions with helpful messages

## Development Workflow

1. **Plan**: Review specifications in `specs/` directory
2. **Test**: Write unit tests for the planned feature
3. **Implement**: Write code to pass the tests
4. **Verify**: Run `make check` and `make test` to ensure code quality and test coverage
5. **Document**: Update docstrings and documentation
6. **Review**: Ensure code meets style guidelines and best practices
7. **Update Status**: Update implementation status in `.states/implementation_status.md`
8. **Archive**: Move completed specs to `ai_docs/completed/`

## Incremental Implementation

1. Start with core functionality (presentation creation, saving, loading)
2. Add slide manipulation capabilities
3. Implement text, image, and shape elements
4. Add table and chart support
5. Implement templating and styling features
6. Add responsive positioning and grid layout
7. Add advanced formatting options

## Key Libraries

### python-pptx

Primary library for PowerPoint manipulation. Key concepts:
- Presentation object - represents a PowerPoint file
- Slide objects - represents individual slides
- Shape objects - elements on slides (text boxes, images, etc.)
- PlaceholderFormat - working with placeholders in templates
- Table objects - creating and manipulating tables

Documentation: https://python-pptx.readthedocs.io/

### Pillow (PIL)

Used for image processing:
- Resizing and cropping images
- Converting between formats
- Adjusting image properties

### pandas

For data handling:
- Converting DataFrames to tables
- Data manipulation for charts

## Directory Structure

```
easypptx/
├── .devcontainer/       # Development container configuration
├── .github/             # GitHub workflows and templates
├── .states/             # Implementation status tracking
│   ├── checkpoints/     # Serialized intermediate states
│   ├── implementation_status.md  # Feature implementation status
│   └── progress/        # Development progress tracking
├── ai_docs/             # Knowledge repository for AI tools
│   └── completed/       # Archived completed specifications
├── docs/                # Documentation files for mkdocs
├── specs/               # Specifications and feature documents
├── src/                 # Source code
│   └── easypptx/        # Main package
├── tests/               # Test directory
├── CONTRIBUTING.md      # Contribution guidelines
├── CLAUDE.md            # This file - guidance for AI assistance
├── LICENSE              # MIT License
├── README.md            # Project overview
├── mkdocs.yml           # MkDocs configuration
└── pyproject.toml       # Project configuration and dependencies
```

## Code Style Guidelines

- Follow PEP 8 for code style
- Use Google-style docstrings for function and class documentation
- Keep functions focused and single-purpose
- Use type annotations throughout
- Prefer percentage-based positioning for responsive layouts
- Use the Grid layout system for complex positioning arrangements
- Follow consistent naming conventions:
  - Classes: `CamelCase`
  - Functions/methods: `snake_case`
  - Constants: `UPPER_SNAKE_CASE`
  - Private attributes/methods: `_prefixed_with_underscore`

### Docstring Format

```python
def function_name(param1: type, param2: type) -> return_type:
    """Short description of function.

    More detailed description if needed.

    Args:
        param1: Description of param1
        param2: Description of param2

    Returns:
        Description of return value

    Raises:
        ExceptionType: When and why this exception is raised

    Examples:
        ```python
        result = function_name("value", 42)
        ```
    """
    # Implementation
```

## Best Practices

1. Keep interfaces simple and intuitive
2. Design for composability and reuse
3. Document extensively, especially API boundaries
4. Handle errors gracefully with clear messages
5. Maintain backward compatibility when possible
6. Follow the principle of least surprise in API design
7. Use percentage-based positioning and the h_align parameter for responsive layouts
8. Use the Grid layout system for complex layouts with nested elements
9. Optimize for readability and maintainability

## Model Context Protocol (MCP)

The Model Context Protocol (MCP) allows Claude to access specific external context.

### Using context7 for Documentation

When working with this project, use context7 MCP for retrieving the latest documentation:

```
<context7>
retrieve the latest documentation for [LIBRARY/FUNCTION], focused on [SPECIFIC ASPECT]
</context7>
```

Example:
```
<context7>
retrieve the latest documentation for python-pptx.Presentation.slides, focused on adding new slides
</context7>
```

### MCP for Development Assistance

- Use context7 for retrieving documentation about Python libraries
- Use context7 for understanding best practices for PowerPoint manipulation
- Use context7 for exploring python-pptx APIs and capabilities

## Completed Plans Management

When a specification is fully implemented:

1. **Verification**: Run `make check` and `make test` to ensure all tests pass, coding standards are met, and documentation is complete
2. **Status Update**: Mark as "Completed" in `.states/implementation_status.md`
3. **Move to Archive**: Transfer from `specs/` to `ai_docs/completed/`
4. **Implementation Notes**: Add implementation notes at the end of the spec document
5. **Reference**: Update any references to the completed spec in current development

A plan is considered "completed" when:
- All specified functionality is implemented
- `make test` passes with adequate coverage of the implementation
- `make check` passes with no linting or type errors
- Documentation is complete and accurate
- Code has been reviewed and meets style guidelines

## Licensing

This project is licensed under the MIT License - see the LICENSE file for details.
