# EasyPPTX Template Examples

This directory contains examples demonstrating how to use templates in EasyPPTX. Templates provide a powerful way to create consistent, professional presentations.

## Examples

### 001_template_basic.py
Basic template usage showing how to:
- Initialize a presentation with a TOML template
- Add slides that use a default template
- Override templates for individual slides
- Create slides without any template

### 002_template_presets.py
Using built-in template presets to create slides with:
- Title slides
- Content slides
- Section divider slides
- Image, table, and chart slides
- Comparison slides

### 003_template_toml_manager.py
Working with TOML template files:
- Creating custom templates
- Exporting templates to TOML and JSON formats
- Importing templates from files
- Using template manager to apply templates

### 004_template_manager.py
Template Manager API advanced features:
- Creating and registering custom templates
- Saving templates to and loading templates from files
- Using custom templates in presentations
- Listing available templates

### 005_template_toml_reference.py
Using TOML templates with reference PPTX files:
- Loading a template that specifies a reference PPTX file
- Creating slides based on this template
- Leveraging the reference PPTX for layouts and styling

## Running the Examples

Each example can be run directly:

```bash
python 001_template_basic.py
```

The presentations will be saved to the `output` directory at the root of the project.
