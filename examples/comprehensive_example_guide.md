# Comprehensive Example Guide

This document provides a detailed walkthrough of the `comprehensive_example.py` script, which demonstrates how to create a complete business presentation using EasyPPTX.

## Overview

The example creates a 10-slide business presentation that resembles a professional annual report. Each slide demonstrates different capabilities of the EasyPPTX library.

## Step-by-Step Explanation

### Setup and Data Preparation

The script begins by setting up necessary directories and preparing sample data:

```python
# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Ensure the image directory exists
image_dir = output_dir / "images"
image_dir.mkdir(exist_ok=True)
```

It then generates sample images using PIL if available:

```python
def create_sample_image(name, size=(800, 600), color=(200, 200, 200)):
    """Create a sample image for the example."""
    # Implementation details omitted for brevity
```

Next, it creates sample data for the presentation:

```python
# Monthly sales data
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
sales_values = [42, 85, 53, 62, 75, 92, 98, 87, 76, 83, 95, 110]

# Product data, regional data, satisfaction data, etc.
# ...

# Create pandas DataFrames
product_df = pd.DataFrame({
    "Product": products,
    "Price": prices,
    # other columns
})
```

### Slide 1: Title Slide

The first slide is a title slide with a company logo:

```python
slide1 = pres.add_slide()
text = Text(slide1)
text.add_title("Annual Business Report")
text.add_paragraph("Fiscal Year 2023", x=1, y=2, font_size=24)
text.add_paragraph("Created with EasyPPTX", x=1, y=3, font_size=18, font_italic=True)

# Add company logo
img = Image(slide1)
img.add(str(logo_path), x=7, y=0.5, width=2.5)
```

Key features demonstrated:
- Adding a title
- Formatting text with different sizes
- Using italic formatting
- Positioning elements precisely on the slide
- Adding an image with specific dimensions

### Slide 2: Agenda

The second slide presents an agenda with bullet points:

```python
slide2 = pres.add_slide()
text = Text(slide2)
text.add_title("Agenda")

# Add agenda items with bullet points (simulated)
agenda_items = [
    "Company Overview",
    "Annual Sales Performance",
    # ...
]

for i, item in enumerate(agenda_items):
    text.add_paragraph(f"• {item}", x=1, y=2 + (i * 0.6), font_size=24)
```

Key features demonstrated:
- Creating a list of items
- Using a loop to add multiple text elements
- Calculating positions dynamically
- Simulating bullet points in text

### Slide 3: Company Overview

This slide combines an image with a data table:

```python
slide3 = pres.add_slide()
text = Text(slide3)
text.add_title("Company Overview")

# Add team photo
img = Image(slide3)
img.add(str(team_photo_path), x=1, y=1.5, width=8)

# Add key metrics as a table
metrics = [
    ("Founded", "2010"),
    ("Employees", "250+"),
    # ...
]

table = Table(slide3)
table_data = [["Metric", "Value"]] + metrics
table.add(table_data, x=1, y=4.5, width=8, first_row_header=True)
```

Key features demonstrated:
- Adding a full-width image
- Converting a list of tuples to a table
- Setting the first row as a header

### Slide 4: Annual Sales Performance

This slide features a line chart with analysis text:

```python
slide4 = pres.add_slide()
text = Text(slide4)
text.add_title("Annual Sales Performance")

# Add line chart for monthly sales
chart = Chart(slide4)
chart.add(
    chart_type="line",
    categories=months,
    values=sales_values,
    x=1,
    y=1.5,
    width=8,
    height=4,
    title="Monthly Sales (in $1,000s)"
)

# Add summary text with calculated values
total_sales = sum(sales_values)
avg_sales = total_sales / len(sales_values)
# Add formatted text with the calculated values
```

Key features demonstrated:
- Creating a line chart
- Setting chart dimensions and title
- Performing calculations on the data
- Adding formatted results as text

### Slide 5: Product Analysis - Table

This slide uses a pandas DataFrame to create a table:

```python
slide5 = pres.add_slide()
text = Text(slide5)
text.add_title("Product Analysis - Performance")

# Convert the DataFrame to a table
table = Table(slide5)
table.from_dataframe(
    product_df,
    x=0.5,
    y=1.5,
    width=9,
    first_row_header=True
)
```

Key features demonstrated:
- Converting a pandas DataFrame directly to a table
- Setting table width for better readability
- Combining a table with an image

### Slide 6: Product Analysis - Charts

This slide showcases multiple charts from pandas data:

```python
slide6 = pres.add_slide()
text = Text(slide6)
text.add_title("Product Analysis - Visualization")

# Add a bar chart for product revenue
chart = Chart(slide6)
chart.from_dataframe(
    product_df,
    chart_type="bar",
    category_column="Product",
    value_column="Revenue",
    x=0.5,
    y=1.5,
    width=4.5,
    height=3,
    title="Revenue by Product"
)

# Add a column chart for product profitability
chart.from_dataframe(
    # parameters
)
```

Key features demonstrated:
- Creating multiple charts on one slide
- Using different chart types (bar and column)
- Positioning charts side by side
- Creating charts directly from DataFrame columns
- Adding insight text with data-driven analysis

### Slide 7: Regional Performance

This slide combines a pie chart with a percentage table:

```python
slide7 = pres.add_slide()
text = Text(slide7)
text.add_title("Regional Sales Breakdown")

# Add a pie chart
chart = Chart(slide7)
chart.from_dataframe(
    regional_df,
    chart_type="pie",
    category_column="Region",
    value_column="Sales",
    x=2,
    y=1.5,
    width=6,
    height=4,
    title="Sales by Region"
)

# Calculate percentages and create a formatted table
# ...
```

Key features demonstrated:
- Creating a pie chart
- Calculating percentages from raw data
- Formatting numbers with commas and percentage symbols
- Creating a formatted data table to accompany a chart

### Slide 8: Customer Satisfaction

This slide features a column chart with conditional formatting:

```python
slide8 = pres.add_slide()
text = Text(slide8)
text.add_title("Customer Satisfaction")

# Add a column chart
chart = Chart(slide8)
chart.from_dataframe(
    satisfaction_df,
    chart_type="column",
    category_column="Category",
    value_column="Percentage",
    x=1.5,
    y=1.5,
    width=7,
    height=3.5,
    title="Customer Satisfaction Survey Results (%)"
)

# Calculate the satisfaction score
score = sum([satisfaction_values[i] * (5-i) for i in range(5)]) / sum(satisfaction_values)

# Add conditional formatting based on the score
if score >= 4.0:
    interpretation = "Excellent customer satisfaction!"
    color = (0, 128, 0)  # Green
elif score >= 3.0:
    # ...
```

Key features demonstrated:
- Visualizing survey data with a column chart
- Calculating a weighted score from raw data
- Using conditional logic to determine text content
- Applying different text colors based on data values

### Slide 9: Future Outlook

This slide combines bullet points with a projection chart:

```python
slide9 = pres.add_slide()
text = Text(slide9)
text.add_title("Future Outlook")

# Add bullet points as paragraph texts
future_plans = [
    "Expand into two new markets: Asia and South America",
    # ...
]

for i, plan in enumerate(future_plans):
    text.add_paragraph(f"• {plan}", x=1, y=1.8 + (i * 0.7), font_size=18)

# Add projection chart
# Generate projected sales with 15% annual growth
current_yearly_sales = sum(sales_values)
projected_years = ["2023", "2024", "2025", "2026", "2027"]
projected_sales = [current_yearly_sales]

for i in range(4):
    projected_sales.append(projected_sales[-1] * 1.15)  # 15% growth
```

Key features demonstrated:
- Creating bullet point lists
- Generating future projections based on existing data
- Visualizing projections with a chart
- Combining text and chart elements

### Slide 10: Thank You / Q&A

The final slide provides a clean conclusion:

```python
slide10 = pres.add_slide()
text = Text(slide10)
text.add_title("Thank You!")
text.add_paragraph("Questions?", x=3.5, y=3.5, font_size=40, font_bold=True)

# Add contact information
text.add_paragraph("For more information:", x=3, y=5, font_size=14)
text.add_paragraph("email@example.com", x=3, y=5.4, font_size=14)
text.add_paragraph("www.example.com", x=3, y=5.7, font_size=14)

# Add a small logo
img = Image(slide10)
img.add(str(logo_path), x=7, y=6.5, width=2)
```

Key features demonstrated:
- Creating a clean conclusion slide
- Using large text for emphasis
- Adding contact information
- Positioning a logo as a branding element

## Running the Example

To run this example, ensure you have EasyPPTX installed and execute:

```bash
python examples/comprehensive_example.py
```

The script will:
1. Create necessary directories
2. Generate sample images if PIL is available
3. Create and format all 10 slides
4. Save the presentation to `output/comprehensive_example.pptx`
5. Print a summary of what was created

## Customization Ideas

You can customize this example in several ways:

1. **Use your own data**: Replace the sample data with real data from your business
2. **Add your own images**: Replace the generated images with actual photos or graphics
3. **Adjust colors and formatting**: Modify colors, fonts, and sizes to match your branding
4. **Add more slides**: Extend the presentation with additional content
5. **Create different chart types**: Experiment with different chart types for your data

## Learning from the Example

This comprehensive example demonstrates:

1. **Best practices** for organizing a PowerPoint presentation
2. **Data-driven content** generation
3. **Consistent formatting** across slides
4. **Visual storytelling** techniques
5. **Combining multiple elements** on slides effectively

By studying and adapting this example, you can create your own professional-quality presentations programmatically using EasyPPTX.
