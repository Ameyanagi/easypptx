"""
Comprehensive example of EasyPPTX capabilities.

This example creates a detailed presentation that demonstrates the
full range of features available in the EasyPPTX library, including:

- Different slide types and layouts
- Text formatting and positioning
- Image handling with aspect ratio control
- Table creation from data and pandas DataFrames
- Chart generation (column, bar, pie, line) from data and pandas DataFrames
- Styling and formatting options

The presentation is organized as a business presentation with real-world
examples of how to use EasyPPTX in various scenarios.
"""

from datetime import date
from pathlib import Path

import pandas as pd

from easypptx import Chart, Image, Presentation, Table

# -----------------------------------------------------------------------------
# Setup
# -----------------------------------------------------------------------------

# Create output directory
output_dir = Path("output")
output_dir.mkdir(exist_ok=True)

# Ensure the image directory exists
image_dir = output_dir / "images"
image_dir.mkdir(exist_ok=True)


# Create sample images if they don't exist
def create_sample_image(name, size=(800, 600), color=(200, 200, 200)):
    """Create a sample image for the example."""
    try:
        from PIL import Image as PILImage
        from PIL import ImageDraw, ImageFont

        img = PILImage.new("RGB", size, color=color)
        draw = ImageDraw.Draw(img)

        # Add text to the image
        try:
            font = ImageFont.truetype("arial.ttf", 32)
        except OSError:
            font = ImageFont.load_default()

        draw.text((size[0] // 3, size[1] // 2), name, fill=(0, 0, 0), font=font)

        # Save the image
        img_path = image_dir / f"{name}.png"
        img.save(img_path)
    except ImportError:
        print("PIL not installed. Using placeholder paths.")
        return f"images/{name}.png"

    return img_path


# Create some sample images
logo_path = create_sample_image("company_logo", (400, 100), (255, 255, 255))
product_image_path = create_sample_image("product_image", (800, 600), (240, 240, 255))
team_photo_path = create_sample_image("team_photo", (800, 400), (255, 240, 240))
graph_image_path = create_sample_image("graph_image", (600, 400), (240, 255, 240))

# Create sample data for charts and tables
# Monthly sales data
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
sales_values = [42, 85, 53, 62, 75, 92, 98, 87, 76, 83, 95, 110]

# Product data
products = ["Product A", "Product B", "Product C", "Product D", "Product E"]
prices = [199, 149, 299, 99, 249]
units_sold = [120, 250, 75, 360, 90]
revenue = [p * u for p, u in zip(prices, units_sold, strict=False)]
costs = [p * 0.6 * u for p, u in zip(prices, units_sold, strict=False)]
profits = [r - c for r, c in zip(revenue, costs, strict=False)]

# Customer satisfaction data
satisfaction_categories = ["Very Satisfied", "Satisfied", "Neutral", "Unsatisfied", "Very Unsatisfied"]
satisfaction_values = [45, 30, 15, 7, 3]

# Regional sales data
regions = ["North", "South", "East", "West", "Central"]
region_sales = [28500, 19700, 31200, 34800, 22300]

# Create pandas DataFrames
product_df = pd.DataFrame({
    "Product": products,
    "Price": prices,
    "Units Sold": units_sold,
    "Revenue": revenue,
    "Cost": costs,
    "Profit": profits,
})

monthly_sales_df = pd.DataFrame({"Month": months, "Sales": sales_values})

regional_df = pd.DataFrame({"Region": regions, "Sales": region_sales})

satisfaction_df = pd.DataFrame({"Category": satisfaction_categories, "Percentage": satisfaction_values})

# -----------------------------------------------------------------------------
# Create the presentation
# -----------------------------------------------------------------------------
pres = Presentation()

# -----------------------------------------------------------------------------
# Slide 1: Title Slide
# -----------------------------------------------------------------------------
slide1 = pres.add_slide()

# Add centered title and subtitle with responsive positioning
slide1.add_text(
    text="Annual Business Report",
    x="50%",
    y="30%",
    width="80%",
    height="15%",
    font_size=44,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)
slide1.add_text(
    text="Fiscal Year 2023",
    x="50%",
    y="45%",
    width="60%",
    height="10%",
    font_size=24,
    align="center",
    h_align="center",  # Enable responsive positioning
)
slide1.add_text(
    text="Created with EasyPPTX",
    x="50%",
    y="55%",
    width="60%",
    height="10%",
    font_size=18,
    font_italic=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add company logo
img = Image(slide1)
img.add(str(logo_path), x="80%", y="10%", width="15%", h_align="center")

# Add date
today = date.today().strftime("%B %d, %Y")
slide1.add_text(f"Date: {today}", x="80%", y="85%", width="15%", height="5%", font_size=10, align="right")

# -----------------------------------------------------------------------------
# Slide 2: Agenda
# -----------------------------------------------------------------------------
slide2 = pres.add_slide()

# Add centered title with responsive positioning
slide2.add_text(
    text="Agenda",
    x="50%",
    y="10%",
    width="80%",
    height="15%",
    font_size=44,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add agenda items with bullet points using percentage-based positioning
agenda_items = [
    "Company Overview",
    "Annual Sales Performance",
    "Product Analysis",
    "Regional Breakdown",
    "Customer Satisfaction",
    "Future Outlook",
]

for i, item in enumerate(agenda_items):
    slide2.add_text(text=f"• {item}", x="30%", y=f"{25 + (i * 10)}%", width="60%", height="8%", font_size=24)

# -----------------------------------------------------------------------------
# Slide 3: Company Overview
# -----------------------------------------------------------------------------
slide3 = pres.add_slide()
# Centered title with responsive positioning
slide3.add_text(
    text="Company Overview",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add team photo
img = Image(slide3)
img.add(str(team_photo_path), x=1, y=1.5, width=8)

# Add key metrics
metrics = [
    ("Founded", "2010"),
    ("Employees", "250+"),
    ("Revenue", "$36.5M"),
    ("Growth", "+15% YoY"),
    ("Locations", "12 Countries"),
]

# Create a simple table for metrics
table = Table(slide3)
table_data = [["Metric", "Value"], *metrics]
table.add(table_data, x=1, y=4.5, width=8, first_row_header=True)

# -----------------------------------------------------------------------------
# Slide 4: Annual Sales Performance
# -----------------------------------------------------------------------------
slide4 = pres.add_slide()
# Centered title with responsive positioning
slide4.add_text(
    text="Annual Sales Performance",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

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
    title="Monthly Sales (in $1,000s)",
)

# Add summary text
total_sales = sum(sales_values)
avg_sales = total_sales / len(sales_values)
slide4.add_text(
    text=f"Total Annual Sales: ${total_sales}k",
    x="25%",
    y="80%",
    width="40%",
    height="10%",
    font_size=18,
    font_bold=True,
    align="center",
)
slide4.add_text(
    text=f"Average Monthly Sales: ${avg_sales:.1f}k",
    x="75%",
    y="80%",
    width="40%",
    height="10%",
    font_size=18,
    align="center",
)

# -----------------------------------------------------------------------------
# Slide 5: Product Analysis - Table
# -----------------------------------------------------------------------------
slide5 = pres.add_slide()
# Centered title with responsive positioning
slide5.add_text(
    text="Product Analysis - Performance",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Convert the DataFrame to a table
table = Table(slide5)
table.from_dataframe(product_df, x=0.5, y=1.5, width=9, first_row_header=True)

# Add a product image
img = Image(slide5)
img.add(str(product_image_path), x=7.5, y=5, width=2)

# -----------------------------------------------------------------------------
# Slide 6: Product Analysis - Charts
# -----------------------------------------------------------------------------
slide6 = pres.add_slide()
# Centered title with responsive positioning
slide6.add_text(
    text="Product Analysis - Visualization",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

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
    title="Revenue by Product",
)

# Add a column chart for product profitability
chart.from_dataframe(
    product_df,
    chart_type="column",
    category_column="Product",
    value_column="Profit",
    x=5.5,
    y=1.5,
    width=4,
    height=3,
    title="Profit by Product",
)

# Add insight text
most_profitable = product_df.loc[product_df["Profit"].idxmax()]["Product"]
slide6.add_text(
    text=f"Key Insight: {most_profitable} is our most profitable product.",
    x="50%",
    y="80%",
    width="80%",
    height="10%",
    font_size=16,
    font_bold=True,
    color=(0, 100, 0),  # Dark green
    align="center",
    h_align="center",
)

# -----------------------------------------------------------------------------
# Slide 7: Regional Performance
# -----------------------------------------------------------------------------
slide7 = pres.add_slide()
# Centered title with responsive positioning
slide7.add_text(
    text="Regional Sales Breakdown",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

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
    title="Sales by Region",
)

# Calculate percentages
total_regional = sum(region_sales)
region_percentages = [round((s / total_regional) * 100, 1) for s in region_sales]

# Create a table with percentages
percentage_data = [["Region", "Sales", "Percentage"]]
for i, region in enumerate(regions):
    percentage_data.append([region, f"${region_sales[i]:,}", f"{region_percentages[i]}%"])

table = Table(slide7)
table.add(percentage_data, x=1.5, y=5.5, width=7, first_row_header=True)

# -----------------------------------------------------------------------------
# Slide 8: Customer Satisfaction
# -----------------------------------------------------------------------------
slide8 = pres.add_slide()
# Centered title with responsive positioning
slide8.add_text(
    text="Customer Satisfaction",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

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
    title="Customer Satisfaction Survey Results (%)",
)

# Calculate the satisfaction score
score = sum([satisfaction_values[i] * (5 - i) for i in range(5)]) / sum(satisfaction_values)

# Add summary text
slide8.add_text(
    text=f"Overall satisfaction score: {score:.1f}/5.0",
    x="50%",
    y="75%",
    width="60%",
    height="10%",
    font_size=20,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add interpretation
if score >= 4.0:
    interpretation = "Excellent customer satisfaction!"
    color = (0, 128, 0)  # Green
elif score >= 3.0:
    interpretation = "Good customer satisfaction"
    color = (0, 0, 255)  # Blue
else:
    interpretation = "Needs improvement"
    color = (200, 0, 0)  # Red

slide8.add_text(
    text=interpretation,
    x="50%",
    y="85%",
    width="60%",
    height="10%",
    font_size=20,
    color=color,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# -----------------------------------------------------------------------------
# Slide 9: Future Outlook
# -----------------------------------------------------------------------------
slide9 = pres.add_slide()
# Centered title with responsive positioning
slide9.add_text(
    text="Future Outlook",
    x="50%",
    y="10%",
    width="80%",
    height="10%",
    font_size=32,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add bullet points as text objects with percentage-based positioning
future_plans = [
    "Expand into two new markets: Asia and South America",
    "Launch 3 new product lines in Q2 2024",
    "Improve customer satisfaction score to 4.5+",
    "Increase operational efficiency by 12%",
    "Invest in employee training and development",
]

for i, plan in enumerate(future_plans):
    slide9.add_text(text=f"• {plan}", x="30%", y=f"{25 + (i * 10)}%", width="60%", height="8%", font_size=18)

# Add projection chart - projecting future sales based on current growth
# Generate projected sales with 15% annual growth
current_yearly_sales = sum(sales_values)
projected_years = ["2023", "2024", "2025", "2026", "2027"]
projected_sales = [current_yearly_sales]

for _ in range(4):
    projected_sales.append(projected_sales[-1] * 1.15)  # 15% growth

chart = Chart(slide9)
chart.add(
    chart_type="column",
    categories=projected_years,
    values=projected_sales,
    x=5.5,
    y=2,
    width=4,
    height=3,
    title="Projected Annual Sales ($k)",
)

# -----------------------------------------------------------------------------
# Slide 10: Thank You / Q&A
# -----------------------------------------------------------------------------
slide10 = pres.add_slide()
# Centered title with responsive positioning
slide10.add_text(
    text="Thank You!",
    x="50%",
    y="20%",
    width="80%",
    height="15%",
    font_size=44,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)
# Centered subtitle with responsive positioning
slide10.add_text(
    text="Questions?",
    x="50%",
    y="50%",
    width="60%",
    height="15%",
    font_size=40,
    font_bold=True,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add contact information with responsive positioning
slide10.add_text(
    text="For more information:",
    x="50%",
    y="70%",
    width="60%",
    height="5%",
    font_size=14,
    align="center",
    h_align="center",  # Enable responsive positioning
)
slide10.add_text(
    text="email@example.com",
    x="50%",
    y="75%",
    width="60%",
    height="5%",
    font_size=14,
    align="center",
    h_align="center",  # Enable responsive positioning
)
slide10.add_text(
    text="www.example.com",
    x="50%",
    y="80%",
    width="60%",
    height="5%",
    font_size=14,
    align="center",
    h_align="center",  # Enable responsive positioning
)

# Add a small logo with responsive positioning
img = Image(slide10)
img.add(str(logo_path), x="80%", y="85%", width="15%", h_align="center")

# -----------------------------------------------------------------------------
# Save the presentation
# -----------------------------------------------------------------------------
output_path = output_dir / "comprehensive_example.pptx"
pres.save(output_path)
print(f"Presentation saved to {output_path}")

# Print a summary of what was created
print(f"\nCreated a presentation with {len(pres.slides)} slides:")
print("1. Title Slide")
print("2. Agenda")
print("3. Company Overview")
print("4. Annual Sales Performance")
print("5. Product Analysis - Table")
print("6. Product Analysis - Charts")
print("7. Regional Sales Breakdown")
print("8. Customer Satisfaction")
print("9. Future Outlook")
print("10. Thank You / Q&A")
