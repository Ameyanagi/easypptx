"""
Example demonstrating the advanced templates in EasyPPTX.

This example shows:
1. Using the quote_slide template
2. Using the bullets_slide template
3. Using the agenda_slide template
4. Using the team_slide template
5. Using the statement_slide template
6. Using the dashboard_slide template
7. Using the timeline_slide template
"""

from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

from easypptx import Chart, Image, Presentation, Pyplot, Text

# Create a folder for outputs if it doesn't exist
output_dir = Path("outputs")
output_dir.mkdir(exist_ok=True)

# Create a new presentation
pres = Presentation()

# Add a title slide
title_slide = pres.add_title_slide(
    title="Advanced Templates in EasyPPTX", subtitle="Showcasing specialized slide layouts"
)

# 1. Quote slide
quote_slide = pres.add_slide_from_template("quote_slide")

# Update the quote and author
quote_shapes = [shape for shape in quote_slide.shapes if shape.has_text_frame]
for shape in quote_shapes:
    if "Your quote goes here" in shape.text_frame.text:
        shape.text_frame.text = "The best way to predict the future is to create it."
    elif "Author Name" in shape.text_frame.text:
        shape.text_frame.text = "Abraham Lincoln"

# 2. Bullets slide
bullets_slide = pres.add_slide_from_template("bullets_slide")

# Update the title and bullet points
bullet_shapes = [shape for shape in bullets_slide.shapes if shape.has_text_frame]
for shape in bullet_shapes:
    if "Key Points" in shape.text_frame.text:
        shape.text_frame.text = "Key Benefits"
    elif "Point" in shape.text_frame.text:
        shape.text_frame.text = """• Consistent slide layouts with minimal code
• Professional-looking presentations out of the box
• Percentage-based positioning for all elements
• Customizable templates for specific needs
• Support for matplotlib and seaborn visualizations"""

# 3. Agenda slide
agenda_slide = pres.add_slide_from_template("agenda_slide")

# Update the title and agenda items
agenda_shapes = [shape for shape in agenda_slide.shapes if shape.has_text_frame]
for shape in agenda_shapes:
    if "Agenda" in shape.text_frame.text:
        shape.text_frame.text = "Today's Topics"
    elif "Introduction" in shape.text_frame.text:
        shape.text_frame.text = """1. Template Overview
2. Basic Templates
3. Advanced Templates
4. Custom Templates
5. Styling Options
6. Integration with Matplotlib
7. Q&A"""

# Add an image to the agenda slide
try:
    # Try to add an image if available
    Image.add(
        slide=agenda_slide,
        image_path="examples/assets/sample_image.jpg",
        position={"x": "55%", "y": "20%", "width": "40%", "height": "70%"},
        maintain_aspect_ratio=True,
        center=True,
    )
except Exception:
    # If image is not available, add a placeholder text
    Text.add(
        slide=agenda_slide,
        text="[Image Placeholder]",
        position={"x": "55%", "y": "20%", "width": "40%", "height": "70%"},
        font_size=24,
        align="center",
        vertical_align="middle",
        color="gray",
    )

# 4. Team slide
team_slide = pres.add_slide_from_template("team_slide")

# Update the title
team_title = next(shape for shape in team_slide.shapes if shape.has_text_frame)
team_title.text_frame.text = "Our Development Team"

# Add team members
member_positions = [
    {"x": "10%", "y": "20%", "width": "20%", "height": "30%"},
    {"x": "40%", "y": "20%", "width": "20%", "height": "30%"},
    {"x": "70%", "y": "20%", "width": "20%", "height": "30%"},
    {"x": "25%", "y": "55%", "width": "20%", "height": "30%"},
    {"x": "55%", "y": "55%", "width": "20%", "height": "30%"},
]

member_names = ["John Smith", "Jane Doe", "Alex Chen", "Maria Garcia", "Sam Wilson"]
member_titles = ["Project Lead", "Senior Developer", "UI/UX Designer", "Data Scientist", "QA Engineer"]

# Add placeholder images and names
for _i, (position, name, title) in enumerate(zip(member_positions, member_names, member_titles, strict=False)):
    # Add placeholder for person image
    try:
        # Try to add an image if available
        Image.add(
            slide=team_slide,
            image_path="examples/assets/sample_image.jpg",
            position=position,
            maintain_aspect_ratio=True,
            center=True,
        )
    except Exception:
        # If image is not available, add a placeholder shape
        placeholder = team_slide.add_shape(position=position, fill_color="lightgray")

    # Add name and title
    name_position = {
        "x": position["x"],
        "y": str(float(position["y"].replace("%", "")) + float(position["height"].replace("%", "")) + 2) + "%",
        "width": position["width"],
        "height": "5%",
    }

    title_position = {
        "x": position["x"],
        "y": str(float(name_position["y"].replace("%", "")) + 5) + "%",
        "width": position["width"],
        "height": "5%",
    }

    Text.add(slide=team_slide, text=name, position=name_position, font_size=16, font_bold=True, align="center")

    Text.add(slide=team_slide, text=title, position=title_position, font_size=12, align="center", color="darkgray")

# 5. Statement slide
statement_slide = pres.add_slide_from_template("statement_slide")

# Update the statement
statement_shape = next(shape for shape in statement_slide.shapes if shape.has_text_frame)
statement_shape.text_frame.text = "Creating Professional Presentations Has Never Been Easier"

# 6. Dashboard slide
dashboard_slide = pres.add_slide_from_template("dashboard_slide")

# Update the title
dashboard_title = next(shape for shape in dashboard_slide.shapes if shape.has_text_frame)
dashboard_title.text_frame.text = "Performance Dashboard"

# Get the chart positions from the template
template = pres.template.get_preset("dashboard_slide")
chart_positions = template["charts"]

# Create some sample data
dates = pd.date_range(start="1/1/2023", periods=12, freq="M")
# Using numpy's random for secure random number generation
revenue = np.random.randint(900, 1500, size=12).tolist()
expenses = np.random.randint(700, 1200, size=12).tolist()
profit = [r - e for r, e in zip(revenue, expenses, strict=False)]

# Add a line chart to the top left
data1 = pd.DataFrame({"Month": [d.strftime("%b") for d in dates], "Revenue": revenue, "Expenses": expenses})

Chart.add(
    slide=dashboard_slide,
    data=data1,
    chart_type="line",
    position=chart_positions["top_left"]["position"],
    category_column="Month",
    value_columns=["Revenue", "Expenses"],
    has_legend=True,
    has_title=True,
    chart_title="Monthly Revenue & Expenses",
)

# Add a column chart to the top right
data2 = pd.DataFrame({
    "Quarter": ["Q1", "Q2", "Q3", "Q4"],
    "Profit": [sum(profit[0:3]), sum(profit[3:6]), sum(profit[6:9]), sum(profit[9:12])],
})

Chart.add(
    slide=dashboard_slide,
    data=data2,
    chart_type="column",
    position=chart_positions["top_right"]["position"],
    category_column="Quarter",
    value_columns="Profit",
    has_legend=False,
    has_title=True,
    chart_title="Quarterly Profit",
)

# Add a pie chart to the bottom left
data3 = pd.DataFrame({
    "Category": ["Product A", "Product B", "Product C", "Product D"],
    "Sales": np.random.randint(100, 500, size=4).tolist(),
})

Chart.add(
    slide=dashboard_slide,
    data=data3,
    chart_type="pie",
    position=chart_positions["bottom_left"]["position"],
    category_column="Category",
    value_columns="Sales",
    has_legend=True,
    has_title=True,
    chart_title="Sales by Product",
)

# Create a matplotlib heatmap for the bottom right
plt.figure(figsize=(6, 4))
corr_data = np.random.rand(5, 5)
corr_data = (corr_data + corr_data.T) / 2  # Make symmetric
np.fill_diagonal(corr_data, 1)  # Set diagonal to 1
sns.heatmap(
    corr_data, annot=True, cmap="YlGnBu", xticklabels=["A", "B", "C", "D", "E"], yticklabels=["A", "B", "C", "D", "E"]
)
plt.title("Correlation Matrix")

# Add the matplotlib figure to the bottom right
Pyplot.add(
    slide=dashboard_slide,
    figure=plt.gcf(),
    position=chart_positions["bottom_right"]["position"],
    style={"border": True, "border_color": "gray"},
)

# 7. Timeline slide
timeline_slide = pres.add_slide_from_template("timeline_slide")

# Update the title
timeline_title = next(shape for shape in timeline_slide.shapes if shape.has_text_frame)
timeline_title.text_frame.text = "Project Timeline"

# Get timeline area from the template
timeline_area = template["timeline_area"]["position"]
step_colors = template["step_colors"]

# Create timeline steps
steps = [
    {"title": "Planning", "date": "Jan 2023", "description": "Initial planning and requirements gathering"},
    {"title": "Development", "date": "Mar 2023", "description": "Core functionality implementation"},
    {"title": "Testing", "date": "Jun 2023", "description": "Quality assurance and bug fixes"},
    {"title": "Release", "date": "Sep 2023", "description": "Public release and deployment"},
    {"title": "Maintenance", "date": "Dec 2023", "description": "Ongoing support and updates"},
]

# Calculate positions for timeline steps
step_width = "15%"
step_height = "12%"
connector_height = "2%"
y_pos = timeline_area["y"]
spacing = (100 - 2 * float(timeline_area["y"].replace("%", "")) - 5 * float(step_height.replace("%", ""))) / 4

# Add the steps with connecting lines
for i, step in enumerate(steps):
    # Calculate position
    if i == 0:
        y = y_pos
    else:
        y = str(float(y_pos.replace("%", "")) + i * (float(step_height.replace("%", "")) + spacing)) + "%"

    # Alternate between left and right
    if i % 2 == 0:
        x = timeline_area["x"]
    else:
        x = str(100 - float(timeline_area["x"].replace("%", "")) - float(step_width.replace("%", ""))) + "%"

    # Add step box
    step_box = timeline_slide.add_shape(
        position={"x": x, "y": y, "width": step_width, "height": step_height},
        fill_color=step_colors[i % len(step_colors)],
    )

    # Add step title
    Text.add(
        slide=timeline_slide,
        text=step["title"],
        position={"x": x, "y": y, "width": step_width, "height": str(float(step_height.replace("%", "")) / 2) + "%"},
        font_name="Meiryo",
        font_size=16,
        font_bold=True,
        align="center",
        vertical_align="middle",
        color="white",
    )

    # Add step date
    date_y = str(float(y.replace("%", "")) + float(step_height.replace("%", "")) / 2) + "%"
    Text.add(
        slide=timeline_slide,
        text=step["date"],
        position={
            "x": x,
            "y": date_y,
            "width": step_width,
            "height": str(float(step_height.replace("%", "")) / 4) + "%",
        },
        font_name="Meiryo",
        font_size=12,
        align="center",
        vertical_align="middle",
        color="white",
    )

    # Add description
    desc_y = str(float(date_y.replace("%", "")) + float(step_height.replace("%", "")) / 4) + "%"
    Text.add(
        slide=timeline_slide,
        text=step["description"],
        position={
            "x": x,
            "y": desc_y,
            "width": step_width,
            "height": str(float(step_height.replace("%", "")) / 4) + "%",
        },
        font_name="Meiryo",
        font_size=10,
        align="center",
        vertical_align="middle",
        color="white",
    )

    # Add connector line (except for the last step)
    if i < len(steps) - 1:
        connector_x = str(float(timeline_area["x"].replace("%", "")) + float(step_width.replace("%", ""))) + "%"
        connector_y = str(float(y.replace("%", "")) + float(step_height.replace("%", "")) / 2) + "%"
        connector_width = (
            str(100 - 2 * float(timeline_area["x"].replace("%", "")) - 2 * float(step_width.replace("%", ""))) + "%"
        )

        # Adjust connector for alternating pattern
        if i % 2 == 0:
            connector_x = str(float(timeline_area["x"].replace("%", "")) + float(step_width.replace("%", ""))) + "%"
        else:
            connector_x = timeline_area["x"]
            connector_width = step_width

        timeline_slide.add_shape(
            position={"x": connector_x, "y": connector_y, "width": connector_width, "height": connector_height},
            fill_color=template["connector_color"],
        )

# 8. Add a "thank you" slide
thank_you_slide = pres.add_slide_from_template("thank_you_slide")

# Save the presentation
pres.save(output_dir / "advanced_templates_example.pptx")
print(f"Presentation saved to {output_dir / 'advanced_templates_example.pptx'}")
