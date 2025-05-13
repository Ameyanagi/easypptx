"""
Example showing how to fix title alignment issues when using TOML templates with grid slides.
"""

from easypptx import Presentation

# Create a simple TOML template file
with open("left_aligned_template.toml", "w") as f:
    f.write("""
# Template with left-aligned title

[title]
text = "Presentation Title"
position = { x = "10%", y = "5%", width = "80%", height = "10%" }
font = { name = "Meiryo", size = 44, bold = true }
align = "left"
vertical = "middle"
color = "black"
""")

# Method 1: Without fix - alignment from template is ignored
print("Method 1: Without explicitly setting title_align (before fix)")
pres1 = Presentation(template_toml="left_aligned_template.toml")
slide1, grid1 = pres1.add_grid_slide(title="Title Not Respecting Template", cols=2, rows=2)
grid1[0, 0].add_text("This slide's title will be center-aligned despite template")
pres1.save("template_issue_demo.pptx")

# Method 2: Workaround 1 - Explicitly specify the alignment
print("Method 2: Explicitly setting title_align to match template")
pres2 = Presentation(template_toml="left_aligned_template.toml")
slide2, grid2 = pres2.add_grid_slide(
    title="Title With Explicit Left Alignment",
    cols=2,
    rows=2,
    title_align="left",  # Explicitly override the default center alignment
)
grid2[0, 0].add_text("This slide's title will be left-aligned")
pres2.save("template_manual_fix.pptx")

# Method 3: Workaround 2 - Use add_slide_from_template + add_grid
print("Method 3: Using add_slide_from_template + add_grid")
pres3 = Presentation()
template_name = pres3.template_manager.load("left_aligned_template.toml")
slide3 = pres3.add_slide_from_template(template_name)
slide3.title = "Title From Template Properly Applied"
grid3 = pres3.add_grid(
    slide=slide3,
    x="0%",
    y="20%",  # Position below the title
    width="100%",
    height="80%",
    rows=2,
    cols=2,
)
grid3[0, 0].add_text("This slide's title will be left-aligned from template")
pres3.save("template_alternate_fix.pptx")

print("With the patch applied, all three methods should use the template's left alignment")
