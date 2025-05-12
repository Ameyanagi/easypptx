"""
Dynamic Grid Example

This example demonstrates the new dynamic grid features:
1. Using the append method to add content to the grid and auto-expand it
2. Accessing out-of-bounds cells with auto-expansion
3. Using the flat indexing with auto-expansion

These features make it easier to work with grids by allowing
more flexible content addition without manually managing grid dimensions.
"""

from easypptx import Presentation

# Create a new presentation
pres = Presentation()

# First slide - Demonstrate append method
slide1, grid1 = pres.add_grid_slide(rows=2, cols=2, title="Grid Append Method Demo", title_align="center")


# Define append functions for dynamic content
def create_text1(**kwargs):
    return slide1.add_text(
        text="Item 1 - Appended",
        font_size=18,
        font_bold=True,
        align="center",
        vertical="middle",
        bg_color="lightblue",
        **kwargs,
    )


def create_text2(**kwargs):
    return slide1.add_text(
        text="Item 2 - Appended",
        font_size=18,
        font_bold=True,
        align="center",
        vertical="middle",
        bg_color="lightgreen",
        **kwargs,
    )


def create_text3(**kwargs):
    return slide1.add_text(
        text="Item 3 - Appended",
        font_size=18,
        font_bold=True,
        align="center",
        vertical="middle",
        bg_color="lightyellow",
        **kwargs,
    )


def create_text4(**kwargs):
    return slide1.add_text(
        text="Item 4 - Appended",
        font_size=18,
        font_bold=True,
        align="center",
        vertical="middle",
        bg_color="lightpink",
        **kwargs,
    )


def create_text5(**kwargs):
    return slide1.add_text(
        text="Item 5 - Auto-expanded!",
        font_size=18,
        font_bold=True,
        align="center",
        vertical="middle",
        bg_color="orange",
        **kwargs,
    )


# Append content to the grid - this will fill the grid
grid1.append(create_text1)
grid1.append(create_text2)
grid1.append(create_text3)
grid1.append(create_text4)

# Append one more item - this will auto-expand the grid by adding a row
grid1.append(create_text5)

# Second slide - Demonstrate out-of-bounds access with auto-expansion
slide2, grid2 = pres.add_grid_slide(rows=2, cols=2, title="Out-of-Bounds Access Demo", title_align="center")

# Access cells in the initial grid
grid2[0, 0].add_text(
    text="Cell [0, 0]", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="lightgray"
)

grid2[0, 1].add_text(
    text="Cell [0, 1]", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="lightblue"
)

# Access out-of-bounds cells - this will auto-expand the grid
grid2[2, 2].add_text(
    text="Cell [2, 2]\nAuto-expanded!",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
    bg_color="lightgreen",
)

grid2[3, 3].add_text(
    text="Cell [3, 3]\nAuto-expanded!",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
    bg_color="lightyellow",
)

# Third slide - Demonstrate flat indexing with auto-expansion
slide3, grid3 = pres.add_grid_slide(rows=2, cols=3, title="Flat Indexing with Auto-Expansion", title_align="center")

# Fill the initial grid
for i in range(6):
    grid3[i].add_text(
        text=f"grid[{i}]",
        font_size=18,
        font_bold=True,
        align="center",
        vertical="middle",
        bg_color=["lightgray", "lightblue", "lightgreen", "pink", "lightyellow", "lightcyan"][i],
    )

# Access beyond the grid size using flat indexing
grid3[10].add_text(
    text="grid[10]\nAuto-expanded!", font_size=18, font_bold=True, align="center", vertical="middle", bg_color="orange"
)

grid3[15].add_text(
    text="grid[15]\nAuto-expanded!",
    font_size=18,
    font_bold=True,
    align="center",
    vertical="middle",
    bg_color="lavender",
)

# Save the presentation
pres.save("dynamic_grid_example.pptx")
print("Presentation saved as 'dynamic_grid_example.pptx'")
