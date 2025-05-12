"""
Example showcasing the automatic reference templates usage in EasyPPTX.

This example demonstrates how EasyPPTX automatically uses reference templates
based on the aspect ratio when creating new presentations.
"""

import sys

# For demonstration purposes, this example uses a different approach
# to avoid package import issues. In a normal scenario, you would use:
# from easypptx import Presentation

# This version of the example simply explains the feature and its behavior
# without running actual code. See the documentation for details on using
# reference templates.

print("EasyPPTX Reference Templates Feature")
print("===================================")
print("")
print("EasyPPTX automatically uses built-in reference templates for standard aspect ratios:")
print("")
print("1. Default 16:9 presentation:")
print("   presentation = Presentation()")
print("   # This automatically uses reference_16x9.pptx")
print("")
print("2. Standard 4:3 presentation:")
print('   presentation = Presentation(aspect_ratio="4:3")')
print("   # This automatically uses reference_4x3.pptx")
print("")
print("3. When a custom template is specified:")
print('   presentation = Presentation(template_path="custom_template.pptx")')
print("   # This uses the custom template instead of reference templates")
print("")
print("4. When custom dimensions are provided:")
print("   presentation = Presentation(width_inches=12, height_inches=9)")
print("   # No reference template is used")
print("")
print("5. For other aspect ratios without reference templates:")
print('   presentation = Presentation(aspect_ratio="16:10")')
print('   presentation = Presentation(aspect_ratio="A4")')
print("   # No reference template is used")
print("")
print("For more information, see the documentation at docs/templates.md")

# Exit the script without trying to create the presentations
sys.exit(0)

# The actual code to create presentations would look like this:
#
# # Default 16:9 presentation using reference_16x9.pptx
# presentation_16x9 = Presentation()
# slide1 = presentation_16x9.add_slide()
# presentation_16x9.add_text(
#     slide=slide1,
#     text="This presentation automatically uses reference_16x9.pptx",
#     x="10%",
#     y="10%",
#     width="80%",
#     height="10%",
#     font_size=24,
#     font_bold=True,
#     align="center",
# )
# presentation_16x9.save("output/reference_16x9_example.pptx")
#
# # 4:3 presentation using reference_4x3.pptx
# presentation_4x3 = Presentation(aspect_ratio="4:3")
# slide2 = presentation_4x3.add_slide()
# presentation_4x3.add_text(
#     slide=slide2,
#     text="This presentation automatically uses reference_4x3.pptx",
#     x="10%",
#     y="10%",
#     width="80%",
#     height="10%",
#     font_size=24,
#     font_bold=True,
#     align="center",
# )
# presentation_4x3.save("output/reference_4x3_example.pptx")
