from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

# Create a new presentation
prs = Presentation()

# Set slide size to 16:9 (13.33 x 7.5 inches)
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

# Use a blank slide layout (usually index 6 in the default slide master)
blank_layout = prs.slide_layouts[6]  # Blank layout

# Function to add the first slide with the 2x2 layout
def add_slide1_with_2x2_layout():
    slide = prs.slides.add_slide(blank_layout)

    # Shape 1: Top-left title (50% width, smaller height for title)
    text_box_title1 = slide.shapes.add_textbox(
        left=Inches(0.5),
        top=Inches(0.3),
        width=Inches(6.0),
        height=Inches(0.5)
    )
    text_box_title1.name = "Slide1TopLeftTitle"
    text_frame_title1 = text_box_title1.text_frame
    text_frame_title1.text = "Slide1 Top Left Title Placeholder"
    text_frame_title1.word_wrap = True  # Enable word wrapping
    for paragraph in text_frame_title1.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        for run in paragraph.runs:
            run.font.size = Pt(18)
            run.font.bold = True

    # Shape 2: Top-left content (50% width, adjusted height)
    text_box_content1 = slide.shapes.add_textbox(
        left=Inches(0.5),
        top=Inches(0.9),  # Adjusted to make space for the title
        width=Inches(6.0),
        height=Inches(2.6)  # Reduced height to fit the title
    )
    text_box_content1.name = "Slide1TopLeftContent"
    text_frame_content1 = text_box_content1.text_frame
    text_frame_content1.text = "Slide1 Top Left Content Placeholder"
    text_frame_content1.word_wrap = True  # Enable word wrapping
    for paragraph in text_frame_content1.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT  # Left-align the content
        for run in paragraph.runs:
            run.font.size = Pt(14)

    # Shape 3: Top-right image placeholder (50% width, 50% height)
    shape3 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Inches(6.83),
        top=Inches(0.5),
        width=Inches(6.0),
        height=Inches(3.0)
    )
    shape3.name = "Slide1TopRightImage"
    shape3.text = "Slide1 Top Right Image Placeholder"
    text_frame3 = shape3.text_frame
    for paragraph in text_frame3.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(14)

    # Shape 4: Bottom-left image placeholder (50% width, 50% height)
    shape4 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Inches(0.5),
        top=Inches(4.0),
        width=Inches(6.0),
        height=Inches(3.0)
    )
    shape4.name = "Slide1BottomLeftImage"
    shape4.text = "Slide1 Bottom Left Image Placeholder"
    text_frame4 = shape4.text_frame
    for paragraph in text_frame4.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(14)

    # Shape 5: Bottom-right title (50% width, smaller height for title)
    text_box_title2 = slide.shapes.add_textbox(
        left=Inches(6.83),
        top=Inches(3.8),
        width=Inches(6.0),
        height=Inches(0.5)
    )
    text_box_title2.name = "Slide1BottomRightTitle"
    text_frame_title2 = text_box_title2.text_frame
    text_frame_title2.text = "Slide1 Bottom Right Title Placeholder"
    text_frame_title2.word_wrap = True  # Enable word wrapping
    for paragraph in text_frame_title2.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        for run in paragraph.runs:
            run.font.size = Pt(18)
            run.font.bold = True

    # Shape 6: Bottom-right content (50% width, adjusted height)
    text_box_content2 = slide.shapes.add_textbox(
        left=Inches(6.83),
        top=Inches(4.4),  # Adjusted to make space for the title
        width=Inches(6.0),
        height=Inches(2.6)  # Reduced height to fit the title
    )
    text_box_content2.name = "Slide1BottomRightContent"
    text_frame_content2 = text_box_content2.text_frame
    text_frame_content2.text = "Slide1 Bottom Right Content Placeholder"
    text_frame_content2.word_wrap = True  # Enable word wrapping
    for paragraph in text_frame_content2.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT  # Left-align the content
        for run in paragraph.runs:
            run.font.size = Pt(14)

# Function to add a slide with an image on the left and two title+content sections on the right
def add_slide_with_image_text_layout(slide_prefix):
    slide = prs.slides.add_slide(blank_layout)

    # Shape 1: Full-height image placeholder on the left (50% width)
    shape_image = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Inches(0.5),
        top=Inches(0.5),
        width=Inches(6.0),
        height=Inches(6.5)  # Full height minus margins
    )
    shape_image.name = f"{slide_prefix}LeftImage"
    shape_image.text = f"{slide_prefix} Left Image Placeholder"
    text_frame_image = shape_image.text_frame
    for paragraph in text_frame_image.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(14)

    # Shape 2: First title on the right (50% width, smaller height for title)
    text_box_title1 = slide.shapes.add_textbox(
        left=Inches(6.83),
        top=Inches(0.5),
        width=Inches(6.0),
        height=Inches(0.5)
    )
    text_box_title1.name = f"{slide_prefix}RightTitle1"
    text_frame_title1 = text_box_title1.text_frame
    text_frame_title1.text = f"{slide_prefix} Right Title 1 Placeholder"
    text_frame_title1.word_wrap = True  # Enable word wrapping
    for paragraph in text_frame_title1.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        for run in paragraph.runs:
            run.font.size = Pt(18)
            run.font.bold = True

    # Shape 3: First content on the right (50% width, adjusted height)
    text_box_content1 = slide.shapes.add_textbox(
        left=Inches(6.83),
        top=Inches(1.1),
        width=Inches(6.0),
        height=Inches(2.9)  # Adjusted to fit above the second section
    )
    text_box_content1.name = f"{slide_prefix}RightContent1"
    text_frame_content1 = text_box_content1.text_frame
    text_frame_content1.text = f"{slide_prefix} Right Content 1 Placeholder"
    text_frame_content1.word_wrap = True  # Enable word wrapping
    for paragraph in text_frame_content1.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        for run in paragraph.runs:
            run.font.size = Pt(14)

    # Shape 4: Second title on the right (50% width, smaller height for title)
    text_box_title2 = slide.shapes.add_textbox(
        left=Inches(6.83),
        top=Inches(4.2),
        width=Inches(6.0),
        height=Inches(0.5)
    )
    text_box_title2.name = f"{slide_prefix}RightTitle2"
    text_frame_title2 = text_box_title2.text_frame
    text_frame_title2.text = f"{slide_prefix} Right Title 2 Placeholder"
    text_frame_title2.word_wrap = True  # Enable word wrapping
    for paragraph in text_frame_title2.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        for run in paragraph.runs:
            run.font.size = Pt(18)
            run.font.bold = True

    # Shape 5: Second content on the right (50% width, adjusted height)
    text_box_content2 = slide.shapes.add_textbox(
        left=Inches(6.83),
        top=Inches(4.8),
        width=Inches(6.0),
        height=Inches(2.2)  # Adjusted to fit below the first section
    )
    text_box_content2.name = f"{slide_prefix}RightContent2"
    text_frame_content2 = text_box_content2.text_frame
    text_frame_content2.text = f"{slide_prefix} Right Content 2 Placeholder"
    text_frame_content2.word_wrap = True  # Enable word wrapping
    for paragraph in text_frame_content2.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        for run in paragraph.runs:
            run.font.size = Pt(14)

# Function to add the third slide with four images on the left and title+content on the right
def add_slide3_with_four_images_and_content():
    slide = prs.slides.add_slide(blank_layout)

    # Left Side: Four images
    # Shape 1: Large image at the top (50% width, ~50% height)
    shape_image1 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Inches(0.5),
        top=Inches(0.5),
        width=Inches(6.0),
        height=Inches(3.5)
    )
    shape_image1.name = "Slide3TopLeftImage"
    shape_image1.text = "Slide3 Top Left Image Placeholder"
    text_frame_image1 = shape_image1.text_frame
    for paragraph in text_frame_image1.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(14)

    # Shape 2: First small image below (50% width, ~16% height)
    shape_image2 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Inches(0.5),
        top=Inches(4.1),
        width=Inches(6.0),
        height=Inches(1.0)
    )
    shape_image2.name = "Slide3BottomLeftImage1"
    shape_image2.text = "Slide3 Bottom Left Image 1 Placeholder"
    text_frame_image2 = shape_image2.text_frame
    for paragraph in text_frame_image2.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(12)

    # Shape 3: Second small image below (50% width, ~16% height)
    shape_image3 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Inches(0.5),
        top=Inches(5.2),
        width=Inches(6.0),
        height=Inches(1.0)
    )
    shape_image3.name = "Slide3BottomLeftImage2"
    shape_image3.text = "Slide3 Bottom Left Image 2 Placeholder"
    text_frame_image3 = shape_image3.text_frame
    for paragraph in text_frame_image3.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(12)

    # Shape 4: Third small image below (50% width, ~16% height)
    shape_image4 = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        left=Inches(0.5),
        top=Inches(6.3),
        width=Inches(6.0),
        height=Inches(1.0)
    )
    shape_image4.name = "Slide3BottomLeftImage3"
    shape_image4.text = "Slide3 Bottom Left Image 3 Placeholder"
    text_frame_image4 = shape_image4.text_frame
    for paragraph in text_frame_image4.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(12)

    # Right Side: Title and content
    # Shape 5: Title at the top (50% width, smaller height for title)
    text_box_title = slide.shapes.add_textbox(
        left=Inches(6.83),
        top=Inches(0.5),
        width=Inches(6.0),
        height=Inches(0.5)
    )
    text_box_title.name = "Slide3RightTitle"
    text_frame_title = text_box_title.text_frame
    text_frame_title.text = "Slide3 Right Title Placeholder"
    text_frame_title.word_wrap = True
    for paragraph in text_frame_title.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        for run in paragraph.runs:
            run.font.size = Pt(18)
            run.font.bold = True

    # Shape 6: Content below the title (50% width, full height minus title)
    text_box_content = slide.shapes.add_textbox(
        left=Inches(6.83),
        top=Inches(1.1),
        width=Inches(6.0),
        height=Inches(6.1)  # Full height minus title and margins
    )
    text_box_content.name = "Slide3RightContent"
    text_frame_content = text_box_content.text_frame
    text_frame_content.text = "Slide3 Right Content Placeholder"
    text_frame_content.word_wrap = True
    for paragraph in text_frame_content.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        for run in paragraph.runs:
            run.font.size = Pt(14)

# Add first slide (Slide1) with 2x2 layout
add_slide1_with_2x2_layout()

# Add second slide (Slide2) with image on the left and two title+content sections on the right
add_slide_with_image_text_layout("Slide2")

# Add third slide (Slide3) with four images on the left and title+content on the right
add_slide3_with_four_images_and_content()

# Save the presentation as a template (.potx)
prs.save("custom_2x2_template.potx")
print("Template created successfully!")