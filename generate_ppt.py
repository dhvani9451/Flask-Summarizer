from pptx import Presentation
from pptx.util import Pt, Inches
import io
import textwrap
import re
import os
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.dml.color import RGBColor

def clean_text(text):
    """Cleans and structures the extracted text."""
    if isinstance(text, list):
        text = "\n".join(text)

    text = re.sub(r'[*#]', '', text)  # Remove unwanted characters
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)  # Keep only letters, numbers, and punctuation
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces
    text = re.sub(r'\n+', '\n', text)  # Remove extra newlines

    # Split into paragraphs or thoughts based on line breaks
    paragraphs = text.split("\n")
    structured_text = [para.strip() for para in paragraphs if para]

    return structured_text

def add_slide(prs, title, content):
    """Adds a slide ensuring bullet points are applied to new thoughts and text fits within the textbox."""
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)

    # Set slide title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Set slide content
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True  # Enable word wrapping

    # Add bullet points for each new thought
    for i, paragraph in enumerate(content):
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.font.size = Pt(16)  # Slightly smaller font for better fit
        p.space_after = Pt(8)  # Adjust spacing for readability

        if i == 0:
            p.level = 0  # PowerPoint's default bullet
            # Explicitly set bullet properties for the first item
            p.paragraph_format.bullet.type = 0  # Auto bullet (usually a circle)
            # Add color code if you want the bullet to have a color
            #p.paragraph_format.bullet.font.color.rgb = RGBColor(0xFF, 0x00, 0x00) #Red color

        else:
            p.level = 1
            # Explicitly remove bullet properties for subsequent items
            p.paragraph_format.bullet.type = 4 #PP_BULLET_TYPE.NONE
            #text_frame.paragraphs[i].alignment = PP_PARAGRAPH_ALIGNMENT.LEFT #adjust to left alignment
            p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT #adjust to left alignment

def create_presentation(file_texts):
    """Creates a PowerPoint presentation and returns it as a file object."""
    template_path = "Ion.pptx"
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # Add Title Slide
    slide_layout = prs.slide_layouts[0]  # Title Slide Layout
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Vehicle Rental System - Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"

    max_lines_per_slide = 6  # Maximum lines per slide
    max_chars_per_line = 80  # Maximum characters per line

    for filename, text in file_texts.items():
        structured_text = clean_text(text)
        current_slide_text = []
        first_slide = True

        # Add a Slide for Each File Title
        slide_layout = prs.slide_layouts[5]
        file_title_slide = prs.slides.add_slide(slide_layout)
        file_title_slide.shapes.title.text = f"Summary of {filename}"  # More descriptive title

        for paragraph in structured_text:
            # Split large paragraphs into smaller chunks
            wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line)
            for wrapped_line in wrapped_lines:
                current_slide_text.append(wrapped_line)

                # If the slide is full, add a new slide
                if len(current_slide_text) == max_lines_per_slide:
                    if first_slide:
                        add_slide(prs, f"Key Points from {filename}", current_slide_text)
                        first_slide = False
                    else:
                        add_slide(prs, f"Cont. Key Points from {filename}", current_slide_text)
                    current_slide_text = []

        # Add remaining text to a new slide
        if current_slide_text:
            if first_slide:
                add_slide(prs, f"Key Points from {filename}", current_slide_text)
                first_slide = False
            else:
                add_slide(prs, f"Cont. Key Points from {filename}", current_slide_text)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
