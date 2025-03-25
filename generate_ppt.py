from pptx import Presentation
from pptx.util import Pt, Inches
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures the extracted text while preserving sentence boundaries."""
    if isinstance(text, list):
        text = "\n".join(text)

    # Remove unwanted characters but keep sentence terminators and newlines
    text = re.sub(r'[*#]', '', text)
    text = re.sub(r'[^A-Za-z0-9.,;!?\s\n]', '', text)  # Keep sentence punctuation
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces
    text = re.sub(r'\n\s*\n', '\n\n', text)  # Preserve intentional line breaks

    # Split into paragraphs based on double newlines
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    
    # Further split paragraphs that contain sentence terminators
    structured_text = []
    for para in paragraphs:
        # Split at sentence boundaries but preserve the terminator
        sentences = re.split(r'(?<=[.!?])\s+', para)
        structured_text.extend(s for s in sentences if s)
    
    return structured_text

def add_slide(prs, title, content):
    """Adds a slide with controlled bullet point application."""
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
    text_frame.word_wrap = True

    # Add content with controlled bullet points
    for i, paragraph in enumerate(content):
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.font.size = Pt(16)
        p.space_after = Pt(8)
        # Only add bullet if it's the first item or starts with capital letter (new sentence)
        if i == 0 or (paragraph and paragraph[0].isupper()):
            p.level = 0  # Add bullet point
        else:
            p.level = 1  # No bullet point (indented continuation)

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

        # Add a Slide for Each File Title
        file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only Layout
        file_title_slide.shapes.title.text = f"Summary of {filename}"

        for paragraph in structured_text:
            # Split large paragraphs into smaller chunks
            wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line)
            for i, wrapped_line in enumerate(wrapped_lines):
                # Only consider as new thought if it's the first line or starts with capital
                if i == 0 or (wrapped_line and wrapped_line[0].isupper()):
                    current_slide_text.append(wrapped_line)
                else:
                    # Append to last line if it's a continuation
                    if current_slide_text:
                        current_slide_text[-1] += " " + wrapped_line
                    else:
                        current_slide_text.append(wrapped_line)

                # If the slide is full, add a new slide
                if len(current_slide_text) >= max_lines_per_slide:
                    add_slide(prs, f"Key Points from {filename}", current_slide_text)
                    current_slide_text = []

        # Add remaining text to a new slide
        if current_slide_text:
            add_slide(prs, f"Additional Info from {filename}", current_slide_text)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
