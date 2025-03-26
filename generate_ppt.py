from pptx import Presentation
from pptx.util import Pt, Inches
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures the extracted text while retaining essential punctuation."""
    if isinstance(text, list):
        text = "\n".join(text)

    # Remove unwanted characters but retain essential punctuation
    text = re.sub(r'[*#]', '', text)  # Remove unwanted characters
    # Retain letters, numbers, punctuation, and common symbols
    text = re.sub(r'[^\w\s.,:;!?\'"()-]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces
    text = re.sub(r'\n+', '\n', text)  # Remove extra newlines

    # Split into paragraphs based on line breaks
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

    # Add bullet points for each sentence/idea
    for sentence in content:
        p = text_frame.add_paragraph()
        p.text = sentence.strip()  # Remove leading/trailing whitespace
        p.font.size = Pt(16)  # Slightly smaller font for better fit
        p.space_after = Pt(8)  # Adjust spacing for readability
        p.level = 0  # PowerPoint's default bullet


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
        file_title_slide.shapes.title.text = f"Summary of {filename}"  # More descriptive title

        for paragraph in structured_text:
            # Split large paragraphs into sentences
            sentences = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?|!)\s', paragraph)  #splits paragraph into sentences
            for sentence in sentences:
                if sentence.strip():  # Ensure non-empty sentences
                    # Add the sentence to the current slide text
                    current_slide_text.append(sentence.strip())

                    # If the slide is full, add a new slide
                    if len(current_slide_text) == max_lines_per_slide:
                        add_slide(prs, f"Key Points from {filename}", current_slide_text)
                        current_slide_text = []

        # Add remaining text to a new slide
        if current_slide_text:
            add_slide(prs, f"Additional Info from {filename}", current_slide_text)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
