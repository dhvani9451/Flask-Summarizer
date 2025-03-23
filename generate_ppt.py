from pptx import Presentation
from pptx.util import Pt
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures the extracted text based on semantic patterns."""
    if isinstance(text, list):
        text = "\n".join(text)

    # Remove unwanted characters
    text = re.sub(r'[*#]', '', text)
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)  # Keep only letters, numbers, and punctuation
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces
    text = re.sub(r'\n+', '\n', text)  # Remove extra newlines

    # Define semantic patterns to identify new thoughts
    patterns = ["Firstly", "Secondly", "In addition", "Moreover", "However", "Therefore", "Finally", "On the other hand"]

    # Split text into sentences based on semantic patterns
    structured_text = []
    current_sentence = ""
    for word in text.split():
        current_sentence += word + " "
        # Check if the current word matches any of the semantic patterns
        if any(word.startswith(pattern) for pattern in patterns):
            structured_text.append(current_sentence.strip())
            current_sentence = ""
    # Add the last sentence if it exists
    if current_sentence:
        structured_text.append(current_sentence.strip())

    return structured_text

def add_slide(prs, title, content):
    """Adds a slide ensuring bullet points are applied only to new thoughts."""
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

    # Define semantic patterns to identify new thoughts
    patterns = ["Firstly", "Secondly", "In addition", "Moreover", "However", "Therefore", "Finally", "On the other hand"]

    # Add bullet points only for new thoughts
    for i, paragraph in enumerate(content):
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.font.size = Pt(18)
        p.space_after = Pt(10)

        # Apply bullet only if it's a new thought (matches semantic patterns)
        if i == 0 or any(paragraph.startswith(pattern) for pattern in patterns):
            p.level = 0  # PowerPoint's default bullet
        else:
            p.level = 1  # No bullet for wrapped lines or continuous text

def create_presentation(file_texts):
    """Creates a PowerPoint presentation and returns it as a file object."""
    # Use a template if available
    template_path = "Ion.pptx"
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # Add Title Slide
    slide_layout = prs.slide_layouts[0]  # Title Slide Layout
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"

    max_lines_per_slide = 6

    # Process each file's text
    for filename, text in file_texts.items():
        structured_text = clean_text(text)
        current_slide_text = []

        # Add a Slide for Each File Title
        file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only Layout
        file_title_slide.shapes.title.text = f"📄 {filename}"  # Show filename as slide title

        # Add slides for the content
        for line in structured_text:
            wrapped_lines = textwrap.wrap(line, width=80)
            for wrapped_line in wrapped_lines:
                current_slide_text.append(wrapped_line)

                # Add a new slide if the maximum lines per slide is reached
                if len(current_slide_text) == max_lines_per_slide:
                    add_slide(prs, f"Key Points - {filename}", current_slide_text)
                    current_slide_text = []

        # Add remaining content to a new slide
        if current_slide_text:
            add_slide(prs, f"Additional Info - {filename}", current_slide_text)

    # Save the presentation to a BytesIO object
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
