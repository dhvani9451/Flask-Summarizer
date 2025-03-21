from pptx import Presentation
from pptx.util import Pt
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures the extracted text into distinct sentences and paragraphs."""
    if isinstance(text, list):
        text = "\n".join(text)

    text = re.sub(r'[*#]', '', text)  # Remove unwanted characters
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)  # Keep only letters, numbers, and punctuation
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces
    text = re.sub(r'\n+', '\n', text)  # Remove extra newlines

    # ✅ Split into sentences while keeping full stops
    sentences = re.split(r'(?<=[.!?])\s+', text)

    structured_text = [sentence.strip() for sentence in sentences if sentence]
    return structured_text

def add_slide(prs, title, content):
    """Adds a slide ensuring proper bullet point formatting for new sentences or paragraphs."""
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)

    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()

    is_new_point = True  # ✅ Ensures bullets are only for new points

    for paragraph in content:
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.font.size = Pt(18)
        p.space_after = Pt(10)

        # ✅ Apply bullet points **only for new sentences or paragraphs**
        if is_new_point:
            p.level = 0  # ✅ Default PowerPoint bullet
        else:
            p.level = 1  # ✅ No bullet for continuous lines

        # ✅ Only apply a bullet if it's a new sentence (ends with full stop)
        is_new_point = paragraph.endswith(("."))

def create_presentation(file_texts):
    """Creates a PowerPoint presentation and returns it as a file object."""
    template_path = "Ion.pptx"
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # ✅ Add Title Slide
    slide_layout = prs.slide_layouts[0]  # Title Slide Layout
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"

    max_lines_per_slide = 6

    for filename, text in file_texts.items():
        structured_text = clean_text(text)
        current_slide_text = []

        # ✅ Add a Slide for Each File Title
        file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only Layout
        file_title_slide.shapes.title.text = f"📄 {filename}"  # Show filename as slide title

        for line in structured_text:
            wrapped_lines = textwrap.wrap(line, width=80)
            for wrapped_line in wrapped_lines:
                current_slide_text.append(wrapped_line)

                if len(current_slide_text) == max_lines_per_slide:
                    add_slide(prs, f"Key Points - {filename}", current_slide_text)
                    current_slide_text = []

        if current_slide_text:
            add_slide(prs, f"Additional Info - {filename}", current_slide_text)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
