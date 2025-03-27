from pptx import Presentation
from pptx.util import Pt, Inches
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures the text into complete paragraphs with preserved punctuation."""
    if isinstance(text, list):
        text = "\n".join(text)

    # Preserve essential punctuation and symbols
    text = re.sub(r'[^\w\s.,!?;:\'’"()%#—-]', '', text)  # Keep relevant punctuation
    text = re.sub(r'\s+', ' ', text).strip()  # Collapse multiple spaces
    text = re.sub(r'\n+', '\n', text)  # Normalize newlines

    # Split and merge into complete paragraphs
    paragraphs = []
    current_para = []
    
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            if current_para:
                paragraphs.append(" ".join(current_para))
                current_para = []
            continue
            
        # Merge continuation lines
        if current_para and (line[0].islower() or not current_para[-1].endswith(('.', '!', '?'))):
            current_para[-1] += ' ' + line
        else:
            current_para.append(line)
    
    if current_para:
        paragraphs.append(" ".join(current_para))
    
    return paragraphs

def add_slide(prs, title, content):
    """Adds a slide with proper bullet points for full paragraphs."""
    slide_layout = prs.slide_layouts[1]  # Title & Content
    slide = prs.slides.add_slide(slide_layout)

    # Set title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Set content with paragraph bullets
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True

    for paragraph in content:
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.font.size = Pt(16)
        p.space_after = Pt(8)
        p.level = 0  # Top-level bullet

def create_presentation(file_texts):
    """Creates presentation with proper paragraph bullets."""
    template_path = "Ion.pptx"
    prs = Presentation(template_path) if os.path.exists(template_path) else Presentation()

    # Title Slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Vehicle Rental System - Summary Presentation"
    title_slide.placeholders[1].text = "Created by AI PPT Generator"

    max_bullets_per_slide = 5  # Paragraphs per slide
    max_chars_per_bullet = 500  # Let PowerPoint handle wrapping

    for filename, text in file_texts.items():
        structured_text = clean_text(text)
        current_bullets = []

        # File title slide
        title_slide = prs.slides.add_slide(prs.slide_layouts[5])
        title_slide.shapes.title.text = f"Summary of {filename}"

        for paragraph in structured_text:
            current_bullets.append(paragraph)
            
            if len(current_bullets) >= max_bullets_per_slide:
                add_slide(prs, f"Key Points from {filename}", current_bullets)
                current_bullets = []

        if current_bullets:
            add_slide(prs, f"Key Points from {filename}", current_bullets)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
