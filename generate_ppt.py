from pptx import Presentation
from pptx.util import Pt, Inches
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures the extracted text based on semantic patterns."""
    if isinstance(text, list):
        text = "\n".join(text)
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text).strip()
    sentences = text.split('. ')
    return [s.strip() for s in sentences if s]

def add_slide(prs, title, content):
    """Adds a slide ensuring bullet points are applied only to new thoughts."""
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    
    max_chars_per_line = 80
    max_lines_per_slide = 6
    current_lines = 0
    
    for paragraph in content:
        wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line)
        for wrapped_line in wrapped_lines:
            if current_lines >= max_lines_per_slide:
                return  # Prevent overflow
            p = text_frame.add_paragraph()
            p.text = wrapped_line
            p.font.size = Pt(18)
            p.space_after = Pt(10)
            current_lines += 1

def create_presentation(file_texts):
    """Creates a PowerPoint presentation and returns it as a file object."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"
    
    for filename, text in file_texts.items():
        structured_text = clean_text(text)
        file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])
        file_title_slide.shapes.title.text = f"📄 {filename}"
        current_slide_text = []
        
        for line in structured_text:
            if len(current_slide_text) >= 6:
                add_slide(prs, f"Key Points - {filename}", current_slide_text)
                current_slide_text = []
            current_slide_text.append(line)
        
        if current_slide_text:
            add_slide(prs, f"Additional Info - {filename}", current_slide_text)
    
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
