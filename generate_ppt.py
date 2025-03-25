from pptx import Presentation
from pptx.util import Pt, Inches
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures the extracted text while preserving meaningful breaks."""
    if isinstance(text, list):
        text = "\n".join(text)

    # Preserve sentence terminators, bullets, and meaningful line breaks
    text = re.sub(r'[*#]', '', text)
    text = re.sub(r'[^A-Za-z0-9.,;!?:\-\s\n]', '', text)  # Keep essential punctuation
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r'\n\s*\n', '\n\n', text)  # Preserve paragraph breaks

    # Split into paragraphs first (double newlines)
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    
    # Then split each paragraph into sentences
    structured_text = []
    for para in paragraphs:
        # Split at sentence boundaries but keep the delimiter
        sentences = re.split(r'(?<=[.!?])\s+', para)
        structured_text.extend(s for s in sentences if s)
    
    return structured_text

def add_slide(prs, title, content):
    """Adds a slide with proper text wrapping and bullet control."""
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)

    # Set slide title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Set slide content with controlled bullets
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.auto_size = True  # Allow textbox to expand

    # Track if we're continuing a thought
    current_thought = []
    for i, paragraph in enumerate(content):
        # Start new bullet for proper sentences
        if i == 0 or (paragraph and paragraph[0].isupper()):
            if current_thought:  # Add accumulated thought first
                p = text_frame.add_paragraph()
                p.text = " ".join(current_thought)
                p.font.size = Pt(16)
                p.space_after = Pt(8)
                p.level = 1  # No bullet for continuations
                current_thought = []
            
            p = text_frame.add_paragraph()
            p.text = paragraph
            p.font.size = Pt(16)
            p.space_after = Pt(8)
            p.level = 0  # Add bullet
        else:
            current_thought.append(paragraph)
    
    # Add any remaining text
    if current_thought:
        p = text_frame.add_paragraph()
        p.text = " ".join(current_thought)
        p.font.size = Pt(16)
        p.space_after = Pt(8)
        p.level = 1  # No bullet

def create_presentation(file_texts):
    """Creates a PowerPoint presentation with proper text handling."""
    template_path = "Ion.pptx"
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # Add Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Vehicle Rental System - Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"

    max_chars_per_slide = 1000  # Character limit per slide
    max_lines_per_slide = 8     # Line limit per slide

    for filename, text in file_texts.items():
        structured_text = clean_text(text)
        current_slide_content = []
        current_char_count = 0

        # Add section title slide
        section_slide = prs.slides.add_slide(prs.slide_layouts[5])
        section_slide.shapes.title.text = f"Summary of {filename}"

        for paragraph in structured_text:
            # Check if adding this would exceed limits
            if (current_char_count + len(paragraph) > max_chars_per_slide or 
                len(current_slide_content) >= max_lines_per_slide):
                add_slide(prs, f"Key Points from {filename}", current_slide_content)
                current_slide_content = []
                current_char_count = 0
            
            current_slide_content.append(paragraph)
            current_char_count += len(paragraph)

        # Add remaining content
        if current_slide_content:
            add_slide(prs, f"Additional Info from {filename}", current_slide_content)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
