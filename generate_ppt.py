from pptx import Presentation
from pptx.util import Pt, Inches
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures the extracted text while preserving meaningful structure."""
    if isinstance(text, list):
        text = "\n".join(text)

    # Remove markdown characters but preserve most punctuation
    text = re.sub(r'[*#]', '', text)
    
    # Normalize whitespace (convert multiple spaces/tabs to single space)
    text = re.sub(r'[ \t]+', ' ', text)
    
    # Normalize line breaks (convert multiple newlines to double newlines)
    text = re.sub(r'\n\s*\n', '\n\n', text.strip())
    
    # Split into paragraphs based on double newlines
    paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
    
    # Further split long paragraphs into coherent chunks
    structured_text = []
    for para in paragraphs:
        # Split at sentence boundaries if paragraph is too long
        if len(para) > 120:
            sentences = re.split(r'(?<=[.!?])\s+', para)
            current_chunk = []
            for sentence in sentences:
                if sum(len(s) for s in current_chunk) + len(sentence) < 100:
                    current_chunk.append(sentence)
                else:
                    structured_text.append(' '.join(current_chunk))
                    current_chunk = [sentence]
            if current_chunk:
                structured_text.append(' '.join(current_chunk))
        else:
            structured_text.append(para)
    
    return structured_text

def add_slide(prs, title, content):
    """Adds a slide with proper bullet point structure."""
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

    # Add content with bullet points only for main points
    for paragraph in content:
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.font.size = Pt(16)
        p.space_after = Pt(8)
        
        # Only add bullet if it's a complete thought (not a continuation)
        if len(paragraph.split()) > 3 and paragraph[-1] in '.!?':
            p.level = 0  # Add bullet point
        else:
            p.level = 1  # Sub-point without bullet

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

        # Group content into logical chunks for slides
        current_chunk = []
        char_count = 0
        
        for paragraph in structured_text:
            # Estimate if adding this paragraph would exceed slide limits
            if (len(current_chunk) >= max_lines_per_slide or 
                char_count + len(paragraph) > max_chars_per_line * max_lines_per_slide):
                add_slide(prs, f"Key Points from {filename}", current_chunk)
                current_chunk = []
                char_count = 0
            
            current_chunk.append(paragraph)
            char_count += len(paragraph)

        # Add remaining content
        if current_chunk:
            add_slide(prs, f"Additional Info from {filename}", current_chunk)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
