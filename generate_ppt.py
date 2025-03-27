from pptx import Presentation
from pptx.util import Pt, Inches
import io
import re
import os

def clean_text(text):
    """Cleans and structures the extracted text while preserving punctuation."""
    if isinstance(text, list):
        text = "\n".join(text)
    
    text = re.sub(r'[*#]', '', text)  # Remove unwanted characters
    text = re.sub(r'[^A-Za-z0-9.,!?;:\-\s]', '', text)  # Keep relevant punctuation
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces
    text = re.sub(r'\n+', '\n', text)  # Remove extra newlines
    
    # Ensure full stops are preserved and split text correctly
    paragraphs = re.split(r'(?<=\.)\s+', text)
    structured_text = [para.strip() for para in paragraphs if para]
    
    print("Debug - Structured Text:", structured_text)  # Debugging
    
    return structured_text

def add_slide(prs, title, content):
    """Adds a slide ensuring bullet points are applied to full paragraphs while preventing overflow."""
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)
    
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)
    
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True  # Enable word wrapping
    
    max_chars_per_slide = 600  # Limit characters per slide
    current_text = ""
    
    for paragraph in content:
        temp_text = paragraph.strip()
        if len(current_text) + len(temp_text) > max_chars_per_slide:
            break  # Stop adding text if exceeding limit
        
        p = text_frame.add_paragraph()
        p.text = temp_text
        p.font.size = Pt(14)  # Reduce font size slightly for better fit
        p.space_after = Pt(6)  # Improve readability
        p.level = 0  # Default bullet level
        current_text += temp_text + " "

def create_presentation(file_texts):
    """Creates a PowerPoint presentation with structured slides while preventing text overflow."""
    template_path = "Ion.pptx"
    prs = Presentation(template_path) if os.path.exists(template_path) else Presentation()
    
    # Add Title Slide
    slide_layout = prs.slide_layouts[0]  # Title Slide Layout
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Vehicle Rental System - Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"
    
    max_paragraphs_per_slide = 6  # Control slide content size
    
    for filename, text in file_texts.items():
        structured_text = clean_text(text)
        current_slide_text = []
        
        # Add a Slide for Each File Title
        file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only Layout
        file_title_slide.shapes.title.text = f"Summary of {filename}"
        
        for paragraph in structured_text:
            if paragraph:
                current_slide_text.append(paragraph)
                
                if len(current_slide_text) == max_paragraphs_per_slide:
                    add_slide(prs, f"Key Points from {filename}", current_slide_text)
                    current_slide_text = []
        
        # Add remaining text to a new slide
        if current_slide_text:
            add_slide(prs, f"Additional Info from {filename}", current_slide_text)
    
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
