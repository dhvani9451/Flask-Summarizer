from pptx import Presentation
from pptx.util import Pt, Inches
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures the extracted text while preserving meaningful punctuation and logical grouping."""
    if isinstance(text, list):
        text = "\n".join(text)

    # Remove unwanted formatting characters but preserve essential punctuation
    text = re.sub(r'[*#]', '', text)
    text = re.sub(r'[^A-Za-z0-9.,!?;:\'\-\s]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r'\n+', '\n', text)

    # Split into paragraphs based on double newlines or sentence boundaries
    paragraphs = []
    current_para = []
    
    for line in text.split('\n'):
        line = line.strip()
        if line:  # If line has content
            # If line ends with sentence-ending punctuation, consider it a complete thought
            if re.search(r'[.!?]\s*$', line):
                current_para.append(line)
                paragraphs.append(' '.join(current_para))
                current_para = []
            else:
                current_para.append(line)
    
    # Add any remaining content
    if current_para:
        paragraphs.append(' '.join(current_para))

    # Split long paragraphs at natural break points
    final_paragraphs = []
    for para in paragraphs:
        if len(para) > 100:  # Split long paragraphs
            # Split at sentence boundaries or conjunctions
            sentences = re.split(r'(?<=[.!?])\s+', para)
            for sentence in sentences:
                if sentence:
                    # Further split at conjunctions if still long
                    if len(sentence) > 80:
                        clauses = re.split(r',\s+(?:and|or)\s+', sentence)
                        final_paragraphs.extend(clauses)
                    else:
                        final_paragraphs.append(sentence)
        else:
            final_paragraphs.append(para)

    return final_paragraphs

def add_slide(prs, title, content):
    """Adds a slide with properly grouped bullet points."""
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)

    # Set slide title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Set slide content with proper bullet hierarchy
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True

    # Track bullet levels for nested structure
    current_level = 0
    
    for paragraph in content:
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.font.size = Pt(18)  # Slightly larger for better readability
        p.space_after = Pt(12)  # Increased spacing
        p.level = current_level
        
        # Detect if this should be a sub-bullet (simple heuristic)
        if len(paragraph) < 60 and not re.search(r'[.!?]$', paragraph):
            current_level = 1
        else:
            current_level = 0

def create_presentation(file_texts):
    """Creates a PowerPoint presentation with proper formatting."""
    template_path = "Ion.pptx"
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # Add Title Slide
    slide_layout = prs.slide_layouts[0]  # Title Slide Layout
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "Vehicle Rental System - Summary Presentation"
    subtitle = slide.placeholders[1]
    subtitle.text = "Created by AI PPT Generator"

    # Slide content parameters
    max_lines_per_slide = 7  # Optimal for readability
    max_chars_per_line = 70  # Better for projection

    for filename, text in file_texts.items():
        structured_text = clean_text(text)
        current_slide_text = []

        # Add section title slide
        section_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only
        section_slide.shapes.title.text = f"Summary of {filename}"

        # Process content in logical chunks
        for paragraph in structured_text:
            # Wrap long lines while preserving punctuation
            wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line, 
                                         break_long_words=False, 
                                         break_on_hyphens=True)
            
            # Add to current slide or create new one
            if len(current_slide_text) + len(wrapped_lines) <= max_lines_per_slide:
                current_slide_text.extend(wrapped_lines)
            else:
                add_slide(prs, f"Key Points from {filename}", current_slide_text)
                current_slide_text = wrapped_lines

        # Add remaining content
        if current_slide_text:
            add_slide(prs, f"Additional Info from {filename}", current_slide_text)

    # Save to memory
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
