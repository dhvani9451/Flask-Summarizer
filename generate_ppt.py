from pptx import Presentation
from pptx.util import Pt, Inches
import io
import re
import os
from typing import Dict, List

def clean_text(text: str) -> List[str]:
    """
    Cleans and structures text while preserving all punctuation.
    Properly splits into bullet points at sentence boundaries and paragraphs.
    """
    if isinstance(text, list):
        text = "\n".join(text)

    # Preserve all punctuation and normalize whitespace
    text = re.sub(r'[ \t]+', ' ', text.strip())
    text = re.sub(r'\n\s*\n', '\n\n', text)  # Preserve paragraph breaks

    # Split into meaningful chunks for bullet points
    chunks = []
    paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]

    for para in paragraphs:
        # Handle numbered lists (1. item 2. item)
        if re.search(r'^\d+\.\s', para):
            numbered_items = re.split(r'(?<=\d\.)\s', para)
            chunks.extend([item.strip() for item in numbered_items if item.strip()])
        else:
            # Split at sentence boundaries while preserving punctuation
            sentences = re.findall(r'[^.!?]+[.!?]', para)
            if sentences:
                chunks.extend([s.strip() for s in sentences if s.strip()])
            else:
                chunks.append(para)  # Fallback for text without sentence endings

    return chunks

def add_slide(prs: Presentation, title: str, content: List[str]):
    """Adds a slide with proper text formatting and bullet points."""
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)

    # Set slide title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Set content with proper text wrapping
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.auto_size = True  # Allow text box to expand

    # Configure margins to prevent overflow
    text_frame.margin_left = Inches(0.5)
    text_frame.margin_right = Inches(0.5)
    text_frame.margin_top = Inches(0.5)
    text_frame.margin_bottom = Inches(0.5)

    for item in content:
        p = text_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(16)
        p.space_after = Pt(8)
        p.level = 0  # Always add bullet points since we pre-split properly

        # Handle text that's too long by splitting into multiple paragraphs
        if len(item) > 120:
            words = item.split()
            current_line = words[0]
            for word in words[1:]:
                if len(current_line) + len(word) + 1 < 80:  # 80 char line limit
                    current_line += " " + word
                else:
                    cont_p = text_frame.add_paragraph()
                    cont_p.text = current_line
                    cont_p.font.size = Pt(16)
                    cont_p.space_after = Pt(8)
                    cont_p.level = 1  # Continuation as sub-bullet
                    current_line = word
            if current_line:
                cont_p = text_frame.add_paragraph()
                cont_p.text = current_line
                cont_p.font.size = Pt(16)
                cont_p.space_after = Pt(8)
                cont_p.level = 1

def create_presentation(file_texts: Dict[str, str]):
    """Creates PowerPoint presentation with proper formatting."""
    prs = Presentation("Ion.pptx") if os.path.exists("Ion.pptx") else Presentation()

    # Title Slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Vehicle Rental System - Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"

    for filename, text in file_texts.items():
        # File title slide
        title_slide = prs.slides.add_slide(prs.slide_layouts[5])
        title_slide.shapes.title.text = f"Summary of {filename}"

        # Process content with proper punctuation and bullet points
        structured_text = clean_text(text)
        
        # Split into slides with max 5 bullet points (to prevent overflow)
        for i in range(0, len(structured_text), 5):
            slide_title = (
                f"Key Points from {filename}" 
                if i == 0 else 
                f"Continued Points from {filename}"
            )
            add_slide(
                prs, 
                slide_title,
                structured_text[i:i+5]  # Fewer points per slide
            )

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
