from pptx import Presentation
from pptx.util import Pt
import io
import re
import os
import textwrap
from typing import List

# Semantic patterns to identify new thoughts
THOUGHT_STARTERS = [
    "First", "Firstly", "Second", "Secondly", "Third", "Thirdly",
    "Next", "Then", "Additionally", "Moreover", "Furthermore",
    "However", "Therefore", "Thus", "Hence", "Consequently",
    "Meanwhile", "Specifically", "For example", "In conclusion"
]

def is_new_thought(text: str) -> bool:
    """Check if text starts with a thought-starter phrase"""
    first_word = text.split()[0].rstrip('.,;:').lower()
    return any(starter.lower() == first_word for starter in THOUGHT_STARTERS)

def clean_text(text: str) -> List[str]:
    """Clean and split text into meaningful sentences"""
    if isinstance(text, list):
        text = " ".join(text)

    # Preserve sentence boundaries
    text = re.sub(r'\s+', ' ', text).strip()
    sentences = re.split(r'(?<=[.!?])\s+', text)
    return [s.strip() for s in sentences if s.strip()]

def add_slide(prs, title: str, content: List[str]):
    """Add slide with proper bullet hierarchy"""
    slide_layout = prs.slide_layouts[1]  # Title + Content layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Set title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Configure content
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True  # Enable word wrap

    MAX_LINES = 6  # Optimal for slide readability
    MAX_WIDTH = 80  # Characters per line
    
    current_thought = None
    lines_added = 0

    for sentence in content:
        if lines_added >= MAX_LINES:
            break  # Prevent overflow
            
        if is_new_thought(sentence) or current_thought is None:
            # New thought = new bullet
            p = text_frame.add_paragraph()
            p.text = sentence
            p.level = 0  # Top-level bullet
            current_thought = sentence
            lines_added += 1
        else:
            # Continuation of thought
            p = text_frame.add_paragraph()
            p.text = sentence
            p.level = 1  # Sub-bullet
            lines_added += 1

        # Handle long sentences
        if len(sentence) > MAX_WIDTH:
            wrapped = textwrap.wrap(sentence, width=MAX_WIDTH)
            if len(wrapped) > 1:
                p.text = wrapped[0]
                for line in wrapped[1:]:
                    if lines_added >= MAX_LINES:
                        break
                    cont_p = text_frame.add_paragraph()
                    cont_p.text = line
                    cont_p.level = 1  # Continuation as sub-bullet
                    lines_added += 1

def create_presentation(file_texts):
    """Create presentation with proper formatting"""
    prs = Presentation()
    
    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"

    for filename, text in file_texts.items():
        # Clean and structure text
        sentences = clean_text(text)
        
        # Split into chunks of 4-6 sentences per slide
        chunk_size = 5
        for i in range(0, len(sentences), chunk_size):
            chunk = sentences[i:i + chunk_size]
            title = f"Key Points - {filename}" if i == 0 else f"Continued - {filename}"
            add_slide(prs, title, chunk)

    # Save to BytesIO
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
