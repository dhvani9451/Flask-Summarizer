from pptx import Presentation
from pptx.util import Pt
import io
import re
import os
import spacy

# Load English NLP model
nlp = spacy.load("en_core_web_sm")

def group_related_sentences(text):
    """Group sentences into logical paragraphs using NLP"""
    doc = nlp(text)
    paragraphs = []
    current_para = []
    
    for sent in doc.sents:
        sent_text = sent.text.strip()
        if not sent_text:
            continue
            
        # Start new paragraph for transition words or long pauses
        if (any(sent_text.lower().startswith(word) for word in ["first", "second", "however", "additionally"]) \
           or len(current_para) >= 3:
            if current_para:
                paragraphs.append(" ".join(current_para))
            current_para = [sent_text]
        else:
            current_para.append(sent_text)
    
    if current_para:
        paragraphs.append(" ".join(current_para))
    
    return paragraphs

def add_slide(prs, title, paragraphs):
    """Add slide with proper bullet grouping"""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    
    # Title setup
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Content setup
    body_shape = slide.shapes.placeholders[1]
    text_frame = body_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True

    MAX_LINES = 6
    current_lines = 0

    for para in paragraphs:
        if current_lines >= MAX_LINES:
            break
            
        # Add main bullet point
        p = text_frame.add_paragraph()
        p.text = para[:80]  # Trim long paragraphs
        p.level = 0
        p.font.size = Pt(16)
        current_lines += 1

        # Handle overflow with sub-bullets
        if len(para) > 80:
            wrapped = textwrap.wrap(para[80:], width=80)
            for line in wrapped:
                if current_lines >= MAX_LINES:
                    break
                sub_p = text_frame.add_paragraph()
                sub_p.text = line
                sub_p.level = 1  # Sub-bullet
                sub_p.font.size = Pt(14)
                current_lines += 1

def create_presentation(file_texts):
    prs = Presentation()
    
    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Vehicle Rental System - Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"

    for filename, text in file_texts.items():
        # Process text
        paragraphs = group_related_sentences(text)
        
        # Split into slide-sized chunks
        for i in range(0, len(paragraphs), 3):  # 3 paragraphs per slide
            chunk = paragraphs[i:i+3]
            title = f"Key Points from {filename}" if i == 0 else f"{filename} (Continued)"
            add_slide(prs, title, chunk)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io