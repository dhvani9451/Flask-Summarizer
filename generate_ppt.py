from pptx import Presentation
from pptx.util import Pt
import io
import re
import os

def clean_text(text):
    """Splits text into logical paragraphs with sentence-aware splitting."""
    if isinstance(text, list):
        text = "\n".join(text)

    # Preserve punctuation and clean text
    text = re.sub(r'[^\w\s.,!?;:\'’"()%#—-]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Split into sentences
    sentences = re.split(r'(?<=[.!?])\s+(?=[A-Z])', text)  # Sentence boundary detection
    
    # Merge into paragraphs of 1-3 sentences
    paragraphs = []
    current_para = []
    for sentence in sentences:
        current_para.append(sentence)
        if len(current_para) >= 2 or len(' '.join(current_para)) > 300:
            paragraphs.append(' '.join(current_para))
            current_para = []
    if current_para:
        paragraphs.append(' '.join(current_para))
    
    return paragraphs

def add_slide(prs, title, content):
    """Adds slide with controlled text density."""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)

    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    content_shape = slide.shapes.placeholders[1]
    tf = content_shape.text_frame
    tf.clear()

    for para in content:
        p = tf.add_paragraph()
        p.text = para
        p.font.size = Pt(18)
        p.space_after = Pt(12)
        p.level = 0

def create_presentation(file_texts):
    prs = Presentation("Ion.pptx") if os.path.exists("Ion.pptx") else Presentation()

    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Vehicle Rental System - Summary Presentation"
    title_slide.placeholders[1].text = "Created by AI PPT Generator"

    for filename, text in file_texts.items():
        paragraphs = clean_text(text)
        chunk_size = max(2, min(4, len(paragraphs)))  # 2-4 paragraphs per slide

        # Split paragraphs into slide-sized chunks
        for i in range(0, len(paragraphs), chunk_size):
            slide_paragraphs = paragraphs[i:i+chunk_size]
            title = f"Key Points from {filename}" if i == 0 else f"{filename} Cont."
            add_slide(prs, title, slide_paragraphs)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
