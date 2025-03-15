from pptx import Presentation
from pptx.util import Pt
import io
import re

def clean_text(text):
    if isinstance(text, list):
        text = "\n".join(text)  

    text = re.sub(r'[*#]', '', text)  
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)  
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text

def generate_numbered_bullets(text, lines_per_bullet=6):
    """
    Splits text into numbered bullet points every 6-7 sentences.
    """
    sentences = re.split(r'(?<=[.!?])\s+', text)  # Split sentences properly
    bullets = []
    buffer = []
    bullet_number = 1  # Start numbering from 1

    for i, sentence in enumerate(sentences):
        buffer.append(sentence.strip())

        # ✅ Every 6-7 sentences, create a new numbered bullet point
        if (i + 1) % lines_per_bullet == 0 or i == len(sentences) - 1:
            bullets.append(f"{bullet_number}. " + " ".join(buffer))  # Convert to single bullet
            buffer = []  # Reset buffer
            bullet_number += 1  # Increment bullet number

    return bullets

def add_slide(prs, title, bullet_points):
    """
    Adds a slide with a title and grouped bullet points.
    """
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)

    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)  # Big readable title

    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()  # Remove default text

    for bullet in bullet_points:
        p = text_frame.add_paragraph()
        p.text = bullet
        p.font.size = Pt(22)  # Bigger font for readability
        p.space_after = Pt(10)  # Space between bullet points
        p.level = 0  # Keep all at same bullet level

def create_presentation(title, text):
    """
    Generates a PowerPoint presentation with properly formatted numbered bullet points.
    """
    prs = Presentation("Ion.pptx")  # Use Ion Theme for design

    # ✅ Title Slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = "Generated by AI"

    # ✅ Convert text into numbered bullet points (6-7 lines per bullet)
    bullet_points = generate_numbered_bullets(text, lines_per_bullet=6)

    # ✅ Add slides with bullets
    max_bullets_per_slide = 4  # Limit to prevent overcrowding
    current_slide_bullets = []

    for bullet in bullet_points:
        current_slide_bullets.append(bullet)

        if len(current_slide_bullets) >= max_bullets_per_slide:
            add_slide(prs, "Key Points", current_slide_bullets)
            current_slide_bullets = []  # Reset for next slide

    # ✅ Add any remaining bullets
    if current_slide_bullets:
        add_slide(prs, "Additional Information", current_slide_bullets)

    # ✅ Save presentation to memory
    pptx_io = "/tmp/Generated_Summary_Presentation.pptx"
    prs.save(pptx_io)
    return pptx_io
