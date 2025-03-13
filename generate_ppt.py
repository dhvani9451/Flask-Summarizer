from pptx import Presentation
from pptx.util import Inches, Pt
import io
import textwrap
import re

def clean_text(text):
    """Remove special characters and extra bullet points."""
    text = re.sub(r'[*]+', '', text)  # Remove *, **, ***
    text = re.sub(r'•+', '•', text)  # Remove duplicate bullet points
    text = re.sub(r'\n+', '\n', text.strip())  # Remove extra new lines
    return text

def split_into_paragraphs_and_bullets(text):
    """
    Splits text into:
    - Bullet points for short sentences.
    - Paragraphs for long continuous text.
    """
    sentences = text.split("\n")  # Split by lines
    structured_content = []
    
    for sentence in sentences:
        sentence = sentence.strip()
        
        if len(sentence) == 0:
            continue  # Skip empty lines
        
        # If the text has more than 2 sentences without full stops, treat it as a paragraph
        if len(sentence.split(".")) <= 2 and len(sentence) < 120:
            structured_content.append(f"• {sentence}")  # Make it a bullet point
        else:
            structured_content.append(sentence)  # Keep it as a paragraph
    
    return structured_content

def add_slide(prs, title, content):
    """Add a slide with bullet points and paragraphs."""
    slide_layout = prs.slide_layouts[5]  # Title Only layout
    slide = prs.slides.add_slide(slide_layout)
    
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(24)

    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8.5), Inches(5))
    text_frame = textbox.text_frame
    text_frame.word_wrap = True

    for paragraph in content:
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.space_after = Pt(10)  # ✅ Add spacing between paragraphs
        p.font.size = Pt(18)  # ✅ Set a readable font size

def create_presentation(title, raw_text):
    """Create a PowerPoint presentation with structured content."""
    prs = Presentation("Ion.pptx")  # ✅ Load Ion theme

    # ✅ Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = "Created by AI-Powered Summary Generator"

    # ✅ Clean and structure text
    cleaned_text = clean_text(raw_text)
    structured_text = split_into_paragraphs_and_bullets(cleaned_text)

    # ✅ Create slides
    max_lines_per_slide = 10
    current_slide_text = []

    for line in structured_text:
        if len(current_slide_text) == max_lines_per_slide:
            add_slide(prs, "Key Points", current_slide_text)
            current_slide_text = []
        
        current_slide_text.append(line)
    
    if current_slide_text:
        add_slide(prs, "Key Points", current_slide_text)

    # ✅ Save PPT as bytes
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
