from pptx import Presentation
from pptx.util import Pt, Inches
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures text for PowerPoint slides"""
    if isinstance(text, list):
        text = "\n".join(text)

    # Basic cleaning
    text = re.sub(r'[*#]', '', text)
    text = re.sub(r'[^A-Za-z0-9.,;\-\s]', '', text)  # Keep basic punctuation
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Split into meaningful chunks
    paragraphs = []
    for chunk in re.split(r'(?<=[.!?])\s+', text):
        if chunk:
            # Split long sentences into smaller parts
            if len(chunk) > 120:
                for sentence in textwrap.wrap(chunk, width=120):
                    paragraphs.append(sentence)
            else:
                paragraphs.append(chunk)
    
    return paragraphs

def add_slide(prs, title, content):
    """Adds a slide with proper formatting"""
    slide_layout = prs.slide_layouts[1]  # Title and Content layout
    slide = prs.slides.add_slide(slide_layout)

    # Set title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Set content
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True

    for i, paragraph in enumerate(content):
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.font.size = Pt(18)
        p.space_after = Pt(12)
        p.level = 0  # Top-level bullet point

def create_presentation(file_texts):
    """Creates PowerPoint presentation from processed text"""
    # Try to use template if available
    template_path = "template.pptx"
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()
        # Set default slide size (16:9)
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)

    # Add title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "AI Generated Summary"
    subtitle = title_slide.placeholders[1]
    subtitle.text = "Created with Gemini AI"

    # Process each file's content
    for filename, text in file_texts.items():
        print(f"Processing content for: {filename}")
        
        # Add section title slide
        section_slide = prs.slides.add_slide(prs.slide_layouts[5])
        section_slide.shapes.title.text = f"Summary: {filename}"

        # Clean and structure text
        structured_text = clean_text(text)
        
        # Split content into slides (5-7 points per slide)
        points_per_slide = 6
        for i in range(0, len(structured_text), points_per_slide):
            slide_content = structured_text[i:i + points_per_slide]
            slide_title = f"Key Points {i//points_per_slide + 1}"
            add_slide(prs, slide_title, slide_content)

    # Save to bytes buffer
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
