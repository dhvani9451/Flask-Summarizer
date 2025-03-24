from pptx import Presentation
from pptx.util import Pt
import io
import re
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def clean_text(text):
    """Clean and split text into logical chunks."""
    if isinstance(text, list):
        text = "\n".join(text)
    
    # Normalize whitespace and preserve sentence structure
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Split into bullet points if already formatted
    if '•' in text or '-' in text:
        return [line.strip(' •-') for line in text.split('\n') if line.strip()]
    
    # Fallback: Split by periods or newlines
    return [s.strip() for s in re.split(r'(?<=[.!?])\s+|\n', text) if s.strip()]

def add_slide(prs, title, content):
    """Add a slide with properly formatted bullet points."""
    slide_layout = prs.slide_layouts[1]  # Title + Content layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Set title
    slide.shapes.title.text = title
    slide.shapes.title.text_frame.paragraphs[0].font.bold = True
    slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(24)
    
    # Configure content
    content_box = slide.shapes.placeholders[1]
    text_frame = content_box.text_frame
    text_frame.clear()
    text_frame.word_wrap = True  # Prevent overflow
    
    # Add content with smart bullets
    for i, line in enumerate(content):
        p = text_frame.add_paragraph()
        p.text = line
        p.font.size = Pt(16)
        p.space_after = Pt(8)
        
        # Apply bullets only to new thoughts (first line or after a blank line)
        p.level = 0 if (i == 0 or not content[i-1]) else 1

def create_presentation(file_texts):
    """Generate PowerPoint from processed text."""
    try:
        prs = Presentation()
        
        # Title slide
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_slide.shapes.title.text = "Summary Presentation"
        title_slide.placeholders[1].text = "Created by AI"
        
        # Process each file
        for filename, text in file_texts.items():
            # Clean and split text
            chunks = clean_text(text)
            
            # Split into slide-sized groups (max 4 chunks per slide)
            for i in range(0, len(chunks), 4):
                slide_title = f"{filename}" if i == 0 else f"{filename} (Continued)"
                add_slide(prs, slide_title, chunks[i:i+4])
        
        # Save to memory
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)
        return ppt_io
        
    except Exception as e:
        logger.error(f"PPT generation failed: {str(e)}")
        raise  # Ensure Flask catches the error
