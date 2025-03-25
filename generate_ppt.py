from pptx import Presentation
from pptx.util import Pt
import io
import re
import os
import logging
from datetime import datetime

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def clean_text(text):
    """Cleans text while preserving sentence structure and meaningful breaks."""
    try:
        if isinstance(text, list):
            text = "\n".join(text)

        # Preserve essential formatting
        text = re.sub(r'[*#]', '', text)
        text = re.sub(r'[^A-Za-z0-9.,;!?:\-\s\n]', '', text)
        text = re.sub(r'\s+', ' ', text).strip()
        text = re.sub(r'\n\s*\n', '\n\n', text)  # Preserve paragraphs

        # Split into paragraphs then sentences
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
        structured_text = []
        
        for para in paragraphs:
            # Split at sentence boundaries but keep the delimiter
            sentences = re.split(r'(?<=[.!?])\s+', para)
            structured_text.extend(s for s in sentences if s)
        
        return structured_text
    except Exception as e:
        logger.error(f"Error in clean_text: {str(e)}")
        raise

def add_slide(prs, title, content):
    """Adds a slide with proper text handling and bullet points."""
    try:
        slide_layout = prs.slide_layouts[1]  # Title & Content
        slide = prs.slides.add_slide(slide_layout)

        # Set title
        title_shape = slide.shapes.title
        title_shape.text = title[:100]  # Limit title length
        title_shape.text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].font.size = Pt(28)

        # Configure content
        content_shape = slide.shapes.placeholders[1]
        text_frame = content_shape.text_frame
        text_frame.clear()
        text_frame.word_wrap = True

        current_thought = []
        for i, paragraph in enumerate(content):
            paragraph = paragraph[:500]  # Limit paragraph length
            
            # New bullet for new sentences
            if i == 0 or (paragraph and paragraph[0].isupper()):
                if current_thought:
                    p = text_frame.add_paragraph()
                    p.text = " ".join(current_thought)
                    p.font.size = Pt(16)
                    p.level = 1  # No bullet
                    current_thought = []
                
                p = text_frame.add_paragraph()
                p.text = paragraph
                p.font.size = Pt(16)
                p.level = 0  # With bullet
            else:
                current_thought.append(paragraph)

        # Add any remaining text
        if current_thought:
            p = text_frame.add_paragraph()
            p.text = " ".join(current_thought)
            p.font.size = Pt(16)
            p.level = 1
    except Exception as e:
        logger.error(f"Error in add_slide: {str(e)}")
        raise

def create_presentation(file_texts):
    """Creates PowerPoint with robust error handling."""
    try:
        # Initialize presentation
        template_path = "Ion.pptx"
        prs = Presentation(template_path) if os.path.exists(template_path) else Presentation()

        # Add title slide
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_slide.shapes.title.text = "Vehicle Rental System - Summary"
        title_slide.placeholders[1].text = f"Generated on {datetime.now().strftime('%Y-%m-%d')}"

        # Process content
        for filename, text in file_texts.items():
            try:
                content = clean_text(text)
                
                # Add section header
                section_slide = prs.slides.add_slide(prs.slide_layouts[5])
                section_slide.shapes.title.text = f"Summary of {os.path.basename(filename)}"
                
                # Split content into slides
                current_slide_content = []
                char_count = 0
                
                for paragraph in content:
                    if char_count + len(paragraph) > 2000 or len(current_slide_content) >= 10:
                        add_slide(prs, f"Key Points from {os.path.basename(filename)}", current_slide_content)
                        current_slide_content = []
                        char_count = 0
                    
                    current_slide_content.append(paragraph)
                    char_count += len(paragraph)
                
                if current_slide_content:
                    add_slide(prs, f"More from {os.path.basename(filename)}", current_slide_content)
            
            except Exception as e:
                logger.error(f"Error processing {filename}: {str(e)}")
                continue

        # Save to bytes buffer
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)
        return ppt_io

    except Exception as e:
        logger.error(f"Fatal error in create_presentation: {str(e)}")
        raise
