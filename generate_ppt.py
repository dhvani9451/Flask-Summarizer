from pptx import Presentation
from pptx.util import Pt
import io
import re
import os
import textwrap
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def clean_text(text):
    """Cleans and structures the extracted text."""
    try:
        if isinstance(text, list):
            text = "\n".join(text)

        text = re.sub(r'[*#]', '', text)
        text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)
        text = re.sub(r'\s+', ' ', text).strip()
        text = re.sub(r'\n+', '\n', text)  # Remove extra newlines

        # Split text into sentences
        sentences = re.split(r'(?<=[.!?])\s+', text)
        structured_text = [sentence.strip() for sentence in sentences if sentence]

        return structured_text
    except Exception as e:
        logging.error(f"Error cleaning text: {e}")
        return []

def add_slide(prs, title, content):
    """Adds a new slide with title and formatted content."""
    try:
        slide_layout = prs.slide_layouts[1]  # Title & Content layout
        slide = prs.slides.add_slide(slide_layout)

        title_shape = slide.shapes.title
        title_shape.text = title
        title_shape.text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].font.size = Pt(28)

        content_shape = slide.shapes.placeholders[1]
        text_frame = content_shape.text_frame
        text_frame.clear()

        for paragraph in content:
            p = text_frame.add_paragraph()
            p.text = paragraph
            p.font.size = Pt(18)
            p.space_after = Pt(10)
            p.level = 0  # Reset bullet level
            p.bullet = True  # Apply bullet point
    except Exception as e:
        logging.error(f"Error adding slide: {e}")

def create_presentation(file_texts):
    """Creates a PowerPoint presentation and returns it as a file object."""
    try:
        template_path = "Ion.pptx"
        if os.path.exists(template_path):
            prs = Presentation(template_path)
        else:
            prs = Presentation()

        slide_layout = prs.slide_layouts[0]  # Title Slide Layout
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = "Summary Presentation"
        slide.placeholders[1].text = "Created by AI PPT Generator"

        max_lines_per_slide = 6

        for filename, text in file_texts.items():
            structured_text = clean_text(text)
            current_slide_text = []

            file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only Layout
            file_title_slide.shapes.title.text = f"📄 {filename}"

            for sentence in structured_text:
                wrapped_lines = textwrap.wrap(sentence, width=80)
                current_slide_text.append("• " + " ".join(wrapped_lines))  # Add bullet point

                if len(current_slide_text) == max_lines_per_slide:
                    add_slide(prs, f"Key Points - {filename}", current_slide_text)
                    current_slide_text = []

            if current_slide_text:
                add_slide(prs, f"Additional Info - {filename}", current_slide_text)

        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)
        return ppt_io
    except Exception as e:
        logging.error(f"Error creating presentation: {e}")
        raise

# Example usage
# file_texts = {"example.txt": "Your text content here."}
# create_presentation(file_texts)
