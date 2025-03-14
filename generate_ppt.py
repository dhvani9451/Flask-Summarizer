from pptx import Presentation
from pptx.util import Inches, Pt
import io
import textwrap
import re


def clean_text(text):
    if isinstance(text, list):
        text = "\n".join(text)  

    # ✅ Remove all special characters except letters, numbers, spaces, periods, and commas
    text = re.sub(r'[*\"\'@#%^&()_+=<>?/|{}[\]\\]', '', text)

    # ✅ Ensure proper sentence structure and spacing
    text = re.sub(r'\s+', ' ', text).strip()
    lines = text.split(". ")  # Keep sentences separate

    structured_text = []
    for line in lines:
        line = line.strip()
        if line:  
            structured_text.append(line)

    return structured_text

def add_slide(prs, title, content):
    slide_layout = prs.slide_layouts[5]  # Title Only layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Title Formatting
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(24)

    # Text Box for Content
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8.5), Inches(5))
    text_frame = textbox.text_frame
    text_frame.word_wrap = True

    # Formatting content with bullet points & spacing
    for i, paragraph in enumerate(content):
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.font.size = Pt(16)
        
        if len(paragraph.split()) <= 20 and "." not in paragraph:
            p.level = 0  # Bullet Point
        else:
            p.level = 1  # Paragraph (No bullet)
        
        # Add a new line after every 3 points for readability
        if (i + 1) % 3 == 0:
            text_frame.add_paragraph().text = ""


def create_presentation(title, text_data):
    prs = Presentation("Ion.pptx")  # Load Ion theme

    # Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = "Created by PPT Summ"

    # Organize content into slides with proper spacing
    max_words_per_line = 12
    max_lines_per_slide = 8  # More spacing for readability
    current_slide_text = []

    for line in text_data:
        wrapped_lines = textwrap.wrap(line, width=max_words_per_line * 6)
        for wrapped_line in wrapped_lines:
            current_slide_text.append(wrapped_line)
            
            if len(current_slide_text) == max_lines_per_slide:
                add_slide(prs, "Key Points", current_slide_text)
                current_slide_text = []

    # Add any remaining text
    if current_slide_text:
        add_slide(prs, "Additional Information", current_slide_text)

    # Save PPT to memory for downloading
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
