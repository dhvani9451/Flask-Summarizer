from pptx import Presentation
from pptx.util import Pt
import io
import textwrap
import re

def clean_text(text):
    if isinstance(text, list):
        text = "\n".join(text)  

    text = re.sub(r'[*#]', '', text)  
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)  
    text = re.sub(r'\s+', ' ', text).strip()
    
    lines = text.split(". ")  

    structured_text = []
    for line in lines:
        line = line.strip()
        if line:  
            structured_text.append(line)

    return structured_text

def add_slide(prs, title, content):
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)

    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)  

    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()  

    grouped_text = []
    for i, paragraph in enumerate(content):
        grouped_text.append(paragraph)

        if (i + 1) % 5 == 0 or i == len(content) - 1:  
            p = text_frame.add_paragraph()
            p.text = " ".join(grouped_text)  
            p.font.size = Pt(18)  
            p.space_after = Pt(10)  
            p.level = 0  
            grouped_text = []  

def create_presentation(title, text_data):
    prs = Presentation("Ion.pptx")  

    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = "Created by PPT Summ"

    max_lines_per_slide = 6  
    current_slide_text = []

    for line in text_data:
        wrapped_lines = textwrap.wrap(line, width=80)
        for wrapped_line in wrapped_lines:
            current_slide_text.append(wrapped_line)
            
            if len(current_slide_text) == max_lines_per_slide:
                add_slide(prs, "Key Points", current_slide_text)
                current_slide_text = []

    if current_slide_text:
        add_slide(prs, "Additional Information", current_slide_text)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
