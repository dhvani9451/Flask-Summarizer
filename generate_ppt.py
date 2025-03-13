from pdfminer.high_level import extract_text
from pptx import Presentation
from pptx.util import Inches, Pt
from docx import Document
import textwrap
import re
import io  # ✅ Required for returning file as a byte stream

def extract_text_from_pdf(pdf_path):
    return extract_text(pdf_path)  

def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])  

def extract_text_from_txt(txt_path):
    with open(txt_path, "r", encoding="utf-8") as file:
        return file.read()  

def clean_and_structure_text(text):
    text = re.sub(r'\n+', '\n', text.strip())  # ✅ Remove extra new lines
    lines = text.split("\n")
    structured_text = [line.strip() for line in lines if line.strip()]
    return structured_text

def add_slide(prs, title, content):
    slide_layout = prs.slide_layouts[5]  # ✅ Title Only layout for better space
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(24)  # ✅ Title Font Size 24px
    
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8.5), Inches(5))
    text_frame = textbox.text_frame
    text_frame.word_wrap = True  # ✅ Enable word wrapping for text

    for paragraph in content:
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.space_after = Pt(5)
        p.font.size = Pt(16)  # ✅ Normal text size 16px

def create_presentation(title, text_data):
    prs = Presentation("Ion.pptx")  # ✅ Load Ion Theme

    # ✅ Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = "Created by PPT Summ"

    max_words_per_line = 12
    max_lines_per_slide = 12
    current_slide_text = []

    for line in text_data:
        wrapped_lines = textwrap.wrap(line, width=max_words_per_line * 6)  # ✅ Wrap text properly
        for wrapped_line in wrapped_lines:
            current_slide_text.append(wrapped_line)
            if len(current_slide_text) == max_lines_per_slide:
                add_slide(prs, "Content", current_slide_text)
                current_slide_text = []

    if current_slide_text:
        add_slide(prs, "Content", current_slide_text)  # ✅ Add remaining text

    # ✅ Return PPT file as a byte stream
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io  # ✅ Flask will use this to send the file
