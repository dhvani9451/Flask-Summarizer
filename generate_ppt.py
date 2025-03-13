from pptx import Presentation
from pptx.util import Inches, Pt
import io
import textwrap
import re

# ✅ Function to clean and format text
def clean_text(text):
    if isinstance(text, list):
        text = "\n".join(text)  # Convert list to string if necessary

    text = re.sub(r'[*]+', '', text)  # Remove *, **, ***
    text = re.sub(r'\n+', '\n', text.strip())  # Remove extra new lines
    lines = text.split("\n")

    structured_text = []
    for line in lines:
        line = line.strip()
        if line:  # Only add non-empty lines
            structured_text.append(line)

    return structured_text

# ✅ Function to add slides with bullet points & paragraph mix
def add_slide(prs, title, content):
    slide_layout = prs.slide_layouts[5]  # ✅ Title Only layout
    slide = prs.slides.add_slide(slide_layout)
    
    # ✅ Title Formatting
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(24)

    # ✅ Text Box for Content
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8.5), Inches(5))
    text_frame = textbox.text_frame
    text_frame.word_wrap = True

    # ✅ Formatting content with bullet points & paragraphs
    for paragraph in content:
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.space_after = Pt(10)  # Line spacing after paragraph
        p.font.size = Pt(16)

        # ✅ Use bullet points only for short points (not paragraphs)
        if len(paragraph.split()) <= 20 and "." not in paragraph:
            p.level = 0  # Bullet Point
        else:
            p.level = 1  # Paragraph (No bullet)

# ✅ Function to create a well-structured PowerPoint presentation
def create_presentation(title, text_data):
    prs = Presentation("Ion.pptx")  # ✅ Load Ion theme

    # ✅ Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = "Created by PPT Summ"

    # ✅ Organize content into slides with proper spacing
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

    # ✅ Add any remaining text
    if current_slide_text:
        add_slide(prs, "Additional Information", current_slide_text)

    # ✅ Save PPT to memory for downloading
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
