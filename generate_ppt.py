from pptx import Presentation
from pptx.util import Inches, Pt
import io
import textwrap

def add_slide(prs, title, content):
    slide_layout = prs.slide_layouts[5]  # ✅ Title Only layout
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
        p.text = f"• {paragraph.strip()}"  # ✅ Ensures only one bullet point
        p.space_after = Pt(10)  # ✅ Adds spacing between points
        p.font.size = Pt(18)  # ✅ Slightly larger for readability

def create_presentation(title, text_data):
    prs = Presentation("Ion.pptx")  # ✅ Load Ion theme

    # ✅ Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = "Created by PPT Summ"

    max_words_per_line = 10  # ✅ Reduced words per line for better spacing
    max_lines_per_slide = 8  # ✅ Fewer lines per slide for readability
    current_slide_text = []

    for line in text_data:
        wrapped_lines = textwrap.wrap(line, width=max_words_per_line * 6)
        for wrapped_line in wrapped_lines:
            current_slide_text.append(wrapped_line)
            if len(current_slide_text) >= max_lines_per_slide:
                add_slide(prs, "Key Points", current_slide_text)
                current_slide_text = []

    if current_slide_text:
        add_slide(prs, "Key Points", current_slide_text)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
