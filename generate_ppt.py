from pptx import Presentation
from pptx.util import Pt, Inches
import io
import textwrap
import re
import os

def clean_text(text):
    """Cleans and structures the extracted text to detect new thoughts."""
    if isinstance(text, list):
        text = "\n".join(text)

    text = re.sub(r'[*#]', '', text)  # Remove unwanted characters
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)  # Keep only letters, numbers, and punctuation
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces
    text = re.sub(r'\n+', '\n', text)  # Remove extra newlines

    # ✅ Split into logical chunks to identify new thoughts
    split_phrases = [
        "Example", "Once you", "Provide as much detail", "I can help", 
        "Specify", "Use standard UML", "Additionally", "Furthermore", "In conclusion"
    ]
    pattern = "|".join(re.escape(phrase) for phrase in split_phrases)
    structured_text = re.split(pattern, text)
    structured_text = [sentence.strip() for sentence in structured_text if sentence]

    return structured_text

def add_slide(prs, title, content):
    """Adds a slide ensuring proper bullet point formatting while keeping text within the box."""
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)

    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()

    max_chars_per_line = 80  # ✅ Prevents overflow
    max_bullets_per_slide = 6  # ✅ Prevents excessive text on one slide

    bullet_count = 0  # Track number of bullets

    for paragraph in content:
        if not paragraph.strip():  
            continue  # ✅ Skip empty lines

        wrapped_lines = textwrap.wrap(paragraph, width=max_chars_per_line)

        for i, line in enumerate(wrapped_lines):
            if bullet_count >= max_bullets_per_slide:
                return  # ✅ Stops adding bullets if the slide is full

            p = text_frame.add_paragraph()
            p.text = line
            p.font.size = Pt(18)
            p.space_after = Pt(10)

            # ✅ Apply bullets only at the start of a new thought, not every line
            if i == 0:
                p.level = 0  # ✅ Default PowerPoint bullet
            else:
                p.level = 1  # ✅ Continuation of a paragraph without a bullet

            bullet_count += 1

def create_presentation(file_texts):
    """Creates a PowerPoint presentation and returns it as a file object."""
    template_path = "Ion.pptx"
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # ✅ Add Title Slide
    slide_layout = prs.slide_layouts[0]  # Title Slide Layout
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"

    max_lines_per_slide = 6

    for filename, text in file_texts.items():
        structured_text = clean_text(text)
        current_slide_text = []

        # ✅ Add a Slide for Each File Title
        file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only Layout
        file_title_slide.shapes.title.text = f"📄 {filename}"  # Show filename as slide title

        for line in structured_text:
            current_slide_text.append(line)

            if len(current_slide_text) >= max_lines_per_slide:
                add_slide(prs, f"Key Points - {filename}", current_slide_text)
                current_slide_text = []

        if current_slide_text:
            add_slide(prs, f"Additional Info - {filename}", current_slide_text)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
