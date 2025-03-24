from pptx import Presentation
from pptx.util import Pt
import io
import re
import spacy

# Load English NLP model
nlp = spacy.load("en_core_web_sm")

def split_into_thoughts(text):
    """Split text into logical thoughts without relying on punctuation"""
    doc = nlp(text)
    thoughts = []
    current_thought = []
    
    for token in doc:
        # Detect thought boundaries (customize these rules as needed)
        if (token.is_space and len(current_thought) > 20) or \
           (token.text.lower() in ['first', 'second', 'however', 'additionally']):
            if current_thought:
                thoughts.append("".join(current_thought).strip())
            current_thought = [token.text_with_ws]
        else:
            current_thought.append(token.text_with_ws)
    
    if current_thought:
        thoughts.append("".join(current_thought).strip())
    
    return thoughts

def add_slide(prs, title, thoughts):
    """Add slide with bullet points for each complete thought"""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    
    # Title setup
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Content setup
    body_shape = slide.shapes.placeholders[1]
    text_frame = body_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True

    MAX_LINES = 6
    current_lines = 0

    for thought in thoughts:
        if current_lines >= MAX_LINES:
            break
            
        # Add main bullet point
        p = text_frame.add_paragraph()
        p.text = thought[:100]  # Trim very long thoughts
        p.level = 0
        p.font.size = Pt(16)
        current_lines += 1

        # Handle overflow (no sub-bullets for unpunctuated text)
        if len(thought) > 100:
            remaining = thought[100:]
            p = text_frame.add_paragraph()
            p.text = remaining[:100]
            p.level = 0  # Same level for continuation
            current_lines += 1

def create_presentation(file_texts):
    prs = Presentation()
    
    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"

    for filename, text in file_texts.items():
        # Process unpunctuated text
        thoughts = split_into_thoughts(text)
        
        # Split into slides
        for i in range(0, len(thoughts), 3):  # 3 thoughts per slide
            chunk = thoughts[i:i+3]
            title = f"Key Points from {filename}" if i == 0 else f"{filename} (Continued)"
            add_slide(prs, title, chunk)

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io