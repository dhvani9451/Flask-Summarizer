from pptx import Presentation
from pptx.util import Pt
import io
import textwrap
import re
import os
import spacy

# Load English NLP model
nlp = spacy.load("en_core_web_sm")

def clean_text(text):
    """Enhanced text cleaning with NLP sentence segmentation"""
    if isinstance(text, list):
        text = "\n".join(text)

    # Basic cleaning while preserving sentence structure
    text = re.sub(r'[*#]', '', text)  # Remove markdown markers
    text = re.sub(r'\s+', ' ', text).strip()  # Normalize whitespace

    # Use NLP to split into meaningful sentences
    doc = nlp(text)
    sentences = [sent.text.strip() for sent in doc.sents if sent.text.strip()]
    
    return sentences

def is_new_thought(sentence):
    """Use NLP to detect if sentence starts a new thought"""
    if not sentence:
        return False
    
    # Analyze first few words
    doc = nlp(sentence[:50])  # Only check beginning for efficiency
    
    # Check for transition words or discourse markers
    transition_words = {
        "first", "second", "next", "then", "finally",
        "however", "therefore", "moreover", "additionally",
        "consequently", "specifically", "in conclusion"
    }
    
    if len(doc) > 0:
        first_token = doc[0]
        return (first_token.text.lower() in transition_words or
                first_token.pos_ in ["ADV", "SCONJ"] or
                first_token.dep_ == "advmod")
    return False

def add_slide(prs, title, content):
    """Create slide with NLP-aware bullet points and proper text wrapping"""
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)

    # Set slide title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Configure content area
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True  # Enable automatic word wrapping

    MAX_LINES = 6  # Optimal number of lines per slide
    MAX_WIDTH = 70  # Characters per line (narrower for better readability)

    current_thought = None
    lines_added = 0

    for sentence in content:
        if lines_added >= MAX_LINES:
            break  # Prevent overflow

        # Determine bullet level based on thought structure
        if is_new_thought(sentence) or current_thought is None:
            bullet_level = 0  # Main bullet point
            current_thought = sentence
        else:
            bullet_level = 1  # Sub-point

        # Add primary sentence
        p = text_frame.add_paragraph()
        p.text = sentence[:MAX_WIDTH]  # Truncate if needed
        p.level = bullet_level
        p.font.size = Pt(18)
        p.space_after = Pt(6)
        lines_added += 1

        # Handle text wrapping for long sentences
        if len(sentence) > MAX_WIDTH:
            wrapped_lines = textwrap.wrap(sentence, width=MAX_WIDTH)
            for line in wrapped_lines[1:]:  # Skip first line already added
                if lines_added >= MAX_LINES:
                    break
                cont_p = text_frame.add_paragraph()
                cont_p.text = line
                cont_p.level = bullet_level + 1  # Indent wrapped lines
                cont_p.font.size = Pt(16)
                cont_p.space_after = Pt(4)
                lines_added += 1

def create_presentation(file_texts):
    """Create presentation with NLP-processed content"""
    # Use template if available
    template_path = "Ion.pptx"
    prs = Presentation(template_path) if os.path.exists(template_path) else Presentation()

    # Title slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "Vehicle Rental System - Summary Presentation"
    title_slide.placeholders[1].text = "Created by AI PPT Generator"

    for filename, text in file_texts.items():
        # Process text with NLP
        sentences = clean_text(text)
        
        # Add document title slide
        doc_title_slide = prs.slides.add_slide(prs.slide_layouts[5])
        doc_title_slide.shapes.title.text = f"Analysis of {filename}"

        # Split content into slide-sized chunks
        chunk_size = 5  # Optimal number of thoughts per slide
        for i in range(0, len(sentences), chunk_size):
            chunk = sentences[i:i + chunk_size]
            slide_title = f"Key Points from {filename}" if i == 0 else f"{filename} (Continued)"
            add_slide(prs, slide_title, chunk)

    # Save to in-memory file
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io