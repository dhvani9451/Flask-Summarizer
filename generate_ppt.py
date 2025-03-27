import re
import io
import os
import textwrap
from pptx import Presentation
from pptx.util import Pt

def intelligent_text_processing(text):
    """
    Comprehensive text processing with advanced sentence detection.
    
    Key Features:
    - Preserve all original punctuation
    - Detect complete sentences intelligently
    - Handle complex text structures
    """
    # Convert list to string if needed
    if isinstance(text, list):
        text = " ".join(text)
    
    # Advanced sentence splitting regex
    # Handles abbreviations, quotes, and complex sentence structures
    sentence_pattern = r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?|\!)\s+'
    
    # Split into sentences while preserving punctuation
    sentences = re.split(sentence_pattern, text)
    
    # Clean and validate sentences
    processed_sentences = []
    for sentence in sentences:
        # Trim whitespace, preserve punctuation
        sentence = sentence.strip()
        
        # Capitalize first letter if not already capitalized
        if sentence and not sentence[0].isupper():
            sentence = sentence[0].upper() + sentence[1:]
        
        if sentence:
            processed_sentences.append(sentence)
    
    return processed_sentences

def add_slide(prs, title, content):
    """
    Enhanced slide creation with intelligent, punctuation-aware bullet points.
    """
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)

    # Title formatting
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Content area preparation
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()
    text_frame.word_wrap = True

    # Intelligent bullet point generation
    for paragraph in content:
        # Detect and process sentences
        sentences = intelligent_text_processing(paragraph)
        
        for sentence in sentences:
            # Add each sentence as a bullet point
            p = text_frame.add_paragraph()
            p.text = sentence
            p.font.size = Pt(16)
            p.space_after = Pt(8)
            p.level = 0  # Consistent bullet style

def create_presentation(file_texts):
    """
    Create a PowerPoint with enhanced, context-aware text processing.
    """
    # Template or new presentation
    template_path = "Ion.pptx"
    prs = Presentation(template_path) if os.path.exists(template_path) else Presentation()

    # Title slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Comprehensive Summary Presentation"
    slide.placeholders[1].text = "Generated with Advanced Text Processing"

    # Processing parameters
    max_lines_per_slide = 6
    max_chars_per_line = 80

    for filename, text in file_texts.items():
        # File title slide
        file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])
        file_title_slide.shapes.title.text = f"Summary of {filename}"

        # Prepare slide content
        current_slide_text = []
        
        # Wrap text maintaining readability
        wrapped_lines = textwrap.wrap(text, width=max_chars_per_line)
        
        for wrapped_line in wrapped_lines:
            if wrapped_line.strip():
                current_slide_text.append(wrapped_line.strip())

                # New slide when current is full
                if len(current_slide_text) == max_lines_per_slide:
                    add_slide(prs, f"Key Points from {filename}", current_slide_text)
                    current_slide_text = []

        # Final slide for remaining content
        if current_slide_text:
            add_slide(prs, f"Additional Details from {filename}", current_slide_text)

    # Save presentation
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
