from pptx import Presentation
from pptx.util import Pt
import io
import textwrap
import re
import os

# Define semantic patterns to identify new thoughts (ONCE)
PATTERNS = [
    "Firstly", "Secondly", "Thirdly", "Finally", 
    "In addition", "Moreover", "Furthermore", "Additionally", 
    "However", "Nevertheless", "On the other hand", "In contrast", 
    "Therefore", "Thus", "Hence", "As a result", 
    "For example", "For instance", "Specifically", 
    "In conclusion", "To summarize", "In summary", 
    "Meanwhile", "Subsequently", "Afterwards", 
    "Because", "Since", "As", 
    "Although", "Even though", "Despite", 
    "Not only", "But also", 
    "In order to", "So that", 
    "As well as", "Along with", 
    "In fact", "Indeed", 
    "On the contrary", "In comparison", 
    "In the meantime", "During", 
    "As soon as", "Until", 
    "Unless", "Provided that", 
    "While", "Whereas", 
    "In the first place", "To begin with", 
    "Last but not least", 
    "In other words", "That is to say", 
    "As mentioned earlier", "As stated before", 
    "In the same way", "Similarly", 
    "For this reason", "Due to this", 
    "In the long run", "Over time", 
    "In the short term", "At the same time", 
    "In general", "Overall", 
    "As a matter of fact", "To illustrate", 
    "To clarify", "To emphasize", 
    "To highlight", "To conclude", 
    "To reiterate", "To put it differently", 
    "To put it simply", "In brief", 
    "In essence", "In reality", 
    "In practice", "In theory", 
    "In any case", "In any event", 
    "In the end", "At the end of the day", 
    "By and large", "For the most part", 
    "In most cases", "In some cases", 
    "Under certain circumstances", 
    "With this in mind", "With regard to", 
    "With respect to", "In terms of", 
    "As far as", "As long as", 
    "As much as", "As soon as", 
    "As though", "Even if", 
    "In case", "In the event that", 
    "Only if", "So long as", 
    "Supposing that", "To the extent that", 
    "Whether or not", "While", 
    "Now that", "Given that", 
    "Seeing that", "Considering that", 
    "Provided that", "Assuming that", 
    "Insofar as", "Inasmuch as", 
    "In the same vein", "In like manner", 
    "In a similar fashion", "In a similar vein", 
    "In the same manner", "In the same fashion", 
    "In the same way", "In the same light", 
    "In the same spirit", "In the same context", 
    "In the same regard", "In the same respect", 
    "In the same sense", "In the same tone", 
    "In the same way", "In the same vein", 
    "In the same way", "In the same way", 
    "In the same way", "In the same way"
]

def clean_text(text):
    """Cleans and structures the extracted text based on semantic patterns."""
    if isinstance(text, list):
        text = "\n".join(text)

    # Remove unwanted characters
    text = re.sub(r'[*#]', '', text)
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)  # Keep only letters, numbers, and punctuation
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces
    text = re.sub(r'\n+', '\n', text)  # Remove extra newlines

    # Split text into sentences based on semantic patterns
    structured_text = []
    current_sentence = ""
    for word in text.split():
        current_sentence += word + " "
        # Check if the current word matches any of the semantic patterns
        if any(word.lower().startswith(pattern.lower()) for pattern in PATTERNS):
            structured_text.append(current_sentence.strip())
            current_sentence = ""
    # Add the last sentence if it exists
    if current_sentence:
        structured_text.append(current_sentence.strip())

    return structured_text

def add_slide(prs, title, content):
    """Adds a slide ensuring bullet points are applied only to new thoughts."""
    slide_layout = prs.slide_layouts[1]  # Title & Content layout
    slide = prs.slides.add_slide(slide_layout)

    # Set slide title
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)

    # Set slide content
    content_shape = slide.shapes.placeholders[1]
    text_frame = content_shape.text_frame
    text_frame.clear()

    # Add bullet points only for new thoughts
    for i, paragraph in enumerate(content):
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.font.size = Pt(18)
        p.space_after = Pt(10)

        # Apply bullet only if it's a new thought (matches semantic patterns)
        if i == 0 or any(paragraph.strip().lower().startswith(pattern.lower()) for pattern in PATTERNS):
            p.level = 0  # PowerPoint's default bullet
        else:
            p.level = 1  # No bullet for wrapped lines or continuous text

def create_presentation(file_texts):
    """Creates a PowerPoint presentation and returns it as a file object."""
    # Use a template if available
    template_path = "Ion.pptx"
    if os.path.exists(template_path):
        prs = Presentation(template_path)
    else:
        prs = Presentation()

    # Add Title Slide
    slide_layout = prs.slide_layouts[0]  # Title Slide Layout
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = "Summary Presentation"
    slide.placeholders[1].text = "Created by AI PPT Generator"

    # Define maximum lines per slide and characters per line
    max_lines_per_slide = 6  # Adjust based on slide space
    max_chars_per_line = 80  # Adjust based on font size and slide width

    # Process each file's text
    for filename, text in file_texts.items():
        structured_text = clean_text(text)
        current_slide_text = []

        # Add a Slide for Each File Title
        file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only Layout
        file_title_slide.shapes.title.text = f"📄 {filename}"  # Show filename as slide title

        # Add slides for the content
        for line in structured_text:
            wrapped_lines = textwrap.wrap(line, width=max_chars_per_line)
            for wrapped_line in wrapped_lines:
                current_slide_text.append(wrapped_line)

                # Add a new slide if the maximum lines per slide is reached
                if len(current_slide_text) == max_lines_per_slide:
                    add_slide(prs, f"Key Points - {filename}", current_slide_text)
                    current_slide_text = []

        # Add remaining content to a new slide
        if current_slide_text:
            add_slide(prs, f"Additional Info - {filename}", current_slide_text)

    # Save the presentation to a BytesIO object
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io
