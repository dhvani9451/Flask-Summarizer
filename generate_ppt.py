from pptx import Presentation
from pptx.util import Inches, Pt
import io
import textwrap
import re

# Function to clean and format text
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

# Function to add slides with bullet points & paragraph mix
def add_slide(prs, title, content):
    slide_layout = prs.slide_layouts[5]  # Title Only layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Title Formatting (Heading)
    title_shape = slide.shapes.title
    title_shape.text = title
    title_shape.text_frame.paragraphs[0].font.bold = True
    title_shape.text_frame.paragraphs[0].font.size = Pt(28)  # Larger font for heading
    title_shape.text_frame.paragraphs[0].font.name = "Calibri"  # Set font for heading
    title_shape.text_frame.paragraphs[0].font.color.rgb = (0x1F, 0x3D, 0x5F)  # Dark blue color

    # Text Box for Content (Subheading and Body)
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8.5), Inches(5))
    text_frame = textbox.text_frame
    text_frame.word_wrap = True

    # Formatting content with bullet points & paragraphs
    for paragraph in content:
        p = text_frame.add_paragraph()
        p.text = paragraph
        p.space_after = Pt(12)  # Increased spacing after paragraph
        p.font.size = Pt(18)  # Larger font for better readability
        p.font.name = "Calibri"  # Set font for body text
        p.font.color.rgb = (0x2D, 0x2D, 0x2D)  # Dark gray color

        # Decide if it's a paragraph or a bullet point
        if len(paragraph.split()) > 20 or "\n" in paragraph:  # Long text or multi-line text
            p.level = 1  # Paragraph (No bullet)
            p.space_before = Pt(12)  # Add space before paragraphs
            p.line_spacing = 1.5  # Increase line spacing for paragraphs
        else:
            p.level = 0  # Bullet Point (Short text with full stop)
            p.space_before = Pt(6)  # Less space before bullet points
            p.line_spacing = 1.2  # Slightly less spacing for bullet points

# Function to add a confidentiality watermark
def add_watermark(slide):
    watermark = slide.shapes.add_textbox(Inches(3), Inches(2), Inches(4), Inches(1))
    watermark.text = "Confidential"
    watermark.text_frame.paragraphs[0].font.size = Pt(48)
    watermark.text_frame.paragraphs[0].font.color.rgb = (0xDD, 0xDD, 0xDD)  # Light gray color
    watermark.text_frame.paragraphs[0].font.bold = True
    watermark.text_frame.paragraphs[0].font.italic = True
    watermark.text_frame.paragraphs[0].alignment = 2  # Center alignment

# Function to create a well-structured PowerPoint presentation
def create_presentation(title, text_data):
    prs = Presentation("Ion.pptx")  # Load Ion theme

    # Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    slide.shapes.title.text = title
    slide.shapes.title.text_frame.paragraphs[0].font.name = "Calibri"  # Set font for title slide
    slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(36)  # Large font for title
    slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = (0x1F, 0x3D, 0x5F)  # Dark blue color
    slide.placeholders[1].text = "Created by PPT Summ"
    slide.placeholders[1].text_frame.paragraphs[0].font.name = "Calibri"  # Set font for subtitle
    slide.placeholders[1].text_frame.paragraphs[0].font.size = Pt(18)  # Smaller font for subtitle
    slide.placeholders[1].text_frame.paragraphs[0].font.color.rgb = (0x2D, 0x2D, 0x2D)  # Dark gray color

    # Add a confidentiality watermark to all slides
    for slide in prs.slides:
        add_watermark(slide)

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
