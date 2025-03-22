from pptx import Presentation
from pptx.util import Pt
import io
import textwrap
import re
import os
import logging

# Set up logging to help debug issues
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def clean_text(text):
    """Cleans and structures the extracted text."""
    try:
        if isinstance(text, list):
            text = "\n".join(text)

        text = re.sub(r'[*#]', '', text)
        text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)
        text = re.sub(r'\s+', ' ', text).strip()
        text = re.sub(r'\n+', '\n', text)  # Remove extra newlines

        lines = text.split(". ")
        structured_text = [line.strip() for line in lines if line]

        logger.info(f"Cleaned text into {len(structured_text)} sentences")
        return structured_text
    except Exception as e:
        logger.error(f"Error in clean_text: {str(e)}")
        raise

def add_slide(prs, title, content):
    """Adds a new slide with title and formatted content with bullet points."""
    try:
        # Use a fallback layout if the desired one fails
        slide_layout = prs.slide_layouts[1]  # Title & Content layout
        slide = prs.slides.add_slide(slide_layout)

        # Set the title
        title_shape = slide.shapes.title
        if not title_shape:
            logger.error("Title shape not found in slide layout")
            raise ValueError("Title shape not found in slide layout")
        title_shape.text = title
        title_shape.text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].font.size = Pt(28)

        # Add content with bullet points
        content_shape = slide.shapes.placeholders[1]
        if not content_shape:
            logger.error("Content placeholder not found in slide layout")
            raise ValueError("Content placeholder not found in slide layout")
        text_frame = content_shape.text_frame
        text_frame.clear()

        # Add each sentence as a separate bulleted paragraph
        for sentence in content:
            wrapped_lines = textwrap.wrap(sentence, width=80)

            # Add the first line with a bullet
            p = text_frame.add_paragraph()
            p.text = wrapped_lines[0] if wrapped_lines else ""
            p.font.size = Pt(18)
            p.space_after = Pt(10)
            p.level = 0  # Apply bullet to the first line of the sentence

            # Add subsequent wrapped lines without bullets
            for line in wrapped_lines[1:]:
                p = text_frame.add_paragraph()
                p.text = line
                p.font.size = Pt(18)
                p.space_after = Pt(10)
                p.level = 0
                p.paragraph_format.left_margin = Pt(36)  # Indent to align with bullet text
                p.paragraph_format.bullet = False  # Explicitly disable bullet for wrapped lines

        logger.info(f"Added slide with title: {title}")
    except Exception as e:
        logger.error(f"Error in add_slide: {str(e)}")
        raise

def create_presentation(file_texts):
    """Creates a PowerPoint presentation and returns it as a file object."""
    try:
        # Load template or create a new presentation
        template_path = "Ion.pptx"
        if os.path.exists(template_path):
            logger.info(f"Loading template from {template_path}")
            prs = Presentation(template_path)
        else:
            logger.warning(f"Template {template_path} not found, creating new presentation")
            prs = Presentation()

        # Add Title Slide
        slide_layout = prs.slide_layouts[0]  # Title Slide Layout
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = "Summary Presentation"
        slide.placeholders[1].text = "Created by AI PPT Generator"
        logger.info("Added title slide")

        max_sentences_per_slide = 6

        for filename, text in file_texts.items():
            structured_text = clean_text(text)
            current_slide_sentences = []

            # Add a Slide for Each File Title
            file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only Layout
            file_title_slide.shapes.title.text = f"📄 {filename}"
            logger.info(f"Added title slide for file: {filename}")

            for sentence in structured_text:
                current_slide_sentences.append(sentence)

                if len(current_slide_sentences) == max_sentences_per_slide:
                    add_slide(prs, f"Key Points - {filename}", current_slide_sentences)
                    current_slide_sentences = []

            if current_slide_sentences:
                add_slide(prs, f"Additional Info - {filename}", current_slide_sentences)

        # Save the presentation to a BytesIO object
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)
        logger.info("Presentation saved successfully")
        return ppt_io

    except Exception as e:
        logger.error(f"Error in create_presentation: {str(e)}")
        raise

# Example usage for testing
if __name__ == "__main__":
    try:
        file_texts = {
            "Pr2_VRS_Final.docx": (
                "into existing sections on technical aspects Technology Stack Mention the intended technologies e.g., programming languages, frameworks, databases, cloud. "
                "platform Scalability Address how the system will be designed to handle increasing numbers of users and vehicles. "
                "e.g., data encryption, access control, protection against common web vulnerabilities. "
                "API Integrations Mention potential integrations with thirdparty."
            )
        }
        ppt_file = create_presentation(file_texts)
        with open("output.pptx", "wb") as f:
            f.write(ppt_file.read())
        logger.info("Test presentation generated successfully")
    except Exception as e:
        logger.error(f"Test failed: {str(e)}")
