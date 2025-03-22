from pptx import Presentation
from pptx.util import Pt
import io
import textwrap
import re
import logging
from flask import Flask, send_file, jsonify, request  # Assuming Flask as the API framework

# Set up logging with more detail
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("ppt_generator.log"),  # Log to a file for easier debugging
        logging.StreamHandler()  # Also log to console
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)  # Create a Flask app for API handling

def clean_text(text):
    """Cleans and structures the extracted text."""
    try:
        logger.debug("Starting clean_text function")
        if isinstance(text, list):
            text = "\n".join(text)

        text = re.sub(r'[*#]', '', text)
        text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)
        text = re.sub(r'\s+', ' ', text).strip()
        text = re.sub(r'\n+', '\n', text)

        lines = text.split(". ")
        structured_text = [line.strip() for line in lines if line]
        logger.debug(f"Cleaned text into {len(structured_text)} sentences: {structured_text}")
        return structured_text
    except Exception as e:
        logger.error(f"Error in clean_text: {str(e)}", exc_info=True)
        raise

def add_slide(prs, title, content):
    """Adds a new slide with title and formatted content with bullet points."""
    try:
        logger.debug(f"Adding slide with title: {title}")
        slide_layout = prs.slide_layouts[1]  # Title & Content layout
        slide = prs.slides.add_slide(slide_layout)

        # Set the title
        if not slide.shapes.title:
            logger.error("Title shape not found in slide layout")
            raise ValueError("Title shape not found in slide layout")
        title_shape = slide.shapes.title
        title_shape.text = title
        title_shape.text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].font.size = Pt(28)

        # Add content with bullet points
        if len(slide.shapes.placeholders) < 2:
            logger.error("Content placeholder not found in slide layout")
            raise ValueError("Content placeholder not found in slide layout")
        content_shape = slide.shapes.placeholders[1]
        text_frame = content_shape.text_frame
        text_frame.clear()

        for sentence in content:
            wrapped_lines = textwrap.wrap(sentence, width=80)
            p = text_frame.add_paragraph()
            p.text = wrapped_lines[0] if wrapped_lines else ""
            p.font.size = Pt(18)
            p.space_after = Pt(10)
            p.level = 0

            for line in wrapped_lines[1:]:
                p = text_frame.add_paragraph()
                p.text = line
                p.font.size = Pt(18)
                p.space_after = Pt(10)
                p.level = 0
                p.paragraph_format.left_margin = Pt(36)
                p.paragraph_format.bullet = False

        logger.debug(f"Successfully added slide: {title}")
    except Exception as e:
        logger.error(f"Error in add_slide: {str(e)}", exc_info=True)
        raise

def create_presentation(file_texts):
    """Creates a PowerPoint presentation and returns it as a file object."""
    try:
        logger.debug("Starting create_presentation function")
        # Always create a new presentation to avoid template issues
        prs = Presentation()
        logger.debug("Created new Presentation object")

        # Add Title Slide
        slide_layout = prs.slide_layouts[0]  # Title Slide Layout
        slide = prs.slides.add_slide(slide_layout)
        slide.shapes.title.text = "Summary Presentation"
        slide.placeholders[1].text = "Created by AI PPT Generator"
        logger.debug("Added title slide")

        max_sentences_per_slide = 6

        for filename, text in file_texts.items():
            logger.debug(f"Processing file: {filename}")
            structured_text = clean_text(text)
            current_slide_sentences = []

            # Add a Slide for Each File Title
            file_title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only Layout
            file_title_slide.shapes.title.text = f"📄 {filename}"
            logger.debug(f"Added title slide for file: {filename}")

            for sentence in structured_text:
                current_slide_sentences.append(sentence)

                if len(current_slide_sentences) == max_sentences_per_slide:
                    add_slide(prs, f"Key Points - {filename}", current_slide_sentences)
                    current_slide_sentences = []

            if current_slide_sentences:
                add_slide(prs, f"Additional Info - {filename}", current_slide_sentences)

        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)
        logger.debug("Presentation saved to BytesIO")
        return ppt_io

    except Exception as e:
        logger.error(f"Error in create_presentation: {str(e)}", exc_info=True)
        raise

@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    """API endpoint to generate a PowerPoint presentation."""
    try:
        logger.debug("Received request to /generate-ppt")
        # Expect file_texts to be sent in the request body as JSON
        data = request.get_json()
        if not data or 'file_texts' not in data:
            logger.error("Invalid request: file_texts not provided")
            return jsonify({"error": "file_texts not provided in request body"}), 400

        file_texts = data['file_texts']
        logger.debug(f"Received file_texts: {file_texts}")

        ppt_io = create_presentation(file_texts)
        logger.debug("Generated presentation, sending response")
        return send_file(
            ppt_io,
            download_name="presentation.pptx",
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        logger.error(f"Error in generate_ppt endpoint: {str(e)}", exc_info=True)
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    # Test the code locally
    try:
        file_texts = {
            "Pr2_VRS_Final.docx": (
                "into existing sections on technical aspects Technology Stack Mention the intended technologies e.g., programming languages, frameworks, databases, cloud. "
                "platform Scalability Address how the system will be designed to handle increasing numbers of users and vehicles. "
                "e.g., data encryption, access control, protection against common web vulnerabilities. "
                "API Integrations Mention potential integrations with thirdparty."
            )
        }
        ppt_io = create_presentation(file_texts)
        with open("output.pptx", "wb") as f:
            f.write(ppt_io.read())
        logger.info("Test presentation generated successfully")
    except Exception as e:
        logger.error(f"Local test failed: {str(e)}", exc_info=True)

    # Start the Flask server for API testing
    app.run(debug=True, host='0.0.0.0', port=5000)
