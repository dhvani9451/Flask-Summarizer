from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import google.generativeai as genai
from generate_ppt import create_presentation, clean_text
from werkzeug.utils import secure_filename
import PyPDF2  # For PDF text extraction
from docx import Document  # For DOCX text extraction

# Configure Gemini API
genai.configure(api_key=os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4"))

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # Enable CORS for all origins

# Allowed file extensions for upload
ALLOWED_EXTENSIONS = {'pdf', 'txt', 'docx'}

# Function to check if the file extension is allowed
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Function to extract text from a PDF file
def extract_text_from_pdf(file):
    pdf_reader = PyPDF2.PdfReader(file)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()
    return text

# Function to extract text from a text file
def extract_text_from_txt(file):
    return file.read().decode("utf-8")

# Function to extract text from a DOCX file
def extract_text_from_docx(file):
    doc = Document(file)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"
    return text

# Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")  # Use Gemini AI model
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# Flask route to handle file upload and generate PPT
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
        data = request.json
        text = data.get("text", "")

        if not text:
            return jsonify({"error": "No text provided"}), 400

        # Clean and summarize the text
        cleaned_text = clean_text(text)
        summary = generate_summary(cleaned_text)

        # Generate PowerPoint
        ppt_file = create_presentation("AI-Generated Summary", summary.split("\n"))

        # Return the PPT file as a downloadable attachment
        return send_file(ppt_file, as_attachment=True, download_name="AI_Summary_Presentation.pptx")
    except Exception as e:
        print(f"Error generating PPT: {str(e)}")  # Log the error
        return jsonify({"error": "Internal Server Error", "details": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
