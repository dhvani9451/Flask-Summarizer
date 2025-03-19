from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import re
import google.generativeai as genai
from generate_ppt import create_presentation

# ✅ Load Gemini API key securely from environment variables
api_key = os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4")
if not api_key:
    raise ValueError("⚠️ API Key not found. Set GEMINI_API_KEY in environment variables.")
genai.configure(api_key="AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4")

app = Flask(__name__)
CORS(app, resources={r"/generate-ppt": {"origins": "*"}})  # ✅ Allow CORS for this endpoint

# ✅ Function to clean and structure text
def clean_text(text):
    """Cleans extracted text and structures it for summarization."""
    if isinstance(text, list):
        text = "\n\n".join(text)  # ✅ Handle list-based text

    text = re.sub(r'[*#]', '', text)  # ✅ Remove special characters
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)  # ✅ Allow only alphanumeric, spaces, punctuation
    text = re.sub(r'\s+', ' ', text).strip()  # ✅ Remove extra spaces
    text = re.sub(r'\n+', '\n', text)  # ✅ Remove extra newlines

    lines = text.split(". ")
    structured_text = [line.strip() for line in lines if line]

    return structured_text

# ✅ Function to generate summary using Gemini AI
def generate_summary(text):
    """Uses Gemini AI to generate a summary of the provided text."""
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# ✅ Flask route to generate and return PowerPoint presentation
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
        data = request.json
        file_summaries = {}

        if not data or "file_texts" not in data:
            return jsonify({"error": "Invalid request. Missing 'file_texts' field."}), 400

        # ✅ Process each file separately
        for file_name, text in data["file_texts"].items():
            if not text.strip():
                file_summaries[file_name] = ["No text provided."]
            else:
                summary = generate_summary(text)
                cleaned_summary = clean_text(summary)  # ✅ Ensure clean data for PPT
                file_summaries[file_name] = cleaned_summary

        if not file_summaries:
            return jsonify({"error": "No valid text data provided"}), 400

        # ✅ Generate PowerPoint presentation for all processed summaries
        ppt_file = create_presentation("AI-Generated Summary", file_summaries)

        return send_file(
            ppt_file,
            as_attachment=True,
            download_name="AI_Summary_Presentation.pptx"
        )

    except Exception as e:
        return jsonify({"error": f"Internal Server Error: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
