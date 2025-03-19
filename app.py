from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import re
import google.generativeai as genai
from generate_ppt import create_presentation

# ✅ Configure Gemini API securely (Kept unchanged)
genai.configure(api_key=os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4"))

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # ✅ Enable CORS for all origins

# ✅ Function to clean text (removes special characters)
def clean_text(text):
    if isinstance(text, list):
        text = "\n\n".join(text)

    text = re.sub(r'[*#]', '', text)
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()

    lines = text.split(". ")
    structured_text = [line.strip() for line in lines if line]

    return structured_text

# ✅ Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# ✅ Flask route to generate PowerPoint file
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
        # ✅ Check if JSON is empty or not sent
        if not request.is_json:
            return jsonify({"error": "Request must be in JSON format"}), 400

        data = request.get_json()

        # ✅ Ensure 'file_texts' exists and is not empty
        file_texts = data.get("file_texts")
        if not file_texts or not isinstance(file_texts, dict):
            return jsonify({"error": "Invalid request: 'file_texts' is missing or not a dictionary"}), 400

        file_summaries = {}

        for file_name, text in file_texts.items():
            if not text.strip():
                file_summaries[file_name] = ["No text provided."]
            else:
                summary = generate_summary(text)
                cleaned_summary = clean_text(summary)
                file_summaries[file_name] = cleaned_summary

        # ✅ Ensure at least one valid summary is generated
        if not file_summaries:
            return jsonify({"error": "No valid text data provided"}), 400

        # ✅ Generate PowerPoint for all processed summaries
        ppt_file = create_presentation("AI-Generated Summary", file_summaries)

        return send_file(
            ppt_file,
            as_attachment=True,
            download_name="AI_Summary_Presentation.pptx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
