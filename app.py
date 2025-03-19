from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import re
import google.generativeai as genai
from generate_ppt import create_presentation

# ✅ Load Gemini API key securely
api_key = os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4")

if not api_key:
    raise ValueError("⚠️ API Key not found. Set GEMINI_API_KEY in Render environment variables.")

genai.configure(api_key=api_key)  # ✅ Set API key

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # ✅ Enable CORS globally

# ✅ Function to clean text
def clean_text(text):
    if isinstance(text, list):
        text = "\n\n".join(text)  # ✅ Convert list to properly formatted text

    text = re.sub(r'[*#]', '', text)  # Remove special characters
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)  # Remove non-alphanumeric characters
    text = re.sub(r'\s+', ' ', text).strip()  # Remove extra spaces

    lines = text.split(". ")
    structured_text = [line.strip() for line in lines if line]

    return structured_text

# ✅ Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")  # ✅ Use Gemini AI model
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# ✅ Flask route to generate PowerPoint file
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
        data = request.json
        file_summaries = {}

        # ✅ Process each uploaded file
        for file_name, text in data.get("file_texts", {}).items():
            if not text.strip():
                file_summaries[file_name] = ["No text provided."]
            else:
                summary = generate_summary(text)
                cleaned_summary = clean_text(summary)  # ✅ Clean before generating PPT
                file_summaries[file_name] = cleaned_summary

        if not file_summaries:
            return jsonify({"error": "No valid text data provided"}), 400

        # ✅ Generate PowerPoint with all summaries
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
