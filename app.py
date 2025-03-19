from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io
import os
import re
import google.generativeai as genai
from generate_ppt import create_presentation

# ✅ Configure Gemini API
genai.configure(api_key=os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4"))

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # ✅ Enable CORS for all origins

# ✅ Function to clean text (removes special characters)
def clean_text(text):
    if isinstance(text, list):
        text = "\n\n".join(text)  # ✅ Join multiple texts properly

    text = re.sub(r'[*#]', '', text)
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()

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

# ✅ Flask route to generate and return PowerPoint file
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    data = request.json
    texts = data.get("texts", [])

    if not texts or not isinstance(texts, list):
        return jsonify({"error": "Invalid or missing text data"}), 400

    # ✅ Combine all texts before summarizing
    combined_text = "\n\n".join(texts)

    # ✅ AI Summarization before creating PPT
    summary = generate_summary(combined_text)
    cleaned_summary = clean_text(summary)  # ✅ Clean summary before creating PPT

    # ✅ Generate PowerPoint
    ppt_file = create_presentation("AI-Generated Summary", cleaned_summary)

    return send_file(ppt_file, as_attachment=True, download_name="AI_Summary_Presentation.pptx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
