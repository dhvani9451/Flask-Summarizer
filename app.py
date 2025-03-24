<<<<<<< HEAD
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import re
import google.generativeai as genai
from generate_ppt import create_presentation

# ✅ Configure Gemini API securely
genai.configure(api_key=os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4"))  # ✅ Use environment variable

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # ✅ Enable CORS for all origins

# ✅ Function to clean text (removes special characters)
def clean_text(text):
    if isinstance(text, list):
        text = "\n\n".join(text)  # ✅ Properly format list texts

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
    try:
        data = request.json
        file_summaries = {}

        # ✅ Process each file separately
        for file_name, text in data.get("file_texts", {}).items():
            if not text.strip():
                file_summaries[file_name] = ["No text provided."]
            else:
                summary = generate_summary(text)
                cleaned_summary = clean_text(summary)  # ✅ Clean before PPT generation
                file_summaries[file_name] = cleaned_summary

        if not file_summaries:
            return jsonify({"error": "No valid text data provided"}), 400

        # ✅ Generate PowerPoint for all processed summaries
        ppt_file = create_presentation(file_summaries)

        return send_file(
            ppt_file,
            as_attachment=True,
            download_name="AI_Summary_Presentation.pptx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
=======
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import re
import google.generativeai as genai
from generate_ppt import create_presentation

# ✅ Configure Gemini API securely
genai.configure(api_key=os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4"))  # ✅ Use environment variable

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # ✅ Enable CORS for all origins

# ✅ Function to clean text (removes special characters)
def clean_text(text):
    if isinstance(text, list):
        text = "\n\n".join(text)  # ✅ Properly format list texts

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
    try:
        data = request.json
        file_summaries = {}

        # ✅ Process each file separately
        for file_name, text in data.get("file_texts", {}).items():
            if not text.strip():
                file_summaries[file_name] = ["No text provided."]
            else:
                summary = generate_summary(text)
                cleaned_summary = clean_text(summary)  # ✅ Clean before PPT generation
                file_summaries[file_name] = cleaned_summary

        if not file_summaries:
            return jsonify({"error": "No valid text data provided"}), 400

        # ✅ Generate PowerPoint for all processed summaries
        ppt_file = create_presentation(file_summaries)

        return send_file(
            ppt_file,
            as_attachment=True,
            download_name="AI_Summary_Presentation.pptx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
>>>>>>> b94a9ca662e4d003bf206be789b87968fc9c6eb5
