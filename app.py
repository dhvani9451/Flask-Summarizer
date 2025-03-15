from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import google.generativeai as genai
from generate_ppt import create_presentation, clean_text

# Configure Gemini API
genai.configure(api_key=os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4"))

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # Enable CORS for all origins

# Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")  # Use Gemini AI model
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# Flask route to generate summary
@app.route('/summarize', methods=['POST'])
def summarize_text():
    data = request.json
    text = data.get("text", "")

    if not text:
        return jsonify({"error": "No text provided"}), 400

    summary = generate_summary(text)  # AI Summarization
    cleaned_summary = clean_text(summary)  # Clean summary before returning
    return jsonify({"summary": cleaned_summary})

# Flask route to generate and return PowerPoint file
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
        data = request.json
        text = data.get("text", "")

        if not text:
            return jsonify({"error": "No text provided"}), 400

        # AI Summarization before creating PPT
        summary = generate_summary(text)
        cleaned_summary = clean_text(summary)  # Clean summary before creating PPT

        # Generate PowerPoint
        ppt_file = create_presentation("AI-Generated Summary", cleaned_summary)

        return send_file(ppt_file, as_attachment=True, download_name="AI_Summary_Presentation.pptx")
    except Exception as e:
        print(f"Error generating PPT: {str(e)}")  # Log the error
        return jsonify({"error": "Internal Server Error", "details": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
