from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io
import os
import re  # ✅ For text cleaning
import google.generativeai as genai
from generate_ppt import create_presentation  # ✅ Import PPT generation function

# ✅ Configure Gemini API
genai.configure(api_key=os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4"))

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # ✅ Enable CORS for all origins

# ✅ Function to clean text (removes special characters)
def clean_text(text):
    """Removes special characters like **, *, ***, # from the text."""
    return re.sub(r'[*#]+', '', text).strip()

# ✅ Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")  # ✅ Use Gemini AI model
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# ✅ Flask route to generate summaries for multiple texts
@app.route('/summarize', methods=['POST'])
def summarize_text():
    data = request.json
    texts = data.get("texts", [])  # Expecting a list of texts

    if not texts or not isinstance(texts, list):
        return jsonify({"error": "Invalid input. Expecting a list of texts."}), 400

    summaries = [clean_text(generate_summary(text)) for text in texts]  # ✅ Summarize each text
    return jsonify({"summaries": summaries})

# ✅ Flask route to generate PowerPoint from multiple summaries
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    data = request.json
    texts = data.get("texts", [])  # Expecting a list of texts

    if not texts or not isinstance(texts, list):
        return jsonify({"error": "Invalid input. Expecting a list of texts."}), 400

    # ✅ Summarize each text separately
    summaries = [clean_text(generate_summary(text)) for text in texts]

    # ✅ Combine all summaries for the PowerPoint
    combined_summary = "\n\n".join(summaries)

    # ✅ Generate PowerPoint
    ppt_file = create_presentation("AI-Generated Summary", combined_summary.split("\n"))

    return send_file(ppt_file, as_attachment=True, download_name="AI_Summary_Presentation.pptx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
