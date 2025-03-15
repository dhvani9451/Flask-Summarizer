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
CORS(app, resources={r"/*": {"origins": "*"}})  

# ✅ Function to clean and structure text
def clean_text(text):
    """Removes special characters and structures the text into meaningful bullet points."""
    text = re.sub(r'[*#]+', '', text).strip()
    lines = text.split(". ")  

    structured_text = []
    temp_group = []

    for line in lines:
        line = line.strip()
        if line:
            temp_group.append(line)

        if len(temp_group) >= 5:  # ✅ Group every 5-6 lines into one bullet point
            structured_text.append(" ".join(temp_group))
            temp_group = []

    if temp_group:
        structured_text.append(" ".join(temp_group))

    return structured_text

# ✅ Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")  
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# ✅ Flask route to generate summary
@app.route('/summarize', methods=['POST'])
def summarize_text():
    data = request.json
    text = data.get("text", "")

    if not text:
        return jsonify({"error": "No text provided"}), 400

    summary = generate_summary(text)  
    structured_summary = clean_text(summary)  
    return jsonify({"summary": "\n".join(structured_summary)})

# ✅ Flask route to generate and return PowerPoint file
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    data = request.json
    text = data.get("text", "")

    if not text:
        return jsonify({"error": "No text provided"}), 400

    # ✅ AI Summarization before creating PPT
    summary = generate_summary(text)
    structured_summary = clean_text(summary)  

    # ✅ Generate PowerPoint with proper bullet points
    ppt_file = create_presentation("AI-Generated Summary", structured_summary)

    return send_file(ppt_file, as_attachment=True, download_name="AI_Summary_Presentation.pptx")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
