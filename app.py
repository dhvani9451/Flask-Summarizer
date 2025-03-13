from flask import Flask, request, send_file, jsonify
from flask_cors import CORS  # ✅ Import CORS
import google.generativeai as genai
import os
import io
from generate_ppt import create_presentation, clean_and_structure_text

# ✅ Configure Gemini API
genai.configure(api_key=os.getenv("AIzaSyDLiZW7r215H5zhxaeLGM7bGYJ_CGFHDcg"))

app = Flask(__name__)

# ✅ Set CORS to allow requests from your frontend
CORS(app, origins=["http://127.0.0.1:5500", "https://your-frontend-url.com"], supports_credentials=True)

# ✅ Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# ✅ API Endpoint to Generate PPT
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    data = request.json
    text = data.get("text", "")

    if not text:
        return jsonify({"error": "No text provided"}), 400

    structured_text = clean_and_structure_text(text)
    ppt_file = create_presentation("AI-Generated Summary", structured_text)

    return send_file(
        ppt_file, as_attachment=True, download_name="Generated_Summary_Presentation.pptx"
    )

# ✅ API Endpoint to Generate Summary
@app.route('/summarize', methods=['POST'])
def summarize_text():
    data = request.json
    text = data.get("text", "")
    
    if not text:
        return jsonify({"error": "No text provided"}), 400
    
    summary = generate_summary(text)
    return jsonify({"summary": summary})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
