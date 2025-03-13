from flask import Flask, request, send_file, jsonify
from flask_cors import CORS  # ✅ Enable CORS
import google.generativeai as genai  # ✅ Import Gemini AI
import os
import io
from generate_ppt import create_presentation, clean_and_structure_text  # ✅ Import functions from generate_ppt.py

# ✅ Configure Gemini API
genai.configure(api_key=os.getenv("AIzaSyDLiZW7r215H5zhxaeLGM7bGYJ_CGFHDcg"))  # Make sure to use a valid key

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)  # ✅ Allow all origins

# ✅ API Endpoint to Generate PPT
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    data = request.json
    text = data.get("text", "")

    if not text:
        return jsonify({"error": "No text provided"}), 400

    structured_text = clean_and_structure_text(text)  # ✅ Clean and structure text
    ppt_file = create_presentation("AI-Generated Summary", structured_text)  # ✅ Generate PPT

    response = send_file(
        ppt_file,
        as_attachment=True,
        download_name="Generated_Summary_Presentation.pptx"
    )
    
    response.headers.add("Access-Control-Allow-Origin", "*")  # ✅ Allow frontend to access
    response.headers.add("Access-Control-Allow-Headers", "Content-Type, Authorization")
    response.headers.add("Access-Control-Allow-Methods", "POST, OPTIONS")

    return response

# ✅ API Endpoint to Generate Summary
@app.route('/summarize', methods=['POST'])
def summarize_text():
    data = request.json
    text = data.get("text", "")

    if not text:
        return jsonify({"error": "No text provided"}), 400

    summary = generate_summary(text)  # ✅ Use Gemini AI for summarization

    response = jsonify({"summary": summary})
    response.headers.add("Access-Control-Allow-Origin", "*")  # ✅ Allow frontend access
    response.headers.add("Access-Control-Allow-Headers", "Content-Type, Authorization")
    response.headers.add("Access-Control-Allow-Methods", "POST, OPTIONS")

    return response

# ✅ Handle Preflight Requests
@app.route('/generate-ppt', methods=['OPTIONS'])
@app.route('/summarize', methods=['OPTIONS'])
def handle_preflight():
    response = jsonify({"message": "CORS preflight success"})
    response.headers.add("Access-Control-Allow-Origin", "*")
    response.headers.add("Access-Control-Allow-Headers", "Content-Type, Authorization")
    response.headers.add("Access-Control-Allow-Methods", "POST, OPTIONS")
    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
