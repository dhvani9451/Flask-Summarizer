from flask import Flask, request, send_file, jsonify
from flask_cors import CORS  # ✅ Enable CORS
import google.generativeai as genai  # ✅ Import Gemini AI
import os
import io
from generate_ppt import create_presentation, clean_and_structure_text  # ✅ Import functions from generate_ppt.py

# ✅ Configure Gemini API (Read API Key from environment variable)
API_KEY = os.getenv("AIzaSyDLiZW7r215H5zhxaeLGM7bGYJ_CGFHDcg")
if not API_KEY:
    raise ValueError("❌ Error: GEMINI_API_KEY is not set in environment variables.")
genai.configure(api_key=API_KEY)

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)  # ✅ Allow all origins

# ✅ Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")  # ✅ Using Gemini Model
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# ✅ API Endpoint to Generate PPT
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
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
        
        # ✅ Add CORS Headers
        response.headers.add("Access-Control-Allow-Origin", "*")
        response.headers.add("Access-Control-Allow-Headers", "Content-Type, Authorization")
        response.headers.add("Access-Control-Allow-Methods", "POST, OPTIONS")

        return response

    except Exception as e:
        return jsonify({"error": f"Error generating PPT: {str(e)}"}), 500

# ✅ API Endpoint to Generate Summary
@app.route('/summarize', methods=['POST'])
def summarize_text():
    try:
        data = request.json
        text = data.get("text", "")

        if not text:
            return jsonify({"error": "No text provided"}), 400

        summary = generate_summary(text)  # ✅ Use Gemini AI for summarization
        
        response = jsonify({"summary": summary})
        
        # ✅ Add CORS Headers
        response.headers.add("Access-Control-Allow-Origin", "*")
        response.headers.add("Access-Control-Allow-Headers", "Content-Type, Authorization")
        response.headers.add("Access-Control-Allow-Methods", "POST, OPTIONS")

        return response

    except Exception as e:
        return jsonify({"error": f"Error summarizing text: {str(e)}"}), 500

# ✅ Handle Preflight Requests (CORS)
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
