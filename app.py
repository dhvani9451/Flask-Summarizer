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
        model = genai.GenerativeModel("models/gemini-2.0-flash")  # Using Gemini Pro model
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# Flask route to generate and return PPT file
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
        data = request.json
        text = data.get("text", "")

        if not text:
            return jsonify({"error": "No text provided"}), 400

        cleaned_text = clean_text(text)  # Process and clean text
        summary = generate_summary("\n".join(cleaned_text))  # Generate summary using Gemini
        ppt_file = create_presentation("AI-Generated Summary", summary.split("\n"))
        
        return send_file(ppt_file, as_attachment=True, download_name="AI_Summary_Presentation.pptx")
    except Exception as e:
        print(f"Error generating PPT: {str(e)}")  # Log the error
        return jsonify({"error": "Internal Server Error", "details": str(e)}), 500

# Flask route to generate summary (calls Gemini AI)
@app.route('/summarize', methods=['POST'])
def summarize_text():
    data = request.json
    text = data.get("text", "")
    
    if not text:
        return jsonify({"error": "No text provided"}), 400
    
    summary = generate_summary(text)  # Use Gemini AI for summarization
    return jsonify({"summary": summary})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
