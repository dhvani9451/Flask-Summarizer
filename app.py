from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import re
import google.generativeai as genai
from generate_ppt import create_presentation

# Configure Gemini API securely
api_key = os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4")
if not api_key:
    raise ValueError("No Gemini API key found in environment variables")
genai.configure(api_key=api_key)

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # Enable CORS for all origins

def generate_summary(text):
    """Generate summary using Gemini AI with improved error handling"""
    try:
        print("Generating summary for text:", text[:100] + "...")  # Log first 100 chars
        model = genai.GenerativeModel("gemini-pro")
        response = model.generate_content(
            f"Please summarize the following text for a PowerPoint presentation, keeping key points and maintaining readability:\n\n{text}"
        )
        
        if not response.text:
            raise ValueError("Empty response from Gemini API")
            
        summary = response.text.strip()
        print("Generated summary:", summary[:200] + "...")  # Log first 200 chars
        return summary
        
    except Exception as e:
        print(f"Error generating summary: {str(e)}")
        return f"Summary generation failed: {str(e)}"

@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    """Endpoint to generate PowerPoint from text"""
    try:
        print("Received request with data:", request.json)
        data = request.get_json()
        
        if not data or 'text' not in data:
            return jsonify({"error": "No text data provided"}), 400

        text = data.get("text", "")
        file_name = data.get("filename", "Summary").strip() or "Summary"

        if not text.strip():
            return jsonify({"error": "Text cannot be empty"}), 400

        # Generate and process summary
        summary = generate_summary(text)
        if summary.startswith("Summary generation failed"):
            return jsonify({"error": summary}), 500

        # Prepare data for PPT generation
        file_summaries = {file_name: summary}

        # Generate PowerPoint
        try:
            ppt_file = create_presentation(file_summaries)
            return send_file(
                ppt_file,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name=f"{file_name}_Summary.pptx"
            )
        except Exception as e:
            print(f"PPT generation error: {str(e)}")
            return jsonify({"error": f"PPT generation failed: {str(e)}"}), 500

    except Exception as e:
        print(f"Server error: {str(e)}")
        return jsonify({"error": f"Server error: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
