from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import re
import google.generativeai as genai

# Import the create_presentation function
from generate_ppt import create_presentation

# Configure Gemini API securely
api_key = os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4")
if not api_key:
    print("API key not found in environment variables.")
else:
    genai.configure(api_key=api_key)

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # Enable CORS for all origins

# Function to clean text using refined heuristic-based sentence splitting
def clean_text(text):
    if isinstance(text, list):
        text = "\n".join(text)

    text = re.sub(r'[*#]', '', text)
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()

    # Refined heuristic-based sentence splitting
    sentences = re.split(r'(?<!\w\.\w.)(?<![A-Z][a-z]\.)(?<=\.|\?|\!)\s', text)
    structured_text = [sentence.strip() for sentence in sentences if sentence]

    return structured_text

# Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("gemini-2.0-flash")  # Use Gemini AI model
        response = model.generate_content(text)
        summary = response.text.strip() if response.text else "No summary generated."

        # Log the raw summary for diagnostics
        print("Raw Summary:", summary)

        return summary
    except Exception as e:
        print(f"Error generating summary: {e}")
        return f"Error generating summary: {str(e)}"

# Flask route to generate and return PowerPoint file
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
        data = request.get_json()
        if not data or 'text' not in data:
            return jsonify({"error": "No text data provided in request body"}), 400

        text = data.get("text", "")  # Using .get() to avoid KeyError
        file_name = data.get("filename", "Summary")  # Get the filename or assign default

        if not text.strip():
            return jsonify({"error": "No valid text data provided."}), 400

        # Process and clean summary text
        summary = generate_summary(text)
        cleaned_summary = clean_text(summary)  # Clean before PPT generation

        # file_summaries format is a dictionary with file_name keys and text list
        file_summaries = {file_name: cleaned_summary}

        # Generate PowerPoint for the processed summary
        ppt_file = create_presentation(file_summaries)

        # Improved Error Handling and Logging
        try:
            return send_file(
                ppt_file,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                as_attachment=True,
                download_name="AI_Summary_Presentation.pptx"
            )
        except Exception as e:
            print(f"Error sending file: {e}")
            return jsonify({"error": f"Error sending file: {str(e)}"}), 500
    except Exception as e:
        print(f"Error in generate_ppt function: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
