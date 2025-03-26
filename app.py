from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import re
import google.generativeai as genai

# ✅ Import the create_presentation function
from generate_ppt import create_presentation

# ✅ Configure Gemini API securely
genai.configure(api_key=os.getenv("AIzaSyC5nMYSC6oPwbIkPGWhwKfUnQLqvUf1oR4"))  # ✅ Use environment variable

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # ✅ Enable CORS for all origins

# ✅ Function to clean text (removes special characters)
def clean_text(text):
    if isinstance(text, list):
        text = "\n".join(text)  # ✅ Properly format list texts

    text = re.sub(r'[*#]', '', text)
    text = re.sub(r'[^A-Za-z0-9.,\s]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()

    lines = text.split(". ")
    structured_text = [line.strip() for line in lines if line]

    return structured_text

# ✅ Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("gemini-2.0-flash")  # Use Gemini AI model
        response = model.generate_content(text)
        summary = response.text.strip() if response.text else "No summary generated."

        # Ensure the summary ends with a full stop if missing
        if not summary.endswith("."):
            summary += "."

        return summary
    except Exception as e:
        return f"Error generating summary: {str(e)}"

# ✅ Flask route to generate and return PowerPoint file
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    try:
        data = request.get_json()
        if not data or 'text' not in data:
            return jsonify({"error": "No text data provided in request body"}), 400

        text = data.get("text", "")  # Using .get() to avoid KeyError
        file_name = data.get("filename", "Summary") # get the filename or assign default

        if not text.strip():
            return jsonify({"error": "No valid text data provided."}), 400

        # ✅ Process and clean summary text
        summary = generate_summary(text)
        cleaned_summary = clean_text(summary)  # ✅ Clean before PPT generation

        # file_summaries format is a dictionary with file_name keys and text list
        file_summaries = {file_name : cleaned_summary}

        # ✅ Generate PowerPoint for the processed summary
        ppt_file = create_presentation(file_summaries)

        # ✅ Improved Error Handling and Logging
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
