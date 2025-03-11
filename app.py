from flask import Flask, request, jsonify
from flask_cors import CORS
import google.generativeai as genai
import os

app = Flask(__name__)
CORS(app)  # Enable CORS for frontend integration

# Configure Gemini AI API key
GEMINI_API_KEY = "AIzaSyBr12Wqh__1rPbCqwvyNFyNLgd3yDrniCM"
genai.configure(api_key="AIzaSyBr12Wqh__1rPbCqwvyNFyNLgd3yDrniCM")

# Function to generate a summary
def generate_summary(text):
    try:
        model = genai.GenerativeModel("gemini-pro")

        # Limit summary size to approximately 40% of the input text
        prompt = f"""
        Summarize the following text into concise bullet points while ensuring the summary length is around 4/10th of the original text:
        
        {text}

        Output the summary in bullet points, using clear and professional language.
        """

        response = model.generate_content(prompt)
        summary = response.text.strip()

        # Ensure summary is properly formatted into bullet points
        summary = "\n".join(["• " + line.strip() for line in summary.split("\n") if line.strip()])
        return summary

    except Exception as e:
        return f"Error generating summary: {str(e)}"

# API route to summarize text
@app.route('/summarize', methods=['POST'])
def summarize():
    try:
        data = request.get_json()
        text = data.get("text", "")

        if not text.strip():
            return jsonify({"error": "No text provided"}), 400

        summary = generate_summary(text)

        return jsonify({"summary": summary})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
