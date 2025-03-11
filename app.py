from flask import Flask, request, jsonify
from flask_cors import CORS
import google.generativeai as genai

app = Flask(__name__)

# Allow CORS for all domains and methods
CORS(app, resources={r"/*": {"origins": "*"}})

# Configure Gemini AI API key
GEMINI_API_KEY = "AIzaSyBr12Wqh__1rPbCqwvyNFyNLgd3yDrniCM"
genai.configure(api_key="AIzaSyBr12Wqh__1rPbCqwvyNFyNLgd3yDrniCM")

def generate_summary(text):
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")
        prompt = f"""
        Summarize the following text into concise bullet points while ensuring the summary length is around 4/10th of the original text:

        {text}

        Output the summary in bullet points.
        """
        response = model.generate_content(prompt)
        summary = response.text.strip()
        summary = "\n".join(["• " + line.strip() for line in summary.split("\n") if line.strip()])
        return summary
    except Exception as e:
        return f"Error generating summary: {str(e)}"

@app.route('/summarize', methods=['POST'])
def summarize():
    try:
        data = request.get_json()
        text = data.get("text", "")

        if not text.strip():
            return jsonify({"error": "No text provided"}), 400

        summary = generate_summary(text)

        return jsonify({"summary": summary}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
