from flask import Flask, request, jsonify
from flask_cors import CORS  # Import CORS
import google.generativeai as genai
import os

# Initialize Flask app
app = Flask(__name__)

# Enable CORS for all routes
CORS(app)  # This allows requests from any origin (e.g., frontend on localhost)

# Configure Gemini API Key (Read from Environment Variable)
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")  # Set this in Render
genai.configure(api_key=GEMINI_API_KEY)

# Homepage Route
@app.route("/", methods=["GET"])
def home():
    return "Flask Summarizer API is running!", 200

# Define the summarization endpoint
@app.route("/summarize", methods=["POST"])
def summarize_text():
    data = request.get_json()
    text = data.get("text", "")

    if not text:
        return jsonify({"error": "No text provided"}), 400

    # Call Gemini-2.0-Flash model
    model = genai.GenerativeModel("models/gemini-2.0-flash")
    response = model.generate_content(f"Summarize the following text:\n\n{text}")

    summary = response.text if response and response.text else "Error generating summary."

    return jsonify({"summary": summary})

# Run Flask server
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Use the port provided by Render
    app.run(host="0.0.0.0", port=port, debug=True)
