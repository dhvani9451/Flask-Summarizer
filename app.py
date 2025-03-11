from flask import Flask, request, jsonify
from flask_cors import CORS  # Import CORS
import google.generativeai as genai
import os

# Initialize Flask app
app = Flask(__name__)

# Enable CORS for all routes
CORS(app)  # This allows requests from any origin (e.g., frontend on localhost)

# Configure Gemini API Key (Read from Environment Variable)
GEMINI_API_KEY = os.getenv("AIzaSyAIS2oLi1uxW6BSnCp6pOUT2u3vXVz2JMs")  # Set this in Render
genai.configure(api_key=AIzaSyAIS2oLi1uxW6BSnCp6pOUT2u3vXVz2JMs)

# Homepage Route
from flask import Flask, request, jsonify
import google.generativeai as genai
import re

app = Flask(_name_)

# Configure Gemini API Key
genai.configure(api_key="YOUR_GEMINI_API_KEY")

def summarize_text(text):
    """Summarizes text in bullet points and limits size to 40% of original."""
    
    # Calculate target summary length (40% of original)
    original_length = len(text.split())
    target_length = int(original_length * 0.4)
    
    prompt = f"""
    Summarize the following text into concise bullet points. 
    Ensure that the summary length is approximately 40% of the original text.
    Use proper formatting and structure.

    TEXT: 
    {text}

    OUTPUT FORMAT:
    - Point 1
    - Point 2
    - Point 3
    """

    # Generate summary
    try:
        response = genai.GenerativeModel("gemini-pro").generate_content(prompt)
        summary = response.text.strip()

        # Extract bullet points using regex
        bullet_points = re.findall(r"[-•] (.+)", summary)
        if not bullet_points:
            bullet_points = summary.split("\n")  # Fallback if no bullets detected

        # Trim to match target length
        summarized_text = "\n".join(bullet_points[:target_length])

        return summarized_text
    except Exception as e:
        return f"Error: {str(e)}"

@app.route("/summarize", methods=["POST"])
def summarize():
    data = request.get_json()
    
    if not data or "text" not in data:
        return jsonify({"error": "Invalid request. 'text' field is required."}), 400

    text = data["text"]
    summary = summarize_text(text)
    
    return jsonify({"summary": summary})

if _name_ == "_main_":
    app.run(debug=True)@app.route("/", methods=["POST"])
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
