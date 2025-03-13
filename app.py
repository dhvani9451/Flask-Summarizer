from flask import Flask, request, send_file, jsonify
from flask_cors import CORS  # Import CORS
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import google.generativeai as genai  # Import Gemini API
import os

#  Configure Gemini API
genai.configure(api_key=os.getenv("AIzaSyBr12Wqh__1rPbCqwvyNFyNLgd3yDrniCM"))

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})  # Enable CORS for all origins

#  Function to generate summary using Gemini AI
def generate_summary(text):
    try:
        model = genai.GenerativeModel("models/gemini-2.0-flash")  # Using Gemini Pro model
        response = model.generate_content(text)
        return response.text.strip() if response.text else "No summary generated."
    except Exception as e:
        return f"Error generating summary: {str(e)}"

#  Function to generate a structured PowerPoint presentation
def create_ppt(summary):
    prs = Presentation()
    
    #  Title Slide
    slide_layout = prs.slide_layouts[0]  # Title slide layout
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "AI-Generated Summary"
    subtitle.text = "Created using PPT Summary Maker"

    # Process summary into slides
    bullet_slide_layout = prs.slide_layouts[1]  # Title & Content layout
    paragraphs = summary.split("\n")

    for i in range(0, len(paragraphs), 5):  # Group content per slide
        slide = prs.slides.add_slide(bullet_slide_layout)
        title = slide.shapes.title
        content = slide.placeholders[1].text_frame

        title.text = "Key Points"

        for line in paragraphs[i:i+5]:  # Add bullet points
            p = content.add_paragraph()
            p.text = f"• {line.strip()}"
            p.space_after = Pt(10)  # Adjust spacing
            p.font.size = Pt(24)  # Consistent font size

    #  Save PPT to memory
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

#  Flask route to generate and return PPT file
@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    data = request.json
    text = data.get("text", "")

    if not text:
        return jsonify({"error": "No text provided"}), 400

    summary = generate_summary(text)  # Generate summary using Gemini
    ppt_file = create_ppt(summary)
    return send_file(ppt_file, as_attachment=True, download_name="AI_Summary_Presentation.pptx")

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
