from flask import Flask, request, send_file
from pptx import Presentation
import google.generativeai as genai
import io

app = Flask(__name__)

# Configure Gemini API
genai.configure(api_key="AIzaSyBr12Wqh__1rPbCqwvyNFyNLgd3yDrniCM")  # Replace with your actual Gemini API Key

#  Function to generate AI summary using Gemini
def generate_summary(text):
    model = genai.GenerativeModel("models/gemini-1.5-pro-002")
    response = model.generate_content(text)
    return response.text if response.text else "Summary not available."

#  Function to create PowerPoint presentation
def create_presentation(summary):
    prs = Presentation()

    # Title Slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Generated Summary"
    slide.placeholders[1].text = "Created by AI Summarizer (Gemini)"

    # Content Slide
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Summary Key Points"
    text_frame = slide.placeholders[1].text_frame
    for line in summary.split("\n"):
        p = text_frame.add_paragraph()
        p.text = "• " + line.strip()

    # Save PPT to memory & return
    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

@app.route('/generate-ppt', methods=['POST'])
def generate_ppt():
    data = request.json
    text = data.get("text", "")
    
    if not text:
        return {"error": "No text provided"}, 400

    summary = generate_summary(text)
    ppt_file = create_presentation(summary)

    return send_file(ppt_file, as_attachment=True, download_name="Generated_Summary_Presentation.pptx")

if __name__ == '__main__':
    app.run(debug=True)
