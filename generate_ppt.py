from pptx import Presentation
from pptx.util import Pt
import io
import re
import spacy
import logging
from typing import Dict, List, Union

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PresentationGenerator:
    def __init__(self):
        self.nlp = self._load_nlp_model()

    def _load_nlp_model(self):
        """Safely load the spaCy NLP model with fallback"""
        try:
            nlp = spacy.load("en_core_web_sm")
            logger.info("Successfully loaded spaCy NLP model")
            return nlp
        except Exception as e:
            logger.error(f"Failed to load spaCy model: {str(e)}")
            return None

    def _clean_text(self, text: str) -> str:
        """Basic text cleaning"""
        if not isinstance(text, str):
            return ""
        return re.sub(r'\s+', ' ', text).strip()

    def split_into_thoughts(self, text: str) -> List[str]:
        """Robust text splitting with NLP fallback"""
        try:
            text = self._clean_text(text)
            if not text:
                return []

            # Fallback to simple splitting if NLP fails
            if not self.nlp:
                return [p.strip() for p in text.split('\n') if p.strip()]

            doc = self.nlp(text)
            thoughts = []
            current_thought = []
            
            for token in doc:
                # Detect thought boundaries
                if (token.is_space and len(current_thought) > 15) or \
                   (token.text.lower() in ['first', 'second', 'however', 'moreover']):
                    if current_thought:
                        thoughts.append("".join(current_thought).strip())
                    current_thought = [token.text_with_ws]
                else:
                    current_thought.append(token.text_with_ws)
            
            if current_thought:
                thoughts.append("".join(current_thought).strip())
            
            return thoughts

        except Exception as e:
            logger.error(f"Error in split_into_thoughts: {str(e)}")
            return [text]  # Fallback to original text

    def add_slide(self, prs: Presentation, title: str, thoughts: List[str]) -> None:
        """Safe slide addition with overflow protection"""
        try:
            slide_layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(slide_layout)
            
            # Title setup
            title_shape = slide.shapes.title
            title_shape.text = title[:100]  # Limit title length
            title_shape.text_frame.paragraphs[0].font.size = Pt(28)

            # Content setup
            body_shape = slide.shapes.placeholders[1]
            text_frame = body_shape.text_frame
            text_frame.clear()
            text_frame.word_wrap = True

            MAX_LINES = 6
            current_lines = 0

            for thought in thoughts:
                if current_lines >= MAX_LINES:
                    break
                    
                # Add main bullet point
                p = text_frame.add_paragraph()
                p.text = thought[:200]  # Increased limit for better context
                p.level = 0
                p.font.size = Pt(16)
                current_lines += 1

                # Handle overflow with sub-bullets
                if len(thought) > 200:
                    remaining = thought[200:]
                    p = text_frame.add_paragraph()
                    p.text = remaining[:200]
                    p.level = 1  # Sub-bullet for continuations
                    p.font.size = Pt(14)
                    current_lines += 1

        except Exception as e:
            logger.error(f"Error in add_slide: {str(e)}")
            raise

    def create_presentation(self, file_texts: Dict[str, str]) -> io.BytesIO:
        """Main method to create presentation with comprehensive error handling"""
        try:
            if not isinstance(file_texts, dict):
                raise ValueError("file_texts must be a dictionary")

            prs = Presentation()
            
            # Title slide
            title_slide = prs.slides.add_slide(prs.slide_layouts[0])
            title_slide.shapes.title.text = "Summary Presentation"
            title_slide.placeholders[1].text = "Created by AI PPT Generator"

            for filename, text in file_texts.items():
                try:
                    if not isinstance(text, str):
                        logger.warning(f"Skipping non-string content for {filename}")
                        continue

                    # Process text
                    thoughts = self.split_into_thoughts(text)
                    
                    # Split into slides (3 thoughts per slide)
                    for i in range(0, len(thoughts), 3):
                        chunk = thoughts[i:i+3]
                        title = f"{filename[:30]}" if i == 0 else f"{filename[:30]} (Continued)"
                        self.add_slide(prs, title, chunk)

                except Exception as e:
                    logger.error(f"Error processing {filename}: {str(e)}")
                    continue

            # Save to in-memory file
            ppt_io = io.BytesIO()
            prs.save(ppt_io)
            ppt_io.seek(0)
            return ppt_io

        except Exception as e:
            logger.error(f"Fatal error in create_presentation: {str(e)}")
            raise

# Example usage:
# generator = PresentationGenerator()
# ppt_bytes = generator.create_presentation({"file1.txt": "Your text here..."})