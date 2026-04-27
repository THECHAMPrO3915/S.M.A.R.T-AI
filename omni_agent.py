import os
import json
import requests
import time
import pandas as pd
from datetime import datetime
from groq import Groq
from docx import Document
from pptx import Presentation
from pptx.util import Inches
from fpdf import FPDF

# ==========================================================
# ⚙️ CONFIGURATION
# ==========================================================
# NOTE: Avoid hardcoding keys in public scripts. Use environment variables if possible.
GROQ_API_KEY = "PASTE_YOUR_GROQ_KEY_HERE"

class UniversalAgent:
    def __init__(self, api_key):
        self.client = Groq(api_key=api_key, timeout=30.0)
        self.model = "llama-3.3-70b-versatile"
        self.session = requests.Session()

    def get_text(self, prompt, is_json=False):
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"} if is_json else None
            )
            return response.choices[0].message.content
        except Exception as e:
            return "{}" if is_json else f"Error: {e}"

    # --- 🎨 IMAGE GENERATION ---
    def generate_image(self, prompt):
        print(f"--- 🎨 Generating Image: {prompt} ---")
        try:
            url = f"https://image.pollinations.ai/prompt/{prompt.replace(' ', '_')}?width=1024&height=1024&nologo=true"
            r = self.session.get(url)
            filename = f"IMG_{int(time.time())}.jpg"
            with open(filename, "wb") as f:
                f.write(r.content)
            return filename, f"✅ Image saved as: {filename}"
        except Exception as e:
            return None, f"❌ Image Error: {e}"

    # --- 📝 WORD GENERATION ---
    def create_word(self, topic, filename):
        print(f"--- 📝 Creating Word: {topic} ---")
        try:
            if not filename.endswith(".docx"): filename += ".docx"
            doc = Document()
            doc.add_heading(topic, 0)
            content = self.get_text(f"Write a comprehensive report about {topic}.")
            doc.add_paragraph(content)
            doc.save(filename)
            return f"✅ Word document saved as: {filename}"
        except Exception as e: return f"❌ Word Error: {e}"

    # --- 📊 EXCEL GENERATION ---
    def create_excel(self, topic, filename):
        print(f"--- 📊 Creating Excel: {topic} ---")
        try:
            if not filename.endswith(".xlsx"): filename += ".xlsx"
            data_prompt = f"Provide a dataset for {topic} with 5 rows. Return JSON ONLY: {{'columns': ['Col1', 'Col2'], 'data': [['val1', 'val2'], ['val3', 'val4']]}}"
            raw_data = json.loads(self.get_text(data_prompt, is_json=True))
            df = pd.DataFrame(raw_data['data'], columns=raw_data['columns'])
            df.to_excel(filename, index=False)
            return f"✅ Excel saved as: {filename}"
        except Exception as e: return f"❌ Excel Error: {e}"

    # --- 📄 PDF GENERATION WITH IMAGES ---
    def create_pdf(self, topic, filename):
        print(f"--- 📄 Creating PDF: {topic} ---")
        try:
            if not filename.endswith(".pdf"): filename += ".pdf"
            img_path, _ = self.generate_image(f"Educational illustration of {topic}")
            
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("helvetica", 'B', 16)
            pdf.cell(0, 10, txt=topic, ln=True, align='C')
            
            if img_path:
                pdf.image(img_path, x=10, y=30, w=180)
                pdf.set_y(150) # Move text below image
            
            pdf.set_font("helvetica", size=12)
            content = self.get_text(f"Summarize the core concepts of {topic} in 3 paragraphs.")
            pdf.multi_cell(0, 10, txt=content)
            
            pdf.output(filename)
            if img_path: os.remove(img_path) # Clean up temp image
            return f"✅ PDF with image saved as: {filename}"
        except Exception as e: return f"❌ PDF Error: {e}"

    # --- 📽️ PPT GENERATION WITH IMAGES ---
    def create_ppt(self, topic, filename, slides=5):
        print(f"--- 📽️ Creating {slides}-Slide PPT: {topic} ---")
        try:
            if not filename.endswith(".pptx"): filename += ".pptx"
            prs = Presentation()
            
            for i in range(slides):
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = f"{topic}: Slide {i+1}"
                
                # Add text
                info = self.get_text(f"Write 3 bullet points for a slide about {topic} - part {i+1}.")
                slide.placeholders[1].text = info
                
                # Add image to every slide
                img_path, _ = self.generate_image(f"Visualizing {topic} aspect {i+1}")
                if img_path:
                    slide.shapes.add_picture(img_path, Inches(6), Inches(2), width=Inches(3.5))
                    os.remove(img_path)

            prs.save(filename)
            return f"✅ PPT with images saved as: {filename}"
        except Exception as e: return f"❌ PPT Error: {e}"

    # --- 🧠 DISPATCHER ---
    def handle_request(self, user_prompt):
        brain_p = f"""
        Analyze: "{user_prompt}"
        Identify Tool:
        - 'image': create/generate a picture/visual.
        - 'word': make a word document/report.
        - 'excel': make a spreadsheet/table/excel.
        - 'pdf': create/make a PDF file.
        - 'ppt': create a powerpoint/slides.
        - 'text': general chat.
        Return JSON ONLY: {{"tool": "...", "subject": "...", "file": "filename.ext"}}
        """
        try:
            res = json.loads(self.get_text(brain_p, is_json=True))
            tool, subject, file = res.get('tool'), res.get('subject'), res.get('file', 'document')

            if tool == 'image': return self.generate_image(subject)[1]
            if tool == 'word': return self.create_word(subject, file)
            if tool == 'excel': return self.create_excel(subject, file)
            if tool == 'pdf': return self.create_pdf(subject, file)
            if tool == 'ppt': return self.create_ppt(subject, file)
            
            return self.get_text(user_prompt)
        except: return self.get_text(user_prompt)

# --- 🏁 MAIN ---
if __name__ == "__main__":
    agent = UniversalAgent(GROQ_API_KEY)
    print("--- 🤖 Universal Omni-Agent Online ---")
    while True:
        try:
            inp = input("\nYou: ").strip()
            if not inp or inp.lower() in ['exit', 'quit']: break
            print(f"Agent: {agent.handle_request(inp)}")
        except KeyboardInterrupt: break