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
GROQ_API_KEY = "your_key_here"

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

    # --- 🎨 INTERNAL: IMAGE GEN (FOR PPT/PDF) ---
    def _generate_temp_image(self, prompt):
        try:
            url = f"https://image.pollinations.ai/prompt/{prompt.replace(' ', '_')}?width=800&nologo=true"
            r = self.session.get(url)
            temp_name = f"temp_{int(time.time())}.jpg"
            with open(temp_name, "wb") as f:
                f.write(r.content)
            return temp_name
        except: return None

    # --- 📄 PDF GENERATION ---
    def create_pdf(self, topic, filename):
        try:
            if not filename.endswith(".pdf"): filename = f"{topic.replace(' ', '_')}_{int(time.time())}.pdf"
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("helvetica", 'B', 16)
            pdf.cell(0, 10, txt=topic.upper(), ln=True, align='C')
            
            # Add a visual
            img = self._generate_temp_image(f"visual of {topic}")
            if img:
                pdf.image(img, x=10, y=30, w=180)
                pdf.set_y(150)
                os.remove(img)

            pdf.set_font("helvetica", size=12)
            pdf.multi_cell(0, 10, txt=self.get_text(f"Write a 3-paragraph report on {topic}"))
            pdf.output(filename)
            return {"message": f"✅ PDF Created: {topic}", "file_path": filename}
        except Exception as e: return {"message": f"❌ PDF Error: {e}", "file_path": None}

    # --- 📽️ PPT GENERATION ---
    def create_ppt(self, topic, filename, slides=3):
        try:
            if not filename.endswith(".pptx"): filename = f"{topic.replace(' ', '_')}_{int(time.time())}.pptx"
            prs = Presentation()
            for i in range(slides):
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = f"{topic} - Slide {i+1}"
                slide.placeholders[1].text = self.get_text(f"Bullet points for {topic}, part {i+1}")
            prs.save(filename)
            return {"message": f"✅ PPT Created: {topic}", "file_path": filename}
        except Exception as e: return {"message": f"❌ PPT Error: {e}", "file_path": None}

    # --- 📊 EXCEL GENERATION ---
    def create_excel(self, topic, filename):
        try:
            if not filename.endswith(".xlsx"): filename = f"{topic.replace(' ', '_')}_{int(time.time())}.xlsx"
            data = json.loads(self.get_text(f"Table data for {topic}. JSON: {{'cols':['A','B'],'rows':[['1','2']]}}", is_json=True))
            df = pd.DataFrame(data['rows'], columns=data['cols'])
            df.to_excel(filename, index=False)
            return {"message": f"✅ Excel Created: {topic}", "file_path": filename}
        except Exception as e: return {"message": f"❌ Excel Error: {e}", "file_path": None}

    # --- 📝 WORD GENERATION ---
    def create_word(self, topic, filename):
        try:
            if not filename.endswith(".docx"): filename = f"{topic.replace(' ', '_')}_{int(time.time())}.docx"
            doc = Document()
            doc.add_heading(topic, 0)
            doc.add_paragraph(self.get_text(f"Write a report about {topic}"))
            doc.save(filename)
            return {"message": f"✅ Word Doc Created: {topic}", "file_path": filename}
        except Exception as e: return {"message": f"❌ Word Error: {e}", "file_path": None}

    # --- 🧠 DISPATCHER (RETURNS JSON DATA) ---
    def handle_request(self, user_prompt):
        brain_p = f"""
        User Prompt: "{user_prompt}"
        Identify: tool (pdf, ppt, excel, word, text), subject, filename.
        Return JSON ONLY: {{"tool": "...", "subject": "...", "file": "..."}}
        """
        try:
            res = json.loads(self.get_text(brain_p, is_json=True))
            t, s, f = res.get('tool'), res.get('subject'), res.get('file', 'output')

            # Route to correct tool
            if t == 'pdf': result = self.create_pdf(s, f)
            elif t == 'ppt': result = self.create_ppt(s, f)
            elif t == 'excel': result = self.create_excel(s, f)
            elif t == 'word': result = self.create_word(s, f)
            else: result = {"message": self.get_text(user_prompt), "file_path": None}

            return json.dumps(result) # Return as string for your app to parse
        except:
            return json.dumps({"message": self.get_text(user_prompt), "file_path": None})

# --- 🏁 MAIN ---
if __name__ == "__main__":
    agent = UniversalAgent(GROQ_API_KEY)
    print("--- 🤖 Omni-Agent (Download Ready) ---")
    while True:
        inp = input("\nYou: ").strip()
        if not inp: break
        # The output here is a JSON string containing the file path!
        response_data = json.loads(agent.handle_request(inp))
        print(f"Agent Message: {response_data['message']}")
        if response_data['file_path']:
            print(f"Download Link: ./{response_data['file_path']}")
