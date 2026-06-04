import os
import json
import requests
import time
import pandas as pd
from urllib.parse import quote  # Standard way to encode URLs
from datetime import datetime
from groq import Groq
from docx import Document
from pptx import Presentation
from pptx.util import Inches
from fpdf import FPDF

# ==========================================================
# ⚙️ CONFIGURATION
# ==========================================================
GROQ_API_KEY = "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"

class UniversalAgent:
    def __init__(self, api_key):
        self.client = Groq(api_key=api_key, timeout=30.0)
        self.model = "llama-3.3-70b-versatile"
        self.session = requests.Session()
        # Adding a User-Agent prevents security filters from blocking requests
        self.session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })

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

    # --- 🎨 FIXED: IMAGE GEN (FOR PPT/PDF) ---
    def _generate_temp_image(self, prompt):
        try:
            # UPDATED: New unified 2026 endpoint
            # quote() ensures symbols like commas don't break the URL
            encoded_prompt = quote(prompt)
            url = f"https://gen.pollinations.ai/image/{encoded_prompt}?width=800&nologo=true"
            
            r = self.session.get(url, timeout=15)
            if r.status_code == 200:
                temp_name = f"temp_{int(time.time())}.jpg"
                with open(temp_name, "wb") as f:
                    f.write(r.content)
                return temp_name
            return None
        except Exception as e: 
            print(f"DEBUG: Image Error - {e}")
            return None

    # --- 📄 PDF GENERATION ---
    def create_pdf(self, topic, filename):
        try:
            if not filename.endswith(".pdf"): filename = f"{topic.replace(' ', '_')}_{int(time.time())}.pdf"
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("helvetica", 'B', 16)
            pdf.cell(0, 10, txt=topic.upper(), ln=True, align='C')
            
            # Add a visual
            img = self._generate_temp_image(f"A professional visual representing {topic}")
            if img:
                pdf.image(img, x=10, y=30, w=180)
                pdf.set_y(150)
                os.remove(img)

            pdf.set_font("helvetica", size=12)
            raw_text = self.get_text(f"Write a 3-paragraph report on {topic}")
            # FIX: Cleaning text for FPDF prevents crashes with smart quotes or emojis
            clean_text = raw_text.encode('latin-1', 'ignore').decode('latin-1')
            pdf.multi_cell(0, 10, txt=clean_text)
            
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

    # --- 🧠 FIXED: DISPATCHER ---
    def handle_request(self, user_prompt):
        brain_p = f"""
        User Prompt: "{user_prompt}"
        Identify: tool (pdf, ppt, excel, word, text), subject, filename.
        Return JSON ONLY: {{"tool": "...", "subject": "...", "file": "..."}}
        """
        try:
            # Attempt to get structured instructions
            raw_res = self.get_text(brain_p, is_json=True)
            res = json.loads(raw_res)
            
            t = res.get('tool', 'text')
            s = res.get('subject', user_prompt)
            f = res.get('file', 'output')

            # Route to the correct tool
            if t == 'pdf': result = self.create_pdf(s, f)
            elif t == 'ppt': result = self.create_ppt(s, f)
            elif t == 'excel': result = self.create_excel(s, f)
            elif t == 'word': result = self.create_word(s, f)
            else: 
                # For 'text' or greetings, return normal chat
                result = {"message": self.get_text(user_prompt), "file_path": None}

            return json.dumps(result)
            
        except (json.JSONDecodeError, Exception):
            # FALLBACK: If JSON fails (for "hi", etc.), return a normal chat response
            chat_reply = self.get_text(user_prompt)
            return json.dumps({"message": chat_reply, "file_path": None})

# --- 🏁 MAIN ---
if __name__ == "__main__":
    agent = UniversalAgent(GROQ_API_KEY)
    print("--- 🤖 Omni-Agent (Online) ---")
    while True:
        inp = input("\nYou: ").strip()
        if not inp: break
        
        # Parse the JSON string into a dictionary to access the 'message'
        response_data = json.loads(agent.handle_request(inp))
        print(f"Agent Message: {response_data['message']}")
        
        if response_data.get('file_path'):
            print(f"Download Link: ./{response_data['file_path']}")
