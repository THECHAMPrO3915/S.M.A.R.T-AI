import os
import json
import requests
import time
import re
import pandas as pd
from datetime import datetime
from groq import Groq # Keeping the Groq SDK intact
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
        self.client = Groq(api_key=api_key, timeout=45.0) # Bumped timeout slightly for the 120B model
        # Switched to the OpenAI Open-Weight model hosted on Groq
        self.model = "llama-3.3-70b-versatile"
        self.session = requests.Session()

    def get_text(self, prompt, is_json=False):
        # Reasoning models do best when given a strict developer/system hint to suppress 
        # conversational chatter when building data tables or structural files.
        system_instruction = (
            "You are a precise data extraction utility. Output your answer directly. "
            "Do not include conversational filler, pleasantries, or markdown formatting blocks."
        ) if is_json else "You are a professional report assistant."

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": prompt}
                ],
                response_format={"type": "json_object"} if is_json else None
            )
            content = response.choices[0].message.content
            
            if is_json:
                # Extra guardrail to strip out markdown wrappers if the model includes them
                content = re.sub(r"^```json\s*|\s*```$", "", content.strip(), flags=re.MULTILINE)
            return content
        except Exception as e:
            return "{}" if is_json else f"Error: {e}"

    def _sanitize_latin1(self, text):
        """Converts smart quotes and symbols to standard Latin-1 to prevent FPDF crashes."""
        replacements = {
            '\u2018': "'", '\u2019': "'",  
            '\u201c': '"', '\u201d': '"',  
            '\u2013': '-', '\u2014': '-',  
            '\u2026': '...'                
        }
        for sub, rep in replacements.items():
            text = text.replace(sub, rep)
        return text.encode('latin-1', 'replace').decode('latin-1')

    # --- 🎨 INTERNAL: IMAGE GEN (FOR PPT/PDF) ---
    def _generate_temp_image(self, prompt):
        try:
            url = f"https://image.pollinations.ai/prompt/{prompt.replace(' ', '_')}?width=800&nologo=true"
            r = self.session.get(url, timeout=10)
            if r.status_code == 200:
                temp_name = f"temp_{int(time.time())}.jpg"
                with open(temp_name, "wb") as f:
                    f.write(r.content)
                return temp_name
        except Exception:
            pass
        return None

    # --- 📄 PDF GENERATION ---
    def create_pdf(self, topic, filename):
        img = None
        try:
            if not filename.endswith(".pdf"): 
                filename = f"{topic.replace(' ', '_')}_{int(time.time())}.pdf"
            
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("helvetica", 'B', 16)
            pdf.cell(0, 10, txt=self._sanitize_latin1(topic.upper()), ln=True, align='C')
            
            # Add a visual
            img = self._generate_temp_image(f"visual of {topic}")
            if img:
                pdf.image(img, x=10, y=30, w=180)
                pdf.set_y(150)

            pdf.set_font("helvetica", size=12)
            raw_text = self.get_text(f"Write a clean 3-paragraph report on {topic}. Do not use bullet points or markdown bolding.")
            sanitized_text = self._sanitize_latin1(raw_text)
            
            pdf.multi_cell(0, 10, txt=sanitized_text)
            pdf.output(filename)
            return {"message": f"✅ PDF Created: {topic}", "file_path": filename}
        except Exception as e: 
            return {"message": f"❌ PDF Error: {e}", "file_path": None}
        finally:
            if img and os.path.exists(img):
                os.remove(img)

    # --- 📽️ PPT GENERATION ---
    def create_ppt(self, topic, filename, slides=3):
        try:
            if not filename.endswith(".pptx"): 
                filename = f"{topic.replace(' ', '_')}_{int(time.time())}.pptx"
            prs = Presentation()
            
            for i in range(slides):
                slide = prs.slides.add_slide(prs.slide_layouts[1]) 
                slide.shapes.title.text = f"{topic} - Slide {i+1}"
                
                raw_bullets = self.get_text(f"Provide 3 clean, concise bullet points for a slide deck about {topic} (Part {i+1}). Do not include introductory text, headers, or markdown asterisks.")
                bullet_points = [line.strip().lstrip('*-• ') for line in raw_bullets.split('\n') if line.strip()]
                
                tf = slide.placeholders[1].text_frame
                tf.clear() 
                
                for idx, pt in enumerate(bullet_points):
                    p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
                    p.text = pt
                    p.level = 0
                    
            prs.save(filename)
            return {"message": f"✅ PPT Created: {topic}", "file_path": filename}
        except Exception as e: 
            return {"message": f"❌ PPT Error: {e}", "file_path": None}

    # --- 📊 EXCEL GENERATION ---
    def create_excel(self, topic, filename):
        try:
            if not filename.endswith(".xlsx"): 
                filename = f"{topic.replace(' ', '_')}_{int(time.time())}.xlsx"
            
            prompt = (
                f"Generate a clear data table for {topic}. "
                "Your response must follow this exact structural JSON format layout: "
                '{"cols": ["Header1", "Header2"], "rows": [["Row1_Value1", "Row1_Value2"], ["Row2_Value1", "Row2_Value2"]]}'
            )
            
            raw_json = self.get_text(prompt, is_json=True)
            data = json.loads(raw_json)
            
            if 'rows' in data and 'cols' in data:
                df = pd.DataFrame(data['rows'], columns=data['cols'])
                df.to_excel(filename, index=False)
                return {"message": f"✅ Excel Created: {topic}", "file_path": filename}
            else:
                return {"message": "❌ Excel Error: LLM structural schema mismatch.", "file_path": None}
        except Exception as e: 
            return {"message": f"❌ Excel Error: {e}", "file_path": None}

    # --- 📝 WORD GENERATION ---
    def create_word(self, topic, filename):
        try:
            if not filename.endswith(".docx"): 
                filename = f"{topic.replace(' ', '_')}_{int(time.time())}.docx"
            doc = Document()
            doc.add_heading(topic, 0)
            
            report_text = self.get_text(f"Write a comprehensive, professional report about {topic}.")
            doc.add_paragraph(report_text)
            
            doc.save(filename)
            return {"message": f"✅ Word Doc Created: {topic}", "file_path": filename}
        except Exception as e: 
            return {"message": f"❌ Word Error: {e}", "file_path": None}

    # --- 🧠 DISPATCHER ---
    def handle_request(self, user_prompt):
        brain_p = f"""
        User Request: "{user_prompt}"
        Evaluate the intent and select the single best tool type (pdf, ppt, excel, word, text), identify the subject topic matter, and propose a clean short file name.
        Return JSON ONLY using exactly this scheme format: {{"tool": "...", "subject": "...", "file": "..."}}
        """
        try:
            res = json.loads(self.get_text(brain_p, is_json=True))
            t, s, f = res.get('tool'), res.get('subject'), res.get('file', 'output')

            if t == 'pdf': result = self.create_pdf(s, f)
            elif t == 'ppt': result = self.create_ppt(s, f)
            elif t == 'excel': result = self.create_excel(s, f)
            elif t == 'word': result = self.create_word(s, f)
            else: result = {"message": self.get_text(user_prompt), "file_path": None}

            return json.dumps(result)
        except Exception as e:
            return json.dumps({"message": f"Error parsing routing request: {e}", "file_path": None})

# --- 🏁 MAIN ---
if __name__ == "__main__":
    if GROQ_API_KEY == "your_key_here":
        print("⚠️ Warning: Please replace 'your_key_here' with a valid Groq API key.")
        
    agent = UniversalAgent(GROQ_API_KEY)
    print("--- 🤖 Omni-Agent (Running via Groq & GPT-OSS-120b) ---")
    while True:
        try:
            inp = input("\nYou: ").strip()
            if not inp: 
                break
            
            response_raw = agent.handle_request(inp)
            response_data = json.loads(response_raw)
            
            print(f"Agent Message: {response_data.get('message')}")
            if response_data.get('file_path'):
                print(f"Download Link: ./{response_data['file_path']}")
        except KeyboardInterrupt:
            print("\nExiting...")
            break
