import streamlit as st
import os
import json
import tempfile
from omni_agent import UniversalAgent

st.set_page_config(
    page_title="Omni Agent Assistant",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 Universal Omni Agent")
st.markdown("Upload files, analyze documents or images, or request document generation (Word, Excel, PDF, PPT, Images, Videos).")

# Retrieve server URL from Streamlit Secrets, environment variables, or fallback to local
server_url = st.secrets.get("LLAMA_SERVER_URL", os.getenv("LLAMA_SERVER_URL", "http://localhost:8080/v1"))
agent = UniversalAgent(base_url=server_url)

# File Uploader
uploaded_file = st.file_uploader(
    "Upload context file or image (Optional):", 
    type=['png', 'jpg', 'jpeg', 'webp', 'pdf', 'docx', 'xlsx', 'csv', 'txt', 'pptx']
)

# Text Input
user_prompt = st.text_area("Enter your prompt or request:", height=120)

if st.button("Submit Request", type="primary"):
    if not user_prompt.strip():
        st.warning("Please enter a prompt to proceed.")
    else:
        with st.spinner("Processing request..."):
            file_path = None
            if uploaded_file:
                # Store uploaded file in isolated temporary directory
                temp_dir = tempfile.mkdtemp()
                file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

            # Send prompt and file_path to Omni Agent backend
            raw_response = agent.handle_request(user_prompt, file_path=file_path)

            try:
                result = json.loads(raw_response)
                
                # Display output text / message
                if "message" in result and result["message"]:
                    st.markdown("### Response")
                    st.write(result["message"])

                # Handle download for generated files
                gen_file = result.get("file_path")
                if gen_file and os.path.exists(gen_file):
                    st.markdown("---")
                    file_name = os.path.basename(gen_file)
                    with open(gen_file, "rb") as fp:
                        st.download_button(
                            label=f"📥 Download Output ({file_name})",
                            data=fp,
                            file_name=file_name,
                            mime="application/octet-stream"
                        )
            except Exception as e:
                st.error(f"Error parsing response: {e}")
                st.write(raw_response)