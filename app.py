import streamlit as st
import os
import json
import time
from omni_agent import UniversalAgent

GROQ_API_KEY = "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"

st.set_page_config(page_title="Universal AI Assistant", page_icon="🤖", layout="centered")

# --- INITIALIZE SESSION STATE ---
if "logged_out" not in st.session_state:
    st.session_state.logged_out = False

if "chats" not in st.session_state:
    st.session_state.chats = {}

if "current_chat_id" not in st.session_state:
    new_id = str(int(time.time()))
    st.session_state.current_chat_id = new_id
    st.session_state.chats[new_id] = {"title": "New Chat", "messages": []}

# --- LOGOUT SCREEN ---
if st.session_state.logged_out:
    st.title("👋 Logged Out")
    st.info("You have successfully logged out.")
    if st.button("Log Back In"):
        st.session_state.logged_out = False
        st.session_state.chats = {}
        new_id = str(int(time.time()))
        st.session_state.current_chat_id = new_id
        st.session_state.chats[new_id] = {"title": "New Chat", "messages": []}
        st.rerun()
    st.stop()

agent = UniversalAgent(api_key=GROQ_API_KEY)

# --- SIDEBAR CONTROLS ---
st.sidebar.title("🤖 Chat Control")

# 1. New Chat & Logout
col_new, col_out = st.sidebar.columns(2)
with col_new:
    if st.button("➕ New Chat", use_container_width=True):
        new_chat_id = str(int(time.time()))
        st.session_state.chats[new_chat_id] = {"title": "New Chat", "messages": []}
        st.session_state.current_chat_id = new_chat_id
        st.rerun()

with col_out:
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_out = True
        st.rerun()

st.sidebar.divider()
st.sidebar.subheader("💬 Chat History")

# 2. History List Selector
chat_ids = list(st.session_state.chats.keys())
for cid in reversed(chat_ids):
    c_title = st.session_state.chats[cid]["title"]
    is_active = (cid == st.session_state.current_chat_id)
    btn_label = f"📌 {c_title}" if is_active else f"💬 {c_title}"
    
    if st.sidebar.button(btn_label, key=f"hist_{cid}", use_container_width=True):
        st.session_state.current_chat_id = cid
        st.rerun()

# Get current chat object
current_chat = st.session_state.chats[st.session_state.current_chat_id]

# --- MAIN UI ---
st.title(current_chat["title"])

# Render Chat History
for idx, msg in enumerate(current_chat["messages"]):
    with st.chat_message(msg["role"]):
        st.write(msg["content"])
        if msg.get("attachment_name"):
            st.caption(f"📎 Attached: `{msg['attachment_name']}`")
        if msg.get("file_path") and os.path.exists(msg["file_path"]):
            with open(msg["file_path"], "rb") as f:
                btn_data = f.read()
            st.download_button(
                label=f"📥 Download {os.path.basename(msg['file_path'])}",
                data=btn_data,
                file_name=os.path.basename(msg["file_path"]),
                key=f"dl_{st.session_state.current_chat_id}_{idx}"
            )

# --- PROMPT BAR ATTACHMENT WIDGET ---
st.write("---")
uploaded_file = None
with st.popover("📎 Attach File (Docs, Images, Videos, PPT)", use_container_width=True):
    uploaded_file = st.file_uploader(
        "Supported: PDF, DOCX, TXT, CSV, XLSX, PPTX, JPG, PNG, WEBP, MP4, MOV, MKV, AVI",
        type=["pdf", "docx", "txt", "csv", "xlsx", "pptx", "ppt", "jpg", "jpeg", "png", "webp", "mp4", "mov", "mkv", "avi"],
        key=f"uploader_{st.session_state.current_chat_id}"
    )

temp_file_path = None
if uploaded_file is not None:
    temp_dir = "uploaded_temp"
    os.makedirs(temp_dir, exist_ok=True)
    temp_file_path = os.path.join(temp_dir, uploaded_file.name)
    with open(temp_file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.toast(f"Attached file: {uploaded_file.name}", icon="📎")

# --- CHAT INPUT PROCESSOR ---
if user_prompt := st.chat_input("Ask anything or request a PDF, Word, Excel, PPT, or Image generation..."):
    # Dynamic Auto-Naming for First Message
    if current_chat["title"] == "New Chat":
        auto_title = agent.generate_chat_title(user_prompt)
        current_chat["title"] = auto_title

    # Append User Message
    user_msg_data = {"role": "user", "content": user_prompt}
    if uploaded_file:
        user_msg_data["attachment_name"] = uploaded_file.name
    current_chat["messages"].append(user_msg_data)

    with st.chat_message("user"):
        st.write(user_prompt)
        if uploaded_file:
            st.caption(f"📎 Attached: `{uploaded_file.name}`")

    # Generate Assistant Response
    with st.chat_message("assistant"):
        with st.spinner("Processing request..."):
            raw_response = agent.handle_request(user_prompt, file_path=temp_file_path)
            
            try:
                response_data = json.loads(raw_response)
            except Exception:
                response_data = {"message": raw_response, "file_path": None}
                
            msg_text = response_data.get("message", "Completed.")
            file_path = response_data.get("file_path")
            
            st.write(msg_text)
            
            if file_path and os.path.exists(file_path):
                with open(file_path, "rb") as f:
                    btn_data = f.read()
                
                st.download_button(
                    label=f"📥 Download {os.path.basename(file_path)}",
                    data=btn_data,
                    file_name=os.path.basename(file_path),
                    key=f"dl_new_{int(time.time() * 1000)}"
                )
            
            current_chat["messages"].append({
                "role": "assistant",
                "content": msg_text,
                "file_path": file_path
            })
            st.rerun()