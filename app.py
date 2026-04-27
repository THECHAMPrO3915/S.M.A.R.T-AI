import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader
from omni_agent import UniversalAgent

# 1. Page Configuration
st.set_page_config(layout="wide", page_title="S.M.A.R.T. AI")

# 2. Authentication Setup
with open('config.yaml') as file:
    config = yaml.load(file, Loader=SafeLoader)

authenticator = stauth.Authenticate(
    config['credentials'], config['cookie']['name'], 
    config['cookie']['key'], config['cookie']['expiry_days']
)

# 3. Login Flow
authenticator.login()

if st.session_state["authentication_status"]:
    # --- INITIALIZATION ---
    if "agent" not in st.session_state:
        st.session_state.agent = UniversalAgent("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")
    if "sessions" not in st.session_state:
        st.session_state.sessions = {"Chat 1": []}
    if "current_chat" not in st.session_state:
        st.session_state.current_chat = "Chat 1"

    # --- SIDEBAR ---
    with st.sidebar:
        st.title("S.M.A.R.T. AI Hub")
        if st.button("➕ New Session"):
            new_id = f"Chat {len(st.session_state.sessions) + 1}"
            st.session_state.sessions[new_id] = []
            st.session_state.current_chat = new_id
            st.rerun()
        
        st.divider()
        for chat_name in st.session_state.sessions.keys():
            if st.button(chat_name, key=chat_name, use_container_width=True):
                st.session_state.current_chat = chat_name
                st.rerun()
        
        st.divider()
        authenticator.logout('Logout', 'sidebar')

    # --- MAIN INTERFACE ---
    st.title(f"🤖 {st.session_state.current_chat}")

    # Render History
    for msg in st.session_state.sessions[st.session_state.current_chat]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # Chat Input
    if prompt := st.chat_input("Initiate command..."):
        st.session_state.sessions[st.session_state.current_chat].append({"role": "user", "content": prompt})
        
        # Get response from your agent
        response = st.session_state.agent.handle_request(prompt)
        
        st.session_state.sessions[st.session_state.current_chat].append({"role": "assistant", "content": response})
        st.rerun()

elif st.session_state["authentication_status"] is False:
    st.error("Username/password is incorrect")
elif st.session_state["authentication_status"] is None:
    st.warning("Please enter your credentials to access S.M.A.R.T. AI")