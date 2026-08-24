import streamlit_authenticator as stauth

# Use the static hash method directly
hashed_pw = stauth.Hasher.hash('admin123')
print(hashed_pw)