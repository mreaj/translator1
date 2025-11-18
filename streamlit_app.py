import streamlit as st
from deep_translator import GoogleTranslator

st.set_page_config(page_title="Streamlit Translator", layout="centered")

st.title("🌍 Language Translator")
st.write("Translate text instantly between more than 100 languages.")

languages = {
    "English": "en",
    "Tamil": "ta",
    "Hindi": "hi",
    "French": "fr",
    "German": "de",
    "Spanish": "es",
    "Chinese (Simplified)": "zh-cn",
    "Japanese": "ja",
}

text = st.text_area("Enter text to translate:")

source = st.selectbox("From", list(languages.keys()), index=0)
target = st.selectbox("To", list(languages.keys()), index=1)

if st.button("Translate"):
    try:
        translated = GoogleTranslator(
            source=languages[source],
            target=languages[target]
        ).translate(text)
        
        st.success(translated)
    except Exception as e:
        st.error("Translation failed.")
