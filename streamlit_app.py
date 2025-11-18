import streamlit as st
from deep_translator import GoogleTranslator
from docx import Document
from io import BytesIO
import time

st.set_page_config(page_title="DOCX Translator", layout="centered")
st.title("📄🌍 DOCX File Translator")
st.write("Upload a Word document and translate it to any language.")

# Language codes
languages = {
    "English": "en",
    "Tamil": "ta",
    "Hindi": "hi",
    "French": "fr",
    "Spanish": "es",
    "German": "de",
    "Chinese (Simplified)": "zh-cn",
    "Japanese": "ja",
    "Arabic": "ar",
}

# Upload docx
uploaded_file = st.file_uploader("Upload DOCX file", type=["docx"])

target_language = st.selectbox("Translate to:", list(languages.keys()))

if uploaded_file and st.button("Translate Document"):
    # Load document
    doc = Document(uploaded_file)
    translator = GoogleTranslator(source='auto', target=languages[target_language])

    paragraphs = doc.paragraphs
    total = len(paragraphs)

    st.info(f"Total paragraphs to translate: {total}")

    progress = st.progress(0)
    status = st.empty()

    start_time = time.time()

    # Translate paragraph-by-paragraph
    for i, para in enumerate(paragraphs):
        original = para.text.strip()
        if original:
            try:
                translated_text = translator.translate(original)
                para.text = translated_text
            except Exception:
                para.text = "[Translation Failed]"

        progress.progress((i + 1) / total)
        elapsed = time.time() - start_time
        est_total = (elapsed / (i+1)) * total
        remaining = est_total - elapsed

        status.text(
            f"Translating paragraph {i+1}/{total} • "
            f"Elapsed: {int(elapsed)} sec • "
            f"Remaining: {int(remaining)} sec"
        )

    progress.progress(1.0)
    status.text("Translation Complete ✔️")

    # Export translated file
    output = BytesIO()
    doc.save(output)
    output.seek(0)

    st.success("Your translated document is ready.")

    st.download_button(
        label="⬇ Download Translated DOCX",
        data=output,
        file_name=f"translated_{target_language}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
