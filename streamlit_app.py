import streamlit as st
from deep_translator import GoogleTranslator
from docx import Document
from io import BytesIO
import time
import math

st.set_page_config(page_title="DOCX Translator", layout="centered")
st.title("📄🌍 DOCX File Translator")
st.write("Upload a Word document and translate it to any language.")

# Supported languages
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

# FIXED translate function
def translate_text(text, translator, max_len=300):
    if not text:
        return ""

    text = str(text).strip()
    chunks = [text[i:i+max_len] for i in range(0, len(text), max_len)]
    translated_chunks = []

    for chunk in chunks:
        try:
            result = translator.translate(chunk)

            # Deep translator returns a string directly
            if result:
                translated_chunks.append(str(result))
            else:
                translated_chunks.append(chunk)
        except:
            translated_chunks.append(chunk)

    return " ".join(translated_chunks)


def translate_document(doc, target_code):
    translator = GoogleTranslator(source='auto', target=target_code)

    elements = []

    # Paragraphs
    for p in doc.paragraphs:
        elements.append(p)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    elements.append(p)

    total = len(elements)
    progress = st.progress(0)
    status = st.empty()
    start = time.time()

    for idx, item in enumerate(elements):
        original = item.text
        translated = translate_text(original, translator)
        item.text = translated

        # update progress
        progress.progress((idx + 1) / total)
        elapsed = time.time() - start
        estimated_total = (elapsed / (idx+1)) * total
        remaining = estimated_total - elapsed

        status.text(
            f"Translating {idx+1}/{total} items • "
            f"Elapsed: {int(elapsed)} sec • "
            f"Remaining: {int(remaining)} sec"
        )

    progress.progress(1.0)
    status.text("✔ Translation Completed")

    return doc


# Upload DOCX
uploaded_file = st.file_uploader("Upload DOCX file", type=["docx"])
target_language = st.selectbox("Translate to:", list(languages.keys()))

if uploaded_file and st.button("Translate Document"):
    doc = Document(uploaded_file)
    translated_doc = translate_document(doc, languages[target_language])

    output = BytesIO()
    translated_doc.save(output)
    output.seek(0)

    st.success("Your translated document is ready!")

    st.download_button(
        label="⬇ Download Translated DOCX",
        data=output,
        file_name=f"translated_{target_language}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
