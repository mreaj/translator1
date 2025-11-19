import streamlit as st
from deep_translator import GoogleTranslator
from docx import Document
from io import BytesIO
import time

st.set_page_config(page_title="DOCX Translator", layout="centered")
st.title("📄🌍 DOCX File Translator")
st.write("Upload a Word document and translate it to any language.")

# supported languages
languages = {
    "Hindi": "hi",
    "Tamil": "ta",
    "French": "fr",
    "German": "de",
    "Spanish": "es",
}

uploaded_file = st.file_uploader("Upload DOCX File", type=["docx"])
target_lang = st.selectbox("Translate To:", options=list(languages.keys()))
start_button = st.button("Translate Document")

def translate_text(text, target_lang):
    try:
        return GoogleTranslator(source='auto', target=target_lang).translate(text)
    except:
        return text

def count_blocks(doc):
    total = len(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                total += len(cell.paragraphs)
    return total

if uploaded_file and start_button:

    doc = Document(uploaded_file)
    target = languages[target_lang]

    total_blocks = count_blocks(doc)
    completed = 0
    start_time = time.time()

    progress = st.progress(0)
    eta_text = st.empty()
    info = st.empty()

    info.info("Translating...")

    # paragraphs
    for para in doc.paragraphs:
        for run in para.runs:
            run.text = translate_text(run.text, target)
        completed += 1
        progress.progress(completed / total_blocks)

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.text = translate_text(run.text, target)
                    completed += 1
                    progress.progress(completed / total_blocks)

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    info.success("✔ Translation Complete!")

    st.download_button(
        "⬇ Download Translated DOCX",
        data=output,
        file_name=f"translated_{languages[target_lang]}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
