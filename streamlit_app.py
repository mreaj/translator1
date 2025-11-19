import streamlit as st
import io
import time
from docx import Document
from googletrans import Translator
from io import BytesIO

st.set_page_config(page_title="DOCX Translator", layout="centered")
st.title("📄🌍 DOCX Translator (GoogleTrans Version)")
st.write("Upload a DOCX file, select a language, and download the translated version.")

# ----------------------------
# Helper translation functions
# ----------------------------

def translate_text(text, target_lang, translator):
    try:
        return translator.translate(text, dest=target_lang).text
    except:
        return text


def translate_paragraph(paragraph, target_lang, translator):
    for run in paragraph.runs:
        translated = translate_text(run.text, target_lang, translator)
        run.text = translated


def translate_table(table, target_lang, translator):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                translate_paragraph(paragraph, target_lang, translator)


def count_blocks(doc):
    total = len(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                total += len(cell.paragraphs)
    return total


# ----------------------------
# Streamlit UI
# ----------------------------

uploaded_file = st.file_uploader("Upload DOCX File", type=["docx"])

target_lang = st.selectbox(
    "Translate To:",
    options={
        "Hindi (hi)": "hi",
        "Tamil (ta)": "ta",
        "French (fr)": "fr",
        "German (de)": "de",
        "Spanish (es)": "es"
    }
)

start_button = st.button("Translate Document")


# ----------------------------
# Translation Logic (Streamlit)
# ----------------------------

if uploaded_file and start_button:

    # Load DOCX
    doc = Document(uploaded_file)
    translator = Translator()

    total_blocks = count_blocks(doc)
    completed = 0
    start_time = time.time()

    progress = st.progress(0)
    eta_text = st.empty()
    status_msg = st.empty()

    status_msg.info("Translating... Please wait.")

    # Translate paragraphs
    for paragraph in doc.paragraphs:
        translate_paragraph(paragraph, target_lang, translator)
        completed += 1

        # Update progress
        percent = completed / total_blocks
        progress.progress(percent)

        elapsed = time.time() - start_time
        eta = (elapsed / completed) * (total_blocks - completed)
        eta_text.write(f"⏳ ETA: {eta:.1f} sec")

    # Translate tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    translate_paragraph(paragraph, target_lang, translator)
                    completed += 1

                    # Update progress
                    percent = completed / total_blocks
                    progress.progress(percent)

                    elapsed = time.time() - start_time
                    eta = (elapsed / completed) * (total_blocks - completed)
                    eta_text.write(f"⏳ ETA: {eta:.1f} sec")

    # Save output
    output_buffer = BytesIO()
    doc.save(output_buffer)
    output_buffer.seek(0)

    status_msg.success("✔ Translation Complete!")

    # Download button
    st.download_button(
        label="⬇ Download Translated DOCX",
        data=output_buffer,
        file_name=f"translated_{target_lang}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
