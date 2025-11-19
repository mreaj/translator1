import streamlit as st
from deep_translator import GoogleTranslator
from docx import Document
from io import BytesIO
import time

# -----------------------------
# Streamlit Page Setup
# -----------------------------
st.set_page_config(page_title="DOCX Translator", layout="centered")
st.title("📄🌍 DOCX File Translator")
st.write("Upload a DOCX file, select a language, and download the translated version.")

# -----------------------------
# Language Dictionary
# -----------------------------
languages = {
    "India – Hindi": "hi",
    "France – French": "fr",
    "United Kingdom – English": "en",
    "Poland – Polish": "pl",
    "Sweden – Swedish": "sv",
    "Finland – Finnish": "fi",
    "Italy – Italian": "it",
    "Japan – Japanese": "ja",
    "Netherlands – Dutch": "nl",
    "Germany – German": "de",
    "South Korea – Korean": "ko",
    "Australia – English": "en",
    "USA – English": "en",
    "Greece – Greek": "el",
    "Philippines – Filipino": "tl",
    "Egypt – Arabic": "ar",
    "Austria – German": "de",
    "South Africa – Afrikaans": "af",
    "Canada – English": "en",
    "Ireland – Irish (Gaelic)": "ga",
    "Curaçao – Dutch": "nl",
    "Belgium – Dutch": "nl",
    "International Waters – English": "en",
    "Taiwan – Mandarin Chinese": "zh-TW",
    "China – Chinese (Simplified)": "zh-CN",
    "Czech Republic – Czech": "cs",
    "Spain – Spanish": "es",
    "Mexico – Spanish": "es",
    "Brazil – Portuguese": "pt",
    "Turkey – Turkish": "tr",
    "Argentina – Spanish": "es",
    "Lithuania – Lithuanian": "lt",
    "Portugal – Portuguese": "pt",
    "Romania – Romanian": "ro",
    "Cyprus – Greek": "el",
    "Estonia – Estonian": "et",
    "Denmark – Danish": "da",
    "Croatia – Croatian": "hr",
}

# -----------------------------
# Safe Translation Function
# -----------------------------
def safe_translate(text, target_lang):
    """Translate safely without returning None."""
    if not text or text.strip() == "":
        return text  # empty text stays as is

    try:
        translated = GoogleTranslator(source="auto", target=target_lang).translate(text)
        if translated is None:
            return text
        return str(translated)
    except:
        return text  # fallback


# -----------------------------
# Count total blocks for progress bar
# -----------------------------
def count_blocks(doc):
    total = len(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                total += len(cell.paragraphs)
    return total


# -----------------------------
# Streamlit UI Elements
# -----------------------------
uploaded_file = st.file_uploader("Upload DOCX File", type=["docx"])
target_lang_label = st.selectbox("Translate To:", options=list(languages.keys()))
start_button = st.button("Translate Document")

if uploaded_file and start_button:

    target_lang = languages[target_lang_label]

    # Load DOCX
    doc = Document(uploaded_file)

    total_blocks = count_blocks(doc)
    completed = 0
    start_time = time.time()

    progress = st.progress(0)
    eta_text = st.empty()
    status_msg = st.empty()

    status_msg.info("Translating... Please wait.")

    # -----------------------------
    # Translate Paragraphs
    # -----------------------------
    for para in doc.paragraphs:
        for run in para.runs:
            original = run.text
            translated = safe_translate(original, target_lang)
            run.text = translated if translated else original

        completed += 1

        # Update progress bar and ETA
        percent = completed / total_blocks
        progress.progress(percent)

        elapsed = time.time() - start_time
        eta = (elapsed / max(completed, 1)) * (total_blocks - completed)
        eta_text.write(f"⏳ ETA: {eta:.1f} sec")

    # -----------------------------
    # Translate Tables
    # -----------------------------
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        original = run.text
                        translated = safe_translate(original, target_lang)
                        run.text = translated if translated else original

                    completed += 1

                    percent = completed / total_blocks
                    progress.progress(percent)

                    elapsed = time.time() - start_time
                    eta = (elapsed / max(completed, 1)) * (total_blocks - completed)
                    eta_text.write(f"⏳ ETA: {eta:.1f} sec")

    # -----------------------------
    # Save Output DOCX
    # -----------------------------
    output_buffer = BytesIO()
    doc.save(output_buffer)
    output_buffer.seek(0)

    status_msg.success("✔ Translation Complete!")

    # Download button
    st.download_button(
        label="⬇ Download Translated DOCX",
        data=output_buffer,
        file_name=f"translated_{target_lang}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
