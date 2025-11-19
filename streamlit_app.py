import streamlit as st
from deep_translator import GoogleTranslator
from docx import Document
from io import BytesIO
import time


st.set_page_config(page_title="DOCX Translator", layout="centered")
st.title("📄🌍 DOCX File Translator")
st.write("Upload a DOCX file, select a language, and download the translated version.")


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


def safe_translate(text, target_lang):
    if not text or text.strip() == "":
        return text
    try:
        translated = GoogleTranslator(source="auto", target=target_lang).translate(text)
        return translated if translated else text
    except:
        return text


def count_blocks(doc):
    total = len(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                total += len(cell.paragraphs)
    return total


def format_eta(seconds):
    if seconds < 60:
        return f"{seconds:.1f} sec"
    else:
        return f"{seconds / 60:.1f} min"


uploaded_file = st.file_uploader("Upload DOCX File", type=["docx"])
target_label = st.selectbox("Translate To:", list(languages.keys()))

if st.button("Translate Document") and uploaded_file:
    target = languages[target_label]
    doc = Document(uploaded_file)

    total_blocks = count_blocks(doc)
    completed = 0
    start_time = time.time()

    st.info(f"🔢 Total items to translate: {total_blocks}")

    progress = st.progress(0)
    eta_text = st.empty()
    status_msg = st.empty()

    status_msg.info("Translating... Please wait...")

    for para in doc.paragraphs:
        for run in para.runs:
            run.text = safe_translate(run.text, target)

        completed += 1
        progress.progress(completed / total_blocks)

        elapsed = time.time() - start_time
        eta_text.write(f"⏳ ETA: {format_eta((elapsed / completed) * (total_blocks - completed))}")

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.text = safe_translate(run.text, target)

                    completed += 1
                    progress.progress(completed / total_blocks)

                    elapsed = time.time() - start_time
                    eta_text.write(f"⏳ ETA: {format_eta((elapsed / completed) * (total_blocks - completed))}")

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    status_msg.success("🎉 Translation Complete!")

    st.download_button(
        "⬇ Download Translated DOCX",
        data=output,
        file_name=f"translated_{target}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
