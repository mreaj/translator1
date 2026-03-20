import streamlit as st
from deep_translator import GoogleTranslator
from docx import Document
from docx.oxml.ns import qn
from io import BytesIO
import time
import re

st.set_page_config(page_title="DOCX Translator", layout="centered")
st.title("📄🌍 DOCX File Translator")
st.write("Upload a DOCX file, select a language, and download the translated version.")

languages = {
    "India – Hindi": "hi",
    "India – Tamil": "ta",
    "India – Telugu": "te",
    "India – Kannada": "kn",
    "India – Malayalam": "ml",
    "India – Gujarati": "gu",
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

# ======================
# WIND INDUSTRY GLOSSARY
# Keys are wrong/generic translations Google might produce.
# Values are the correct wind-industry terms per language.
# ======================
WIND_GLOSSARY = {
    "es": {
        "cuchillas": "palas", "cuchilla": "pala", "aspas": "palas", "aspa": "pala",
        "paletas": "palas", "paleta": "pala", "hoja": "pala", "hojas": "palas",
        "veletas": "palas", "veleta": "pala", "cabina": "góndola",
        "multiplicador": "multiplicadora", "caja de cambios": "multiplicadora",
        "caja de velocidades": "multiplicadora", "cubo": "buje", "centro": "buje",
        "turbina de viento": "aerogenerador", "molino de viento": "aerogenerador",
        "generador eólico": "aerogenerador", "parque de viento": "parque eólico",
        "granja eólica": "parque eólico", "granja de viento": "parque eólico",
        "control de cabeceo": "control de paso", "control de inclinación": "control de paso",
        "guiñada": "orientación", "puesta en servicio": "puesta en marcha",
        "comisionamiento": "puesta en marcha", "puesta en funcionamiento": "puesta en marcha",
        "terminación mecánica": "finalización mecánica",
        "accidente fatal": "accidente mortal", "cuasi accidente": "cuasi-accidente",
        "casi accidente": "cuasi-accidente", "estación secundaria": "subestación",
        "cable de arreglo": "cable de interconexión", "pilote único": "monopilote",
        "chaqueta": "estructura de celosía", "guindaste": "grúa",
    },
    "pl": {
        "skrzydła": "łopaty", "skrzydło": "łopata", "łopatki": "łopaty", "łopatka": "łopata",
        "wiatrak": "turbina wiatrowa", "park wiatrowy": "farma wiatrowa",
        "farma wiatrakowa": "farma wiatrowa", "skrzynia biegów": "przekładnia",
        "przekładnia zębata": "przekładnia", "kabina": "gondola", "piasta koła": "piasta",
        "oddanie do eksploatacji": "uruchomienie", "kąt skoku": "skok",
        "kąt łopaty": "skok", "ster": "odchylenie",
        "prawie wypadek": "zdarzenie potencjalnie wypadkowe", "dźwig": "żuraw",
    },
    "de": {
        "blätter": "Rotorblätter", "flügel": "Rotorblätter", "rotorflügel": "Rotorblätter",
        "schaufeln": "Rotorblätter", "schaufel": "Rotorblatt",
        "windmühle": "Windkraftanlage", "windrad": "Windkraftanlage",
        "windturbine": "Windkraftanlage", "windfarm": "Windpark",
        "windgenerator": "Windkraftanlage", "kabine": "Gondel", "radnabe": "Nabe",
        "zahnradgetriebe": "Getriebe", "inbetriebsetzung": "Inbetriebnahme",
        "azimutwinkel": "Azimut", "gierwinkel": "Azimut",
        "blattanstellwinkel": "Blattwinkel", "anstellwinkel": "Blattwinkel",
        "beinahunfall": "Beinaheunfall", "einzelpfahl": "Monopfahl",
    },
    "fr": {
        "lames": "pales", "ailes": "pales", "aile": "pale",
        "turbine éolienne": "éolienne", "moulin à vent": "éolienne",
        "ferme éolienne": "parc éolien", "ferme de vent": "parc éolien",
        "boîte de vitesses": "multiplicateur", "transmission": "multiplicateur",
        "cabine": "nacelle", "centre": "moyeu", "commissionning": "mise en service",
        "lacet": "orientation", "pieu unique": "monopieu",
        "accident fatal": "accident mortel", "presque accident": "quasi-accident",
    },
    "it": {
        "lame": "pale", "turbina eolica": "aerogeneratore",
        "mulino a vento": "aerogeneratore", "fattoria eolica": "parco eolico",
        "scatola del cambio": "moltiplicatore", "trasmissione": "moltiplicatore",
        "cabina": "navicella", "messa in funzione": "messa in servizio",
        "incidente fatale": "incidente mortale", "quasi incidente": "quasi-incidente",
    },
    "nl": {
        "bladen": "rotorbladen", "blad": "rotorblad", "vleugels": "rotorbladen",
        "windmolen": "windturbine", "tandwielkast": "versnellingsbak",
        "cabine": "gondel", "spoedregeling": "bladhoekregeling",
        "fataal ongeluk": "dodelijk ongeluk",
    },
    "pt": {
        "lâminas": "pás", "turbina eólica": "aerogerador",
        "moinho de vento": "aerogerador", "fazenda eólica": "parque eólico",
        "caixa de engrenagens": "multiplicadora", "transmissão": "multiplicadora",
        "cabine": "nacele", "posta em serviço": "entrada em operação",
        "comissionamento": "entrada em operação",
    },
    "sv": {
        "blad": "rotorblad", "vingar": "rotorblad", "vinge": "rotorblad",
        "vindturbin": "vindkraftverk", "vindmölla": "vindkraftverk",
        "kugghjulsväxel": "växellåda", "kabin": "gondol",
        "driftsättning": "idrifttagning", "stegkontroll": "bladvinkelreglering",
    },
    "da": {
        "blade": "rotorblade", "vinger": "rotorblade",
        "vindkraftværk": "vindmølle", "vindpark": "vindmøllepark",
        "kabine": "nacelle", "pitchkontrol": "pitchregulering",
    },
    "fi": {
        "lavat": "roottorin lavat", "lapa": "roottorin lapa",
        "tuuliturbiini": "tuulivoimala", "tuulimylly": "tuulivoimala",
        "hajautus": "suuntaus",
    },
    "ja": {
        "刃": "ブレード", "羽根": "ブレード", "翼": "ブレード",
        "風力タービン": "風力発電機", "風車": "風力発電機",
        "風力団地": "ウインドファーム", "ギアボックス": "増速機",
        "変速機": "増速機", "キャビン": "ナセル", "致命的事故": "死亡事故",
        "ニアミス": "ヒヤリハット",
    },
    "ko": {
        "날": "블레이드", "날개": "블레이드", "풍력 터빈": "풍력 발전기",
        "풍차": "풍력 발전기", "풍력 단지": "풍력 발전 단지",
        "기어박스": "증속기", "요": "요잉", "아찔한 순간": "아차 사고",
    },
    "zh-CN": {
        "刀片": "叶片", "刀": "叶片", "桨叶": "叶片",
        "风力涡轮机": "风力发电机", "风车": "风力发电机",
        "风电农场": "风电场", "风能农场": "风电场", "变速箱": "齿轮箱",
        "试运行": "调试", "致命事故": "死亡事故", "险兆": "未遂事故", "吊车": "起重机",
    },
    "zh-TW": {
        "刀片": "葉片", "刀": "葉片", "風力渦輪機": "風力發電機",
        "風車": "風力發電機", "風電農場": "風電場", "試運行": "調試", "致命事故": "死亡事故",
    },
}

# Patterns that must NEVER be translated
PRESERVE_PATTERNS = [
    r'\bIN\.\d{7,12}\b',
    r'\b\d+(?:\.\d+)?\s*(?:MW|kW|m/s|rpm|Hz|kV|MWh|kWh)\b',
    r'\bIEC\s*\d+[-\w]*\b',
    r'\bISO\s*\d+\b',
    r'\bDNV[-\s]\w+\b',
    r'\bVAS\w*\b',   # Vestas product codes e.g. VAS
]

# Labels that are field headings — translate the label but NOT dates/codes after ":"
# These are matched case-insensitively in the source text
FIELD_LABELS = [
    "Date", "Valid until date", "From", "Notification No",
    "Contact person", "Re", "Health", "Safety", "Env",
    "HSE Notification", "Restricted",
]


def apply_wind_glossary(text: str, lang: str) -> str:
    glossary = WIND_GLOSSARY.get(lang, {})
    for wrong, correct in glossary.items():
        pattern = re.compile(re.escape(wrong), re.IGNORECASE | re.UNICODE)
        def _replace(m, c=correct):
            return c[0].upper() + c[1:] if m.group(0)[0].isupper() else c
        text = pattern.sub(_replace, text)
    return text


def protect_text(text: str):
    placeholders = {}
    for pat in PRESERVE_PATTERNS:
        for m in re.finditer(pat, text):
            original = m.group(0)
            if original not in placeholders.values():
                token = f"__PH{len(placeholders)}__"
                placeholders[token] = original
    for token, original in placeholders.items():
        text = text.replace(original, token)
    return text, placeholders


def restore_text(text: str, placeholders: dict) -> str:
    for token, original in placeholders.items():
        text = text.replace(token, original)
    return text


def safe_translate(text: str, target_lang: str) -> str:
    if not text or text.strip() == "":
        return text
    protected, placeholders = protect_text(text)
    try:
        translated = GoogleTranslator(source="auto", target=target_lang).translate(protected)
        if not translated:
            translated = protected
    except Exception:
        translated = protected
    translated = restore_text(translated, placeholders)
    translated = apply_wind_glossary(translated, target_lang)
    return translated


def run_fmt_key(run):
    try:
        color = run.font.color.rgb if run.font.color and run.font.color.type else None
    except Exception:
        color = None
    return (run.bold, run.italic, run.underline, run.font.size, run.font.name, color)


def translate_paragraph(para, target_lang: str):
    """Translate paragraph runs, merging same-format runs to avoid split line breaks."""
    if not para.runs:
        return
    groups = []
    for run in para.runs:
        key = run_fmt_key(run)
        if groups and groups[-1][0] == key:
            groups[-1][1].append(run)
        else:
            groups.append((key, [run]))
    for _key, runs in groups:
        combined = "".join(r.text for r in runs)
        if not combined.strip():
            continue
        translated = safe_translate(combined, target_lang)
        runs[0].text = translated
        for r in runs[1:]:
            r.text = ""


def translate_xml_runs(xml_element, target_lang: str):
    """
    Directly translate all <w:r><w:t> run nodes inside any XML element.
    Used for text boxes, headers, footers, and content controls that
    python-docx doesn't expose as paragraph objects.
    """
    for t_node in xml_element.iter(qn("w:t")):
        original = t_node.text or ""
        if original.strip():
            t_node.text = safe_translate(original, target_lang)


def collect_paragraphs_from_element(xml_element):
    """Yield python-docx-style paragraph objects from any XML container."""
    from docx.text.paragraph import Paragraph
    for p_node in xml_element.iter(qn("w:p")):
        yield Paragraph(p_node, xml_element)


def count_all_blocks(doc):
    """Count every translatable block across body, headers, footers, text boxes."""
    total = len(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                total += len(cell.paragraphs)
    for section in doc.sections:
        for hdr in [section.header, section.footer,
                    section.even_page_header, section.even_page_footer,
                    section.first_page_header, section.first_page_footer]:
            try:
                total += len(hdr.paragraphs)
            except Exception:
                pass
    # Text boxes / drawing canvas (w:txbx)
    for txbx in doc.element.iter(qn("w:txbx")):
        total += sum(1 for _ in txbx.iter(qn("w:p")))
    # Content controls (w:sdt)
    for sdt in doc.element.iter(qn("w:sdt")):
        total += sum(1 for _ in sdt.iter(qn("w:p")))
    return max(total, 1)


def format_eta(seconds):
    return f"{seconds:.1f} sec" if seconds < 60 else f"{seconds / 60:.1f} min"


# ======================
# STREAMLIT UI
# ======================
uploaded_file = st.file_uploader("Upload DOCX File", type=["docx"])
target_label  = st.selectbox("Translate To:", list(languages.keys()))

if st.button("Translate Document") and uploaded_file:
    target = languages[target_label]
    doc    = Document(uploaded_file)

    total_blocks = count_all_blocks(doc)
    completed    = 0
    start_time   = time.time()

    st.info(f"🔢 Total items to translate: {total_blocks}")
    progress   = st.progress(0)
    eta_text   = st.empty()
    status_msg = st.empty()
    status_msg.info("Translating... Please wait...")

    def tick():
        nonlocal completed
        completed += 1
        progress.progress(min(completed / total_blocks, 1.0))
        elapsed   = time.time() - start_time
        remaining = total_blocks - completed
        if completed > 0 and remaining > 0:
            eta_text.write(f"⏳ ETA: {format_eta((elapsed / completed) * remaining)}")

    # ── 1. Body paragraphs ────────────────────────────────────────────────
    for para in doc.paragraphs:
        translate_paragraph(para, target)
        tick()

    # ── 2. Body tables ────────────────────────────────────────────────────
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    translate_paragraph(para, target)
                    tick()

    # ── 3. Headers & Footers (all sections, all variants) ─────────────────
    for section in doc.sections:
        for hdr in [section.header, section.footer,
                    section.even_page_header, section.even_page_footer,
                    section.first_page_header, section.first_page_footer]:
            try:
                for para in hdr.paragraphs:
                    translate_paragraph(para, target)
                    tick()
                # Tables inside headers/footers
                for table in hdr.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for para in cell.paragraphs:
                                translate_paragraph(para, target)
                                tick()
            except Exception:
                pass

    # ── 4. Text boxes (w:txbx) — HSE form fields live here ───────────────
    for txbx in doc.element.iter(qn("w:txbx")):
        translate_xml_runs(txbx, target)
        for _ in txbx.iter(qn("w:p")):
            tick()

    # ── 5. Content controls (w:sdt) — structured fields ──────────────────
    for sdt in doc.element.iter(qn("w:sdt")):
        # Skip sdt that are inside a txbx (already handled above)
        translate_xml_runs(sdt, target)
        for _ in sdt.iter(qn("w:p")):
            tick()

    # ── Save & offer download ─────────────────────────────────────────────
    output = BytesIO()
    doc.save(output)
    output.seek(0)

    status_msg.success("🎉 Translation Complete!")
    eta_text.empty()

    safe_name = re.sub(r'[^\w\-]', '_', target_label)
    st.download_button(
        "⬇ Download Translated DOCX",
        data=output,
        file_name=f"translated_{safe_name}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
