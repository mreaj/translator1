import streamlit as st
from deep_translator import GoogleTranslator
from docx import Document
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
    # ── Spanish ───────────────────────────────────────────────────────────
    "es": {
        # Blades — most common wrong translations
        "cuchillas": "palas",
        "cuchilla": "pala",
        "aspas": "palas",
        "aspa": "pala",
        "paletas": "palas",
        "paleta": "pala",
        "hoja": "pala",
        "hojas": "palas",
        "veletas": "palas",
        "veleta": "pala",
        # Nacelle
        "cabina": "góndola",
        # Gearbox
        "multiplicador": "multiplicadora",
        "caja de cambios": "multiplicadora",
        "caja de velocidades": "multiplicadora",
        # Hub
        "cubo": "buje",
        "centro": "buje",
        # Turbine / Farm
        "turbina de viento": "aerogenerador",
        "molino de viento": "aerogenerador",
        "generador eólico": "aerogenerador",
        "parque de viento": "parque eólico",
        "granja eólica": "parque eólico",
        "granja de viento": "parque eólico",
        # Controls
        "control de cabeceo": "control de paso",
        "control de inclinación": "control de paso",
        "guiñada": "orientación",
        # Commissioning
        "puesta en servicio": "puesta en marcha",
        "comisionamiento": "puesta en marcha",
        "puesta en funcionamiento": "puesta en marcha",
        # Mechanical completion
        "terminación mecánica": "finalización mecánica",
        "completar mecánico": "finalización mecánica",
        # Safety
        "accidente fatal": "accidente mortal",
        "accidente fatídico": "accidente mortal",
        "cuasi accidente": "cuasi-accidente",
        "casi accidente": "cuasi-accidente",
        # Electrical
        "estación secundaria": "subestación",
        "cable de arreglo": "cable de interconexión",
        "cable de conjunto": "cable de interconexión",
        # Structural
        "pilote único": "monopilote",
        "chaqueta": "estructura de celosía",
        # Crane (Portuguese leak from guindaste)
        "guindaste": "grúa",
    },

    # ── Polish ────────────────────────────────────────────────────────────
    "pl": {
        "skrzydła": "łopaty",
        "skrzydło": "łopata",
        "łopatki": "łopaty",
        "łopatka": "łopata",
        "wiatrak": "turbina wiatrowa",
        "park wiatrowy": "farma wiatrowa",
        "farma wiatrakowa": "farma wiatrowa",
        "skrzynia biegów": "przekładnia",
        "przekładnia zębata": "przekładnia",
        "kabina": "gondola",
        "piasta koła": "piasta",
        "oddanie do eksploatacji": "uruchomienie",
        "kąt skoku": "skok",
        "kąt łopaty": "skok",
        "ster": "odchylenie",
        "prawie wypadek": "zdarzenie potencjalnie wypadkowe",
        "dźwig": "żuraw",
    },

    # ── German ────────────────────────────────────────────────────────────
    "de": {
        "blätter": "Rotorblätter",
        "flügel": "Rotorblätter",
        "rotorflügel": "Rotorblätter",
        "schaufeln": "Rotorblätter",
        "schaufel": "Rotorblatt",
        "windmühle": "Windkraftanlage",
        "windrad": "Windkraftanlage",
        "windturbine": "Windkraftanlage",
        "windfarm": "Windpark",
        "windgenerator": "Windkraftanlage",
        "kabine": "Gondel",
        "radnabe": "Nabe",
        "zahnradgetriebe": "Getriebe",
        "inbetriebsetzung": "Inbetriebnahme",
        "azimutwinkel": "Azimut",
        "gierwinkel": "Azimut",
        "blattanstellwinkel": "Blattwinkel",
        "anstellwinkel": "Blattwinkel",
        "beinahunfall": "Beinaheunfall",
        "einzelpfahl": "Monopfahl",
    },

    # ── French ────────────────────────────────────────────────────────────
    "fr": {
        "lames": "pales",
        "ailes": "pales",
        "aile": "pale",
        "turbine éolienne": "éolienne",
        "moulin à vent": "éolienne",
        "ferme éolienne": "parc éolien",
        "ferme de vent": "parc éolien",
        "boîte de vitesses": "multiplicateur",
        "transmission": "multiplicateur",
        "cabine": "nacelle",
        "centre": "moyeu",
        "commissionning": "mise en service",
        "lacet": "orientation",
        "pieu unique": "monopieu",
        "accident fatal": "accident mortel",
        "presque accident": "quasi-accident",
    },

    # ── Italian ───────────────────────────────────────────────────────────
    "it": {
        "lame": "pale",
        "pale del rotore": "pale",
        "turbina eolica": "aerogeneratore",
        "mulino a vento": "aerogeneratore",
        "fattoria eolica": "parco eolico",
        "scatola del cambio": "moltiplicatore",
        "trasmissione": "moltiplicatore",
        "cabina": "navicella",
        "messa in funzione": "messa in servizio",
        "incidente fatale": "incidente mortale",
        "quasi incidente": "quasi-incidente",
    },

    # ── Dutch ─────────────────────────────────────────────────────────────
    "nl": {
        "bladen": "rotorbladen",
        "blad": "rotorblad",
        "vleugels": "rotorbladen",
        "windmolen": "windturbine",
        "tandwielkast": "versnellingsbak",
        "cabine": "gondel",
        "spoedregeling": "bladhoekregeling",
        "fataal ongeluk": "dodelijk ongeluk",
    },

    # ── Portuguese ────────────────────────────────────────────────────────
    "pt": {
        "lâminas": "pás",
        "turbina eólica": "aerogerador",
        "moinho de vento": "aerogerador",
        "fazenda eólica": "parque eólico",
        "caixa de engrenagens": "multiplicadora",
        "transmissão": "multiplicadora",
        "cabine": "nacele",
        "posta em serviço": "entrada em operação",
        "comissionamento": "entrada em operação",
    },

    # ── Swedish ───────────────────────────────────────────────────────────
    "sv": {
        "blad": "rotorblad",
        "vingar": "rotorblad",
        "vinge": "rotorblad",
        "vindturbin": "vindkraftverk",
        "vindmölla": "vindkraftverk",
        "kugghjulsväxel": "växellåda",
        "kabin": "gondol",
        "driftsättning": "idrifttagning",
        "gir": "girning",
        "stegkontroll": "bladvinkelreglering",
    },

    # ── Danish ────────────────────────────────────────────────────────────
    "da": {
        "blade": "rotorblade",
        "vinger": "rotorblade",
        "vindkraftværk": "vindmølle",
        "vindpark": "vindmøllepark",
        "kabine": "nacelle",
        "gir": "giring",
        "pitchkontrol": "pitchregulering",
    },

    # ── Finnish ───────────────────────────────────────────────────────────
    "fi": {
        "lavat": "roottorin lavat",
        "lapa": "roottorin lapa",
        "tuuliturbiini": "tuulivoimala",
        "tuulimylly": "tuulivoimala",
        "hajautus": "suuntaus",
    },

    # ── Japanese ──────────────────────────────────────────────────────────
    "ja": {
        "刃": "ブレード",
        "羽根": "ブレード",
        "翼": "ブレード",
        "風力タービン": "風力発電機",
        "風車": "風力発電機",
        "風力団地": "ウインドファーム",
        "ギアボックス": "増速機",
        "変速機": "増速機",
        "キャビン": "ナセル",
        "致命的事故": "死亡事故",
        "ニアミス": "ヒヤリハット",
    },

    # ── Korean ────────────────────────────────────────────────────────────
    "ko": {
        "날": "블레이드",
        "날개": "블레이드",
        "풍력 터빈": "풍력 발전기",
        "풍차": "풍력 발전기",
        "풍력 단지": "풍력 발전 단지",
        "기어박스": "증속기",
        "요": "요잉",
        "아찔한 순간": "아차 사고",
    },

    # ── Chinese Simplified ────────────────────────────────────────────────
    "zh-CN": {
        "刀片": "叶片",
        "刀": "叶片",
        "桨叶": "叶片",
        "风力涡轮机": "风力发电机",
        "风车": "风力发电机",
        "风电农场": "风电场",
        "风能农场": "风电场",
        "变速箱": "齿轮箱",
        "试运行": "调试",
        "致命事故": "死亡事故",
        "险兆": "未遂事故",
        "吊车": "起重机",
    },

    # ── Chinese Traditional ───────────────────────────────────────────────
    "zh-TW": {
        "刀片": "葉片",
        "刀": "葉片",
        "風力渦輪機": "風力發電機",
        "風車": "風力發電機",
        "風電農場": "風電場",
        "試運行": "調試",
        "致命事故": "死亡事故",
    },
}

# Patterns that must NEVER be translated — preserved verbatim
PRESERVE_PATTERNS = [
    r'\bIN\.\d{7,12}\b',                          # Event numbers e.g. IN.0000092127
    r'\b\d+(?:\.\d+)?\s*(?:MW|kW|m/s|rpm|Hz|kV|MWh|kWh)\b',  # Engineering units
    r'\bIEC\s*\d+[-\w]*\b',                       # Standards e.g. IEC 61400
    r'\bISO\s*\d+\b',                             # Standards e.g. ISO 9001
    r'\bDNV[-\s]\w+\b',                           # Certification bodies
]


def apply_wind_glossary(text: str, lang: str) -> str:
    """Replace wrong/generic translations with wind-industry-correct terms."""
    glossary = WIND_GLOSSARY.get(lang, {})
    for wrong, correct in glossary.items():
        pattern = re.compile(re.escape(wrong), re.IGNORECASE | re.UNICODE)
        def _replace(m, c=correct):
            # Preserve sentence-start capitalisation
            return c[0].upper() + c[1:] if m.group(0)[0].isupper() else c
        text = pattern.sub(_replace, text)
    return text


def protect_text(text: str):
    """
    Swap preserve-pattern matches with short placeholders so the
    translator cannot alter them. Returns (protected_text, placeholder_map).
    """
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

    # 1. Protect special tokens (event numbers, units, standards)
    protected, placeholders = protect_text(text)

    # 2. Translate via Google
    try:
        translated = GoogleTranslator(source="auto", target=target_lang).translate(protected)
        if not translated:
            translated = protected
    except Exception:
        translated = protected

    # 3. Restore protected tokens
    translated = restore_text(translated, placeholders)

    # 4. Apply wind-industry glossary corrections
    translated = apply_wind_glossary(translated, target_lang)

    return translated


def run_fmt_key(run):
    """Hashable key representing a run's visual formatting."""
    try:
        color = run.font.color.rgb if run.font.color and run.font.color.type else None
    except Exception:
        color = None
    return (
        run.bold,
        run.italic,
        run.underline,
        run.font.size,
        run.font.name,
        color,
    )


def translate_paragraph(para, target_lang: str):
    """
    Translate a paragraph while preserving all run-level formatting.

    Runs with identical formatting are merged before translating to:
      - Reduce the number of API calls
      - Avoid mid-phrase splits that cause unwanted extra line breaks
    """
    if not para.runs:
        return

    # Group consecutive runs that share identical formatting
    groups = []
    for run in para.runs:
        key = run_fmt_key(run)
        if groups and groups[-1][0] == key:
            groups[-1][1].append(run)
        else:
            groups.append((key, [run]))

    for _fmt_key, runs in groups:
        combined = "".join(r.text for r in runs)
        if not combined.strip():
            continue

        translated = safe_translate(combined, target_lang)

        # Write full translated text into the first run; clear the rest
        runs[0].text = translated
        for r in runs[1:]:
            r.text = ""


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
    return f"{seconds / 60:.1f} min"


# ======================
# STREAMLIT UI
# ======================
uploaded_file = st.file_uploader("Upload DOCX File", type=["docx"])
target_label  = st.selectbox("Translate To:", list(languages.keys()))

if st.button("Translate Document") and uploaded_file:
    target = languages[target_label]
    doc    = Document(uploaded_file)

    total_blocks = count_blocks(doc)
    completed    = 0
    start_time   = time.time()

    st.info(f"🔢 Total items to translate: {total_blocks}")
    progress   = st.progress(0)
    eta_text   = st.empty()
    status_msg = st.empty()
    status_msg.info("Translating... Please wait...")

    # ── Body paragraphs ───────────────────────────────────────────────────
    for para in doc.paragraphs:
        translate_paragraph(para, target)
        completed += 1
        progress.progress(completed / total_blocks)
        elapsed = time.time() - start_time
        remaining = total_blocks - completed
        if completed > 0 and remaining > 0:
            eta_text.write(f"⏳ ETA: {format_eta((elapsed / completed) * remaining)}")

    # ── Table cells ───────────────────────────────────────────────────────
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    translate_paragraph(para, target)
                    completed += 1
                    progress.progress(min(completed / total_blocks, 1.0))
                    elapsed = time.time() - start_time
                    remaining = total_blocks - completed
                    if completed > 0 and remaining > 0:
                        eta_text.write(f"⏳ ETA: {format_eta((elapsed / completed) * remaining)}")

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
