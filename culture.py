import streamlit as st
import pdfplumber
import pandas as pd
import re
import os
import tempfile

EXCEL_FILE = "culture_database.xlsx"

# =========================
# ANTIBIOTICS LIST
# =========================
ANTIBIOTICS = [
    "Penicillin","Ampicillin","Clindamycin","Amoxicillin",
    "Erythromycin","Cephradine","Ceftriaxone","Vancomycin",
    "Piperacillin","Levofloxacin","Azithromycin","Amikacin",
    "Netilmycin","Ceftazidime","Imipenem","Tigecycline",
    "Moxifloxacin","Aztreonam","Clarithromycin",
    "Chloramphenicol","Cefotaxime","Trimethoprim",
    "Meropenem","Co-trimoxazole","Gentamycin",
    "Cefuroxime","Doxycycline","Nitrofurantoin",
    "Amoxiclav","Ciprofloxacin","Pefloxacin",
    "Cephalexin","Metronidazole","Cefixime",
    "Linezolid","Cefaclor","Oxacillin",
    "Ofloxacin","Cefoxitine","Cephoxitine",
    "Cefepime","Sulphamethoxazole","Tetracycline"
]

ALL_COLUMNS = ["ID","Name","Age","Sex","Specimen","Organism"] + ANTIBIOTICS

# Known constants in your medical reports
KNOWN_SPECIMENS = ["Sputum"]
KNOWN_ORGANISMS = ["Staphylococcus", "Pseudomonas"]

# =========================
# SESSION INIT
# =========================
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(columns=ALL_COLUMNS)

# =========================
# RESET
# =========================
def reset_all():
    st.session_state.df = pd.DataFrame(columns=ALL_COLUMNS)
    if os.path.exists(EXCEL_FILE):
        os.remove(EXCEL_FILE)

# =========================
# CLEAN TEXT
# =========================
def clean(text):
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text

# =========================
# SMART FIELDS PARSER (FIXED)
# =========================
def extract_specimen_robust(text):
    # Fallback scan for known text blocks
    for spec in KNOWN_SPECIMENS:
        if spec.lower() in text.lower():
            return spec
    # Regex fallback if a new specimen appears
    match = re.search(r"Specimen\s*[:\-]?\s*([A-Za-z]+)", text, re.I)
    return match.group(1).strip().capitalize() if match else ""

def extract_organism_robust(text):
    # Fallback scan for known micro-organisms 
    for org in KNOWN_ORGANISMS:
        if org.lower() in text.lower():
            return org
    # Regex fallback if text pattern fits cleanly
    match = re.search(r"Growth\s*:\s*([A-Za-z]+)", text, re.I)
    return match.group(1).strip().capitalize() if match else ""

# =========================
# PDF EXTRACTION
# =========================
def extract_pdf(path):
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += " " + t

    # Capture uncleaned text segment for drug boundaries, clean up basic info
    cleaned_text = clean(text)

    data = {col: "" for col in ALL_COLUMNS}

    # -------------------------
    # BASIC INFO
    # -------------------------
    id_match = re.search(r'ID\s*No\.?\s*[:\-]?\s*(\d+)', cleaned_text, re.I)
    name_match = re.search(r'Name\s+Age\s+([A-Za-z\s\.\,\_]+?)\s+\d+Y', cleaned_text, re.I)
    age_match = re.search(r'(\d+)\s*Y', cleaned_text, re.I)
    sex_match = re.search(r'Sex\s*(Female|Male)', cleaned_text, re.I)

    data["ID"] = id_match.group(1) if id_match else ""
    
    if name_match:
        data["Name"] = name_match.group(1).replace("MST.", "").replace("MRS.", "").replace("MR.", "").strip()
    else:
        # Fallback Name parser
        n_match = re.search(r'(MST\.|MRS\.|MR\.)\s*[A-Za-z\s]+', cleaned_text)
        data["Name"] = n_match.group(0).strip() if n_match else ""

    data["Age"] = age_match.group(1) if age_match else ""
    data["Sex"] = sex_match.group(1) if sex_match else ""

    # FIXED DETERMINISTIC SPECIMEN/ORGANISM PARSERS
    data["Specimen"] = extract_specimen_robust(cleaned_text)
    data["Organism"] = extract_organism_robust(cleaned_text)

    # -------------------------
    # ANTIBIOTICS (DETERMINISTIC BOUNDARY)
    # -------------------------
    for drug in ANTIBIOTICS:
        # This handles cases where PDF tables read as "Oxacillin S" or "Cefixime S" or "Ceftriaxone S"
        pattern = rf"{re.escape(drug)}\s*\"?\,\"?\s*([SIR])\b"
        m = re.search(pattern, cleaned_text, re.I)
        if m:
            data[drug] = m.group(1).upper()
        else:
            # Secondary check for stacked cell alignment variations
            pattern_stacked = rf"{re.escape(drug)}\s+([SIR])\b"
            m2 = re.search(pattern_stacked, cleaned_text, re.I)
            data[drug] = m2.group(1).upper() if m2 else ""

    return data

# =========================
# SAVE (SAFE)
# =========================
def save(row):
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
        except:
            df = pd.DataFrame(columns=ALL_COLUMNS)
    else:
        df = pd.DataFrame(columns=ALL_COLUMNS)

    for col in ALL_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    if row.get("ID"):
        df = df[df["ID"].astype(str) != str(row["ID"])]

    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)
    return df

# =========================
# UI
# =========================
st.title("🧪 Culture Report Extractor")

col1, col2 = st.columns(2)

with col1:
    if st.button("🔄 RESET DATABASE"):
        reset_all()
        st.success("Database cleared")

with col2:
    st.write("Upload PDF below")

uploaded = st.file_uploader("Upload PDF", type=["pdf"])

# =========================
# PROCESS
# =========================
if uploaded:
    if st.button("🚀 EXTRACT REPORT"):
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        tmp.write(uploaded.getvalue())
        tmp.close()

        try:
            result = extract_pdf(tmp.name)
            st.session_state.df = save(result)
            st.success("Extraction successful!")
        finally:
            os.remove(tmp.name)

# =========================
# DISPLAY
# =========================
st.subheader("📊 Database Preview")
st.dataframe(st.session_state.df)

# =========================
# DOWNLOAD EXCEL
# =========================
if not st.session_state.df.empty:
    file_name = "culture_export.xlsx"
    st.session_state.df.to_excel(file_name, index=False)

    with open(file_name, "rb") as f:
        st.download_button(
            "📥 Download Excel File",
            f,
            file_name="culture_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
