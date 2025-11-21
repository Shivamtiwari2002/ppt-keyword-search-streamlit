import streamlit as st
import pandas as pd
from pptx import Presentation
import zipfile
import tempfile
import os
import re
from io import BytesIO
from rapidfuzz import fuzz

# ------------------- PAGE CONFIG -------------------
st.set_page_config(page_title="PPT Keyword Search Tool", layout="wide")

# ------------------- MODERN UI CSS -------------------
st.markdown("""
<style>

body {
    background-color: #F5F7FF;
}

.header-box {
    background: linear-gradient(135deg, #0047FF, #3F8CFF);
    padding: 25px;
    border-radius: 12px;
    text-align: center;
    color: white;
    font-size: 34px;
    font-weight: 800;
    margin-bottom: 30px;
    box-shadow: 0px 4px 14px rgba(0,0,0,0.12);
}

.section-card {
    background: white;
    padding: 22px;
    border-radius: 16px;
    box-shadow: 0px 4px 12px rgba(0,0,0,0.08);
    margin-bottom: 25px;
}

.stFileUploader>div>div {
    border: 2px dashed #0047FF !important;
    background: #EFF3FF !important;
    border-radius: 14px !important;
}

.upload-label {
    font-size: 18px;
    font-weight: 700;
    color: #0047FF;
}

.keyword-box input {
    border: 2px solid #0047FF !important;
    border-radius: 12px !important;
    padding: 12px !important;
    background: #EFF3FF !important;
    font-size: 17px !important;
    color: #0033CC !important;
    font-weight: 600 !important;
}

.stButton>button {
    background: #0047FF !important;
    color: white !important;
    border-radius: 10px !important;
    padding: 10px 26px !important;
    font-size: 17px !important;
    border: none !important;
    font-weight: 600 !important;
    box-shadow: 0px 3px 10px rgba(0,0,0,0.15);
}

.table-container {
    width: 100%;
    border-collapse: collapse;
    margin-top: 20px;
}

.table-container th {
    background: #0047FF;
    color: white;
    padding: 10px;
    text-align: left;
}

.table-container tr:hover {
    background: #E8EEFF;
}

.table-container td {
    padding: 10px;
    border-bottom: 1px solid #DDD;
    font-size: 15px;
}

.footer {
    text-align:center;
    margin-top: 40px;
    padding: 15px;
    font-size: 14px;
    font-weight: 600;
    color: #0047FF;
}

</style>
""", unsafe_allow_html=True)

# ------------------- HEADER -------------------
st.markdown("<div class='header-box'>PPT Keyword Search Tool</div>", unsafe_allow_html=True)

# ------------------- LOAD KEYWORDS FROM EXCEL -------------------
@st.cache_data
def load_keywords_from_excel(excel_bytes, col_name="Keyword"):
    df = pd.read_excel(excel_bytes)
    col = df.columns[0] if col_name not in df.columns else col_name
    keywords = df[col].dropna().astype(str).str.strip().unique().tolist()
    return sorted([k for k in keywords if k])

# Local file inside repository
LOCAL_EXCEL_PATH = "data/keywords.xlsx"

# ------------------- EXCEL LOAD OPTIONS -------------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("### Load Dropdown Values from Excel")

uploaded_excel = st.file_uploader("Upload Excel for Dropdown (optional)", type=["xlsx", "xls"])
use_local = st.checkbox(f"Use Local Excel: {LOCAL_EXCEL_PATH}")

dropdown_keywords = []

try:
    if uploaded_excel:
        dropdown_keywords = load_keywords_from_excel(uploaded_excel)
    elif use_local and os.path.exists(LOCAL_EXCEL_PATH):
        dropdown_keywords = load_keywords_from_excel(LOCAL_EXCEL_PATH)
except Exception as e:
    st.error(f"Error loading Excel: {e}")

if dropdown_keywords:
    keyword_options = ["-- Select --"] + dropdown_keywords
else:
    keyword_options = ["-- No keywords loaded --"]

st.markdown("</div>", unsafe_allow_html=True)

# ------------------- FILE UPLOADER -------------------
with st.container():
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.markdown("<div class='upload-label'>Upload PPTX / ZIP Files</div>", unsafe_allow_html=True)
    uploaded_files = st.file_uploader("", type=["pptx", "zip"], accept_multiple_files=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ------------------- KEYWORD INPUT -------------------
with st.container():
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.markdown("### Select or Enter Keyword")

    selected_keyword = st.selectbox("Choose from dropdown", keyword_options)

    st.markdown("<div class='keyword-box'>", unsafe_allow_html=True)
    custom_keyword = st.text_input("", placeholder="Or type your own keyword…")
    st.markdown("</div>", unsafe_allow_html=True)

    # Final keyword logic
    if custom_keyword.strip():
        keyword = custom_keyword.strip()
    elif selected_keyword not in ["-- Select --", "-- No keywords loaded --"]:
        keyword = selected_keyword
    else:
        keyword = ""

    search_btn = st.button("🔍 Search")
    st.markdown("</div>", unsafe_allow_html=True)

# ------------------- FUNCTIONS -------------------
def clean_text(text):
    if text is None:
        return ""
    return re.sub(r"[\000-\010\013\014\016-\037]", "", str(text))

def extract_text_from_pptx(file_path, keyword):
    prs = Presentation(file_path)
    matches = []

    for slide_num, slide in enumerate(prs.slides, start=1):

        title_text = ""
        try:
            if slide.shapes.title and slide.shapes.title.text:
                title_text = slide.shapes.title.text.strip()
        except:
            pass

        slide_text = ""
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text:
                slide_text += shape.text + " "

        similarity = fuzz.partial_ratio(keyword.lower(), slide_text.lower())

        if similarity > 80:
            matches.append({
                "File": os.path.basename(file_path),
                "Slide Number": slide_num,
                "Matched Text": title_text,
                "Similarity": similarity
            })

    return matches

def extract_zip(file):
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(file, 'r') as z:
        z.extractall(temp_dir)

    pptx_files = []
    for root, _, files in os.walk(temp_dir):
        for f in files:
            if f.endswith(".pptx"):
                pptx_files.append(os.path.join(root, f))
    return pptx_files

# ------------------- SEARCH LOGIC -------------------
if search_btn:
    if not uploaded_files:
        st.error("⚠ Please upload PPTX or ZIP files.")
        st.stop()

    if not keyword:
        st.error("⚠ Please select or enter a keyword.")
        st.stop()

    results = []
    with st.spinner("Searching…"):
        for file in uploaded_files:
            if file.name.endswith(".pptx"):
                temp_path = os.path.join(tempfile.gettempdir(), file.name)
                with open(temp_path, "wb") as f:
                    f.write(file.read())
                results.extend(extract_text_from_pptx(temp_path, keyword))

            elif file.name.endswith(".zip"):
                for ppt_path in extract_zip(file):
                    results.extend(extract_text_from_pptx(ppt_path, keyword))

    df = pd.DataFrame(results)

    st.markdown("<div class='section-card'>", unsafe_allow_html=True)

    if df.empty:
        st.warning("No matches found.")
    else:
        df = df.applymap(clean_text)

        st.markdown("### Results")

        def render_table(df):
            html = "<table class='table-container'>"
            html += "<tr>" + "".join(f"<th>{c}</th>" for c in df.columns) + "</tr>"
            for _, row in df.iterrows():
                html += "<tr>" + "".join(f"<td>{x}</td>" for x in row) + "</tr>"
            html += "</table>"
            return html

        st.markdown(render_table(df), unsafe_allow_html=True)

        # Download report
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Results")

        st.download_button(
            "⬇ Download Results (Excel)",
            output.getvalue(),
            file_name="ppt_search_results.xlsx"
        )

    st.markdown("</div>", unsafe_allow_html=True)

# ------------------- FOOTER -------------------
st.markdown("<div class='footer'>Made by SKT</div>", unsafe_allow_html=True)
