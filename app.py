import streamlit as st
import pandas as pd
from pptx import Presentation
import zipfile
import tempfile
import os
import re
from io import BytesIO

# -----------------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------------
st.set_page_config(
    page_title="PPT Keyword Search Tool",
    layout="wide"
)

# -----------------------------------------------------------
# CUSTOM CSS (Blue/White Premium Theme) — updated uploader overlay
# -----------------------------------------------------------
st.markdown("""
<style>

    /* Blue top bar */
    .blue-header {
        background-color: #0000FF;
        padding: 22px;
        text-align: center;
        color: white;
        font-size: 32px;
        font-weight: 700;
        letter-spacing: 1px;
        margin-bottom: 20px;
    }

    /* App background */
    .main {
        background-color: #F4F6FF !important;
    }

    /* Container for uploader to allow overlaying actual input */
    .upload-wrapper {
        position: relative;
        display: block;
        width: 100%;
        max-width: 900px;
        margin-bottom: 12px;
    }

    /* Custom upload box (visual) */
    .custom-upload {
        border: 3px dashed #0000FF;
        background: #E6ECFF;
        padding: 32px;
        border-radius: 12px;
        text-align: center;
        font-size: 18px;
        font-weight: 700;
        color: #0000FF;
        cursor: pointer;
        margin-bottom: 20px;
    }

    /* Ensure the real uploader sits on top of the blue box but remains invisible */
    /* This targets Streamlit file uploader wrapper and places it absolute & full-size */
    [data-testid="stFileUploader"] {
        position: absolute !important;
        left: 0;
        top: 0;
        width: 100% !important;
        height: 100% !important;
        opacity: 0 !important;
        z-index: 10;
        padding: 0 !important;
        margin: 0 !important;
        overflow: hidden !important;
    }

    /* When input is focused, show subtle outline on blue box */
    [data-testid="stFileUploader"]:focus + .custom-upload,
    .custom-upload:focus {
        box-shadow: 0 0 0 4px rgba(0,0,255,0.12);
    }

    /* Keyword box */
    .keyword-box input {
        border: 2px solid #0000FF !important;
        background: #E6ECFF !important;
        color: #0000FF !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        padding: 10px !important;
        font-size: 16px !important;
    }

    /* Buttons */
    .stButton>button {
        background-color: #0000FF !important;
        color: white !important;
        padding: 10px 22px !important;
        font-size: 16px !important;
        border-radius: 10px !important;
        border: none !important;
        font-weight: 600 !important;
        cursor: pointer;
    }

    .stButton>button:hover {
        opacity: 0.95;
    }

    /* Result Card */
    .result-card {
        background-color: white;
        padding: 20px;
        border-radius: 12px;
        border-left: 6px solid #0000FF;
        margin-top: 20px;
        box-shadow: 0px 3px 12px rgba(0,0,0,0.08);
    }

    /* Blue Table Header */
    thead tr th {
        background-color: #0000FF !important;
        color: white !important;
        font-weight: bold !important;
        padding: 10px !important;
    }

    /* Footer */
    .footer {
        text-align:center;
        margin-top: 40px;
        padding: 15px;
        font-size: 14px;
        font-weight: 600;
        color: #0000FF;
    }

</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------
# HEADER
# -----------------------------------------------------------
st.markdown("<div class='blue-header'>PPT Keyword Search Tool</div>", unsafe_allow_html=True)


# -----------------------------------------------------------
# UPLOAD: single blue box with invisible overlaying uploader
# -----------------------------------------------------------
st.markdown("### Upload PPTX or ZIP files")

# wrapper allows absolute positioning of the real uploader input
st.markdown("<div class='upload-wrapper'>", unsafe_allow_html=True)

# Real (invisible) Streamlit uploader placed first so it overlays the next element
uploaded_files = st.file_uploader(
    "hidden_uploader",
    type=["pptx", "zip"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

# Visual blue box shown to the user
st.markdown("""
<div class='custom-upload' id='upload-area'>
    Click or drag & drop PPTX / ZIP files here<br>
    <span style='font-size:14px;font-weight:400;'>Limit 200MB per file • PPTX, ZIP</span>
</div>
""", unsafe_allow_html=True)

# close wrapper
st.markdown("</div>", unsafe_allow_html=True)

# add a small script so clicking the blue box also focuses the hidden uploader (not strictly necessary,
# but improves keyboard focus behaviour)
st.markdown("""
<script>
const area = document.getElementById('upload-area');
if (area) {
  area.onclick = function() {
    const uploader = document.querySelector('[data-testid="stFileUploader"] input');
    if (uploader) uploader.click();
  };
}
</script>
""", unsafe_allow_html=True)


# -----------------------------------------------------------
# KEYWORD INPUT
# -----------------------------------------------------------
st.markdown("### Enter Search Keyword")

st.markdown("<div class='keyword-box'>", unsafe_allow_html=True)
keyword = st.text_input("", placeholder="Enter keyword to search...")
st.markdown("</div>", unsafe_allow_html=True)

search_btn = st.button("🔍 Search")


# -----------------------------------------------------------
# CLEAN TEXT (Fix IllegalCharacterError)
# -----------------------------------------------------------
def clean_text(text):
    if text is None:
        return ""
    return re.sub(r"[\000-\010\013\014\016-\037]", "", str(text))


# -----------------------------------------------------------
# PPTX PROCESSING
# -----------------------------------------------------------
def extract_text_from_pptx(file_path):
    prs = Presentation(file_path)
    matches = []

    for slide_num, slide in enumerate(prs.slides, start=1):
        text = ""
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"

        if keyword.lower() in text.lower():
            matches.append({
                "File": os.path.basename(file_path),
                "Slide Number": slide_num,
                "Matched Text": clean_text(text.strip())
            })

    return matches


# -----------------------------------------------------------
# ZIP HANDLING
# -----------------------------------------------------------
def process_zip(file):
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(file, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)

    pptx_files = []
    for root, _, files in os.walk(temp_dir):
        for f in files:
            if f.endswith(".pptx"):
                pptx_files.append(os.path.join(root, f))

    return pptx_files


# -----------------------------------------------------------
# SEARCH ACTION
# -----------------------------------------------------------
if search_btn:

    if not uploaded_files:
        st.error("⚠ Please upload at least one PPTX or ZIP file.")
        st.stop()

    if not keyword.strip():
        st.error("⚠ Please enter a keyword.")
        st.stop()

    results = []

    with st.spinner("Searching slides... Please wait..."):

        for file in uploaded_files:

            if file.name.endswith(".pptx"):
                temp_path = os.path.join(tempfile.gettempdir(), file.name)
                with open(temp_path, "wb") as f:
                    f.write(file.read())
                results.extend(extract_text_from_pptx(temp_path))

            elif file.name.endswith(".zip"):
                pptx_files = process_zip(file)
                for p in pptx_files:
                    results.extend(extract_text_from_pptx(p))

    df = pd.DataFrame(results)

    if df.empty:
        st.warning("No matches found.")
    else:
        df = df.applymap(clean_text)

        st.markdown("<div class='result-card'>", unsafe_allow_html=True)
        st.subheader("Search Results")
        st.dataframe(df, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        # Excel download
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Results")
        excel_data = output.getvalue()

        st.download_button(
            label="⬇ Download Results (Excel)",
            data=excel_data,
            file_name="ppt_search_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # New Search Button
        if st.button("🔄 New Search"):
            st.session_state.clear()
            st.experimental_rerun()


# -----------------------------------------------------------
# FOOTER
# -----------------------------------------------------------
st.markdown("<div class='footer'>Made by SKT</div>", unsafe_allow_html=True)
