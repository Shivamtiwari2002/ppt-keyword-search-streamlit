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
# CUSTOM CSS (Blue/White Premium Theme)
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
        margin-bottom: 18px;
    }

    /* App background */
    .main {
        background-color: #F4F6FF !important;
    }

    /* File uploader UI */
    .custom-upload {
        border: 3px dashed #0000FF;
        background: #E6ECFF;
        padding: 30px;
        border-radius: 12px;
        text-align: center;
        font-size: 18px;
        font-weight: 600;
        color: #0000FF;
        cursor: pointer;
        margin-bottom: 20px;
    }

    /* Hide default uploader label */
    .stFileUploader label {
        font-size: 0px !important;
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

    /* Primary Buttons */
    .stButton>button {
        background-color: #0000FF !important;
        color: white !important;
        padding: 10px 22px !important;
        font-size: 16px !important;
        border-radius: 10px !important;
        border: none !important;
        font-weight: 600 !important;
    }

    /* White Card for Results */
    .result-card {
        background-color: white;
        padding: 20px;
        border-radius: 12px;
        border-left: 6px solid #0000FF;
        margin-top: 20px;
        box-shadow: 0px 3px 12px rgba(0,0,0,0.08);
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

    /* ------------------------------------------------------ */
    /*  NEW ADDITION #1 : Blue Table Header                   */
    /* ------------------------------------------------------ */
    thead tr th {
        background-color: #0000FF !important;
        color: white !important;
        font-weight: bold !important;
        text-align: left !important;
        padding: 10px !important;
        border-bottom: 2px solid #ffffff !important;
    }

    /* Dataframe border subtle */
    .stDataFrame tbody td {
        border-bottom: 1px solid #D9E1FF !important;
    }

    /* ------------------------------------------------------ */
    /*  NEW ADDITION #2 : Blue Download Button                */
    /* ------------------------------------------------------ */
    .download-btn button {
        background-color: #0000FF !important;
        color: white !important;
        padding: 10px 20px !important;
        border-radius: 10px !important;
        border: none !important;
        font-weight: 600 !important;
        cursor: pointer;
        width: 240px;
    }

</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------
# HEADER
# -----------------------------------------------------------
st.markdown("<div class='blue-header'>PPT Keyword Search Tool</div>", unsafe_allow_html=True)


# -----------------------------------------------------------
# MAIN LAYOUT
# -----------------------------------------------------------
st.markdown("### Upload PPTX or ZIP files")

st.markdown(
    "<div class='custom-upload'>Drag and drop files here<br><span style='font-size:14px;font-weight:400;'>Limit 200MB per file • PPTX, ZIP</span></div>",
    unsafe_allow_html=True
)

uploaded_files = st.file_uploader(
    "",
    type=["pptx", "zip"],
    accept_multiple_files=True
)

st.markdown("<br>", unsafe_allow_html=True)

st.markdown("### Enter Search Keyword")

st.markdown("<div class='keyword-box'>", unsafe_allow_html=True)
keyword = st.text_input("", placeholder="Enter keyword to search...")
st.markdown("</div>", unsafe_allow_html=True)

search_btn = st.button("🔍 Search")


# -----------------------------------------------------------
# CLEAN TEXT BEFORE WRITING TO EXCEL
# -----------------------------------------------------------
def clean_text(text):
    if text is None:
        return ""
    return re.sub(r"[\000-\010\013\014\016-\037]", "", str(text))


# -----------------------------------------------------------
# PROCESS THE PPTX FILES
# -----------------------------------------------------------
def extract_text_from_pptx(file_path):
    prs = Presentation(file_path)
    matches = []

    for slide_num, slide in enumerate(prs.slides, start=1):
        slide_text = ""
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                slide_text += shape.text + "\n"

        if keyword.lower() in slide_text.lower():
            matches.append({
                "File": os.path.basename(file_path),
                "Slide Number": slide_num,
                "Matched Text": slide_text.strip()
            })

    return matches


# -----------------------------------------------------------
# ZIP HANDLING
# -----------------------------------------------------------
def process_zip(file):
    extracted_temp = tempfile.mkdtemp()
    with zipfile.ZipFile(file, 'r') as zip_ref:
        zip_ref.extractall(extracted_temp)

    pptx_files = []
    for root, _, files in os.walk(extracted_temp):
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

    # Convert to DF
    df = pd.DataFrame(results)

    if df.empty:
        st.warning("No matches found.")
    else:

        # Clean illegal chars
        df = df.applymap(clean_text)

        # Results Card
        st.markdown("<div class='result-card'>", unsafe_allow_html=True)
        st.subheader("Search Results")

        st.dataframe(df, use_container_width=True)

        st.markdown("</div>", unsafe_allow_html=True)

        # -------- BLUE Download Button --------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Results")
        excel_data = output.getvalue()

        st.markdown("<div class='download-btn'>", unsafe_allow_html=True)
        st.download_button(
            label="⬇ Download Results (Excel)",
            data=excel_data,
            file_name="ppt_search_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.markdown("</div>", unsafe_allow_html=True)

        # New Search Button
        st.button("🔄 New Search")


# -----------------------------------------------------------
# FOOTER
# -----------------------------------------------------------
st.markdown("<div class='footer'>Made by SKT</div>", unsafe_allow_html=True)
