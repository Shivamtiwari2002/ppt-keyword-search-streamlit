import streamlit as st
import pandas as pd
from pptx import Presentation
import zipfile
import tempfile
import os
import re
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

# -----------------------------------------------------------
# PAGE CONFIG
# -----------------------------------------------------------
st.set_page_config(
    page_title="PPT Keyword Search Tool",
    layout="wide"
)

# -----------------------------------------------------------
# CUSTOM CSS
# -----------------------------------------------------------
st.markdown("""
<style>
    .blue-header {
        background-color: #0000FF;
        padding: 22px;
        text-align: center;
        color: white;
        font-size: 32px;
        font-weight: 700;
        letter-spacing: 1px;
        margin-bottom: 24px;
        border-radius: 0px;
    }

    .main {
        background-color: #F4F6FF !important;
    }

    .stFileUploader>div>div>div>input {
        border: 3px dashed #0000FF !important;
        border-radius: 14px !important;
        background-color: #E6ECFF !important;
        padding: 30px !important;
        cursor: pointer !important;
        color: #0000FF !important;
    }
    .stFileUploader>div>label {
        font-size: 20px !important;
        font-weight: 700 !important;
        color: #0000FF !important;
    }

    .keyword-box input {
        border: 2px solid #0000FF !important;
        background: #E6ECFF !important;
        color: #0000FF !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        padding: 10px !important;
        font-size: 16px !important;
        width: 100% !important;
    }

    .stButton>button {
        background-color: #0000FF !important;
        color: white !important;
        padding: 10px 22px !important;
        font-size: 16px !important;
        border-radius: 10px !important;
        border: none !important;
        font-weight: 600 !important;
        cursor: pointer;
        margin-right: 8px !important;
    }

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
# FILE UPLOADER
# -----------------------------------------------------------
uploaded_files = st.file_uploader(
    "Click or Drag & Drop PPTX / ZIP files here (Limit 200MB per file)", 
    type=["pptx", "zip"], 
    accept_multiple_files=True
)
st.markdown("<br>", unsafe_allow_html=True)

# -----------------------------------------------------------
# KEYWORD INPUT
# -----------------------------------------------------------
st.markdown("### Enter Search Keyword")
st.markdown("<div class='keyword-box'>", unsafe_allow_html=True)
keyword = st.text_input("", placeholder="Enter keyword to search...")
st.markdown("</div>", unsafe_allow_html=True)

# -----------------------------------------------------------
# SEARCH BUTTON
# -----------------------------------------------------------
search_btn = st.button("🔍 Search")

# -----------------------------------------------------------
# CLEAN TEXT
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

    df = pd.DataFrame(results)

    if df.empty:
        st.warning("No matches found.")
    else:
        df = df.applymap(clean_text)
        st.markdown("<div class='result-card'>", unsafe_allow_html=True)
        st.subheader("Search Results")

        # -----------------------
        # AgGrid interactive table
        # -----------------------
        gb = GridOptionsBuilder.from_dataframe(df)
        gb.configure_default_column(editable=False, filter=True, sortable=True)
        gb.configure_selection('single')
        gb.configure_grid_options(domLayout='normal')
        grid_options = gb.build()

        AgGrid(
            df,
            gridOptions=grid_options,
            height=400,
            width='100%',
            update_mode=GridUpdateMode.NO_UPDATE,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            theme='blue'
        )

        st.markdown("</div>", unsafe_allow_html=True)

        # Export Excel
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

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 New Search"):
            st.session_state.clear()
            st.experimental_rerun()

# -----------------------------------------------------------
# FOOTER
# -----------------------------------------------------------
st.markdown("<div class='footer'>Made by SKT</div>", unsafe_allow_html=True)
