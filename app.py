import streamlit as st
import pandas as pd
from pptx import Presentation
import tempfile
import os
import re
from io import BytesIO

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(page_title="PPT Keyword Search Tool", layout="wide")

# -----------------------------
# CUSTOM CSS (A1 THEME)
# -----------------------------
st.markdown("""
<style>

    /* Blue Header */
    .main-header {
        background-color: #0000FF;
        padding: 20px 40px;
        color: white;
        font-size: 32px;
        font-weight: 800;
        border-radius: 0px;
        text-align: left;
        margin-bottom: 25px;
    }

    /* White Card */
    .white-card {
        background-color: white;
        padding: 30px;
        border-radius: 15px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.12);
        margin-bottom: 20px;
    }

    /* Buttons */
    .stButton>button {
        background-color: #0000FF !important;
        color: white !important;
        border-radius: 10px !important;
        padding: 10px 20px;
        font-size: 16px;
        border: none;
    }

    .stButton>button:hover {
        background-color: #0022BB !important;
        color: white !important;
        border: none;
    }

</style>
""", unsafe_allow_html=True)

# -----------------------------
# HEADER BAR
# -----------------------------
st.markdown('<div class="main-header">PPT Keyword Search Tool</div>', unsafe_allow_html=True)

# -----------------------------
# FUNCTIONS
# -----------------------------
def clean_text(text):
    """Remove illegal Excel characters."""
    if not isinstance(text, str):
        return text
    return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', text)

def extract_text_from_ppt(ppt_file):
    prs = Presentation(ppt_file)
    all_text = []

    for slide_num, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                cleaned = clean_text(shape.text)
                all_text.append((slide_num, cleaned))
    return all_text

def create_excel(df):
    """Creates downloadable Excel file with clean text."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")

    return output.getvalue()

# -----------------------------
# WHITE CARD CONTAINER
# -----------------------------
with st.container():
    st.markdown('<div class="white-card">', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Upload one or more PPTX files:",
        type=["pptx"],
        accept_multiple_files=True
    )

    keywords_input = st.text_input(
        "Enter keywords (comma separated):",
        placeholder="e.g., digital, process, automation"
    )

    search_button = st.button("🔍 Search Keywords")

    st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# SEARCH PROCESS
# -----------------------------
if search_button:
    if not uploaded_files:
        st.error("Please upload at least one PPTX file.")
    elif not keywords_input.strip():
        st.error("Please enter at least one keyword.")
    else:
        keywords = [k.strip().lower() for k in keywords_input.split(",")]
        results = []

        for file in uploaded_files:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
                tmp.write(file.read())
                ppt_path = tmp.name

            extracted_data = extract_text_from_ppt(ppt_path)

            for slide_num, text in extracted_data:
                text_lower = text.lower()
                for kw in keywords:
                    if kw in text_lower:
                        highlighted = re.sub(
                            kw, 
                            f"[**{kw.upper()}**]",
                            text, 
                            flags=re.IGNORECASE
                        )
                        results.append([file.name, slide_num, kw, clean_text(highlighted)])

            os.remove(ppt_path)

        # Convert results to DataFrame
        if results:
            df = pd.DataFrame(results, columns=["File Name", "Slide Number", "Keyword", "Text Extract"])

            st.write("### 🔎 Search Results")
            st.dataframe(df, use_container_width=True)

            # -----------------------------
            # EXCEL DOWNLOAD BUTTON
            # -----------------------------
            excel_data = create_excel(df)
            st.download_button(
                label="⬇ Download Results (Excel)",
                data=excel_data,
                file_name="ppt_keyword_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # -----------------------------
            # NEW SEARCH BUTTON (NS2)
            # -----------------------------
            st.markdown("###")
            if st.button("🔄 New Search"):
                st.session_state.clear()
                st.experimental_rerun()

        else:
            st.warning("No matches found for the given keywords.")

# -----------------------------
# FOOTER
# -----------------------------
st.markdown("""
<hr>
<div style="text-align:center; color: gray; padding-top:10px;">
    © 2025 — SKT
</div>
""", unsafe_allow_html=True)
