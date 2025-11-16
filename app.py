import streamlit as st
from pptx import Presentation
import zipfile
import io
import os

# -------------------------------------------------------------
# 🔵 CUSTOM CSS STYLING (Blue-White Premium UI)
# -------------------------------------------------------------
st.markdown("""
    <style>

    /* Page background */
    .stApp {
        background: #EEF3FF;
    }

    /* Top header bar */
    .top-header {
        width: 100%;
        padding: 18px 0;
        background-color: #0000FF;
        text-align: center;
        color: white;
        font-size: 28px;
        font-weight: 700;
        border-radius: 0 0 10px 10px;
        margin-bottom: 25px;
    }

    /* White card container */
    .card {
        background-color: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0px 2px 12px rgba(0,0,0,0.10);
        border: 2px solid #0000FF10;
        margin-bottom: 25px;
    }

    /* Buttons styling */
    .stButton>button {
        background-color: #0000FF;
        color: white;
        border-radius: 8px;
        font-size: 16px;
        padding: 8px 20px;
        border: none;
    }

    .stButton>button:hover {
        background-color: #0000CC;
        color: white;
    }

    /* Search result box */
    .result-box {
        border: 2px solid #0000FF;
        padding: 12px;
        border-radius: 8px;
        background: white;
        font-size: 15px;
        margin-bottom: 10px;
    }

    /* Footer */
    .footer {
        text-align: center;
        margin-top: 40px;
        color: #0000FF;
        padding: 15px 0;
        font-weight: 600;
    }

    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# 🔵 TOP HEADER BAR
# -------------------------------------------------------------
st.markdown("<div class='top-header'>PPT Keyword Search Tool</div>", unsafe_allow_html=True)

# -------------------------------------------------------------
# Functions
# -------------------------------------------------------------
def search_keyword_in_pptx(pptx_file, keyword):
    prs = Presentation(pptx_file)
    results = []
    
    for slide_number, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                if keyword.lower() in shape.text.lower():
                    results.append(f"Slide {slide_number}: {shape.text.strip()}")
    return results


def search_keyword_in_zip(zip_file, keyword):
    results = []
    zip_bytes = io.BytesIO(zip_file.read())

    with zipfile.ZipFile(zip_bytes, 'r') as z:
        for filename in z.namelist():
            if filename.endswith(".pptx"):
                pptx_bytes = z.read(filename)
                pptx_stream = io.BytesIO(pptx_bytes)

                slide_results = search_keyword_in_pptx(pptx_stream, keyword)
                if slide_results:
                    results.append((filename, slide_results))
    return results


# -------------------------------------------------------------
# 🔵 Main UI Card
# -------------------------------------------------------------
with st.container():
    st.markdown("<div class='card'>", unsafe_allow_html=True)

    st.subheader("Upload PPT or ZIP Files")

    uploaded_files = st.file_uploader(
        "Choose PPTX or ZIP files:",
        type=["pptx", "zip"],
        accept_multiple_files=True
    )

    keyword = st.text_input("Enter keyword to search:")

    # Buttons
    search_btn = st.button("🔍 Search")
    new_search_btn = st.button("🔄 New Search")
    clear_btn = st.button("🗑 Clear All")

    if new_search_btn:
        st.experimental_rerun()

    if clear_btn:
        st.experimental_rerun()

    if search_btn:
        if not uploaded_files:
            st.warning("Please upload at least one PPTX or ZIP file.")
        elif not keyword.strip():
            st.warning("Please enter a keyword.")
        else:
            st.markdown("---")
            st.subheader("Search Results")

            found_any = False

            for file in uploaded_files:
                filename = file.name

                if filename.endswith(".pptx"):
                    results = search_keyword_in_pptx(file, keyword)
                    if results:
                        found_any = True
                        st.markdown(f"<div class='result-box'><b>{filename}</b></div>", unsafe_allow_html=True)
                        for r in results:
                            st.markdown(f"<div class='result-box'>{r}</div>", unsafe_allow_html=True)

                elif filename.endswith(".zip"):
                    zip_results = search_keyword_in_zip(file, keyword)
                    if zip_results:
                        found_any = True
                        for ppt_name, slides in zip_results:
                            st.markdown(f"<div class='result-box'><b>{ppt_name}</b></div>", unsafe_allow_html=True)
                            for item in slides:
                                st.markdown(f"<div class='result-box'>{item}</div>", unsafe_allow_html=True)

            if not found_any:
                st.info("No results found for the keyword.")

    st.markdown("</div>", unsafe_allow_html=True)

# -------------------------------------------------------------
# 🔵 FOOTER
# -------------------------------------------------------------
st.markdown("<div class='footer'>Made by SKT</div>", unsafe_allow_html=True)
