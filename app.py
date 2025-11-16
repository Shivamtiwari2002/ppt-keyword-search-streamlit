import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pandas as pd
import io
import re
from datetime import datetime

# =========================================================
#                   🔹 GLOBAL UI STYLING 🔹
# =========================================================

BLUE = "#0000FF"
WHITE = "#FFFFFF"

st.set_page_config(
    page_title="PPT Keyword Search Tool",
    layout="wide",
)

st.markdown(
    f"""
    <style>

    /* Center the main container */
    .main-container {{
        background-color: {WHITE};
        padding: 20px;
        border-radius: 15px;
        border: 2px solid {BLUE};
        box-shadow: 0 0 10px rgba(0,0,0,0.15);
    }}

    /* Headings */
    .main-title {{
        color: {BLUE};
        text-align: center;
        font-size: 32px;
        font-weight: 700;
        margin-bottom: 10px;
    }}

    .section-title {{
        color: {BLUE};
        font-size: 20px;
        font-weight: bold;
        margin-top: 15px;
        margin-bottom: 5px;
    }}

    /* Button Styling */
    .stButton>button {{
        background-color: {BLUE} !important;
        color: {WHITE} !important;
        border-radius: 8px !important;
        padding: 8px 18px !important;
        font-size: 16px !important;
        border: none;
    }}

    .stButton>button:hover {{
        opacity: 0.9;
    }}

    /* Output box styling */
    .output-box {{
        background-color: {WHITE};
        border: 2px solid {BLUE};
        padding: 20px;
        border-radius: 12px;
        margin-top: 20px;
        box-shadow: 0 0 8px rgba(0,0,0,0.1);
    }}

    </style>
    """, unsafe_allow_html=True
)


# =========================================================
#            🔹 CLEAN TEXT (Fix Excel Illegal Characters)
# =========================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    value = str(value)
    value = re.sub(r"[\x00-\x1F\x7F]", "", value)
    return value


# =========================================================
#        🔹 RECURSIVE PPT TEXT EXTRACTION (CHARTS + TABLES)
# =========================================================

def extract_text_recursive(shape):
    text = ""

    try:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for s in shape.shapes:
                text += extract_text_recursive(s) + " "

        elif hasattr(shape, "has_table") and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    text += cell.text.strip() + " "

        elif hasattr(shape, "has_chart") and shape.has_chart:
            chart = shape.chart

            if chart.has_title and chart.chart_title.has_text_frame:
                text += chart.chart_title.text_frame.text.strip() + " "

            try:
                for series in chart.series:
                    if series.has_data_labels:
                        for point in series.points:
                            if point.data_label and point.data_label.has_text_frame:
                                text += point.data_label.text_frame.text.strip() + " "
            except:
                pass

        elif hasattr(shape, "text") and shape.text.strip():
            text += shape.text.strip() + " "

    except:
        pass

    return text.strip()


def extract_text_from_pptx(file):
    prs = Presentation(file)
    slides_data = []

    for i, slide in enumerate(prs.slides):
        text = ""
        candidate_titles = []

        for shape in slide.shapes:
            shape_text = extract_text_recursive(shape)
            if shape_text:
                text += shape_text + " "
                candidate_titles.append((getattr(shape, "top", 999999), shape_text))

        title_text = ""
        if candidate_titles:
            candidate_titles.sort(key=lambda x: x[0])
            title_text = candidate_titles[0][1]

        slides_data.append({
            "slide_num": i + 1,
            "title": title_text,
            "text": text.strip()
        })

    return slides_data


# =========================================================
#                     🔹 SEARCH FUNCTION
# =========================================================

def search_ppt(file, keyword):
    keyword = keyword.lower()
    slides_data = extract_text_from_pptx(file)
    results = []

    for slide in slides_data:
        if keyword in slide["text"].lower():
            results.append({
                "PPT Name": file.name,
                "Slide No": slide["slide_num"],
                "Visualization Title": slide["title"]
            })

    return results


# =========================================================
#                    🔹 UI MAIN CONTAINER
# =========================================================

st.markdown('<div class="main-container">', unsafe_allow_html=True)
st.markdown('<div class="main-title">🔍 PPT Keyword Search Tool</div>', unsafe_allow_html=True)


# =========================================================
#                     🔹 INPUT SECTION
# =========================================================

uploaded_files = st.file_uploader(
    "Upload multiple PPTX files",
    type=["pptx"],
    accept_multiple_files=True
)

keyword = st.text_input("Enter keyword to search")

search_clicked = st.button("🔎 Search")


# =========================================================
#                     🔹 SEARCH ACTION
# =========================================================

if search_clicked:
    if not uploaded_files:
        st.warning("Please upload at least one PPTX file.")
        st.stop()

    if not keyword.strip():
        st.warning("Please enter a keyword.")
        st.stop()

    all_results = []

    for file in uploaded_files:
        slides = search_ppt(file, keyword)
        all_results.extend(slides)

    if not all_results:
        st.error("No results found.")
    else:
        df = pd.DataFrame(all_results)
        df = df.applymap(clean_text)

        st.markdown('<div class="output-box">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Search Results</div>', unsafe_allow_html=True)

        st.dataframe(df, use_container_width=True)

        excel_buffer = io.BytesIO()
        df.to_excel(excel_buffer, index=False, engine="openpyxl")
        excel_buffer.seek(0)

        st.download_button(
            label="⬇ Download Excel",
            data=excel_buffer,
            file_name=f"ppt_search_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.write("")

        st.button("⬆ Upload More PPTX Files", on_click=lambda: st.experimental_rerun())
        st.button("🧹 Clear Results", on_click=lambda: st.experimental_rerun())
        st.button("🔄 New Search", on_click=lambda: st.session_state.clear() or st.experimental_rerun())

        st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
#                           FOOTER
# =========================================================

st.markdown(
    f"""
    <br>
    <center style="color:{BLUE}; font-size:14px;">
        Built by <b>SKT</b> • Designed in Blue & White Theme
    </center>
    """,
    unsafe_allow_html=True
)

st.markdown("</div>", unsafe_allow_html=True)
