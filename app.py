# app.py
import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import tempfile
import os
import traceback

st.set_page_config(page_title="PPT Keyword Search", layout="wide")
st.title("🔍 PPT Keyword Search Tool (Streamlit) — Option A (Multiple PPTX Uploads)")
st.markdown("Upload one or more `.pptx` files, enter a keyword, and download an Excel with matching slides.")

# --------------------------
# Text extraction (robust)
# --------------------------
def extract_text_recursive(shape):
    text = ""
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for s in shape.shapes:
                text += extract_text_recursive(s) + " "

        elif hasattr(shape, "has_table") and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if cell.text:
                        text += cell.text.strip() + " "

        elif hasattr(shape, "has_chart") and shape.has_chart:
            chart = shape.chart
            try:
                if getattr(chart, "has_title", False):
                    title_frame = getattr(chart, "chart_title", None)
                    if title_frame and getattr(title_frame, "text_frame", None):
                        text += title_frame.text_frame.text.strip() + " "
            except Exception:
                pass

            try:
                if getattr(chart, "has_category_axis", False):
                    ax = chart.category_axis
                    if getattr(ax, "has_title", False) and getattr(ax, "axis_title", None):
                        t = ax.axis_title
                        if getattr(t, "has_text_frame", False):
                            text += t.text_frame.text.strip() + " "
            except Exception:
                pass

            try:
                if getattr(chart, "has_value_axis", False):
                    ax = chart.value_axis
                    if getattr(ax, "has_title", False) and getattr(ax, "axis_title", None):
                        t = ax.axis_title
                        if getattr(t, "has_text_frame", False):
                            text += t.text_frame.text.strip() + " "
            except Exception:
                pass

            try:
                if getattr(chart, "has_legend", False) and chart.legend:
                    for entry in chart.legend.entries:
                        try:
                            txt = entry.text
                            if txt and txt.strip():
                                text += txt.strip() + " "
                        except Exception:
                            pass
            except Exception:
                pass

            try:
                for series in chart.series:
                    if getattr(series, "has_data_labels", False):
                        for point in series.points:
                            lbl = getattr(point, "data_label", None)
                            if lbl and getattr(lbl, "has_text_frame", False):
                                text += lbl.text_frame.text.strip() + " "
            except Exception:
                pass

        elif hasattr(shape, "text") and shape.text and shape.text.strip():
            text += shape.text.strip() + " "
    except Exception:
        pass

    return text.strip()

def extract_text_from_presentation_bytes(filelike):
    """
    Accepts a file-like (BytesIO) or path string and returns list of slide dicts.
    Each dict: {'slide_num': int, 'title': str, 'text': str}
    """
    try:
        if isinstance(filelike, (str, os.PathLike)):
            prs = Presentation(str(filelike))
        else:
            prs = Presentation(filelike)
    except Exception:
        return []

    slides_data = []
    for i, slide in enumerate(prs.slides):
        text = ""
        candidate_titles = []
        for shape in slide.shapes:
            shape_text = extract_text_recursive(shape)
            if shape_text:
                text += shape_text + " "
                y_position = getattr(shape, "top", 9999999)
                font_size = 0
                try:
                    if hasattr(shape, "text_frame") and shape.text_frame is not None:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                try:
                                    if run.font and run.font.size:
                                        size = run.font.size.pt
                                        if size:
                                            font_size = max(font_size, size)
                                except Exception:
                                    pass
                except Exception:
                    pass
                candidate_titles.append((y_position, font_size, shape_text))

        title_text = ""
        if candidate_titles:
            candidate_titles.sort(key=lambda x: (x[0], -x[1]))
            title_text = candidate_titles[0][2]

        slides_data.append({"slide_num": i + 1, "title": title_text, "text": text.strip()})
    return slides_data

# --------------------------
# UI: uploads + inputs
# --------------------------
uploaded_files = st.file_uploader(
    "Upload one or more PPTX files",
    type=["pptx"],
    accept_multiple_files=True
)

col1, col2 = st.columns([3,1])
with col1:
    keyword = st.text_input("Keyword (case-insensitive)", "")
with col2:
    search_btn = st.button("Search 🔍")

progress_placeholder = st.empty()

# --------------------------
# Search action
# --------------------------
if search_btn:
    if not uploaded_files:
        st.warning("Please upload at least one .pptx file.")
    elif not keyword or keyword.strip() == "":
        st.warning("Please enter a keyword to search.")
    else:
        kw = keyword.strip().lower()
        results = []
        total = len(uploaded_files)
        with st.spinner("Processing files..."):
            for idx, uploaded in enumerate(uploaded_files, start=1):
                try:
                    fname = uploaded.name
                    b = uploaded.read()
                    bio = io.BytesIO(b)
                    slides = extract_text_from_presentation_bytes(bio)
                    for s in slides:
                        if kw in s["text"].lower():
                            results.append({
                                "PPT Title": fname,
                                "PPT Slide No": s["slide_num"],
                                "Visualization Title": s["title"]
                            })
                except Exception as e:
                    st.error(f"Error processing {uploaded.name}: {e}")
                    st.text(traceback.format_exc())
                progress = int((idx/total) * 100)
                progress_placeholder.progress(progress)

        progress_placeholder.empty()

        if results:
            df = pd.DataFrame(results, columns=["PPT Title", "PPT Slide No", "Visualization Title"])
            st.success(f"Found {len(df)} matching slide(s).")
            st.dataframe(df)

            # prepare Excel in-memory
            towrite = io.BytesIO()
            with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Results")
                writer.save()
            towrite.seek(0)

            # attempt left-align using openpyxl
            try:
                wb = load_workbook(towrite)
                ws = wb.active
                for row in ws.iter_rows():
                    for cell in row:
                        cell.alignment = Alignment(horizontal="left")
                out_stream = io.BytesIO()
                wb.save(out_stream)
                out_stream.seek(0)
            except Exception:
                towrite.seek(0)
                out_stream = towrite

            st.download_button(
                label="⬇️ Download results as Excel",
                data=out_stream,
                file_name="ppt_search_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("No matches found for the keyword in uploaded PPTX files.")
