import streamlit as st
import pandas as pd
from pptx import Presentation
import zipfile
import tempfile
import os
import re
from io import BytesIO
from rapidfuzz import fuzz
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image

# ------------------- PAGE CONFIG -------------------
st.set_page_config(page_title="PPT Keyword Search Tool — Enhanced", layout="wide")

# ------------------- DEFAULTS & SESSION -------------------
if "recent_searches" not in st.session_state:
    st.session_state.recent_searches = []

if "last_results" not in st.session_state:
    st.session_state.last_results = pd.DataFrame()

# ------------------- CUSTOM CSS -------------------
st.markdown("""
<style>
.header {
    background-color: #0b63d6;
    padding: 18px;
    color: white;
    border-radius: 8px;
    font-size: 22px;
    font-weight: 700;
}
.badge {
    display:inline-block;
    padding:4px 8px;
    border-radius:6px;
    color:white;
    font-weight:700;
    font-size:12px;
}
.badge-high { background:#16a34a; }
.badge-medium { background:#f59e0b; }
.badge-low { background:#ef4444; }
.small-muted { font-size:12px; color:#555; }
</style>
""", unsafe_allow_html=True)

# ------------------- HEADER -------------------
st.markdown("<div class='header'>PPT Keyword Search Tool — Enhanced (No OCR Version)</div>", unsafe_allow_html=True)
st.markdown("##")

# ------------------- SIDEBAR -------------------
with st.sidebar:
    st.markdown("### Upload & Settings")
    uploaded_files = st.file_uploader(
        "Upload PPTX or ZIP files",
        type=["pptx", "zip"],
        accept_multiple_files=True
    )

    keywords_input = st.text_input("Enter keyword(s)", placeholder="e.g. job code, role")
    
    exact_match = st.checkbox("Exact match only", value=False)
    fuzzy_threshold = st.slider("Fuzzy threshold", 50, 100, 80)

    show_images = st.checkbox("Extract images (does NOT use OCR)", value=True)

    if st.button("🔍 Search"):
        st.session_state.run_search = True

# ------------------- HELPERS -------------------
def clean_text(text):
    if text is None:
        return ""
    return re.sub(r"[\000-\010\013\014\016-\037]", "", str(text))

def extract_from_all_shapes(shape, collect_images=False):
    text = ""
    images = []

    # Simple text
    if hasattr(shape, "text") and shape.text:
        text += shape.text + " "

    # Table text
    if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
        for row in shape.table.rows:
            for cell in row.cells:
                text += cell.text + " "

    # Grouped shapes
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for shp in shape.shapes:
            t, imgs = extract_from_all_shapes(shp, collect_images=collect_images)
            text += t
            images.extend(imgs)

    # Picture shapes (NO OCR)
    if collect_images and shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        img_blob = shape.image.blob
        images.append(img_blob)

    return text, images

def process_zip(file):
    extracted_temp = tempfile.mkdtemp()
    with zipfile.ZipFile(file, 'r') as z:
        z.extractall(extracted_temp)
    pptx_files = []
    for root, _, files in os.walk(extracted_temp):
        for f in files:
            if f.lower().endswith(".pptx"):
                pptx_files.append(os.path.join(root, f))
    return pptx_files

# ------------------- SEARCH LOGIC -------------------
if st.session_state.get("run_search", False):

    if not uploaded_files:
        st.error("⚠ Upload at least one file.")
        st.stop()

    if not keywords_input.strip():
        st.error("⚠ Enter at least one keyword.")
        st.stop()

    keywords = [k.strip() for k in keywords_input.split(",") if k.strip()]
    st.session_state.recent_searches.append(", ".join(keywords))

    matches = []
    all_pptx = []

    # Collect PPTX paths
    for file in uploaded_files:
        if file.name.lower().endswith(".pptx"):
            tmp_path = os.path.join(tempfile.gettempdir(), file.name)
            with open(tmp_path, "wb") as f:
                f.write(file.read())
            all_pptx.append(tmp_path)
        elif file.name.lower().ends_with(".zip"):
            all_pptx.extend(process_zip(file))

    # Search loop
    for path in all_pptx:
        prs = Presentation(path)

        for slide_num, slide in enumerate(prs.slides, start=1):
            slide_text = ""
            slide_images = []

            for shape in slide.shapes:
                t, imgs = extract_from_all_shapes(shape, collect_images=show_images)
                slide_text += t
                slide_images.extend(imgs)

            cleaned = clean_text(slide_text)

            for kw in keywords:
                kw_lower = kw.lower()

                if exact_match:
                    found = kw_lower in cleaned.lower()
                    similarity = 100 if found else 0
                else:
                    similarity = fuzz.partial_ratio(kw_lower, cleaned.lower())
                    found = similarity >= fuzzy_threshold

                if found:
                    matches.append({
                        "File": os.path.basename(path),
                        "Slide Number": slide_num,
                        "Keyword": kw,
                        "Similarity": similarity,
                        "Excerpt": cleaned[:200] + ("..." if len(cleaned) > 200 else ""),
                        "FullText": cleaned,
                        "Images": slide_images
                    })

    # Display results
    if not matches:
        st.warning("No matches found.")
    else:
        df = pd.DataFrame(matches)
        st.session_state.last_results = df

        st.success(f"Found {len(df)} matches")

        for _, row in df.iterrows():
            sim = row["Similarity"]
            badge = "badge-high" if sim >= 90 else "badge-medium" if sim >= 70 else "badge-low"

            st.markdown(f"""
                **Keyword:** {row['Keyword']}  
                **File:** {row['File']}  
                **Slide:** {row['Slide Number']}  
                <span class='badge {badge}'>{sim}% match</span>  
            """, unsafe_allow_html=True)

            with st.expander("Show full text"):
                st.write(row["FullText"])

            if show_images:
                with st.expander("Images extracted (NO OCR)"):
                    for img_blob in row["Images"]:
                        try:
                            img = Image.open(BytesIO(img_blob))
                            st.image(img, use_column_width=True)
                        except:
                            pass

        # Download Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        st.download_button(
            "⬇ Download Results",
            output.getvalue(),
            "ppt_search_results.xlsx"
        )
