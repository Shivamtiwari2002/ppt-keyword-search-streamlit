# Full upgraded Streamlit app
# Features added:
# - Sidebar with settings
# - Multi-keyword support (comma-separated)
# - Exact match toggle + fuzzy threshold slider
# - Similarity score + color-coded match strength
# - Expanders to preview full extracted text + embedded images
# - Recent searches (session_state)
# - OCR + tables + grouped shapes + pictures extraction (from prior version)

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
import pytesseract
import base64

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
.badge-high { background:#16a34a; }   /* green */
.badge-medium { background:#f59e0b; } /* amber */
.badge-low { background:#ef4444; }    /* red */
.result-row { padding:8px; border-bottom:1px solid #e6eefc; }
.small-muted { font-size:12px; color:#555; }
.uploader { border:2px dashed #0b63d6; padding:14px; border-radius:10px; background:#f2f7ff; }
.keyword { border:1px solid #0b63d6; border-radius:8px; padding:8px; width:100%; }
</style>
""", unsafe_allow_html=True)

# ------------------- HEADER -------------------
st.markdown("<div class='header'>PPT Keyword Search Tool — Enhanced (OCR • SmartArt • Multi-keyword • Preview)</div>", unsafe_allow_html=True)
st.markdown("##")

# ------------------- SIDEBAR -------------------
with st.sidebar:
    st.markdown("### Upload & Settings")
    uploaded_files = st.file_uploader(
        "Click or Drag & Drop PPTX / ZIP files here (Limit 200MB per file)", 
        type=["pptx", "zip"], 
        accept_multiple_files=True,
        help="You can upload multiple PPTX or ZIP (containing PPTX) files."
    )

    st.markdown("---")
    keywords_input = st.text_input("Enter keyword(s) (comma-separated)", placeholder="e.g. job code 1234, Integration BDM")
    st.markdown("Match mode:")
    exact_match = st.checkbox("Exact match only (no fuzzy)", value=False)
    fuzzy_threshold = st.slider("Fuzzy similarity threshold (when fuzzy is ON)", min_value=50, max_value=100, value=80)
    st.markdown("---")
    st.markdown("Advanced options")
    include_ocr = st.checkbox("Enable OCR for images (slower)", value=True)
    show_preview_images = st.checkbox("Show images extracted from slides", value=True)
    st.markdown("---")
    if st.button("🔍 Search"):
        st.session_state.run_search = True
    else:
        if "run_search" not in st.session_state:
            st.session_state.run_search = False

    st.markdown("---")
    st.markdown("Recent searches")
    for rec in reversed(st.session_state.recent_searches[-6:]):
        st.markdown(f"- {rec}")

# ------------------- MAIN: Validation & Controls -------------------
st.markdown("### Controls")
col1, col2 = st.columns([1, 3])
with col1:
    st.write("")  # spacer
with col2:
    st.info("Use the sidebar to upload files and change search settings. Click Search in the sidebar when ready.")

# ------------------- HELPERS -------------------
def clean_text(text):
    if text is None:
        return ""
    return re.sub(r"[\000-\010\013\014\016-\037]", "", str(text))

def extract_from_all_shapes(shape, enable_ocr=True, collect_images=False):
    """
    Recursively extract text from a shape, handle tables, groups.
    Also collect picture blobs if collect_images True.
    Returns (text, list_of_image_bytes)
    """
    text = ""
    images = []

    # Text (most shape types)
    try:
        if hasattr(shape, "text") and shape.text:
            text += shape.text + " "
    except Exception:
        pass

    # Table
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
            tbl = shape.table
            for r in range(len(tbl.rows)):
                for c in range(len(tbl.columns)):
                    try:
                        cell_text = tbl.cell(r, c).text
                        text += cell_text + " "
                    except Exception:
                        pass
    except Exception:
        pass

    # Group
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for shp in shape.shapes:
                t, imgs = extract_from_all_shapes(shp, enable_ocr=enable_ocr, collect_images=collect_images)
                text += t
                images.extend(imgs)
    except Exception:
        pass

    # Picture: optionally capture image bytes for optional OCR or preview
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            img_blob = shape.image.blob
            if collect_images:
                images.append(img_blob)
            # OCR if enabled
            if enable_ocr:
                try:
                    img = Image.open(BytesIO(img_blob))
                    ocr_text = pytesseract.image_to_string(img)
                    if ocr_text:
                        text += " " + ocr_text + " "
                except Exception:
                    pass
    except Exception:
        pass

    return text, images

def process_zip(file):
    extracted_temp = tempfile.mkdtemp()
    with zipfile.ZipFile(file, 'r') as zip_ref:
        zip_ref.extractall(extracted_temp)
    pptx_files = []
    for root, _, files in os.walk(extracted_temp):
        for f in files:
            if f.lower().endswith(".pptx"):
                pptx_files.append(os.path.join(root, f))
    return pptx_files

# ------------------- SEARCH LOGIC -------------------
if st.session_state.run_search:
    # validate
    if not uploaded_files:
        st.error("⚠ Please upload at least one PPTX or ZIP file in the sidebar.")
        st.session_state.run_search = False
    elif not keywords_input.strip():
        st.error("⚠ Please enter at least one keyword in the sidebar.")
        st.session_state.run_search = False
    else:
        keywords = [k.strip() for k in re.split(r",|\n|;", keywords_input) if k.strip()]
        st.session_state.recent_searches.append(", ".join(keywords))
        # counters
        total_files = 0
        total_slides = 0
        matches = []

        progress_text = st.empty()
        progress_bar = st.progress(0)

        # Build list of pptx paths (for direct uploaded pptx we create temp files)
        all_pptx_paths = []
        for uploaded in uploaded_files:
            fname = uploaded.name.lower()
            if fname.endswith(".pptx"):
                total_files += 1
                tmp_path = os.path.join(tempfile.gettempdir(), uploaded.name)
                with open(tmp_path, "wb") as f:
                    f.write(uploaded.read())
                all_pptx_paths.append(tmp_path)
            elif fname.endswith(".zip"):
                pptx_list = process_zip(uploaded)
                for p in pptx_list:
                    total_files += 1
                all_pptx_paths.extend(pptx_list)

        if not all_pptx_paths:
            st.warning("No PPTX files were found inside the uploaded files.")
            st.session_state.run_search = False
        else:
            # iterate files
            file_index = 0
            for file_path in all_pptx_paths:
                file_index += 1
                try:
                    prs = Presentation(file_path)
                except Exception as e:
                    st.warning(f"Could not open {os.path.basename(file_path)}: {e}")
                    continue

                for slide_num, slide in enumerate(prs.slides, start=1):
                    total_slides += 1
                    # accumulate shape text + images
                    slide_text = ""
                    slide_images = []
                    for shape in slide.shapes:
                        t, imgs = extract_from_all_shapes(shape, enable_ocr=include_ocr, collect_images=show_preview_images)
                        slide_text += t
                        if imgs:
                            slide_images.extend(imgs)

                    # normalize
                    cleaned = clean_text(slide_text)
                    if not cleaned.strip():
                        continue

                    # Evaluate matches for each keyword separately (record each match)
                    for kw in keywords:
                        kw_lower = kw.lower()
                        if exact_match:
                            found = kw_lower in cleaned.lower()
                            similarity = 100 if found else 0
                        else:
                            # fuzzy partial ratio - choose higher of partial and token set ratio (robust)
                            similarity = fuzz.partial_ratio(kw_lower, cleaned.lower())
                            found = similarity >= fuzzy_threshold

                        if found:
                            matches.append({
                                "File": os.path.basename(file_path),
                                "Slide Number": slide_num,
                                "Keyword": kw,
                                "Similarity": int(similarity),
                                "Excerpt": (cleaned[:200] + "...") if len(cleaned) > 200 else cleaned,
                                "FullText": cleaned,
                                "Images": slide_images
                            })

                # update progress
                progress = int( (file_index / len(all_pptx_paths)) * 100 )
                progress_bar.progress(progress)
                progress_text.info(f"Processed {file_index}/{len(all_pptx_paths)} files — scanned slides: {total_slides}")

            progress_bar.progress(100)
            progress_text.success(f"Search complete. Files scanned: {total_files}, Slides scanned: {total_slides}, Matches found: {len(matches)}")

            # Present results
            if not matches:
                st.warning("No matches found for the keywords.")
                # clear last results
                st.session_state.last_results = pd.DataFrame()
            else:
                df = pd.DataFrame(matches)
                # keep in session
                st.session_state.last_results = df.copy()

                # Allow filtering / sorting
                st.markdown("### Summary")
                c1, c2, c3 = st.columns(3)
                c1.metric("Files scanned", total_files)
                c2.metric("Slides scanned", total_slides)
                c3.metric("Total matches", len(df))

                st.markdown("### Results")
                # Optionally allow quick filter by keyword or file
                with st.expander("Filters"):
                    filt_col1, filt_col2, filt_col3 = st.columns([2,2,1])
                    with filt_col1:
                        kw_filter = st.selectbox("Filter by Keyword", options=["(all)"] + sorted(df["Keyword"].unique().tolist()))
                    with filt_col2:
                        file_filter = st.selectbox("Filter by File", options=["(all)"] + sorted(df["File"].unique().tolist()))
                    with filt_col3:
                        min_sim = st.slider("Min similarity", min_value=0, max_value=100, value=50)

                df_display = df.copy()
                if kw_filter and kw_filter != "(all)":
                    df_display = df_display[df_display["Keyword"] == kw_filter]
                if file_filter and file_filter != "(all)":
                    df_display = df_display[df_display["File"] == file_filter]
                df_display = df_display[df_display["Similarity"] >= min_sim]
                df_display = df_display.sort_values(by=["Similarity"], ascending=False).reset_index(drop=True)

                # Render results: each row as expander with color badge
                for idx, row in df_display.iterrows():
                    sim = int(row["Similarity"])
                    if sim >= 90:
                        badge_class = "badge-high"
                        badge_label = f"High {sim}%"
                    elif sim >= 75:
                        badge_class = "badge-medium"
                        badge_label = f"Medium {sim}%"
                    else:
                        badge_class = "badge-low"
                        badge_label = f"Low {sim}%"

                    left, right = st.columns([4,1])
                    with left:
                        st.markdown(f"**{row['Keyword']}**  —  `{row['File']}`  •  Slide {row['Slide Number']}")
                        st.markdown(f"<div class='small-muted'>{row['Excerpt']}</div>", unsafe_allow_html=True)
                    with right:
                        st.markdown(f"<div class='badge {badge_class}'>{badge_label}</div>", unsafe_allow_html=True)

                    # Show details in expander
                    with st.expander("Show full slide text & images"):
                        st.write(row["FullText"])
                        # show extracted images if any
                        images = row.get("Images") or []
                        if show_preview_images and images:
                            st.markdown("**Embedded images (extracted)**")
                            img_cols = st.columns(min(3, len(images)))
                            for i, b in enumerate(images):
                                try:
                                    image = Image.open(BytesIO(b))
                                    img_cols[i % 3].image(image, use_column_width=True)
                                except Exception:
                                    pass
                        else:
                            st.info("No images extracted from this slide (or image preview disabled).")

                # Download full results as Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Results")
                st.download_button(
                    label="⬇ Download Full Results (Excel)",
                    data=output.getvalue(),
                    file_name="ppt_search_results_enhanced.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ------------------- FOOTER / LAST RESULTS QUICK VIEW -------------------
st.markdown("---")
st.markdown("### Quick actions")
col_a, col_b, col_c = st.columns(3)
with col_a:
    if st.button("Show last results table"):
        if st.session_state.last_results.empty:
            st.warning("No previous results in this session.")
        else:
            st.dataframe(st.session_state.last_results)
with col_b:
    if st.button("Clear recent searches"):
        st.session_state.recent_searches = []
        st.success("Recent searches cleared.")
with col_c:
    if st.button("Reset app state"):
        st.session_state.clear()
        st.experimental_rerun()

st.markdown("<div style='text-align:center; margin-top:18px; color:#0b63d6; font-weight:700'>Made by SKT — Enhanced</div>", unsafe_allow_html=True)
