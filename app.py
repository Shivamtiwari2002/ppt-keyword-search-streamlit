# app.py
import streamlit as st
import pandas as pd
import os
import re
import zipfile
import tempfile
from io import BytesIO, StringIO
from pptx import Presentation
from rapidfuzz import fuzz
import html

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="PPT → HTML Keyword Search", layout="wide")

# ---------------- UI THEME CSS (Blue & White) ----------------
st.markdown("""
<style>
body { background-color: #F5F7FF; }
.header-box {
    background: linear-gradient(135deg, #0047FF, #3F8CFF);
    padding: 22px;
    border-radius: 12px;
    text-align: center;
    color: white;
    font-size: 28px;
    font-weight: 800;
    margin-bottom: 18px;
    box-shadow: 0px 4px 14px rgba(0,0,0,0.12);
}
.section-card { background: white; padding: 18px; border-radius: 12px; box-shadow: 0px 4px 12px rgba(0,0,0,0.06); margin-bottom: 18px; }
.stFileUploader>div>div { border: 2px dashed #0047FF !important; background: #EFF3FF !important; border-radius: 12px !important; }
.upload-label { font-size: 16px; font-weight: 700; color: #0047FF; }
.keyword-box input { border: 2px solid #0047FF !important; border-radius: 10px !important; padding: 10px !important; background: #EFF3FF !important; font-size: 15px !important; color: #0033CC !important; font-weight: 600 !important; }
.stButton>button { background: #0047FF !important; color: white !important; border-radius: 8px !important; padding: 8px 20px !important; font-size: 15px !important; border: none !important; font-weight: 600 !important; box-shadow: 0px 3px 10px rgba(0,0,0,0.12); }
.table-container { width: 100%; border-collapse: collapse; margin-top: 10px; }
.table-container th { background: #0047FF; color: white; padding: 8px; text-align: left; }
.table-container tr:hover { background: #E8EEFF; }
.table-container td { padding: 8px; border-bottom: 1px solid #DDD; font-size: 14px; }
.footer { text-align:center; margin-top: 28px; padding: 12px; font-size: 13px; font-weight: 600; color: #0047FF; }
.mark { background: #FFF176; }  /* highlight (yellowish) – acceptable with blue theme */
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='header-box'>PPT → HTML Keyword Search Tool</div>", unsafe_allow_html=True)

# ---------------- Constants / Paths ----------------
LOCAL_EXCEL_PATH = "data/keywords.xlsx"  # place your excel here in repo
SAMPLE_PPT_PATH = "/mnt/data/Process Modeling Training Rebranded.pptx"  # user-provided sample path

# ---------------- Utility functions ----------------
@st.cache_data
def load_keywords_from_excel(file_like_or_path, col_name=None):
    try:
        df = pd.read_excel(file_like_or_path)
    except Exception as e:
        raise RuntimeError(f"Failed to read Excel: {e}")
    # choose a column: explicit col_name else first non-empty column
    if col_name and col_name in df.columns:
        series = df[col_name]
    else:
        nonempty_cols = [c for c in df.columns if df[c].dropna().astype(str).str.strip().any()]
        if not nonempty_cols:
            return []
        series = df[nonempty_cols[0]]
    keywords = series.dropna().astype(str).str.strip().unique().tolist()
    keywords = [k for k in keywords if k]
    keywords.sort()
    return keywords

def ppt_to_html_slides(file_path):
    """
    Convert pptx to list of per-slide HTML strings (basic text and simple formatting).
    Returns list of dicts: [{slide_no: int, title: str, html: str}, ...]
    """
    prs = Presentation(file_path)
    slides_out = []
    for i, slide in enumerate(prs.slides, start=1):
        parts = []
        title = ""
        # attempt to read title
        try:
            if slide.shapes.title and slide.shapes.title.text:
                title = slide.shapes.title.text.strip()
        except:
            title = ""
        for shape in slide.shapes:
            # skip if no text
            if not hasattr(shape, "text"):
                continue
            raw = shape.text
            if not raw or not raw.strip():
                continue
            # basic escape and replace line breaks
            txt = html.escape(raw).replace("\r\n", "<br>").replace("\n", "<br>")
            # wrap small headings vs paragraphs
            # If shape seems like a title (large text) we can't detect easily — use position? skip complexity
            parts.append(f"<p>{txt}</p>")
        html_content = "\n".join(parts)
        # create small slide wrapper
        slide_html = f"<div class='slide-block'><h3>Slide {i}</h3><h4>{html.escape(title)}</h4>{html_content}</div>"
        slides_out.append({"slide_no": i, "title": title, "html": slide_html})
    return slides_out

def highlight_terms(html_text, keyword):
    """Simple case-insensitive highlight of keyword in HTML text without breaking tags.
       We'll work on the visible text by using a regex that ignores tags.
    """
    if not keyword:
        return html_text
    def repl(match):
        return f"<mark class='mark'>{match.group(0)}</mark>"
    # build regex for keyword (escape special chars), case-insensitive
    try:
        pattern = re.compile(re.escape(keyword), flags=re.IGNORECASE)
        # naive approach: replace inside text nodes by applying to whole HTML (ok for our simple generated HTML)
        return pattern.sub(repl, html_text)
    except re.error:
        return html_text

def search_slides(slides, keyword, mode="fuzzy", threshold=80):
    results = []
    k = keyword.strip()
    if not k:
        return results
    for s in slides:
        text_for_search = re.sub(r"<[^>]+>", " ", s["html"])  # strip tags for search content
        if mode == "exact":
            if k.lower() in text_for_search.lower():
                snippet = text_for_search.strip()[:300]
                results.append({**s, "matched_snippet": snippet, "score": 100})
        else:
            score = fuzz.partial_ratio(k.lower(), text_for_search.lower())
            if score >= threshold:
                snippet = text_for_search.strip()[:300]
                results.append({**s, "matched_snippet": snippet, "score": score})
    return results

def extract_zip_pptx(zip_file):
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(zip_file, 'r') as z:
            z.extractall(temp_dir)
    except Exception as e:
        raise RuntimeError(f"Error extracting ZIP: {e}")
    pptx_paths = []
    for root, _, files in os.walk(temp_dir):
        for f in files:
            if f.lower().endswith(".pptx"):
                pptx_paths.append(os.path.join(root, f))
    return pptx_paths

# ---------------- Sidebar: Settings & Info ----------------
with st.sidebar:
    st.markdown("### Search Settings")
    search_mode = st.radio("Match mode", options=["fuzzy", "exact"], index=0)
    if search_mode == "fuzzy":
        threshold = st.slider("Fuzzy threshold", min_value=60, max_value=100, value=85, step=1)
    else:
        threshold = 100
    st.markdown("---")
    st.markdown("### Excel Dropdown Source")
    st.markdown("Place a local Excel at `data/keywords.xlsx` or upload one below (it should contain one column of keywords).")
    st.markdown("---")
    st.markdown("Made by SKT")

# ---------------- Top: Excel load UI ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("<div class='upload-label'>Load Keywords for Dropdown</div>", unsafe_allow_html=True)
col1, col2 = st.columns([2, 3])
with col1:
    uploaded_excel = st.file_uploader("Upload Excel (optional)", type=["xlsx", "xls"], key="keywords_excel")
with col2:
    use_local = st.checkbox(f"Use local Excel file (data/keywords.xlsx)", value=False)

# Load keywords list with helpful messages
keywords_list = []
preview_df = None
if uploaded_excel is not None:
    try:
        keywords_list = load_keywords_from_excel(uploaded_excel)
        st.success(f"Loaded {len(keywords_list)} keywords from uploaded file.")
    except Exception as e:
        st.error(f"Failed to load uploaded Excel: {e}")
elif use_local:
    if os.path.exists(LOCAL_EXCEL_PATH):
        try:
            keywords_list = load_keywords_from_excel(LOCAL_EXCEL_PATH)
            st.success(f"Loaded {len(keywords_list)} keywords from local file.")
        except Exception as e:
            st.error(f"Failed to load local Excel: {e}")
    else:
        st.warning(f"Local file not found at: {LOCAL_EXCEL_PATH}")

if keywords_list:
    keyword_options = ["-- Select --"] + keywords_list
else:
    keyword_options = ["-- No keywords loaded --"]

st.markdown("</div>", unsafe_allow_html=True)

# ---------------- File uploader for PPT/ZIP and Sample button ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("<div class='upload-label'>Upload PPTX / ZIP Files (or use sample)</div>", unsafe_allow_html=True)
col1, col2, col3 = st.columns([3, 2, 2])
with col1:
    uploaded_files = st.file_uploader("", type=["pptx", "zip"], accept_multiple_files=True)
with col2:
    use_sample = st.button("Load sample PPT (dev)")
with col3:
    clear_cache = st.button("Clear Cache")

st.markdown("</div>", unsafe_allow_html=True)

# Clear cache action
if clear_cache:
    try:
        st.experimental_memo_clear()
        st.cache_data.clear()
        st.success("Cache cleared.")
    except:
        st.success("Cache cleared (best-effort).")

# ---------------- Keyword selection UI ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("### Select or Enter Keyword")
colA, colB = st.columns([2, 3])
with colA:
    selected_keyword = st.selectbox("Choose from dropdown", keyword_options)
with colB:
    custom_keyword = st.text_input("Or type your own keyword", "")

# final keyword resolution
if custom_keyword and custom_keyword.strip():
    keyword = custom_keyword.strip()
elif selected_keyword and selected_keyword not in ["-- Select --", "-- No keywords loaded --"]:
    keyword = selected_keyword
else:
    keyword = ""

st.markdown("</div>", unsafe_allow_html=True)

# ---------------- Advanced options ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("### Advanced Options")
show_html_preview = st.checkbox("Show matched slide HTML preview", value=True)
download_htmls = st.checkbox("Allow download of matched slide HTMLs", value=True)
highlight_matches = st.checkbox("Highlight matched keyword in preview", value=True)
st.markdown("</div>", unsafe_allow_html=True)

# ---------------- Search button ----------------
search_btn = st.button("🔍 Search in PPTs")

# ---------------- Workflows: prepare list of pptx files to process ----------------
pptx_paths_to_process = []  # list of tuples (display_name, path)

# If sample button clicked, try to use SAMPLE_PPT_PATH
if use_sample:
    if os.path.exists(SAMPLE_PPT_PATH):
        pptx_paths_to_process.append((os.path.basename(SAMPLE_PPT_PATH), SAMPLE_PPT_PATH))
        st.info(f"Loaded sample PPT: {SAMPLE_PPT_PATH}")
    else:
        st.error(f"Sample PPT not found at {SAMPLE_PPT_PATH}")

# If uploaded files present, write them to temp and add
if uploaded_files:
    for uf in uploaded_files:
        # write to temp dir
        temp_path = os.path.join(tempfile.gettempdir(), uf.name)
        with open(temp_path, "wb") as f:
            f.write(uf.read())
        if uf.name.lower().endswith(".pptx"):
            pptx_paths_to_process.append((uf.name, temp_path))
        elif uf.name.lower().endswith(".zip"):
            try:
                extracted = extract_zip_pptx(temp_path)
                for p in extracted:
                    pptx_paths_to_process.append((os.path.basename(p), p))
            except Exception as e:
                st.error(f"Failed to extract ZIP {uf.name}: {e}")

# If nothing selected yet, show a friendly note
if not pptx_paths_to_process and not search_btn:
    st.info("Upload PPTX or ZIP to search, or click 'Load sample PPT (dev)'. You can load keywords from Excel too.")

# ---------------- Search Logic ----------------
results_all = []  # will be list of dicts: file, slide_no, title, html, snippet, score
if search_btn:
    if not pptx_paths_to_process:
        st.error("No PPTX files found to search. Upload a PPTX/ZIP or load sample.")
    elif not keyword:
        st.error("Please select from dropdown or enter a keyword.")
    else:
        with st.spinner("Converting PPTs to HTML and searching..."):
            for display_name, path in pptx_paths_to_process:
                try:
                    slides = ppt_to_html_slides(path)
                except Exception as e:
                    st.error(f"Failed converting {display_name}: {e}")
                    continue
                matches = search_slides(slides, keyword, mode=search_mode, threshold=threshold)
                for m in matches:
                    # compute highlighted html if needed
                    slide_html = m["html"]
                    if highlight_matches:
                        try:
                            slide_html = highlight_terms(slide_html, keyword)
                        except:
                            pass
                    results_all.append({
                        "File": display_name,
                        "Slide Number": m["slide_no"],
                        "Title": m["title"],
                        "Matched Snippet": m.get("matched_snippet", "")[:500],
                        "Score": m.get("score", 0),
                        "Slide HTML": slide_html
                    })
        st.success(f"Search completed. {len(results_all)} matches found.")

# ---------------- Results display and download ----------------
if results_all:
    df = pd.DataFrame(results_all).drop(columns=["Slide HTML"], errors="ignore")
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.markdown("### Search Results")
    # Render table HTML for nicer look
    def render_table(df):
        html_out = "<table class='table-container'>"
        html_out += "<tr>" + "".join(f"<th>{c}</th>" for c in df.columns) + "</tr>"
        for _, row in df.iterrows():
            html_out += "<tr>" + "".join(f"<td>{html.escape(str(x))}</td>" for x in row) + "</tr>"
        html_out += "</table>"
        return html_out

    st.markdown(render_table(df), unsafe_allow_html=True)

    # download all results as Excel
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="SearchResults")
    towrite.seek(0)
    st.download_button("⬇ Download Results (Excel)", towrite.getvalue(), file_name="ppt_search_results.xlsx")

    # Show per-match preview and download options
    if show_html_preview:
        st.markdown("### Matched Slide Previews")
        for idx, r in enumerate(results_all, start=1):
            st.markdown(f"#### {idx}. {r['File']} — Slide {r['Slide Number']} — Score: {r['Score']}")
            st.markdown(f"**Title:** {html.escape(r['Title'])}")
            st.markdown(r["Slide HTML"], unsafe_allow_html=True)
            if download_htmls:
                # prepare download
                html_bytes = r["Slide HTML"].encode("utf-8")
                filename = f"{os.path.splitext(r['File'])[0]}_slide_{r['Slide Number']}.html"
                st.download_button(f"Download Slide {r['Slide Number']} HTML", html_bytes, file_name=filename, key=f"dl_{idx}")
    st.markdown("</div>", unsafe_allow_html=True)

# ---------------- Footer ----------------
st.markdown("<div class='footer'>Made by SKT</div>", unsafe_allow_html=True)
