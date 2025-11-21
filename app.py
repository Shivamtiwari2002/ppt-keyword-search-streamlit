import streamlit as st
import pandas as pd
import os
import re
import zipfile
import tempfile
from io import BytesIO
from pptx import Presentation
from rapidfuzz import fuzz
import html

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="PPT Keyword Search", layout="wide")

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
.mark { background: #FFF176; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='header-box'>PPT → HTML Keyword Search Tool</div>", unsafe_allow_html=True)


# ---------------- PPT → HTML Conversion ----------------
def ppt_to_html_slides(file_path):
    prs = Presentation(file_path)
    slides_out = []
    for i, slide in enumerate(prs.slides, start=1):
        parts = []
        title = ""
        try:
            if slide.shapes.title and slide.shapes.title.text:
                title = slide.shapes.title.text.strip()
        except:
            title = ""

        for shape in slide.shapes:
            if not hasattr(shape, "text"):
                continue
            raw = shape.text
            if raw and raw.strip():
                txt = html.escape(raw).replace("\n", "<br>")
                parts.append(f"<p>{txt}</p>")

        slide_html = f"<div><h3>Slide {i}</h3><h4>{html.escape(title)}</h4>{''.join(parts)}</div>"
        slides_out.append({"slide_no": i, "title": title, "html": slide_html})

    return slides_out


# ---------------- Highlight Matches ----------------
def highlight_terms(html_text, keyword):
    pattern = re.compile(re.escape(keyword), re.IGNORECASE)
    return pattern.sub(lambda m: f"<mark class='mark'>{m.group(0)}</mark>", html_text)


# ---------------- Search Logic ----------------
def search_slides(slides, keyword, mode="exact_phrase", threshold=80):
    results = []
    k = keyword.strip()

    # REGEX for Exact Phrase with punctuation allowed but NO extra words BEFORE
    exact_pattern = re.compile(
        r"(?<!\w)[\s\-\•\(\)]*" + re.escape(k) + r"(?=[\s\-\:\)\]]|$)",
        re.IGNORECASE
    )

    for s in slides:
        raw_text = re.sub(r"<[^>]+>", " ", s["html"])

        if mode == "exact_phrase":
            if exact_pattern.search(raw_text):
                results.append({**s, "score": 100})
        elif mode == "exact":
            if k.lower() in raw_text.lower():
                results.append({**s, "score": 100})
        else:
            score = fuzz.partial_ratio(k.lower(), raw_text.lower())
            if score >= threshold:
                results.append({**s, "score": score})

    return results


def extract_zip_pptx(zip_file):
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(zip_file, "r") as z:
        z.extractall(temp_dir)
    pptx_files = []
    for root, _, files in os.walk(temp_dir):
        for f in files:
            if f.lower().endswith(".pptx"):
                pptx_files.append(os.path.join(root, f))
    return pptx_files


# ---------------- Sidebar ----------------
with st.sidebar:
    st.markdown("### Search Mode")
    search_mode = st.radio(
        "Choose Search Type",
        ["exact_phrase (recommended)", "exact", "fuzzy"],
        index=0
    )

    st.markdown("### Fuzzy Threshold")
    threshold = st.slider("Fuzzy threshold", 60, 100, 85)

    st.markdown("---")
    st.markdown("Made by SKT")


# ---------------- Upload Files ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("<div class='upload-label'>Upload PPTX / ZIP Files</div>", unsafe_allow_html=True)

uploaded_files = st.file_uploader("", type=["pptx", "zip"], accept_multiple_files=True)

st.markdown("</div>", unsafe_allow_html=True)

# ---------------- Keyword Box ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
keyword = st.text_input("Enter Keyword", "", placeholder="e.g. PSD Manager")
st.markdown("</div>", unsafe_allow_html=True)


# ---------------- Search Button ----------------
search_btn = st.button("🔍 Search")


# ---------------- Collect PPT Paths ----------------
pptx_paths = []

if uploaded_files:
    for uf in uploaded_files:
        temp_path = os.path.join(tempfile.gettempdir(), uf.name)
        with open(temp_path, "wb") as f:
            f.write(uf.read())

        if uf.name.endswith(".pptx"):
            pptx_paths.append(temp_path)
        else:
            pptx_paths.extend(extract_zip_pptx(temp_path))


# ---------------- Perform Search ----------------
results_all = []

if search_btn:
    if not pptx_paths:
        st.error("Please upload PPTX files.")
    elif not keyword.strip():
        st.error("Enter a keyword.")
    else:
        mode_clean = search_mode.split(" ")[0]

        with st.spinner("Searching slides…"):
            for p in pptx_paths:
                slides = ppt_to_html_slides(p)
                matches = search_slides(slides, keyword, mode_clean, threshold)

                for m in matches:
                    highlighted = highlight_terms(m["html"], keyword)
                    results_all.append({
                        "File": os.path.basename(p),
                        "Slide": m["slide_no"],
                        "Title": m["title"],
                        "Score": m["score"],
                        "HTML": highlighted
                    })

        st.success(f"{len(results_all)} matches found.")


# ---------------- Display Results ----------------
if results_all:
    df = pd.DataFrame(results_all).drop(columns=["HTML"])

    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.markdown("### Search Results")

    def render_table(df):
        h = "<table class='table-container'>"
        h += "<tr>" + "".join(f"<th>{c}</th>" for c in df.columns) + "</tr>"
        for _, row in df.iterrows():
            h += "<tr>" + "".join(f"<td>{html.escape(str(x))}</td>" for x in row) + "</tr>"
        h += "</table>"
        return h

    st.markdown(render_table(df), unsafe_allow_html=True)

    # Download Excel
    excel_out = BytesIO()
    with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    excel_out.seek(0)

    st.download_button("⬇ Download Results (Excel)", excel_out.getvalue(), "ppt_results.xlsx")

    # Preview
    st.markdown("### Slide Previews")
    for r in results_all:
        st.markdown(f"#### {r['File']} — Slide {r['Slide']}")
        st.markdown(r["HTML"], unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


# ---------------- Footer ----------------
st.markdown("<div class='footer'>Made by SKT</div>", unsafe_allow_html=True)

