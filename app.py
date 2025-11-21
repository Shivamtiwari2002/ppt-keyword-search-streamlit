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
from collections import Counter
import concurrent.futures

# --- AI Connector (OpenAI) ---
from openai import OpenAI
client = OpenAI()

def summarize_text(text):
    """Use OpenAI GPT to summarize slide content."""
    if not text.strip():
        return "No content to summarize."
    prompt = f"Summarize this slide in 1-2 concise sentences:\n{text}"
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role":"user","content":prompt}],
        temperature=0.5
    )
    return response.choices[0].message.content.strip()

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="PPT → HTML Keyword Search",
    layout="wide"
)

# ---------------- UI Theme ----------------
st.markdown("""
<style>
body { background: linear-gradient(140deg, #F7F9FF, #EDF2FF 40%, #FFFFFF 100%); font-family: 'Segoe UI', sans-serif; }
.header-box { background: linear-gradient(135deg, #0047FF, #3F8CFF); padding: 36px; border-radius: 20px; text-align:center; color:white; font-size:36px; font-weight:800; margin-bottom:35px; box-shadow:0px 8px 25px rgba(0,40,140,0.25);}
.section-card { background:white; padding:26px; border-radius:18px; border:1px solid #DFE6FF; box-shadow:0px 10px 22px rgba(0,60,160,0.08); margin-bottom:28px;}
.stFileUploader>div>div { border:2px dashed #0047FF !important; background:#EFF3FF !important; border-radius:16px !important; }
input[type="text"] { border:2px solid #0047FF !important; border-radius:14px !important; padding:14px !important; background:#EFF3FF !important; font-size:16px !important; color:#0033CC !important; font-weight:600 !important; }
.stButton>button { background:linear-gradient(135deg,#0047FF,#2F6BFF)!important; color:white!important; border-radius:12px!important; padding:12px 28px!important; font-size:17px!important; border:none!important; font-weight:600!important; box-shadow:0px 5px 15px rgba(0,0,0,0.17); transition:0.2s ease-in-out;}
.stButton>button:hover { transform:translateY(-2px); box-shadow:0px 7px 18px rgba(0,0,0,0.22); }
.mark { background:#FFF176; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='header-box'>AI-Powered PPT Keyword & Summary Tool</div>", unsafe_allow_html=True)

# ---------------- PPT to HTML ----------------
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
            if hasattr(shape, "text") and shape.text.strip():
                txt = html.escape(shape.text).replace("\n","<br>")
                parts.append(f"<p>{txt}</p>")
        html_content = f"<div><h3>Slide {i}</h3><h4>{html.escape(title)}</h4>{''.join(parts)}</div>"
        slides_out.append({"slide_no": i, "title": title, "html": html_content, "raw_text": " ".join([shape.text for shape in slide.shapes if hasattr(shape,"text")])})
    return slides_out

# ---------------- SEARCH HELPERS ----------------
def highlight_terms(html_text, keyword):
    pattern = re.compile(re.escape(keyword), re.IGNORECASE)
    return pattern.sub(lambda m: f"<mark class='mark'>{m.group(0)}</mark>", html_text)

def search_slides(slides, keyword, mode="exact_phrase", threshold=80):
    exact_pattern = re.compile(r"(?<!\\w)[\\s\\-•(]*"+re.escape(keyword)+r"(?=[\\s\\-:)\]]|$)", re.IGNORECASE)
    results = []
    for s in slides:
        text_for_search = re.sub(r"<[^>]+>", " ", s["html"])
        if mode=="exact_phrase":
            if exact_pattern.search(text_for_search):
                results.append({**s,"score":100})
        elif mode=="exact":
            if keyword.lower() in text_for_search.lower():
                results.append({**s,"score":100})
        else:
            score = fuzz.partial_ratio(keyword.lower(), text_for_search.lower())
            if score>=threshold:
                results.append({**s,"score":score})
    return results

def extract_zip_pptx(zip_file):
    temp_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(zip_file,"r") as z:
        z.extractall(temp_dir)
    pptx_files = []
    for root,_,files in os.walk(temp_dir):
        for f in files:
            if f.lower().endswith(".pptx"):
                pptx_files.append(os.path.join(root,f))
    return pptx_files

# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.markdown("### Search Mode")
    search_mode = st.radio("", ["exact_phrase (recommended)","exact","fuzzy"], index=0)
    threshold = st.slider("Fuzzy threshold", 60, 100, 85)
    st.markdown("---")
    st.markdown("Made by SKT")

# ---------------- UPLOAD ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("### Upload PPTX / ZIP Files")
uploaded_files = st.file_uploader("", type=["pptx","zip"], accept_multiple_files=True)
st.markdown("</div>", unsafe_allow_html=True)

# ---------------- KEYWORD ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
keyword = st.text_input("Enter Keyword","",placeholder="e.g. PSD Manager")
st.markdown("</div>", unsafe_allow_html=True)

# ---------------- SEARCH ----------------
search_btn = st.button("🔍 Search")
pptx_paths = []
if uploaded_files:
    for uf in uploaded_files:
        temp_path = os.path.join(tempfile.gettempdir(), uf.name)
        with open(temp_path,"wb") as f: f.write(uf.read())
        if uf.name.lower().endswith(".pptx"):
            pptx_paths.append(temp_path)
        else:
            pptx_paths.extend(extract_zip_pptx(temp_path))

results_all = []
keyword_counter = Counter()

if search_btn:
    if not pptx_paths:
        st.error("Please upload PPTX or ZIP files.")
    elif not keyword.strip():
        st.error("Please enter a keyword.")
    else:
        mode_clean = search_mode.split()[0]
        progress = st.progress(0)
        total_files = len(pptx_paths)

        def process_file(p, idx):
            slides = ppt_to_html_slides(p)
            matches = search_slides(slides, keyword, mode_clean, threshold)
            local_results = []
            for m in matches:
                highlighted = highlight_terms(m["html"], keyword)
                summary = summarize_text(m["raw_text"])
                local_results.append({
                    "File": os.path.basename(p),
                    "Slide": m["slide_no"],
                    "Title": m["title"],
                    "Score": m["score"],
                    "HTML": highlighted,
                    "Summary": summary
                })
                keyword_counter.update(re.findall(keyword, m["raw_text"], re.IGNORECASE))
            progress.progress((idx+1)/total_files)
            return local_results

        with concurrent.futures.ThreadPoolExecutor() as executor:
            futures = [executor.submit(process_file, p, idx) for idx,p in enumerate(pptx_paths)]
            for f in concurrent.futures.as_completed(futures):
                results_all.extend(f.result())
        st.success(f"{len(results_all)} matches found.")

# ---------------- RESULTS ----------------
if results_all:
    st.markdown("### Slide Previews")
    for r in results_all:
        with st.expander(f"{r['File']} — Slide {r['Slide']}"):
            st.markdown(r['HTML'], unsafe_allow_html=True)
            st.markdown(f"**Summary:** {r['Summary']}")
    st.markdown("### Keyword Frequency Across Slides")
    freq_df = pd.DataFrame(keyword_counter.most_common(), columns=["Keyword","Count"])
    st.table(freq_df)
