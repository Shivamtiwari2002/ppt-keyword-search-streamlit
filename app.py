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
from openai import OpenAI

client = OpenAI()

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Impact Analysis Tool",
    layout="wide"
)

# ---------------- UI THEME WITH IMAGES ----------------
st.markdown("""
<style>

/* BODY: Background image + gradient overlay */
body {
    background:
      linear-gradient(140deg, rgba(247,249,255,0.85) 0%, rgba(237,242,255,0.85) 40%, rgba(255,255,255,0.85) 100%),
      url('bg_main.png') no-repeat center center fixed;
    background-size: cover;
    font-family: 'Segoe UI', sans-serif;
}

/* Animated soft gradient overlay */
body::before {
    content: "";
    position: fixed;
    top: 0; left: 0;
    width: 100%; height: 100%;
    background: radial-gradient(circle at 20% 20%, rgba(0,72,255,0.10), transparent 50%),
                radial-gradient(circle at 80% 80%, rgba(0,72,255,0.08), transparent 50%);
    animation: floatBg 12s ease-in-out infinite alternate;
    z-index: -1;
}

@keyframes floatBg {
    0% { transform: translate(0px, 0px); }
    100% { transform: translate(10px, -10px); }
}

/* HEADER */
.header-box {
    background: linear-gradient(135deg, #0047FF, #3F8CFF);
    padding: 36px;
    border-radius: 20px;
    text-align: center;
    color: white;
    font-size: 36px;
    font-weight: 800;
    margin-bottom: 35px;
    box-shadow: 0px 8px 25px rgba(0,40,140,0.25);
    position: relative;
    overflow: hidden;
}

/* Decorative floating images in header */
.header-box::before {
    content: url('decor_circle1.png');
    position: absolute;
    top: -40px; left: -50px;
    width: 150px;
    opacity: 0.15;
}
.header-box::after {
    content: url('decor_circle2.png');
    position: absolute;
    bottom: -50px; right: -50px;
    width: 180px;
    opacity: 0.12;
}

/* CARDS */
.section-card {
    background: white;
    padding: 26px;
    border-radius: 18px;
    border: 1px solid #DFE6FF;
    box-shadow: 0px 10px 22px rgba(0,60,160,0.08);
    margin-bottom: 28px;
}

/* FILE UPLOADER */
.stFileUploader>div>div {
    border: 2px dashed #0047FF !important;
    background: #EFF3FF !important;
    border-radius: 16px !important;
}

/* TEXT INPUT */
input[type="text"] {
    border: 2px solid #0047FF !important;
    border-radius: 14px !important;
    padding: 14px !important;
    background: #EFF3FF !important;
    font-size: 16px !important;
    color: #0033CC !important;
    font-weight: 600 !important;
}

/* BUTTONS */
.stButton>button {
    background: linear-gradient(135deg, #0047FF, #2F6BFF) !important;
    color: white !important;
    border-radius: 12px !important;
    padding: 12px 28px !important;
    font-size: 17px !important;
    border: none !important;
    font-weight: 600 !important;
    box-shadow: 0px 5px 15px rgba(0,0,0,0.17);
    transition: 0.2s ease-in-out;
}

.stButton>button:hover {
    transform: translateY(-2px);
    box-shadow: 0px 7px 18px rgba(0,0,0,0.22);
}

/* TABLE */
.table-container {
    width: 100%;
    border-collapse: collapse;
    margin-top: 22px;
}

.table-container th {
    background: #0047FF;
    color: white;
    padding: 12px;
    text-align: left;
}

.table-container tr:nth-child(even) { background: #F0F4FF; }
.table-container tr:hover { background: #E5EBFF; }

.table-container td {
    padding: 12px;
    border-bottom: 1px solid #D0D8FF;
    font-size: 15px;
}

/* FOOTER */
.footer {
    text-align: center;
    margin-top: 40px;
    padding: 14px;
    font-size: 14px;
    font-weight: 600;
    color: #0047FF;
}

.mark { background: #FFF176; }

</style>
""", unsafe_allow_html=True)

st.markdown("<div class='header-box'>Impact Analysis Tool</div>", unsafe_allow_html=True)

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

        raw = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                raw.append(shape.text)
                txt = html.escape(shape.text).replace("\n", "<br>")
                parts.append(f"<p>{txt}</p>")

        html_content = f"<div><h3>Slide {i}</h3><h4>{html.escape(title)}</h4>{''.join(parts)}</div>"
        slides_out.append({
            "slide_no": i,
            "title": title,
            "html": html_content,
            "raw_text": " ".join(raw)
        })
    return slides_out


# ---------------- SEARCH HELPERS ----------------
def highlight_terms(html_text, keyword):
    pattern = re.compile(re.escape(keyword), re.IGNORECASE)
    return pattern.sub(lambda m: f"<mark class='mark'>{m.group(0)}</mark>", html_text)


def search_slides(slides, keyword, mode="exact_phrase", threshold=80):
    results = []
    for s in slides:
        text_for_search = s["raw_text"]
        if mode == "exact_phrase" and keyword.lower() in text_for_search.lower():
            results.append({**s, "score": 100})
        elif mode == "exact" and keyword.lower() in text_for_search.lower():
            results.append({**s, "score": 100})
        elif mode == "fuzzy":
            score = fuzz.partial_ratio(keyword.lower(), text_for_search.lower())
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


# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.markdown("### Search Mode")
    search_mode = st.radio("", ["exact_phrase (recommended)", "exact", "fuzzy"], index=0)
    threshold = st.slider("Fuzzy threshold", 60, 100, 85)
    st.markdown("---")
    st.markdown("Made by SKT")


# ---------------- UPLOAD ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("### Upload PPTX / ZIP Files")
uploaded_files = st.file_uploader("", type=["pptx", "zip"], accept_multiple_files=True)
st.markdown("</div>", unsafe_allow_html=True)


# ---------------- PROCESS FILES ----------------
pptx_paths = []
all_slides_text = []

if uploaded_files:
    for uf in uploaded_files:
        temp_path = os.path.join(tempfile.gettempdir(), uf.name)
        with open(temp_path, "wb") as f:
            f.write(uf.read())

        if uf.name.lower().endswith(".pptx"):
            pptx_paths.append(temp_path)
        else:
            pptx_paths.extend(extract_zip_pptx(temp_path))

    # extract text for chatbot
    for p in pptx_paths:
        slides = ppt_to_html_slides(p)
        for s in slides:
            all_slides_text.append({
                "file": os.path.basename(p),
                "slide_no": s["slide_no"],
                "title": s["title"],
                "text": s["raw_text"]
            })


# ---------------- AI CHATBOT FUNCTIONS ----------------

def retrieve_relevant_slides(question, slides, top_n=5):
    scores = []
    for s in slides:
        score = fuzz.partial_ratio(question.lower(), s["text"].lower())
        scores.append((score, s))
    scores.sort(reverse=True, key=lambda x: x[0])
    return [s for score, s in scores[:top_n] if score > 20]


def ask_chatgpt(question, context):
    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system",
             "content": "Answer ONLY using the PPT context provided. Mention slide numbers used."},
            {"role": "user",
             "content": f"Context from PPT:\n{context}\n\nQuestion: {question}"}
        ]
    )
    return response.choices[0].message.content


# ---------------- AI CHAT UI ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
st.markdown("### 🤖 Ask AI (Chat with Your PPT)")
user_question = st.text_input("Ask anything about the PPT:", "", 
                              placeholder="e.g., In which visualisation is this Job used?")
ask_ai_btn = st.button("✨ Ask AI")
st.markdown("</div>", unsafe_allow_html=True)


if ask_ai_btn:
    if not pptx_paths:
        st.error("Please upload PPTX or ZIP files first.")
    elif not user_question.strip():
        st.error("Please enter a question.")
    else:
        with st.spinner("Thinking… Searching slides…"):
            relevant = retrieve_relevant_slides(user_question, all_slides_text, top_n=5)

            if not relevant:
                st.warning("No relevant slide found.")
            else:
                context = ""
                for s in relevant:
                    context += f"\nSlide {s['slide_no']} ({s['file']}):\n{s['text']}\n"

                answer = ask_chatgpt(user_question, context)

                st.markdown("<div class='section-card'>", unsafe_allow_html=True)
                st.markdown("### 🤖 AI Answer")
                st.write(answer)
                st.markdown("</div>", unsafe_allow_html=True)

                st.markdown("### 📌 Relevant Slides Used:")
                for s in relevant:
                    st.markdown(f"**Slide {s['slide_no']} – {s['title']}**")
                    st.write(s["text"][:400] + ("..." if len(s["text"]) > 400 else ""))


# ---------------- KEYWORD SEARCH ----------------
st.markdown("<div class='section-card'>", unsafe_allow_html=True)
keyword = st.text_input("Enter Keyword", "", placeholder="e.g. PSD Manager")
st.markdown("</div>", unsafe_allow_html=True)

search_btn = st.button("🔍 Search")


# ---------------- SEARCH FUNCTION ----------------
results_all = []

if search_btn:
    if not pptx_paths:
        st.error("Please upload PPTX or ZIP files.")
    elif not keyword.strip():
        st.error("Please enter a keyword.")
    else:
        mode_clean = search_mode.split()[0]
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


# ---------------- SHOW RESULTS ----------------
if results_all:
    df = pd.DataFrame(results_all).drop(columns=["HTML"])
    st.markdown("<div class='section-card'>", unsafe_allow_html=True)
    st.markdown("### Search Results")

    def render_table(df):
        table = "<table class='table-container'>"
        table += "<tr>" + "".join(f"<th>{c}</th>" for c in df.columns) + "</tr>"
        for _, row in df.iterrows():
            table += "<tr>" + "".join(f"<td>{html.escape(str(x))}</td>" for x in row) + "</tr>"
        table += "</table>"
        return table

    st.markdown(render_table(df), unsafe_allow_html=True)

    # Download Excel
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    excel_buffer.seek(0)

    st.download_button("⬇ Download Results (Excel)", excel_buffer.getvalue(),
                       "ppt_search_results.xlsx")

    st.markdown("### Slide Previews")
    for r in results_all:
        st.markdown(f"""
        <div style="border: 2px solid #0047FF; border-radius: 18px; padding: 20px;
            margin-bottom: 20px; background: #F7F9FF;
            box-shadow: 0px 6px 18px rgba(0,0,140,0.12);">
            <h4 style='color:#0047FF; margin-bottom:12px;'>{r['File']} — Slide {r['Slide']}</h4>
            {r['HTML']}
        </div>
        """, unsafe_allow_html=True)


# ---------------- FOOTER ----------------
st.markdown("<div class='footer'>Made by SKT</div>", unsafe_allow_html=True)
