# ---------------------------------------------------------
# 🔍 PPT Keyword Search + Auto Excel Save (Full Enhanced Version)
# ---------------------------------------------------------
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os
import pandas as pd
import ipywidgets as widgets
from IPython.display import display
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# --------------------------
# Recursive Text Extraction (now includes charts)
# --------------------------
def extract_text_recursive(shape):
    """Recursively extract text from all shape types, including charts and tables."""
    text = ""

    try:
        # 1️⃣ Grouped shapes → recurse through sub-shapes
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            for s in shape.shapes:
                text += extract_text_recursive(s) + " "

        # 2️⃣ Tables → extract text from all cells
        elif hasattr(shape, "has_table") and shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    text += cell.text.strip() + " "

        # 3️⃣ Charts → extract title, axis labels, legend, and data labels
        elif hasattr(shape, "has_chart") and shape.has_chart:
            chart = shape.chart

            # Chart title
            if chart.has_title and chart.chart_title.has_text_frame:
                text += chart.chart_title.text_frame.text.strip() + " "

            # Axis titles
            if chart.has_category_axis and chart.category_axis.has_title:
                axis_title = chart.category_axis.axis_title
                if axis_title.has_text_frame:
                    text += axis_title.text_frame.text.strip() + " "

            if chart.has_value_axis and chart.value_axis.has_title:
                axis_title = chart.value_axis.axis_title
                if axis_title.has_text_frame:
                    text += axis_title.text_frame.text.strip() + " "

            # Legend entries
            if chart.has_legend:
                for legend_entry in chart.legend.entries:
                    if legend_entry.text.strip():
                        text += legend_entry.text.strip() + " "

            # Series data labels (if text available)
            try:
                for series in chart.series:
                    if series.has_data_labels:
                        for point in series.points:
                            if point.data_label and point.data_label.has_text_frame:
                                text += point.data_label.text_frame.text.strip() + " "
            except Exception:
                pass

        # 4️⃣ Regular text shapes
        elif hasattr(shape, "text") and shape.text.strip():
            text += shape.text.strip() + " "

    except Exception:
        pass

    return text.strip()


# --------------------------
# Extract Text + Visualization Title
# --------------------------
def extract_text_from_pptx(file_path):
    """Extract all text and probable visualization title from each slide."""
    prs = Presentation(file_path)
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
                    if hasattr(shape, "text_frame"):
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if run.font.size:
                                    font_size = max(font_size, run.font.size.pt)
                except Exception:
                    pass

                candidate_titles.append((y_position, font_size, shape_text))

        title_text = ""
        if candidate_titles:
            candidate_titles.sort(key=lambda x: (x[0], -x[1]))
            title_text = candidate_titles[0][2]

        slides_data.append({
            "slide_num": i + 1,
            "title": title_text,
            "text": text.strip()
        })

    return slides_data


# --------------------------
# Search Functions (case-insensitive)
# --------------------------
def search_pptx(file_path, keyword):
    """Return slide numbers and visualization titles containing the keyword (case-insensitive)."""
    key_lower = keyword.lower()
    slides_data = extract_text_from_pptx(file_path)
    results = []

    for slide in slides_data:
        if key_lower in slide["text"].lower():
            results.append({
                "PPT Title": os.path.basename(file_path),
                "PPT Slide No": slide["slide_num"],
                "Visualization Title": slide["title"]
            })
    return results


def search_folder(folder_path, keyword):
    """Search all PPTX files recursively (excluding OUTPUT RESULT) and return results."""
    results = {}
    ppt_files = []

    for root, dirs, files in os.walk(folder_path):
        if "OUTPUT RESULT" in root:
            continue
        for f in files:
            if f.lower().endswith(".pptx"):
                ppt_files.append(os.path.join(root, f))

    for file_path in ppt_files:
        slides = search_pptx(file_path, keyword)
        if slides:
            results[os.path.basename(file_path)] = slides
    return results


# --------------------------
# Widgets UI
# --------------------------
folder_box = widgets.Text(
    placeholder="Enter folder path containing PPTX files",
    description="PPT Folder:"
)

keyword_box = widgets.Text(
    placeholder="Enter word to search",
    description="Keyword:"
)

search_button = widgets.Button(description="Search 🔍")
download_excel_button = widgets.Button(description="Download Excel ⬇️", disabled=True)
output = widgets.Output()

search_results = {}
ppt_folder_global = ""
output_folder_path = ""


# --------------------------
# Helper: Clean Only Files Inside OUTPUT Folder
# --------------------------
def clean_output_folder():
    global output_folder_path, ppt_folder_global
    output_folder_path = os.path.join(ppt_folder_global, "OUTPUT RESULT")

    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)
        with output:
            print(f"📁 Created output folder: {output_folder_path}")
    else:
        for file in os.listdir(output_folder_path):
            file_path = os.path.join(output_folder_path, file)
            try:
                if os.path.isfile(file_path):
                    os.remove(file_path)
            except Exception as e:
                print(f"Error deleting file {file_path}: {e}")
        with output:
            print(f"🧹 Cleaned existing files in: {output_folder_path}")


# --------------------------
# Save to Excel (Auto)
# --------------------------
def save_results_to_dataframe():
    data = []
    for ppt, slides in search_results.items():
        for slide in slides:
            data.append(slide)
    df = pd.DataFrame(data, columns=["PPT Title", "PPT Slide No", "Visualization Title"])
    return df


def save_results_to_excel():
    df = save_results_to_dataframe()
    excel_path = os.path.join(output_folder_path, "ppt_search_results.xlsx")
    df.to_excel(excel_path, index=False)

    # Left-align Excel text
    wb = load_workbook(excel_path)
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal="left")
    wb.save(excel_path)
    return excel_path


# --------------------------
# Button Actions
# --------------------------
def on_search_click(b):
    global search_results, ppt_folder_global
    output.clear_output()
    download_excel_button.disabled = True

    folder_path = folder_box.value.strip()
    keyword = keyword_box.value.strip()
    ppt_folder_global = folder_path

    if not folder_path:
        with output:
            print("⚠️ Please enter the folder path.")
        return
    if not keyword:
        with output:
            print("⚠️ Please enter a keyword.")
        return
    if not os.path.exists(folder_path):
        with output:
            print("❌ Folder path does not exist.")
        return

    clean_output_folder()
    search_results = search_folder(folder_path, keyword)

    with output:
        if search_results:
            print(f'🔍 Found matches for "{keyword}" (case-insensitive):')
            for ppt, slides in search_results.items():
                print(f"\n📄 {ppt}:")
                for item in slides:
                    print(f"   • Slide {item['PPT Slide No']}: {item['Visualization Title']}")

            excel_path = save_results_to_excel()
            print(f"\n✅ Excel automatically saved at: {excel_path}")

            download_excel_button.disabled = False
        else:
            print(f'❌ No results found for "{keyword}".')


def on_download_excel_click(b):
    if search_results:
        excel_path = os.path.join(output_folder_path, "ppt_search_results.xlsx")
        with output:
            print(f"📥 Excel file ready at: {excel_path}")


# --------------------------
# Link Buttons & Display UI
# --------------------------
search_button.on_click(on_search_click)
download_excel_button.on_click(on_download_excel_click)

display(widgets.VBox([
    widgets.HTML("<h3>🔍 PPT Keyword Search Tool (Excel Auto Save + Chart Support)</h3>"),
    folder_box,
    keyword_box,
    search_button,
    download_excel_button,
    output
]))