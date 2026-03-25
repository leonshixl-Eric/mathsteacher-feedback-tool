import sys
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pptx import Presentation
from pptx.util import Cm as PptxCm, Pt as PptxPt
import zipfile
import os
import re
from io import BytesIO
from PIL import Image

# --- FIX FOR PYTHON 3.13 & PYMUPDF ---
try:
    import imghdr
except ImportError:
    import filetype
    class MockImghdr:
        def what(self, file, h=None):
            kind = filetype.guess(file)
            return kind.extension if kind else None
    sys.modules['imghdr'] = MockImghdr()

try:
    import fitz  # PyMuPDF for PDF cropping
except ImportError:
    fitz = None

st.set_page_config(page_title="Maths Feedback Pro", layout="centered", page_icon="📊")

st.title("📊 High-Fidelity Feedback Generator")
st.write("Auto-crops your original exam paper and allows you to manually fine-tune the images before generating.")

# --- 1. THE UPLOADERS ---
uploaded_csv = st.file_uploader("1. Upload Marks (CSV or Excel)", type=["csv", "xlsx"])
uploaded_pdf = st.file_uploader("2. Upload Original Exam PDF (To crop the questions)", type="pdf")
uploaded_mapping = st.file_uploader("3. Upload Topic Mapping (CSV or Excel)", type=["csv", "xlsx"])

# --- BRANDING SETTINGS ---
st.markdown("---")
st.subheader("📝 Document Branding")
col_brand1, col_brand2 = st.columns(2)
with col_brand1:
    unit_title = st.text_input("Unit/Topic Title", value="Algebraic Manipulation")
    class_name = st.text_input("Class Name", value="9y2 Maths")
with col_brand2:
    uploaded_logo = st.file_uploader("Upload School Logo (Optional)", type=["png", "jpg", "jpeg"])

# --- DOCUMENT SETTINGS ---
st.markdown("---")
st.subheader("⚙️ Document Settings")

col_setting1, col_setting2, col_setting3 = st.columns(3)
with col_setting1:
    selected_threshold = st.slider("Reteach Threshold (%)", min_value=0, max_value=100, value=55, step=5)
with col_setting2:
    selected_margin = st.slider("Page Margin (cm)", min_value=0.5, max_value=3.0, value=1.3, step=0.1)
with col_setting3:
    generate_ppt = st.checkbox("📽️ Generate Reteach PowerPoint (PPTX)", value=True)
st.markdown("---")

threshold_decimal = selected_threshold / 100.0

# --- 2. DATA PROCESSING ---
def process_data(uploaded_csv, uploaded_mapping):
    df_marks = pd.read_csv(uploaded_csv, header=None) if uploaded_csv.name.endswith('.csv') else pd.read_excel(uploaded_csv, header=None)
    row0, row1 = df_marks.iloc[0].astype(str).tolist(), df_marks.iloc[1].astype(str).tolist()
    q_labels = ["Surname", "Forename"]
    current_q = ""
    for i in range(2, len(row0)):
        r0, r1 = row0[i].strip(), row1[i].strip()
        if 'total' in r0.lower() or 'total' in r1.lower(): break
        if r0 != 'nan' and r0 != '':
            m = re.search(r'\d+', r0)
            if m: current_q = m.group()
        q_labels.append((current_q + r1) if r1 != 'nan' and r1 != '' else current_q)

    full_marks_row = df_marks.iloc[2]
    percentage_idx = next(i for i in range(len(df_marks)) if 'percentage' in str(df_marks.iloc[i, 0]).lower())
    percentage_row = df_marks.iloc[percentage_idx]
    
    # Skip Absent Students
    raw_students = df_marks.iloc[3:percentage_idx].dropna(subset=[0, 1], how='all')
    student_rows = raw_students[raw_students.iloc[:, 2:len(q_labels)].notnull().any(axis=1)].reset_index(drop=True)
    
    df_map = pd.read_csv(uploaded_mapping, header=None) if uploaded_mapping.name.endswith('.csv') else pd.read_excel(uploaded_mapping, header=None)
    if 'topic' in str(df_map.iloc[0, 0]).lower(): df_map = df_map.iloc[1:]

    dynamic_areas = []
    for _, map_row in df_map.iterrows():
        if pd.isna(map_row.iloc[0]): continue
        topic, qs, last_num = str(map_row.iloc[0]).strip(), [], ""
        for cell in map_row.iloc[1:]:
            if pd.isna(cell): continue
            for t in str(cell).lower().replace('and', ',').replace('&', ',').split(','):
                t = t.strip()
                n = "".join([c for c in t if c.isdigit()])
                l = "".join([c for c in t if c.isalpha()])
                if n: last_num = n
                cand = (n or last_num) + l
                if cand in q_labels and cand not in qs: qs.append(cand)
        idxs = [q_labels.index(q) for q in qs]
        if idxs: dynamic_areas.append((topic, idxs))
        
    return student_rows, percentage_row, full_marks_row, q_labels, dynamic_areas

# --- 3. PDF SCANNER & CROPPER ---
def find_base_crops(pdf_file, required_qs):
    """Auto-guesses the location of the questions to save the teacher time."""
    base_crops = {}
    if not pdf_file or not fitz: return base_crops
    
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    pdf_file.seek(0)
    
    current_num = None
    for page_num in range(len(doc)):
        page = doc[page_num]
        blocks = page.get_text("blocks")
        blocks = [b for b in blocks if (b[6] == 0 and str(b[4]).strip()) or b[6] == 1]
        blocks.sort(key=lambda b: b[1])
        
        for idx, b in enumerate(blocks):
            if b[6] == 1: continue
            text = str(b[4]).strip()
            if not text: continue
            
            num_match = re.match(r"^\s*(?:Question\s+|Q)?(\d+)", text, re.IGNORECASE)
            if num_match: current_num = num_match.group(1)
                
            for q in required_qs:
                if q in base_crops: continue
                m = re.match(r"(\d+)([a-zA-Z]*)", q)
                if not m: continue
                num, let = m.groups()
                
                full_pat = re.compile(rf"(?:^|\n)\s*(?:Question\s+|Q)?0*{num}\s*[\(\.\-]?\s*{let}\s*(?:[\)\.\-\s]|$)", re.IGNORECASE)
                if full_pat.search(text) or (current_num == num and let and re.compile(rf"(?:^|\n)\s*[\(\.]?\s*{let}\s*(?:[\)\.\-\s]|$)", re.IGNORECASE).search(text)):
                    
                    y0 = max(0, b[1] - 15)
                    # Guess the bottom of the question (approx 120 pixels down)
                    y1 = min(page.rect.height, b[1] + 120)
                    base_crops[q] = {"page": page_num, "y0": y0, "y1": y1, "width": page.rect.width, "height": page.rect.height}
                    break
    return base_crops

def get_adjusted_crop(pdf_file, q_code, base_crop, top_adj, bot_adj):
    """Renders the actual image based on the teacher's slider adjustments."""
    if not pdf_file or not fitz or not base_crop: return None
    
    pdf_file.seek(0)
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    page = doc[base_crop["page"]]
    
    # Apply the user's manual adjustments (+ expands the crop, - shrinks it)
    final_y0 = max(0, base_crop["y0"] - top_adj)
    final_y1 = min(base_crop["height"], base_crop["y1"] + bot_adj)
    
    # Ensure y0 is not greater than y1
    if final_y0 >= final_y1: final_y0 = final_y1 - 10
        
    rect = fitz.Rect(0, final_y0, base_crop["width"], final_y1)
    pix = page.get_pixmap(clip=rect, dpi=150) # Use 150 DPI for preview to keep it fast
    
    img_name = f"crop_{q_code}.png"
    pix.save(img_name)
    return img_name

def add_tight_picture(doc, img_path, max_width_cm):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Cm(0.3)
    paragraph.paragraph_format.space_after = Cm(0.3)
    run = paragraph.add_run()
    try:
        with Image.open(img_path) as img:
            dpi = 150.0 
            img_width_cm = (img.size[0] / dpi) * 2.54
            final_width = min(img_width_cm, max_width_cm.inches * 2.54) 
    except:
        final_width = 14.0 
    run.add_picture(img_path, width=Cm(final_width))
    return paragraph

# --- 4. STATE MANAGEMENT & UI FLOW ---
if "step" not in st.session_state:
    st.session_state.step = 1

if "base_crops" not in st.session_state:
    st.session_state.base_crops = {}

if "adjustments" not in st.session_state:
    st.session_state.adjustments = {}

if uploaded_csv and uploaded_pdf and uploaded_mapping:
    student_rows, percentage_row, full_marks_row, q_labels, dynamic_areas = process_data(uploaded_csv, uploaded_mapping)
    
    # Calculate exactly which questions we actually need images for to save time
    required_qs = set()
    for _, row in student_rows.iterrows():
        for _, idxs in dynamic_areas:
            for idx in idxs:
                if pd.to_numeric(row[idx], errors='coerce') < pd.to_numeric(full_marks_row[idx], errors='coerce'):
                    required_qs.add(q_labels[idx])
    
    # Sort the required questions logically
    required_qs = sorted(list(required_qs), key=lambda x: q_labels.index(x))

    if st.session_state.step == 1:
        if st.button("1. Analyze PDF & Prepare Crops", type="primary", use_container_width=True):
            with st.spinner("Finding questions in the PDF..."):
                st.session_state.base_crops = find_base_crops(uploaded_pdf, required_qs)
                
                # Initialize sliders to 0
                for q in required_qs:
                    st.session_state.adjustments[q] = {"top": 0, "bot": 0}
                    
                st.session_state.step = 2
                st.rerun()

    if st.session_state.step == 2:
        st.success("✅ PDF Analyzed! Review and manually adjust your crops below.")
        st.info("💡 **Pro Tip:** Use the sliders to grab missing diagrams or cut out blank answering space.")
        
        # Display the Crop Editor for each required question
        for q in required_qs:
            st.markdown(f"### Question: **{q}**")
            
            if q in st.session_state.base_crops:
                col_img, col_sliders = st.columns([2, 1])
                
                with col_sliders:
                    # Sliders update the session state directly
                    st.session_state.adjustments[q]["top"] = st.slider(f"Adjust Top (px)", -100, 200, st.session_state.adjustments[q]["top"], key=f"top_{q}")
                    st.session_state.adjustments[q]["bot"] = st.slider(f"Adjust Bottom (px)", -200, 400, st.session_state.adjustments[q]["bot"], key=f"bot_{q}")
                
                with col_img:
                    # Render the live preview of the crop
                    img_path = get_adjusted_crop(
                        uploaded_pdf, 
                        q, 
                        st.session_state.base_crops[q], 
                        st.session_state.adjustments[q]["top"], 
                        st.session_state.adjustments[q]["bot"]
                    )
                    st.image(img_path, caption=f"Live Preview: {q}", use_container_width=True)
            else:
                st.error(f"Could not auto-locate Question {q}. It will be skipped.")
                
            st.markdown("---")

        if st.button("2. Generate Final Feedback Documents", type="primary", use_container_width=True):
            with st.spinner(f"Generating reports for {len(student_rows)} students..."):
                logo_path = None
                if uploaded_logo:
                    logo_path = "temp_logo.png"
                    with open(logo_path, "wb") as f: f.write(uploaded_logo.getbuffer())

                doc = Document()
                for section in doc.sections:
                    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(selected_margin)

                for _, row in student_rows.iterrows():
                    name = f"{row[1]} {row[0]}"
                    header = doc.add_paragraph()
                    if logo_path: header.add_run().add_picture(logo_path, width=Cm(1.5)); header.add_run("    ")
                    title_run = header.add_run(f"{unit_title} Feedback: {name}   |   Class: {class_name}")
                    title_run.bold, title_run.font.size = True, Pt(14)
                    
                    table = doc.add_table(rows=1, cols=3)
                    table.style = 'Table Grid'
                    hdr = table.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text = "Area", "What Went Well", "Even Better If"
                    
                    student_ebi = []
                    for title, idxs in dynamic_areas:
                        w, e = [], []
                        for idx in idxs:
                            if pd.to_numeric(row[idx], errors='coerce') >= pd.to_numeric(full_marks_row[idx], errors='coerce'):
                                w.append(q_labels[idx])
                            else: e.append(q_labels[idx]); student_ebi.append(q_labels[idx])
                        r = table.add_row().cells
                        r[0].text, r[1].text, r[2].text = title, ", ".join(w), ", ".join(e)

                    reteach_qs = [q for q in student_ebi if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= threshold_decimal]
                    personal_qs = [q for q in student_ebi if q not in reteach_qs]
                    
                    if personal_qs:
                        doc.add_heading("Personal correction", 2)
                        for q in personal_qs: 
                            if q in st.session_state.base_crops:
                                img_path = f"crop_{q}.png"
                                add_tight_picture(doc, img_path, Cm(14))
                    
                    doc.add_page_break()
                    doc.add_heading(f"Whole-class reteaching - {name}", 1)
                    if reteach_qs:
                        for q in reteach_qs: 
                            if q in st.session_state.base_crops:
                                img_path = f"crop_{q}.png"
                                add_tight_picture(doc, img_path, Cm(15))
                    else: doc.add_paragraph("Excellent mastery of class-wide topics.")
                    doc.add_page_break()

                target_docx = BytesIO()
                doc.save(target_docx)
                
                prs = Presentation()
                global_reteach = [q for q in q_labels[2:] if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= threshold_decimal]
                for q in global_reteach:
                    if q in st.session_state.base_crops:
                        slide = prs.slides.add_slide(prs.slide_layouts[6])
                        img_path = f"crop_{q}.png"
                        try:
                            with Image.open(img_path) as img: aspect = img.size[0] / img.size[1]
                        except: aspect = 2.0
                        if aspect > 1.5: slide.shapes.add_picture(img_path, PptxCm(2), PptxCm(2.5), width=PptxCm(21.4))
                        else: slide.shapes.add_picture(img_path, PptxCm(2), PptxCm(2.5), height=PptxCm(13.0))
                target_pptx = BytesIO()
                prs.save(target_pptx)

                zip_buffer = BytesIO()
                safe_name = str(class_name).replace(" ", "_")
                with zipfile.ZipFile(zip_buffer, "w") as z:
                    z.writestr(f"{safe_name}_Reports.docx", target_docx.getvalue())
                    if global_reteach: z.writestr(f"{safe_name}_Reteach.pptx", target_pptx.getvalue())
                
                st.success(f"✅ Feedback Pack Ready!")
                st.download_button("📦 Download All (ZIP)", zip_buffer.getvalue(), file_name=f"{safe_name}_Pack.zip", type="primary")
                
                # Cleanup
                if logo_path and os.path.exists(logo_path): os.remove(logo_path)
                for f in os.listdir():
                    if f.startswith("crop_") and f.endswith(".png"): os.remove(f)

        if st.button("Restart & Upload New Files"):
            st.session_state.step = 1
            st.rerun()

else:
    st.info("Please upload all three files (Marks, PDF, and Mapping) to begin.")
