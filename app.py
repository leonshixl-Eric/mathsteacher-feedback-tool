import sys
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from pptx import Presentation
from pptx.util import Cm as PptxCm, Pt as PptxPt
import zipfile
import os
import re
from io import BytesIO
from PIL import Image

# --- FIX FOR PYTHON 3.13 & IMAGE HANDLING ---
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
    import fitz # PyMuPDF
except ImportError:
    st.error("PyMuPDF not found. Please ensure 'pymupdf' is in requirements.txt")

from streamlit_cropper import st_cropper 

st.set_page_config(page_title="Maths Feedback Pro", layout="wide", page_icon="📊")

st.title("📊 High-Fidelity Feedback Generator")
st.write("Draw a box to crop questions. Change page numbers to navigate the PDF.")

# --- 1. UPLOADERS ---
col_u1, col_u2, col_u3 = st.columns(3)
with col_u1:
    uploaded_csv = st.file_uploader("1. Upload Marks (CSV/Excel)", type=["csv", "xlsx"])
with col_u2:
    uploaded_pdf = st.file_uploader("2. Upload Original Exam PDF", type="pdf")
with col_u3:
    uploaded_mapping = st.file_uploader("3. Upload Topic Mapping (CSV/Excel)", type=["csv", "xlsx"])

# --- SIDEBAR SETTINGS ---
st.sidebar.header("📝 Branding & Settings")
unit_title = st.sidebar.text_input("Unit Title", value="Algebraic Manipulation")
class_name = st.sidebar.text_input("Class Name", value="9y2")
uploaded_logo = st.sidebar.file_uploader("Upload Logo", type=["png", "jpg"])
selected_threshold = st.sidebar.slider("Reteach Threshold (%)", 0, 100, 55)
page_margin = st.sidebar.slider("Page Margin (cm)", 0.5, 3.0, 1.3)
threshold_decimal = selected_threshold / 100.0

# --- 2. CORE FUNCTIONS ---
def process_data(csv, mapping):
    df = pd.read_csv(csv, header=None) if csv.name.endswith('.csv') else pd.read_excel(csv, header=None)
    row0, row1 = df.iloc[0].astype(str).tolist(), df.iloc[1].astype(str).tolist()
    q_labels = ["Surname", "Forename"]
    curr = ""
    for i in range(2, len(row0)):
        r0, r1 = row0[i].strip(), row1[i].strip()
        if 'total' in r0.lower() or 'total' in r1.lower(): break
        if r0 != 'nan' and r0 != '':
            m = re.search(r'\d+', r0)
            if m: curr = m.group()
        q_labels.append((curr + r1) if r1 != 'nan' and r1 != '' else curr)

    perc_idx = next(i for i in range(len(df)) if 'percentage' in str(df.iloc[i, 0]).lower())
    perc_row = df.iloc[perc_idx]
    full_marks = df.iloc[2]
    
    students = df.iloc[3:perc_idx].dropna(subset=[0, 1], how='all')
    students = students[students.iloc[:, 2:len(q_labels)].notnull().any(axis=1)].reset_index(drop=True)
    
    df_map = pd.read_csv(mapping, header=None) if mapping.name.endswith('.csv') else pd.read_excel(mapping, header=None)
    if 'topic' in str(df_map.iloc[0, 0]).lower(): df_map = df_map.iloc[1:]
    
    dyn_areas = []
    for _, m_row in df_map.iterrows():
        if pd.isna(m_row.iloc[0]): continue
        topic, qs, last_n = str(m_row.iloc[0]).strip(), [], ""
        for cell in m_row.iloc[1:]:
            if pd.isna(cell): continue
            for t in str(cell).lower().replace('and', ',').replace('&', ',').split(','):
                t = t.strip()
                n, l = "".join([c for c in t if c.isdigit()]), "".join([c for c in t if c.isalpha()])
                if n: last_n = n
                cand = (n or last_n) + l
                if cand in q_labels and cand not in qs: qs.append(cand)
        idxs = [q_labels.index(q) for q in qs]
        if idxs: dyn_areas.append((topic, idxs))
        
    return students, perc_row, full_marks, q_labels, dyn_areas

def scan_pdf_for_metadata(pdf_bytes, required_qs):
    """One-pass scanner to find BOTH the page number and the instruction for each question."""
    pages_dict = {}
    titles_dict = {}
    page_count = 1
    
    if fitz:
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            page_count = max(1, len(doc))
            
            for page_num in range(len(doc)):
                text = doc[page_num].get_text("text")
                clean_text = text.lower().replace(" ", "")
                
                for q in required_qs:
                    m = re.match(r"(\d+)([a-zA-Z]*)", q)
                    if not m: continue
                    num, let = m.groups()
                    
                    # 1. Find the Page
                    if q not in pages_dict:
                        if let:
                            if f"{num}{let})" in clean_text or f"{num}({let})" in clean_text or f"{num}.{let}" in clean_text:
                                pages_dict[q] = page_num
                        else:
                            if f"question{num}" in clean_text or f"q{num}" in clean_text or f"\n{num})" in clean_text:
                                pages_dict[q] = page_num
                                
                    # 2. Find the Instruction Title
                    if q not in titles_dict:
                        pat = rf"^(?:Question\s+|Q)?0*{num}\s*[\.\-\)]?\s*([A-Za-z].*)"
                        for line in text.split('\n'):
                            match = re.search(pat, line.strip(), re.IGNORECASE)
                            if match:
                                instr = match.group(1).strip()
                                instr = re.sub(r'\s*\[\d+.*\]', '', instr, flags=re.IGNORECASE)
                                titles_dict[q] = f"Question {q}) {instr}"
        except:
            pass
            
    # Fallbacks if scanner missed anything
    for q in required_qs:
        if q not in pages_dict: pages_dict[q] = 0
        if q not in titles_dict: titles_dict[q] = f"Question {q}"
        
    return pages_dict, titles_dict, page_count

def get_page_img(pdf_bytes, p_num):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[p_num]
    pix = page.get_pixmap(dpi=150)
    return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

# --- 3. APP FLOW ---
if "step" not in st.session_state: st.session_state.step = 1
if "q_pages" not in st.session_state: st.session_state.q_pages = {}
if "q_titles" not in st.session_state: st.session_state.q_titles = {}
if "saved_crops" not in st.session_state: st.session_state.saved_crops = {}
if "pdf_pages" not in st.session_state: st.session_state.pdf_pages = 1

if uploaded_csv and uploaded_pdf and uploaded_mapping:
    students, perc_row, full_marks, q_labels, dyn_areas = process_data(uploaded_csv, uploaded_mapping)
    req_qs = sorted(list({q_labels[i] for i in range(2, len(q_labels))}), key=lambda x: q_labels.index(x))
    pdf_b = uploaded_pdf.getvalue()

    if st.session_state.step == 1:
        if st.button("🚀 Scan PDF & Open Cropper", use_container_width=True):
            with st.spinner("Analyzing PDF pages and extracting instructions..."):
                pages_dict, titles_dict, total_pages = scan_pdf_for_metadata(pdf_b, req_qs)
                st.session_state.q_pages = pages_dict
                st.session_state.q_titles = titles_dict
                st.session_state.pdf_pages = total_pages
                st.session_state.step = 2
                st.rerun()

    if st.session_state.step == 2:
        st.success(f"Loaded {len(req_qs)} questions. Save each crop below.")
        
        for q in req_qs:
            st.markdown(f"### ✂️ Question {q}")
            
            # 1. Editable Label
            st.session_state.q_titles[q] = st.text_input(f"Document Label for {q}", value=st.session_state.q_titles.get(q, ""), key=f"t_{q}")
            
            # 2. Safely initialize the Page Number Widget state so it maps correctly
            widget_key = f"page_input_{q}"
            if widget_key not in st.session_state:
                st.session_state[widget_key] = st.session_state.q_pages.get(q, 0) + 1
            
            col_a, col_b = st.columns([2, 1])
            
            with col_a:
                # Page Selector
                p_val = st.number_input("Select Page", min_value=1, max_value=st.session_state.pdf_pages, key=widget_key)
                page_idx = p_val - 1
                
                st.caption(f"👀 Now viewing Page {p_val}")
                
                # Fetch image and apply dynamic key to force unmount on page change
                img = get_page_img(pdf_b, page_idx)
                cropper_key = f"cropper_{q}_pg_{page_idx}"
                cropped = st_cropper(img, realtime_update=True, box_color='#FF0000', aspect_ratio=None, key=cropper_key)
                
            with col_b:
                st.write("Live Preview:")
                if cropped:
                    st.image(cropped, use_container_width=True)
                    if st.button(f"✅ Save Crop for {q}", key=f"save_{q}"):
                        buf = BytesIO()
                        cropped.save(buf, format="PNG")
                        st.session_state.saved_crops[q] = buf.getvalue()
                        st.toast(f"Question {q} Saved!", icon="✅")
            
            st.divider()

        if st.button("📦 Generate All Documents (Word & PPTX)", type="primary", use_container_width=True):
            if len(st.session_state.saved_crops) < len(req_qs):
                st.error(f"Please click 'Save Crop' for all questions first! ({len(st.session_state.saved_crops)}/{len(req_qs)})")
            else:
                progress_bar = st.progress(0)
                doc = Document()
                prs = Presentation()
                
                for section in doc.sections:
                    section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Cm(page_margin)

                logo_p = None
                if uploaded_logo:
                    logo_p = "temp_logo.png"
                    with open(logo_p, "wb") as f: f.write(uploaded_logo.getbuffer())

                for i, (_, s_row) in enumerate(students.iterrows()):
                    name = f"{s_row[1]} {s_row[0]}"
                    
                    head = doc.add_paragraph()
                    if logo_p: head.add_run().add_picture(logo_p, width=Cm(1.5))
                    title = head.add_run(f"  {unit_title} Feedback: {name} | Class: {class_name}")
                    title.bold, title.font.size = True, Pt(14)
                    
                    table = doc.add_table(rows=1, cols=3)
                    table.style = 'Table Grid'
                    hdr = table.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text = "Area", "What Went Well", "Even Better If"
                    
                    s_ebi = []
                    for t_name, idxs in dyn_areas:
                        w, e = [], []
                        for idx in idxs:
                            score = pd.to_numeric(s_row[idx], errors='coerce')
                            full = pd.to_numeric(full_marks[idx], errors='coerce')
                            if score >= full: w.append(q_labels[idx])
                            else: 
                                e.append(q_labels[idx])
                                s_ebi.append(q_labels[idx])
                        r = table.add_row().cells
                        r[0].text, r[1].text, r[2].text = t_name, ", ".join(w), ", ".join(e)

                    reteach_qs = [q for q in s_ebi if pd.to_numeric(perc_row[q_labels.index(q)], errors='coerce') <= threshold_decimal]
                    personal_qs = [q for q in s_ebi if q not in reteach_qs]
                    
                    if personal_qs:
                        doc.add_heading("Personal correction", 2)
                        for q in personal_qs:
                            doc.add_paragraph().add_run(st.session_state.q_titles[q]).bold = True
                            img_data = BytesIO(st.session_state.saved_crops[q])
                            doc.add_paragraph().add_run().add_picture(img_data, width=Cm(14))

                    doc.add_page_break()
                    doc.add_heading(f"Whole-class reteaching - {name}", 1)
                    if reteach_qs:
                        for q in reteach_qs:
                            doc.add_paragraph().add_run(st.session_state.q_titles[q]).bold = True
                            img_data = BytesIO(st.session_state.saved_crops[q])
                            doc.add_paragraph().add_run().add_picture(img_data, width=Cm(15))
                    doc.add_page_break()
                    progress_bar.progress((i + 1) / len(students))

                global_reteach = [q for q in q_labels[2:] if pd.to_numeric(perc_row[q_labels.index(q)], errors='coerce') <= threshold_decimal]
                for q in global_reteach:
                    slide = prs.slides.add_slide(prs.slide_layouts[6])
                    txBox = slide.shapes.add_textbox(PptxCm(2), PptxCm(1), PptxCm(20), PptxCm(1.5))
                    txBox.text_frame.paragraphs[0].text = st.session_state.q_titles[q]
                    img_data = BytesIO(st.session_state.saved_crops[q])
                    slide.shapes.add_picture(img_data, PptxCm(2), PptxCm(3), width=PptxCm(21))

                word_buf, ppt_buf, zip_buf = BytesIO(), BytesIO(), BytesIO()
                doc.save(word_buf)
                prs.save(ppt_buf)
                
                with zipfile.ZipFile(zip_buf, "w") as zf:
                    zf.writestr(f"{class_name}_Feedback.docx", word_buf.getvalue())
                    zf.writestr(f"{class_name}_Reteach.pptx", ppt_buf.getvalue())
                
                st.success("✅ Pack Ready!")
                st.download_button("📦 Download Feedback Pack (ZIP)", zip_buf.getvalue(), f"{class_name}_Pack.zip", type="primary")

else:
    st.info("Please upload all three files to begin.")
