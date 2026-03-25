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
import textwrap
from io import BytesIO

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
    import fitz  # PyMuPDF for reading text only (no cropping!)
except ImportError:
    fitz = None

st.set_page_config(page_title="Maths Feedback Pro", layout="centered", page_icon="📊")

st.title("📊 High-Fidelity Feedback Generator")
st.write("Auto-reads questions from your PDF and reconstructs them with beautiful math formatting. Absent students are skipped.")

# --- 1. THE UPLOADERS ---
uploaded_csv = st.file_uploader("1. Upload Marks (CSV or Excel)", type=["csv", "xlsx"])
uploaded_pdf = st.file_uploader("2. Upload Original Exam PDF (To read the text)", type="pdf")
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
    selected_font_size = st.slider("Question Font Size", min_value=10, max_value=14, value=12, step=1)
with col_setting2:
    selected_threshold = st.slider("Reteach Threshold (%)", min_value=0, max_value=100, value=55, step=5)
with col_setting3:
    selected_margin = st.slider("Page Margin (cm)", min_value=0.5, max_value=3.0, value=1.3, step=0.1)

generate_ppt = st.checkbox("📽️ Also generate Whole-Class Reteach PowerPoint (PPTX)", value=True)
st.markdown("---")

threshold_decimal = selected_threshold / 100.0

# --- 2. THE RECONSTRUCTION ENGINE ---

def build_dynamic_db(pdf_file, q_labels):
    """Reads the PDF text, matches it to your questions, and applies Math formatting."""
    db = {}
    valid_qs = [q for q in q_labels if q not in ["Surname", "Forename"]]
    
    if not pdf_file or not fitz:
        return {q: f"Question {q} (PDF missing)" for q in valid_qs}
        
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    pdf_file.seek(0)
    
    current_q = None
    current_text = []
    
    for page in doc:
        blocks = page.get_text("blocks")
        blocks.sort(key=lambda b: b[1])
        
        for b in blocks:
            if b[6] == 1: continue  # Completely ignore images/diagrams
            text = b[4].strip()
            if not text: continue
            
            found_q = None
            for q in valid_qs:
                m = re.match(r"(\d+)([a-zA-Z]*)", q)
                if not m: continue
                num, let = m.groups()
                
                # Check if this text block looks like the start of a question
                if let:
                    if re.search(rf"^\s*(?:Question\s+|Q)?{num}\s*[\.\-\)]?\s*\(?{let}\)?", text, re.IGNORECASE):
                        found_q = q; break
                    elif current_q and re.match(r"(\d+)", current_q).group(1) == num and re.search(rf"^\s*\(?{let}\)?", text, re.IGNORECASE):
                        found_q = q; break
                else:
                    if re.search(rf"^\s*(?:Question\s+|Q)?{num}\s*[\.\-\)]", text, re.IGNORECASE):
                        found_q = q; break
                        
            if found_q:
                # Save the previous question's text before starting a new one
                if current_q and current_text:
                    db[current_q] = "\n".join(current_text)
                current_q = found_q
                current_text = [text]
            elif current_q:
                # Stop grabbing text if we hit a marks indicator (e.g., "[2]")
                if re.search(r"^\[\d+\]$|^\(\d+\s*marks?\)$|total.*marks", text, re.IGNORECASE):
                    db[current_q] = "\n".join(current_text)
                    current_q = None
                    current_text = []
                else:
                    current_text.append(text)
                    
    # Catch the very last question on the paper
    if current_q and current_text:
        db[current_q] = "\n".join(current_text)
        
    # Format the extracted text beautifully
    for q in valid_qs:
        if q in db:
            raw = db[q]
            # Fix PDF line breaks
            raw = re.sub(r'(?<![\.\?\!\:])\n(?!\n)', ' ', raw)
            # Math Auto-Fixer: Converts x2 to $x^2$, y3 to $y^3$ for the math engine
            clean = re.sub(r'\b([a-zA-Z])([2345])\b', r'$\1^\2$', raw)
            
            # Ensure the question number is clearly stated at the start
            if not clean.lower().startswith(q.lower()):
                clean = f"{q}) {clean}"
                
            db[q] = textwrap.fill(clean, width=65)
        else:
            db[q] = f"Question {q} (Could not locate in PDF text)"
            
    return db

def create_text_image(q_code, font_size, dynamic_db):
    """Generates a pure text, high-fidelity snippet using Matplotlib."""
    text = dynamic_db.get(q_code, f"Question {q_code}")
    line_count = text.count('\n') + 1
    base_padding = 0.24 
    height_per_line = font_size * 0.035 
    fig_height = base_padding + (line_count * height_per_line)
    
    plt.figure(figsize=(7, fig_height))
    plt.text(0.01, 0.5, text, fontsize=font_size, verticalalignment='center', fontfamily='serif')
    plt.axis('off')
    plt.tight_layout(pad=0)
    
    img_name = f"txt_{q_code}.png"
    plt.savefig(img_name, dpi=200, bbox_inches='tight')
    plt.close()
    return img_name

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
    
    # Filter out absent students (blank marks)
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

def add_tight_picture(doc, img_path, width):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Cm(0.3)
    paragraph.paragraph_format.space_after = Cm(0.3)
    run = paragraph.add_run()
    run.add_picture(img_path, width=width)
    return paragraph

# --- 3. UI LOGIC ---
col1, col2 = st.columns(2)
with col1: preview_clicked = st.button("👀 Preview Sample Student", use_container_width=True)
with col2: generate_clicked = st.button("📄 Generate All Feedback", type="primary", use_container_width=True)

if preview_clicked or generate_clicked:
    if not (uploaded_csv and uploaded_pdf and uploaded_mapping):
        st.error("Missing Marks, PDF, or Mapping files.")
    else:
        try:
            student_rows, percentage_row, full_marks_row, q_labels, dynamic_areas = process_data(uploaded_csv, uploaded_mapping)
            
            if student_rows.empty:
                st.warning("No students with marks were found.")
            else:
                with st.spinner("Reading Text from PDF..."):
                    dynamic_db = build_dynamic_db(uploaded_pdf, q_labels)

                if preview_clicked:
                    row = student_rows.iloc[0]
                    st.markdown(f"### Preview Feedback: **{row[1]} {row[0]}**")
                    student_ebi = []
                    preview_data = []
                    for title, idxs in dynamic_areas:
                        w, e = [], []
                        for idx in idxs:
                            if pd.to_numeric(row[idx], errors='coerce') >= pd.to_numeric(full_marks_row[idx], errors='coerce'):
                                w.append(q_labels[idx])
                            else: e.append(q_labels[idx]); student_ebi.append(q_labels[idx])
                        preview_data.append({"Topic": title, "What Went Well": ", ".join(w), "Even Better If": ", ".join(e)})
                    st.table(pd.DataFrame(preview_data))

                    st.markdown("#### Preview Reconstructed Questions (EBI)")
                    for q in student_ebi[:3]:
                        img = create_text_image(q, selected_font_size, dynamic_db)
                        st.image(img)
                        os.remove(img)

                if generate_clicked:
                    with st.spinner(f"Generating pack for {len(student_rows)} students..."):
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
                                    img = create_text_image(q, selected_font_size, dynamic_db)
                                    add_tight_picture(doc, img, width=Cm(14))
                                    os.remove(img)
                            
                            doc.add_page_break()
                            doc.add_heading(f"Whole-class reteaching - {name}", 1)
                            if reteach_qs:
                                for q in reteach_qs: 
                                    img = create_text_image(q, selected_font_size, dynamic_db)
                                    add_tight_picture(doc, img, width=Cm(15))
                                    os.remove(img)
                            else: doc.add_paragraph("Excellent mastery of class-wide topics.")
                            doc.add_page_break()

                        target_docx = BytesIO()
                        doc.save(target_docx)
                        
                        prs = Presentation()
                        global_reteach = [q for q in q_labels[2:] if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= threshold_decimal]
                        for q in global_reteach:
                            slide = prs.slides.add_slide(prs.slide_layouts[6])
                            txt_img = create_text_image(q, selected_font_size, dynamic_db)
                            slide.shapes.add_picture(txt_img, PptxCm(1), PptxCm(2), width=PptxCm(20))
                            os.remove(txt_img)
                        target_pptx = BytesIO()
                        prs.save(target_pptx)

                        zip_buffer = BytesIO()
                        safe_name = str(class_name).replace(" ", "_")
                        with zipfile.ZipFile(zip_buffer, "w") as z:
                            z.writestr(f"{safe_name}_Reports.docx", target_docx.getvalue())
                            if global_reteach: z.writestr(f"{safe_name}_Reteach.pptx", target_pptx.getvalue())
                        
                        st.success(f"✅ Feedback Pack Ready!")
                        st.download_button("📦 Download All (ZIP)", zip_buffer.getvalue(), file_name=f"{safe_name}_Pack.zip", type="primary")
                        if logo_path and os.path.exists(logo_path): os.remove(logo_path)

        except Exception as e:
            st.error(f"Error: {e}")
