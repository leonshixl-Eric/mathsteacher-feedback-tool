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

# --- FIX FOR PYTHON 3.13 ---
try:
    import imghdr
except ImportError:
    import filetype
    class MockImghdr:
        def what(self, file, h=None):
            kind = filetype.guess(file)
            return kind.extension if kind else None
    sys.modules['imghdr'] = MockImghdr()

st.set_page_config(page_title="Maths Feedback Pro", layout="centered", page_icon="📊")

st.title("📊 High-Fidelity Feedback Generator")
st.write("Generating structured feedback with WWW/EBI tables and reconstructed math snippets.")

# --- 1. THE UPLOADERS ---
uploaded_csv = st.file_uploader("1. Upload Marks (CSV or Excel)", type=["csv", "xlsx"])
uploaded_mapping = st.file_uploader("2. Upload Topic Mapping (CSV or Excel)", type=["csv", "xlsx"])

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

# --- 2. THE RECONSTRUCTION DATABASE ---
questions_db = {
    "1a": r"1a) Expand $3(x + 5)$",
    "1b": r"1b) Expand $4(2y - 3)$",
    "1c": r"1c) Expand $a(a + 4)$",
    "2a": r"2a) Factorise $6x + 9$",
    "2b": r"2b) Factorise $10y - 15$",
    "3":  r"3) Expand and simplify $(x + 3)(x + 5)$",
    "4":  r"4) Expand and simplify $(2x + 1)(x + 4)$",
    "5":  r"5) Factorise $x^2 + 7x + 10$",
    "6":  r"6) Expand and simplify $3(x + 2) + 2(x - 1)$",
    "7":  r"7) Solve $x^2 + 5x + 6 = 0$",
    "8":  r"8) The length of a rectangle is $(x+5)$ and the width is $(x+2)$." + "\n" + r"   Write an expression for the Area.",
    "9":  r"9) Expand $(x + 1)(x + 2)(x + 3)$"
}

def create_question_image(q_code, font_size):
    text = questions_db.get(q_code, f"Question {q_code}\n(Please update the code with the question text)")
    line_count = text.count('\n') + 1
    base_padding = 0.24 
    height_per_line = font_size * 0.035 
    fig_height = base_padding + (line_count * height_per_line)
    
    plt.figure(figsize=(7, fig_height))
    plt.text(0.01, 0.5, text, fontsize=font_size, verticalalignment='center', fontfamily='serif')
    plt.axis('off')
    plt.tight_layout(pad=0)
    
    img_name = f"q_{q_code}.png"
    plt.savefig(img_name, dpi=200, bbox_inches='tight')
    plt.close()
    return img_name

def process_data(uploaded_csv, uploaded_mapping):
    df_marks = pd.read_csv(uploaded_csv, header=None) if uploaded_csv.name.endswith('.csv') else pd.read_excel(uploaded_csv, header=None)
    
    row0 = df_marks.iloc[0].astype(str).tolist()
    row1 = df_marks.iloc[1].astype(str).tolist()
    
    q_labels = ["Surname", "Forename"]
    current_q = ""
    
    for i in range(2, len(row0)):
        r0 = row0[i].strip()
        r1 = row1[i].strip()
        if r0.lower() == 'total' or r1.lower() == 'total': break
        if r0 != 'nan' and r0 != '':
            m = re.search(r'\d+', r0)
            if m: current_q = m.group()
        if r1 != 'nan' and r1 != '': q_labels.append(current_q + r1)
        else: q_labels.append(current_q)

    full_marks_row = df_marks.iloc[2]
    
    percentage_idx = None
    for i in range(len(df_marks)):
        if 'percentage' in str(df_marks.iloc[i, 0]).lower():
            percentage_idx = i
            break
    
    if percentage_idx is None:
        raise ValueError("Could not find row starting with 'Percentage'")
        
    percentage_row = df_marks.iloc[percentage_idx]
    
    # --- SKIP STUDENTS WITH BLANK MARKS ---
    # We grab the student rows and filter: keep only rows where there is at least one non-null value in the question columns
    raw_student_rows = df_marks.iloc[3:percentage_idx].dropna(subset=[0, 1], how='all')
    
    # The marks columns start from column index 2
    student_rows = raw_student_rows[raw_student_rows.iloc[:, 2:len(q_labels)].notnull().any(axis=1)].reset_index(drop=True)
    # --------------------------------------
    
    df_map = pd.read_csv(uploaded_mapping, header=None) if uploaded_mapping.name.endswith('.csv') else pd.read_excel(uploaded_mapping, header=None)
    if 'topic' in str(df_map.iloc[0, 0]).lower():
        df_map = df_map.iloc[1:].reset_index(drop=True)

    dynamic_areas = []
    for _, map_row in df_map.iterrows():
        if pd.isna(map_row.iloc[0]): continue
        topic = str(map_row.iloc[0]).strip()
        qs = []
        last_num = ""
        for cell in map_row.iloc[1:]:
            if pd.isna(cell): continue
            tokens = str(cell).lower().replace('and', ',').replace('&', ',').split(',')
            for t in tokens:
                t = t.strip().replace(" ", "")
                if not t: continue
                num_part = "".join([c for c in t if c.isdigit()])
                let_part = "".join([c for c in t if c.isalpha()])
                if num_part: last_num = num_part
                candidate = (num_part or last_num) + let_part
                if candidate in q_labels and candidate not in qs:
                    qs.append(candidate)
        
        indices = [q_labels.index(q) for q in qs]
        if indices: dynamic_areas.append((topic, indices))
            
    return student_rows, percentage_row, full_marks_row, q_labels, dynamic_areas

def add_tight_picture(doc, img_path, width):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Cm(0.3)
    paragraph.paragraph_format.space_after = Cm(0.3)
    run = paragraph.add_run()
    run.add_picture(img_path, width=width)
    return paragraph

# --- 3. UI BUTTONS ---
col1, col2 = st.columns(2)
with col1:
    preview_clicked = st.button("👀 Preview Sample Student", use_container_width=True)
with col2:
    generate_clicked = st.button("📄 Generate All Feedback", type="primary", use_container_width=True)

# --- 4. LOGIC ---
if preview_clicked or generate_clicked:
    if not (uploaded_csv and uploaded_mapping):
        st.error("Please upload the Marks and Mapping files.")
    else:
        try:
            student_rows, percentage_row, full_marks_row, q_labels, dynamic_areas = process_data(uploaded_csv, uploaded_mapping)
            
            if student_rows.empty:
                st.warning("No students found with recorded marks in the question columns.")
            else:
                if preview_clicked:
                    row = student_rows.iloc[0]
                    name = f"{row[1]} {row[0]}"
                    st.markdown(f"### Preview Feedback: **{name}**")
                    
                    student_ebi = []
                    preview_data = []
                    for title, idxs in dynamic_areas:
                        w, e = [], []
                        for idx in idxs:
                            if pd.to_numeric(row[idx], errors='coerce') >= pd.to_numeric(full_marks_row[idx], errors='coerce'):
                                w.append(q_labels[idx])
                            else:
                                e.append(q_labels[idx]); student_ebi.append(q_labels[idx])
                        preview_data.append({"Topic": title, "WWW": ", ".join(w), "EBI": ", ".join(e)})
                    
                    st.table(pd.DataFrame(preview_data))
                    
                    reteach_list = [q for q in student_ebi if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= threshold_decimal]
                    personal_list = [q for q in student_ebi if q not in reteach_list]
                    
                    st.markdown("#### 🎯 Personal Corrections (EBI)")
                    for q in personal_list:
                        img = create_question_image(q, selected_font_size)
                        st.image(img)
                        os.remove(img)
                    
                    st.markdown(f"#### 🏫 Whole-Class Reteaching (Class Avg ≤ {selected_threshold}%)")
                    for q in reteach_list:
                        img = create_question_image(q, selected_font_size)
                        st.image(img)
                        os.remove(img)

                if generate_clicked:
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
                            if logo_path:
                                run = header.add_run()
                                run.add_picture(logo_path, width=Cm(1.5))
                                run.add_text("    ")
                            title_run = header.add_run(f"{unit_title} Feedback: {name}   |   Class: {class_name}")
                            title_run.bold = True
                            title_run.font.size = Pt(14)
                            
                            table = doc.add_table(rows=1, cols=3)
                            table.style = 'Table Grid'
                            hdr = table.rows[0].cells
                            hdr[0].text, hdr[1].text, hdr[2].text = "Area", "what went well", "even better if"
                            
                            student_ebi = []
                            for title, idxs in dynamic_areas:
                                w, e = [], []
                                for idx in idxs:
                                    if pd.to_numeric(row[idx], errors='coerce') >= pd.to_numeric(full_marks_row[idx], errors='coerce'):
                                        w.append(q_labels[idx])
                                    else:
                                        e.append(q_labels[idx]); student_ebi.append(q_labels[idx])
                                r = table.add_row().cells
                                r[0].text, r[1].text, r[2].text = title, ", ".join(w), ", ".join(e)

                            reteach_qs = [q for q in student_ebi if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= threshold_decimal]
                            personal_qs = [q for q in student_ebi if q not in reteach_qs]
                            
                            if personal_qs:
                                doc.add_heading("Personal correction", 2)
                                for q in personal_qs:
                                    img = create_question_image(q, selected_font_size)
                                    add_tight_picture(doc, img, width=Cm(14))
                                    os.remove(img)
                            
                            doc.add_page_break()
                            
                            doc.add_heading(f"Whole-class reteaching - {name}", 1)
                            if reteach_qs:
                                for q in reteach_qs:
                                    img = create_question_image(q, selected_font_size)
                                    add_tight_picture(doc, img, width=Cm(15))
                                    os.remove(img)
                            else:
                                doc.add_paragraph("Excellent mastery of class-wide topics.")
                            
                            doc.add_page_break()

                        target_docx = BytesIO()
                        doc.save(target_docx)
                        
                        prs = Presentation()
                        global_reteach = [q for q in q_labels[2:] if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= threshold_decimal]
                        for q in global_reteach:
                            slide = prs.slides.add_slide(prs.slide_layouts[6])
                            img = create_question_image(q, selected_font_size)
                            prs_img = slide.shapes.add_picture(img, PptxCm(1), PptxCm(2), width=PptxCm(23))
                            os.remove(img)
                        
                        target_pptx = BytesIO()
                        prs.save(target_pptx)

                        zip_buffer = BytesIO()
                        safe_class = str(class_name).replace(" ", "_")
                        with zipfile.ZipFile(zip_buffer, "w") as z:
                            z.writestr(f"{safe_class}_Feedback_Reports.docx", target_docx.getvalue())
                            if global_reteach:
                                z.writestr(f"{safe_class}_Reteach_Slides.pptx", target_pptx.getvalue())
                        
                        st.success(f"✅ Success! Feedback generated for {len(student_rows)} students.")
                        st.download_button("📦 Download All (ZIP)", zip_buffer.getvalue(), file_name=f"{safe_class}_Feedback_Pack.zip", type="primary")

                        if logo_path and os.path.exists(logo_path): os.remove(logo_path)

        except Exception as e:
            st.error(f"Error: {e}")
