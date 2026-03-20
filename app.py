import sys
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
import streamlit as st
import pandas as pd
from docx import Document
# NEW: Import necessary WordML formatting tools
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
from io import BytesIO

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
st.write("Upload all three files. Questions are fully reconstructed with perfect mathematical formatting.")

# --- 1. THE UPLOADERS ---
uploaded_csv = st.file_uploader("1. Upload Marks (CSV or Excel)", type=["csv", "xlsx"])
uploaded_pdf = st.file_uploader("2. Upload Original Exam PDF (Reference)", type="pdf")
uploaded_mapping = st.file_uploader("3. Upload Topic Mapping (CSV or Excel)", type=["csv", "xlsx"])

st.markdown("---")
st.subheader("⚙️ Document Settings")

col_setting1, col_setting2, col_setting3 = st.columns(3)
with col_setting1:
    selected_font_size = st.slider("Question Font Size", min_value=10, max_value=14, value=11, step=1)
with col_setting2:
    selected_threshold = st.slider("Reteach Threshold (%)", min_value=0, max_value=100, value=55, step=5)
with col_setting3:
    selected_margin = st.slider("Page Margin (cm)", min_value=0.5, max_value=3.0, value=1.27, step=0.1)
st.markdown("---")

threshold_decimal = selected_threshold / 100.0

# --- 2. THE RECONSTRUCTION ENGINE ---
questions_db = {
    "1a": r"Write each number as a power of 10." + "\n" + r"1a) 1000",
    "1b": r"Write each number as a power of 10." + "\n" + r"1b) 0.01",
    "2a": r"Write each power of 10 as an ordinary number." + "\n" + r"2a) $10^5$",
    "2b": r"Write each power of 10 as an ordinary number." + "\n" + r"2b) $10^{-3}$",
    "3a": r"Write each number in standard form as an ordinary number." + "\n" + r"3a) $5 \times 10^6$",
    "3b": r"Write each number in standard form as an ordinary number." + "\n" + r"3b) $3.7 \times 10^3$",
    "4a": r"Write each number in standard form as an ordinary number." + "\n" + r"4a) $7 \times 10^{-3}$",
    "4b": r"Write each number in standard form as an ordinary number." + "\n" + r"4b) $8.39 \times 10^{-5}$",
    "5a": r"5a) The diameter of Mars is approximately 7000 km." + "\n" + r"      Write the diameter of Mars in standard form.",
    "5b": r"5b) The diameter of Uranus is approximately 50,720,000 m." + "\n" + r"      Write the diameter of Uranus in standard form.",
    "6a": r"Write each number in standard form." + "\n" + r"6a) 0.0005",
    "6b": r"Write each number in standard form." + "\n" + r"6b) 0.0201",
    "7a": r"Write <, > or = to make the statements correct." + "\n" + r"7a) 810,000 [   ] $8.1 \times 10^4$",
    "7b": r"Write <, > or = to make the statements correct." + "\n" + r"7b) $3 \times 10^{-4}$ [   ] 0.0003",
    "8a": r"Write each number in standard form." + "\n" + r"8a) $64 \times 10^7$",
    "8b": r"Write each number in standard form." + "\n" + r"8b) $360.7 \times 10^{-5}$",
    "9a": r"Work out the following." + "\n" + r"Give your answers in standard form." + "\n" + r"9a) $(3 \times 10^4) + (6 \times 10^3)$",
    "9b": r"Work out the following." + "\n" + r"Give your answers in standard form." + "\n" + r"9b) $(1.5 \times 10^{-5}) \div (5 \times 10^{-1})$",
    "10": r"10) The distance from Earth to Venus is approximately $4.5 \times 10^7$ km." + "\n" +
          r"      A spacecraft travels at a speed of $5 \times 10^8$ km/h." + "\n" +
          r"      Work out how many hours it will take the spacecraft to reach Venus." + "\n" +
          r"      Give your answer in standard form."
}

def create_question_image(q_code, text, font_size):
    line_count = text.count('\n') + 1
    base_padding = 0.3
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
    student_rows = df_marks.iloc[3:29].reset_index(drop=True)
    percentage_row = df_marks.iloc[34]
    q_labels = ["", "1a", "1b", "2a", "2b", "3a", "3b", "4a", "4b", "5a", "5b", "6a", "6b", "7a", "7b", "8a", "8b", "9a", "9b", "10"]

    df_map = pd.read_csv(uploaded_mapping, header=None) if uploaded_mapping.name.endswith('.csv') else pd.read_excel(uploaded_mapping, header=None)
    
    first_cell = str(df_map.iloc[0, 0]).lower()
    if 'topic' in first_cell or 'area' in first_cell:
        df_map = df_map.iloc[1:].reset_index(drop=True)

    dynamic_areas = []
    for _, map_row in df_map.iterrows():
        if pd.isna(map_row.iloc[0]): continue
        topic = str(map_row.iloc[0]).strip()
        
        qs = []
        last_num = ""
        
        for col_idx in range(1, len(map_row)):
            cell_val = map_row.iloc[col_idx]
            if pd.isna(cell_val): continue
            
            raw_str = str(cell_val).lower().replace('and', ',').replace('&', ',')
            raw_str = "".join(raw_str.split())
            tokens = raw_str.split(',')
            
            for t in tokens:
                if not t: continue
                num_part = "".join([c for c in t if c.isdigit()])
                letter_part = "".join([c for c in t if c.isalpha()])
                
                if num_part:
                    last_num = num_part
                    candidate = num_part + letter_part
                else:
                    candidate = last_num + letter_part
                    
                if candidate in q_labels and candidate not in qs:
                    qs.append(candidate)
        
        indices = [q_labels.index(q) for q in qs]
        if indices:
            dynamic_areas.append((topic, indices))
            
    return student_rows, percentage_row, q_labels, dynamic_areas

# NEW FORMATTING HELPERS
def add_tight_picture(doc, img_path, width):
    """Adds a picture in its own paragraph with ZERO space after it."""
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Cm(0)
    paragraph.paragraph_format.space_before = Cm(0)
    run = paragraph.add_run()
    run.add_picture(img_path, width=width)
    return paragraph

def add_spacer(doc, height_cm):
    """Adds a paragraph forced to have exactly the specified height (the gap)."""
    paragraph = doc.add_paragraph()
    # Lock the exact height using WordprocessingML formatting
    paragraph.paragraph_format.space_before = Cm(0)
    paragraph.paragraph_format.space_after = Cm(0)
    
    # Access internal XML for precise font height control
    p_format = paragraph.paragraph_format
    run = paragraph.add_run()
    run.font.size = Pt(1) # Make font tiny so it doesn't affect space
    
    # Define exact line height (Word treats 12pt font as ~0.42cm usually)
    # 0.5cm is approximately 14.17 points.
    line_spacing_pts = height_per_line_in_pts = height_cm * 28.35
    p_format.line_spacing = Pt(line_spacing_pts)
    p_format.line_spacing_rule = 3 # Exact line spacing rule

# --- 3. BUTTON LAYOUT ---
col1, col2 = st.columns(2)
with col1:
    preview_clicked = st.button("👀 Preview Sample Student", use_container_width=True)
with col2:
    generate_clicked = st.button("📄 Generate All Feedback", type="primary", use_container_width=True)

# --- 4. PREVIEW LOGIC ---
if preview_clicked:
    if not (uploaded_csv and uploaded_pdf and uploaded_mapping):
        st.warning("Please upload all three files to see a preview.")
    else:
        with st.spinner("Generating preview..."):
            student_rows, percentage_row, q_labels, dynamic_areas = process_data(uploaded_csv, uploaded_mapping)
            
            with st.expander("🛠️ DEBUG: See how the app read your Multi-Column Mapping File", expanded=True):
                st.write("The app now scans across all columns in your Excel file.")
                for title, idxs in dynamic_areas:
                    mapped_qs = [q_labels[i] for i in idxs]
                    st.success(f"**{title}**: {mapped_qs}")

            first_student = None
            for _, row in student_rows.iterrows():
                if str(row[0]) != 'nan' and str(row[0]) != 'Name':
                    first_student = row
                    break
            
            if first_student is not None:
                name = str(first_student[0])
                st.markdown(f"### 📋 Preview for: **{name}**")
                
                preview_table = []
                student_ebi = []
                for title, idxs in dynamic_areas:
                    w, e = [], []
                    for idx in idxs:
                        score = pd.to_numeric(first_student[idx], errors='coerce')
                        if score > 0: w.append(q_labels[idx])
                        else: e.append(q_labels[idx]); student_ebi.append(q_labels[idx])
                    preview_table.append({"Area": title, "What Went Well": ", ".join(w), "Even Better If": ", ".join(e)})
                
                st.table(pd.DataFrame(preview_table))
                
                reteach = [q for q in student_ebi if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= threshold_decimal]
                personal = [q for q in student_ebi if q not in reteach]
                
                st.markdown("#### 🎯 Personal Corrections")
                if personal:
                    for q in personal:
                        img_path = create_question_image(q, questions_db[q], selected_font_size)
                        st.image(img_path)
                        os.remove(img_path)
                else:
                    st.success("No personal corrections needed!")
                
                st.markdown(f"#### 🏫 Whole-Class Reteaching (≤ {selected_threshold}%)")
                if reteach:
                    for q in reteach:
                        img_path = create_question_image(q, questions_db[q], selected_font_size)
                        st.image(img_path)
                        os.remove(img_path)
                else:
                    st.success("No whole-class reteaching needed!")
            else:
                st.error("Could not find any valid students in the CSV.")

# --- 5. GENERATE LOGIC ---
if generate_clicked:
    if not (uploaded_csv and uploaded_pdf and uploaded_mapping):
        st.error("Please upload all three files (Marks, PDF, Mapping).")
    else:
        try:
            with st.spinner(f'Reconstructing questions and applying exact 0.5cm image gaps...'):
                student_rows, percentage_row, q_labels, dynamic_areas = process_data(uploaded_csv, uploaded_mapping)
                q_images = {q: create_question_image(q, txt, selected_font_size) for q, txt in questions_db.items()}

                doc = Document()
                
                available_width_cm = 21.0 - (2 * selected_margin)
                area_col_width = available_width_cm - 7.0 
                col_widths = [Cm(area_col_width), Cm(3.5), Cm(3.5)]
                
                personal_img_width = Cm(min(12.0, available_width_cm))
                reteach_img_width = Cm(min(14.0, available_width_cm))

                for section in doc.sections:
                    section.page_width = Cm(21.0)
                    section.page_height = Cm(29.7)
                    section.top_margin, section.bottom_margin = Cm(selected_margin), Cm(selected_margin)
                    section.left_margin, section.right_margin = Cm(selected_margin), Cm(selected_margin)

                for _, row in student_rows.iterrows():
                    name = str(row[0])
                    if name == 'nan' or name == 'Name': continue
                    
                    doc.add_heading(f"Feedback Report: {name}", 1)
                    
                    table = doc.add_table(rows=1, cols=3)
                    table.style = 'Table Grid'
                    
                    for i in range(3): table.columns[i].width = col_widths[i]
                    
                    hdr = table.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text = "Area", "what went well", "even better if"
                    for i in range(3): hdr[i].width = col_widths[i]
                    
                    student_ebi = []
                    for title, idxs in dynamic_areas:
                        w, e = [], []
                        for idx in idxs:
                            score = pd.to_numeric(row[idx], errors='coerce')
                            if score > 0: w.append(q_labels[idx])
                            else: e.append(q_labels[idx]); student_ebi.append(q_labels[idx])
                        
                        r = table.add_row().cells
                        r[0].text, r[1].text, r[2].text = str(title), ", ".join(w), ", ".join(e)
                        for i in range(3): r[i].width = col_widths[i]

                    reteach = [q for q in student_ebi if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= threshold_decimal]
                    personal = [q for q in student_ebi if q not in reteach]
                    
                    # --- APPLIYNG IMAGE SPACING (PERSONAL) ---
                    if personal:
                        doc.add_heading("Personal correction", 2)
                        for q in personal:
                            # 1. Add picture in a ZERO-SPACE paragraph
                            add_tight_picture(doc, q_images[q], width=personal_img_width)
                            # 2. Add an EXACT 0.5cm spacer
                            add_spacer(doc, 0.5)

                    doc.add_page_break()
                    doc.add_heading(f"Whole-class reteaching - {name}", 1)
                    
                    # --- APPLIYNG IMAGE SPACING (RETEACH) ---
                    if reteach:
                        for q in reteach: 
                            # 1. Add picture in a ZERO-SPACE paragraph
                            add_tight_picture(doc, q_images[q], width=reteach_img_width)
                            # 2. Add an EXACT 0.5cm spacer
                            add_spacer(doc, 0.5)
                    else: doc.add_paragraph("Excellent mastery of class topics.")
                    doc.add_page_break()

                target = BytesIO()
                doc.save(target)
                st.success(f"✅ Feedback Pack Ready! (Exact image gaps: 0.5cm)")
                st.download_button("📥 Download Document", data=target.getvalue(), file_name="Feedback_Final.docx")
                
                for f in os.listdir():
                    if f.startswith("q_") and f.endswith(".png"): os.remove(f)

        except Exception as e:
            st.error(f"Error: {e}")
