import sys
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Cm
import os
from io import BytesIO

# --- STEP 1: PYTHON 3.13 FIX ---
try:
    import imghdr
except ImportError:
    import filetype
    class MockImghdr:
        def what(self, file, h=None):
            kind = filetype.guess(file)
            return kind.extension if kind else None
    sys.modules['imghdr'] = MockImghdr()

st.set_page_config(page_title="Dynamic Feedback Pro", layout="centered")

st.title("📊 Dynamic Exam Feedback Generator")
st.write("This app reads the PDF to extract text and reconstructs questions as clean images.")

# --- STEP 2: THE THREE UPLOADERS ---
uploaded_csv = st.file_uploader("1. Upload Marks (CSV or Excel)", type=["csv", "xlsx"])
uploaded_pdf = st.file_uploader("2. Upload Original Exam PDF", type="pdf")
uploaded_mapping = st.file_uploader("3. Upload Topic Mapping (CSV or Excel)", type=["csv", "xlsx"])

def load_data(file):
    if file.name.endswith('.csv'): return pd.read_csv(file, header=None)
    else: return pd.read_excel(file, header=None)

def create_reconstructed_img(q_code, text):
    # Determine height based on text length
    line_count = text.count('\n') + 1
    if len(text) > 100: line_count += 1
    fig_height = 0.6 + (line_count * 0.45)
    
    plt.figure(figsize=(7, fig_height))
    plt.text(0.01, 0.5, text, fontsize=12, verticalalignment='center', fontfamily='serif', wrap=True)
    plt.axis('off')
    plt.tight_layout(pad=0)
    
    img_name = f"recon_{q_code}.png"
    plt.savefig(img_name, dpi=200, bbox_inches='tight')
    plt.close()
    return img_name

# --- STEP 3: CORE GENERATOR ---
if st.button("Generate Feedback Pack"):
    if not (uploaded_csv and uploaded_pdf and uploaded_mapping):
        st.error("Please upload all three files.")
    else:
        try:
            with st.spinner('Reading PDF and reconstructing questions...'):
                # 1. Load Marks & Mapping
                df_marks = load_data(uploaded_csv)
                student_rows = df_marks.iloc[3:29].reset_index(drop=True)
                percentage_row = df_marks.iloc[34]
                q_labels = ["", "1a", "1b", "2a", "2b", "3a", "3b", "4a", "4b", "5a", "5b", "6a", "6b", "7a", "7b", "8a", "8b", "9a", "9b", "10"]

                df_map = load_data(uploaded_mapping)
                if len(df_map.columns) < 2:
                    # Handle cases where mapping might not have headers
                    df_map = load_data(uploaded_mapping)
                
                # 2. Extract Text from PDF
                pdf_bytes = uploaded_pdf.read()
                doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
                full_text = ""
                for page in doc_pdf:
                    full_text += page.get_text()

                # 3. Dynamic Question Logic
                # This part looks for the question text in the PDF based on the mapping
                q_images = {}
                dynamic_areas = []
                
                for _, map_row in df_map.iterrows():
                    topic = str(map_row.iloc[0])
                    qs = str(map_row.iloc[1]).split(',')
                    indices = []
                    for q in qs:
                        q_code = q.strip()
                        if q_code in q_labels:
                            indices.append(q_labels.index(q_code))
                            # Attempt to find text for this question code in PDF
                            # (Heuristic: search for "1a)" or "1 a)")
                            search_term = f"{q_code})"
                            if q_code not in q_images:
                                # We extract a snippet of text around the question label
                                start_idx = full_text.find(search_term)
                                if start_idx != -1:
                                    # Take the next 300 characters as the question
                                    excerpt = full_text[start_idx:start_idx+300].split('\n\n')[0]
                                    q_images[q_code] = create_reconstructed_img(q_code, excerpt)
                                else:
                                    # Fallback if text not found
                                    q_images[q_code] = create_reconstructed_img(q_code, f"Question {q_code}")
                    dynamic_areas.append((topic, indices))

                # 4. Build Document
                doc = Document()
                for section in doc.sections:
                    section.top_margin, section.bottom_margin = Cm(0.5), Cm(0.5)
                    section.left_margin, section.right_margin = Cm(0.5), Cm(0.5)

                for _, row in student_rows.iterrows():
                    name = str(row[0])
                    if name == 'nan' or name == 'Name': continue
                    
                    doc.add_heading(f"Feedback: {name}", 1)
                    table = doc.add_table(rows=1, cols=3); table.style = 'Table Grid'
                    hdr = table.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text = "Area", "what went well", "even better if"
                    table.columns[0].width, table.columns[1].width, table.columns[2].width = Cm(11), Cm(3.25), Cm(3.25)
                    
                    student_ebi = []
                    for title, idxs in dynamic_areas:
                        w, e = [], []
                        for idx in idxs:
                            score = pd.to_numeric(row[idx], errors='coerce')
                            q_code = q_labels[idx]
                            if score > 0: w.append(q_code)
                            else:
                                e.append(q_code)
                                student_ebi.append(q_code)
                        r = table.add_row().cells
                        r[0].text, r[1].text, r[2].text = title, ", ".join(w), ", ".join(e)

                    reteach = [q for q in student_ebi if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= 0.55]
                    personal = [q for q in student_ebi if q not in reteach]
                    
                    if personal:
                        doc.add_heading("Personal correction", 2)
                        for q in personal: doc.add_picture(q_images[q], width=Cm(12))

                    doc.add_page_break()
                    doc.add_heading(f"Whole-class reteaching - {name}", 1)
                    if reteach:
                        for q in reteach: 
                            doc.add_picture(q_images[q], width=Cm(14))
                            doc.add_paragraph()
                    doc.add_page_break()

                # 5. Output
                target = BytesIO()
                doc.save(target)
                st.success("✅ Success!")
                st.download_button("📥 Download Feedback Pack", data=target.getvalue(), file_name="Dynamic_Feedback.docx")
                
                # Cleanup
                doc_pdf.close()
                for f in os.listdir():
                    if f.startswith("recon_") and f.endswith(".png"): os.remove(f)

        except Exception as e:
            st.error(f"Error: {e}")
