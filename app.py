import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Cm
import os
from io import BytesIO

st.set_page_config(page_title="Exam Feedback Tool", layout="centered")

st.title("📊 Exam Feedback Generator")
st.write("Upload your files below. You can now use either CSV or Excel files for all inputs.")

# 1. FLEXIBLE FILE UPLOADERS
# Added 'xlsx' and 'xls' to all types so they aren't greyed out
uploaded_csv = st.file_uploader("1. Upload Marks (CSV or Excel)", type=["csv", "xlsx", "xls"])
uploaded_pdf = st.file_uploader("2. Upload Exam PDF", type="pdf")
uploaded_mapping = st.file_uploader("3. Upload Topic Mapping (CSV or Excel)", type=["csv", "xlsx", "xls"])

# Helper function to read either CSV or Excel
def load_data(file):
    if file.name.endswith('.csv'):
        return pd.read_csv(file, header=None)
    else:
        return pd.read_excel(file, header=None)

if st.button("Generate Feedback Pack"):
    if not (uploaded_csv and uploaded_pdf and uploaded_mapping):
        st.error("Missing files! Please upload all three files.")
    else:
        try:
            with st.spinner('Processing...'):
                # Process Marks with the new helper
                df_marks = load_data(uploaded_csv)
                
                # Standard locations from your template
                full_marks_row = df_marks.iloc[2]
                student_rows = df_marks.iloc[3:29].reset_index(drop=True)
                percentage_row = df_marks.iloc[34]
                
                q_labels = ["", "1a", "1b", "2a", "2b", "3a", "3b", "4a", "4b", "5a", "5b", "6a", "6b", "7a", "7b", "8a", "8b", "9a", "9b", "10"]

                # Process Mapping
                if uploaded_mapping.name.endswith('.csv'):
                    df_map = pd.read_csv(uploaded_mapping)
                else:
                    df_map = pd.read_excel(uploaded_mapping)
                
                # Ensure mapping has the right columns
                dynamic_areas = []
                for _, map_row in df_map.iterrows():
                    topic = map_row.iloc[0] # Topic name
                    qs = str(map_row.iloc[1]).split(',') # Question list
                    indices = [q_labels.index(q.strip()) for q in qs if q.strip() in q_labels]
                    dynamic_areas.append((topic, indices))

                # Process PDF
                pdf_bytes = uploaded_pdf.read()
                doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")

                # Coordinates
                crops = {
                    "1a": (0, fitz.Rect(50, 130, 550, 240)), "1b": (0, fitz.Rect(50, 200, 550, 310)),
                    "2a": (0, fitz.Rect(50, 330, 550, 440)), "2b": (0, fitz.Rect(50, 400, 550, 510)),
                    "3a": (0, fitz.Rect(50, 530, 550, 640)), "3b": (0, fitz.Rect(50, 600, 550, 710)),
                    "4a": (1, fitz.Rect(50, 100, 550, 240)), "4b": (1, fitz.Rect(50, 210, 550, 350)),
                    "5a": (1, fitz.Rect(50, 350, 550, 500)), "5b": (1, fitz.Rect(50, 480, 550, 630)),
                    "6a": (1, fitz.Rect(50, 630, 550, 730)), "6b": (1, fitz.Rect(50, 700, 550, 800)),
                    "7a": (2, fitz.Rect(50, 80, 550, 240)),  "7b": (2, fitz.Rect(50, 220, 550, 380)),
                    "8a": (2, fitz.Rect(50, 380, 550, 530)), "8b": (2, fitz.Rect(50, 530, 550, 680)),
                    "9a": (3, fitz.Rect(50, 80, 550, 250)),  "9b": (3, fitz.Rect(50, 250, 550, 420)),
                    "10": (3, fitz.Rect(50, 420, 550, 750)) 
                }

                doc = Document()
                for section in doc.sections:
                    section.top_margin, section.bottom_margin = Cm(0.5), Cm(0.5)
                    section.left_margin, section.right_margin = Cm(0.5), Cm(0.5)

                for _, row in student_rows.iterrows():
                    name = str(row[0])
                    if name == 'nan' or name == 'Name': continue
                    
                    doc.add_heading(f"Feedback: {name}", level=1)
                    table = doc.add_table(rows=1, cols=3); table.style = 'Table Grid'
                    hdr = table.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text = "Area", "what went well", "even better if"
                    table.columns[0].width, table.columns[1].width, table.columns[2].width = Cm(11), Cm(3.25), Cm(3.25)
                    
                    student_ebi = []
                    for title, idxs in dynamic_areas:
                        www, ebi = [], []
                        for idx in idxs:
                            score = pd.to_numeric(row[idx], errors='coerce')
                            if score > 0: www.append(q_labels[idx])
                            else:
                                ebi.append(q_labels[idx]); student_ebi.append(q_labels[idx])
                        r = table.add_row().cells
                        r[0].text, r[1].text, r[2].text = title, ", ".join(www), ", ".join(ebi)

                    reteach = [q for q in student_ebi if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= 0.55]
                    personal = [q for q in student_ebi if q not in reteach]
                    
                    if personal:
                        doc.add_heading("Personal correction", level=2)
                        for q in personal:
                            pix = doc_pdf[crops[q][0]].get_pixmap(clip=crops[q][1], matrix=fitz.Matrix(2, 2))
                            pix.save(f"temp_{q}.png")
                            doc.add_picture(f"temp_{q}.png", width=Cm(12))

                    doc.add_page_break()
                    doc.add_heading(f"Whole-class reteaching - {name}", level=1)
                    if reteach:
                        for q in reteach:
                            pix = doc_pdf[crops[q][0]].get_pixmap(clip=crops[q][1], matrix=fitz.Matrix(2, 2))
                            pix.save(f"temp_{q}.png")
                            doc.add_picture(f"temp_{q}.png", width=Cm(14))
                    doc.add_page_break()

                # Clean images
                for f in os.listdir():
                    if f.startswith("temp_") and f.endswith(".png"): os.remove(f)

                target = BytesIO()
                doc.save(target)
                st.success("✅ Pack Ready!")
                st.download_button("📥 Download Document", data=target.getvalue(), file_name="Feedback.docx")

        except Exception as e:
            st.error(f"Error: {e}")
