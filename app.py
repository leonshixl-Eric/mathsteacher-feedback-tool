import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Cm
import os
from io import BytesIO

st.set_page_config(page_title="Universal Exam Feedback Tool", layout="centered")

st.title("📊 Exam Feedback Generator")
st.write("Upload your marks, exam paper, and topic mapping to generate booklets.")

# 1. File Uploaders
uploaded_csv = st.file_uploader("1. Upload Marks CSV", type="csv")
uploaded_pdf = st.file_uploader("2. Upload Exam PDF", type="pdf")
uploaded_mapping = st.file_uploader("3. Upload Topic Mapping (CSV or Excel)", type=["csv", "xlsx"])

if st.button("Generate Feedback Pack"):
    if not uploaded_csv or not uploaded_pdf or not uploaded_mapping:
        st.error("Please upload all three files (Marks, PDF, and Mapping) to proceed.")
    else:
        try:
            with st.spinner('Processing...'):
                # Load Marks Data
                df_marks = pd.read_csv(uploaded_csv, header=None)
                full_marks_row = df_marks.iloc[2]
                student_rows = df_marks.iloc[3:29].reset_index(drop=True)
                percentage_row = df_marks.iloc[34]
                
                # Question labels from column headers
                q_labels = ["", "1a", "1b", "2a", "2b", "3a", "3b", "4a", "4b", "5a", "5b", "6a", "6b", "7a", "7b", "8a", "8b", "9a", "9b", "10"]

                # Process Mapping File
                if uploaded_mapping.name.endswith('.csv'):
                    df_map = pd.read_csv(uploaded_mapping)
                else:
                    df_map = pd.read_excel(uploaded_mapping)
                
                # Convert Mapping to the list format the code needs
                # Assumes columns are named 'Topic' and 'Questions'
                dynamic_areas = []
                for _, map_row in df_map.iterrows():
                    topic = map_row['Topic']
                    qs = str(map_row['Questions']).split(',')
                    indices = []
                    for q in qs:
                        q_clean = q.strip()
                        if q_clean in q_labels:
                            indices.append(q_labels.index(q_clean))
                    dynamic_areas.append((topic, indices))

                # Process PDF
                with open("temp_exam.pdf", "wb") as f:
                    f.write(uploaded_pdf.getbuffer())
                doc_pdf = fitz.open("temp_exam.pdf")

                # Snippet coordinates
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
                                ebi.append(q_labels[idx])
                                student_ebi.append(q_labels[idx])
                        r = table.add_row().cells
                        r[0].text, r[1].text, r[2].text = title, ", ".join(www), ", ".join(ebi)
                        r[0].width, r[1].width, r[2].width = Cm(11), Cm(3.25), Cm(3.25)

                    reteach, personal = [], []
                    for q in student_ebi:
                        idx = q_labels.index(q)
                        if pd.to_numeric(percentage_row[idx], errors='coerce') <= 0.55:
                            reteach.append(q)
                        else: personal.append(q)
                    
                    if personal:
                        doc.add_heading("Personal correction", level=2)
                        for q in personal:
                            pg, rect = crops[q]
                            pix = doc_pdf[pg].get_pixmap(clip=rect, matrix=fitz.Matrix(2, 2))
                            pix.save(f"temp_{q}.png")
                            doc.add_picture(f"temp_{q}.png", width=Cm(12))

                    doc.add_page_break()
                    doc.add_heading(f"Whole-class reteaching - {name}", level=1)
                    if reteach:
                        for q in reteach:
                            pg, rect = crops[q]
                            pix = doc_pdf[pg].get_pixmap(clip=rect, matrix=fitz.Matrix(2, 2))
                            pix.save(f"temp_{q}.png")
                            doc.add_picture(f"temp_{q}.png", width=Cm(14))
                    doc.add_page_break()

                target = BytesIO()
                doc.save(target)
                st.success("Pack Ready!")
                st.download_button(label="📥 Download Word Document", data=target.getvalue(), file_name="Feedback.docx")
                doc_pdf.close()
                os.remove("temp_exam.pdf")
                for f in os.listdir():
                    if f.startswith("temp_") and f.endswith(".png"): os.remove(f)

        except Exception as e:
            st.error(f"Error: {e}")
