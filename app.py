import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Cm
from PIL import Image, ImageChops
import os
from io import BytesIO

st.set_page_config(page_title="Maths Feedback Automator", layout="centered")

st.title("📊 Universal Exam Feedback Tool")
st.write("Upload your files. This version uses 'Smart-Crop' to capture perfect, tight snippets from the PDF.")

# 1. THE THREE UPLOADERS
uploaded_csv = st.file_uploader("1. Upload Marks (CSV or Excel)", type=["csv", "xlsx"])
uploaded_pdf = st.file_uploader("2. Upload Original Exam PDF", type="pdf")
uploaded_mapping = st.file_uploader("3. Upload Topic Mapping (CSV or Excel)", type=["csv", "xlsx"])

# Helper function to trim white space
def trim_white_space(img_path):
    im = Image.open(img_path)
    bg = Image.new(im.mode, im.size, im.getpixel((0,0)))
    diff = ImageChops.difference(im, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return im.crop(bbox)
    return im

if st.button("Generate Feedback Pack"):
    if not (uploaded_csv and uploaded_pdf and uploaded_mapping):
        st.error("Please upload all THREE files to proceed.")
    else:
        try:
            with st.spinner('Smart-cropping questions from PDF...'):
                # Load Data
                df_marks = pd.read_csv(uploaded_csv, header=None) if uploaded_csv.name.endswith('.csv') else pd.read_excel(uploaded_csv, header=None)
                student_rows = df_marks.iloc[3:29].reset_index(drop=True)
                percentage_row = df_marks.iloc[34]
                q_labels = ["", "1a", "1b", "2a", "2b", "3a", "3b", "4a", "4b", "5a", "5b", "6a", "6b", "7a", "7b", "8a", "8b", "9a", "9b", "10"]

                # Process PDF
                pdf_bytes = uploaded_pdf.read()
                doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")

                # The "Master List" of Question Coordinates (The real question parts)
                # These coordinates are slightly taller to catch the instructions above.
                crops = {
                    "1a": (0, fitz.Rect(50, 105, 550, 185)), "1b": (0, fitz.Rect(50, 105, 550, 240)),
                    "2a": (0, fitz.Rect(50, 260, 550, 360)), "2b": (0, fitz.Rect(50, 260, 550, 430)),
                    "3a": (0, fitz.Rect(50, 460, 550, 560)), "3b": (0, fitz.Rect(50, 460, 550, 630)),
                    "4a": (1, fitz.Rect(50, 80, 550, 200)),  "4b": (1, fitz.Rect(50, 80, 550, 300)),
                    "5a": (1, fitz.Rect(50, 320, 550, 450)), "5b": (1, fitz.Rect(50, 320, 550, 580)),
                    "6a": (1, fitz.Rect(50, 610, 550, 710)), "6b": (1, fitz.Rect(50, 610, 550, 780)),
                    "7a": (2, fitz.Rect(50, 60, 550, 180)),  "7b": (2, fitz.Rect(50, 60, 550, 320)),
                    "8a": (2, fitz.Rect(50, 360, 550, 480)), "8b": (2, fitz.Rect(50, 360, 550, 630)),
                    "9a": (3, fitz.Rect(50, 70, 550, 230)),  "9b": (3, fitz.Rect(50, 70, 550, 380)),
                    "10": (3, fitz.Rect(50, 410, 550, 680)) 
                }

                # Create Smart-Cropped Images
                q_images = {}
                for q, (pg, rect) in crops.items():
                    pix = doc_pdf[pg].get_pixmap(clip=rect, matrix=fitz.Matrix(2, 2))
                    temp_path = f"raw_{q}.png"
                    pix.save(temp_path)
                    
                    # Apply the "Smart-Crop" (Trim white space)
                    trimmed_img = trim_white_space(temp_path)
                    final_path = f"smart_{q}.png"
                    trimmed_img.save(final_path)
                    q_images[q] = final_path
                    os.remove(temp_path)

                # Process Mapping
                df_map = pd.read_csv(uploaded_mapping) if uploaded_mapping.name.endswith('.csv') else pd.read_excel(uploaded_mapping)
                dynamic_areas = []
                for _, map_row in df_map.iterrows():
                    topic, qs = str(map_row.iloc[0]), str(map_row.iloc[1]).split(',')
                    indices = [q_labels.index(q.strip()) for q in qs if q.strip() in q_labels]
                    dynamic_areas.append((topic, indices))

                # Build Doc
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
                            if score > 0: w.append(q_labels[idx])
                            else: e.append(q_labels[idx]); student_ebi.append(q_labels[idx])
                        r = table.add_row().cells
                        r[0].text, r[1].text, r[2].text = str(title), ", ".join(w), ", ".join(e)

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

                target = BytesIO(); doc.save(target)
                st.success("✅ Document Ready!")
                st.download_button("📥 Download Feedback Pack", data=target.getvalue(), file_name="Feedback_Final.docx")
                
                # Cleanup
                doc_pdf.close()
                for f in os.listdir():
                    if f.startswith("smart_") and f.endswith(".png"): os.remove(f)

        except Exception as e:
            st.error(f"Error: {e}")
