import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Inches
import matplotlib.pyplot as plt
import os
from io import BytesIO

st.set_page_config(page_title="Exam Feedback Tool", layout="centered")

st.title("📊 Exam Feedback Generator")
st.write("Upload your files. This version uses reconstructed high-resolution question images for a perfect fit.")

# 1. Flexible File Uploaders
uploaded_csv = st.file_uploader("1. Upload Marks (CSV or Excel)", type=["csv", "xlsx", "xls"])
uploaded_pdf = st.file_uploader("2. Upload Exam PDF (Used for reference only)", type="pdf")
uploaded_mapping = st.file_uploader("3. Upload Topic Mapping (CSV or Excel)", type=["csv", "xlsx", "xls"])

# Helper to read data
def load_data(file):
    if file.name.endswith('.csv'):
        return pd.read_csv(file, header=None)
    else:
        return pd.read_excel(file, header=None)

# 2. Mathematical Question Reconstructions
# This matches the exam paper's formatting exactly but in a compact size
questions_text = {
    "1a": r"1a) Write 1000 as a power of 10",
    "1b": r"1b) Write 0.01 as a power of 10",
    "2a": r"2a) Write the power of 10 as an ordinary number: $10^5$",
    "2b": r"2b) Write the power of 10 as an ordinary number: $10^{-3}$",
    "3a": r"3a) Write the standard form as an ordinary number: $5 \times 10^6$",
    "3b": r"3b) Write the standard form as an ordinary number: $3.7 \times 10^3$",
    "4a": r"4a) Write the standard form as an ordinary number: $7 \times 10^{-3}$",
    "4b": r"4b) Write the standard form as an ordinary number: $8.39 \times 10^{-5}$",
    "5a": r"5a) The diameter of Mars is approx 7000 km. Write in standard form.",
    "5b": r"5b) The diameter of Uranus is approx 50,720,000 m. Write in standard form.",
    "6a": r"6a) Write 0.0005 in standard form.",
    "6b": r"6b) Write 0.0201 in standard form.",
    "7a": r"7a) Write <, > or = to compare: 810000 [  ] $8.1 \times 10^4$",
    "7b": r"7b) Write <, > or = to compare: $3 \times 10^{-4}$ [  ] 0.0003",
    "8a": r"8a) Write $64 \times 10^7$ in standard form.",
    "8b": r"8b) Write $360.7 \times 10^{-5}$ in standard form.",
    "9a": r"9a) $(3 \times 10^4) + (6 \times 10^3)$. Give answer in standard form.",
    "9b": r"9b) $(1.5 \times 10^{-5}) \div (5 \times 10^{-1})$. Give answer in standard form.",
    "10": "10) The distance from Earth to Venus is approximately $4.5 \\times 10^7$ km.\n      A spacecraft travels at a speed of $5 \\times 10^8$ km/h.\n      Work out how many hours it will take to reach Venus.\n      Give your answer in standard form."
}

def make_reconstructed_img(q_code, text):
    num_lines = text.count('\n') + 1
    height = 0.7 + (num_lines - 1) * 0.45
    plt.figure(figsize=(7, height))
    plt.text(0.01, 0.5, text, fontsize=12, verticalalignment='center', fontfamily='serif')
    plt.axis('off')
    plt.tight_layout(pad=0)
    img_name = f"recon_{q_code}.png"
    plt.savefig(img_name, dpi=150, bbox_inches='tight')
    plt.close()
    return img_name

if st.button("Generate Feedback Pack"):
    if not (uploaded_csv and uploaded_mapping):
        st.error("Please upload the Marks and Mapping files.")
    else:
        try:
            with st.spinner('Reconstructing questions and building report...'):
                # Generate Images
                snippet_imgs = {q: make_reconstructed_img(q, txt) for q, txt in questions_text.items()}

                # Load Data
                df_marks = load_data(uploaded_csv)
                student_rows = df_marks.iloc[3:29].reset_index(drop=True)
                percentage_row = df_marks.iloc[34]
                q_labels = ["", "1a", "1b", "2a", "2b", "3a", "3b", "4a", "4b", "5a", "5b", "6a", "6b", "7a", "7b", "8a", "8b", "9a", "9b", "10"]

                # Process Mapping
                df_map = load_data(uploaded_mapping)
                dynamic_areas = []
                for _, map_row in df_map.iterrows():
                    topic = map_row.iloc[0]
                    qs = str(map_row.iloc[1]).split(',')
                    indices = [q_labels.index(q.strip()) for q in qs if q.strip() in q_labels]
                    dynamic_areas.append((topic, indices))

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
                            if pd.notnull(score) and score > 0: w.append(q_labels[idx])
                            else: e.append(q_labels[idx]); student_ebi.append(q_labels[idx])
                        r = table.add_row().cells
                        r[0].text, r[1].text, r[2].text = str(title), ", ".join(w), ", ".join(e)
                        r[0].width, r[1].width, r[2].width = Cm(11), Cm(3.25), Cm(3.25)

                    reteach = [q for q in student_ebi if pd.to_numeric(percentage_row[q_labels.index(q)], errors='coerce') <= 0.55]
                    personal = [q for q in student_ebi if q not in reteach]
                    
                    if personal:
                        doc.add_heading("Personal correction", 2)
                        for q in personal: doc.add_picture(snippet_imgs[q], width=Cm(13))
                    
                    doc.add_page_break()
                    doc.add_heading(f"Whole-class reteaching - {name}", 1)
                    if reteach:
                        for q in reteach: 
                            doc.add_picture(snippet_imgs[q], width=Cm(14))
                            doc.add_paragraph()
                    else: doc.add_paragraph("Mastered class topics.")
                    doc.add_page_break()

                target = BytesIO()
                doc.save(target)
                st.success("✅ Document Generated!")
                st.download_button("📥 Download Feedback Pack", data=target.getvalue(), file_name="Feedback_Report.docx")
                
                # Cleanup
                for f in os.listdir():
                    if f.startswith("recon_") and f.endswith(".png"): os.remove(f)

        except Exception as e:
            st.error(f"Error: {e}")
