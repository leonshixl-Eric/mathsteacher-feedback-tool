import sys
import matplotlib
matplotlib.use('Agg') # Crucial for Streamlit servers
import matplotlib.pyplot as plt
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm
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

st.set_page_config(page_title="Maths Feedback Pro", layout="centered")

st.title("📊 High-Fidelity Feedback Generator")
st.write("Upload all three files. Questions are fully reconstructed with perfect mathematical formatting.")

# --- 1. THE THREE UPLOADERS ---
uploaded_csv = st.file_uploader("1. Upload Marks (CSV or Excel)", type=["csv", "xlsx"])
uploaded_pdf = st.file_uploader("2. Upload Original Exam PDF", type="pdf")
uploaded_mapping = st.file_uploader("3. Upload Topic Mapping (CSV or Excel)", type=["csv", "xlsx"])

# --- 2. THE RECONSTRUCTION ENGINE (Perfect Math Formatting) ---
# This ensures the full instruction AND the question are drawn perfectly.
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

def create_question_image(q_code, text):
    """Draws the text perfectly with adaptive height."""
    line_count = text.count('\n') + 1
    fig_height = 0.5 + (line_count * 0.4) # Adapts to fit full instructions
    
    plt.figure(figsize=(7, fig_height))
    plt.text(0.01, 0.5, text, fontsize=12, verticalalignment='center', fontfamily='serif')
    plt.axis('off')
    plt.tight_layout(pad=0)
    
    img_name = f"q_{q_code}.png"
    plt.savefig(img_name, dpi=200, bbox_inches='tight')
    plt.close()
    return img_name

# --- 3. CORE GENERATOR ---
if st.button("Generate Feedback Pack"):
    if not (uploaded_csv and uploaded_pdf and uploaded_mapping):
        st.error("Please upload all three files (Marks, PDF, Mapping).")
    else:
        try:
            with st.spinner('Reconstructing perfect questions and building report...'):
                # Read Marks
                df_marks = pd.read_csv(uploaded_csv, header=None) if uploaded_csv.name.endswith('.csv') else pd.read_excel(uploaded_csv, header=None)
                student_rows = df_marks.iloc[3:29].reset_index(drop=True)
                percentage_row = df_marks.iloc[34]
                q_labels = ["", "1a", "1b", "2a", "2b", "3a", "3b", "4a", "4b", "5a", "5b", "6a", "6b", "7a", "7b", "8a", "8b", "9a", "9b", "10"]

                # Read Mapping
                df_map = pd.read_csv(uploaded_mapping) if uploaded_mapping.name.endswith('.csv') else pd.read_excel(uploaded_mapping)
                dynamic_areas = []
                for _, map_row in df_map.iterrows():
                    topic, qs = str(map_row.iloc[0]), str(map_row.iloc[1]).split(',')
                    indices = [q_labels.index(q.strip()) for q in qs if q.strip() in q_labels]
                    dynamic_areas.append((topic, indices))

                # Generate Images (No more screenshots!)
                q_images = {q: create_question_image(q, txt) for q, txt in questions_db.items()}

                doc = Document()
                for section in doc.sections:
                    section.top_margin, section.bottom_margin = Cm(0.9), Cm(0.9)
                    section.left_margin, section.right_margin = Cm(0.9), Cm(0.9)

                for _, row in student_rows.iterrows():
                    name = str(row[0])
                    if name == 'nan' or name == 'Name': continue
                    
                    doc.add_heading(f"Feedback Report: {name}", 1)
                    table = doc.add_table(rows=1, cols=3); table.style = 'Table Grid'
                    hdr = table.rows[0].cells
                    hdr[0].text, hdr[1].text, hdr[2].text = "Area", "what went well", "even better if"
                    table.columns[0].width, table.columns[1].width, table.columns[2].width = Cm(14), Cm(1.25), Cm(1.25)
                    
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
                    else: doc.add_paragraph("Excellent mastery of class topics.")
                    doc.add_page_break()

                # Output
                target = BytesIO()
                doc.save(target)
                st.success("✅ Feedback Pack Ready!")
                st.download_button("📥 Download Document", data=target.getvalue(), file_name="Feedback_Final.docx")
                
                # Cleanup
                for f in os.listdir():
                    if f.startswith("q_") and f.endswith(".png"): os.remove(f)

        except Exception as e:
            st.error(f"Error: {e}")
