import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.shared import Pt
from string import ascii_lowercase

st.set_page_config(page_title="Reference List Generator", layout="centered")
st.title("📄 Reference List Generator (MVP)")

uploaded_file = st.file_uploader("Upload a spreadsheet (.xlsx or .ods)", type=["xlsx", "ods"])
project_name = st.text_input("Project Name (for file name)", "Example Project")
prefix = st.text_input("Global Prefix (e.g., CEC, TT)", "CEC")
agency_name = st.text_input("Full Agency Name", "California Energy Commission")
proceeding = st.text_input("CEC Proceeding Code (e.g. 24-OPT-05)", "24-OPT-05")

def generate_suffixes(n):
    suffixes = []
    reps = 1
    while len(suffixes) < n:
        for ch in ascii_lowercase:
            suffixes.append(ch * reps)
            if len(suffixes) == n:
                break
        reps += 1
    return suffixes

def generate_references(df, prefix, agency_name, proceeding):
    df = df[['TN #', 'Docketed Date', 'Document Title']].dropna()
    df['Year'] = pd.to_datetime(df['Docketed Date']).dt.year
    df['Formatted Title'] = df['Document Title'].apply(lambda x: str(x).split('\n')[0].strip())
    df = df.sort_values(by=['Year', 'Docketed Date']).reset_index(drop=True)
    df['GroupIndex'] = df.groupby('Year').cumcount()
    suffix_dict = {y: generate_suffixes(len(g)) for y, g in df.groupby('Year')}
    df['Suffix'] = df.apply(lambda r: suffix_dict[r['Year']][r['GroupIndex']], axis=1)
    df['Formatted Date'] = pd.to_datetime(df['Docketed Date']).dt.strftime('%B %-d, %Y')

    base_url = f"https://efiling.energy.ca.gov/Lists/DocketLog.aspx?docketnumber={proceeding}"
    references = [
        f"{prefix} {row['Year']}{row['Suffix']} – {agency_name} (TN {row['TN #']}). {row['Formatted Title']}. Docketed {row['Formatted Date']}. Accessed online at: {base_url}"
        for _, row in df.iterrows()
    ]
    return references

def build_docx(references):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    for ref in references:
        p = doc.add_paragraph(ref)
        p.paragraph_format.first_line_indent = Pt(-18)
        p.paragraph_format.left_indent = Pt(18)
    return doc

if uploaded_file and prefix and agency_name and proceeding:
    try:
        df = pd.read_excel(uploaded_file, engine="odf" if uploaded_file.name.endswith(".ods") else None)
        refs = generate_references(df, prefix, agency_name, proceeding)
        doc = build_docx(refs)
        buffer = BytesIO()
        doc.save(buffer)
        st.success("✅ Reference list ready!")
        st.download_button("📥 Download .docx Reference List", buffer.getvalue(), file_name=f"{project_name}_References.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        st.error(f"❌ Error: {e}")
