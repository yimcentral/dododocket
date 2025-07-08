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

def add_hyperlink(paragraph, url, text):
    # Create the relationship to the URL
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )

    # Create the w:hyperlink tag
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    # Create a run with your style
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    # Remove underline and set color to black
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "none")

    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")

    font = OxmlElement("w:rFonts")
    font.set(qn("w:ascii"), "Tahoma")

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "24")  # 12pt font

    rPr.append(font)
    rPr.append(sz)
    rPr.append(color)
    rPr.append(u)
    new_run.append(rPr)

    text_elem = OxmlElement("w:t")
    text_elem.text = text
    new_run.append(text_elem)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def build_docx(references):
    doc = Document()

    # Set default Normal style
    style = doc.styles['Normal']
    style.font.name = 'Tahoma'
    style.font.size = Pt(12)
    pformat = style.paragraph_format
    pformat.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pformat.space_before = Pt(6)
    pformat.space_after = Pt(6)

    # Heading style
    heading_style = doc.styles['Heading 2']
    heading_style.font.name = 'Tahoma'
    heading_style.font.size = Pt(12)
    heading_style.paragraph_format.space_before = Pt(12)
    heading_style.paragraph_format.space_after = Pt(6)

    for ref_text, url in references:
        p = doc.add_paragraph()
        run = p.add_run(ref_text + " ")
        run.font.name = 'Tahoma'
        run.font.size = Pt(12)
        run.font.underline = False
        run.font.color.rgb = None
        p.paragraph_format.first_line_indent = Pt(-18)
        p.paragraph_format.left_indent = Pt(18)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)

        # Add styled hyperlink
        add_hyperlink(p, url, url)

    return doc

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
