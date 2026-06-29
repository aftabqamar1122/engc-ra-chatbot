import streamlit as st
import anthropic
import os
import json
import io
from datetime import date
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    import fitz
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

st.set_page_config(
    page_title="ADOSH Risk Bot - ENGC",
    page_icon="🦺",
    layout="wide"
)

st.title("🦺 ADOSH Risk Assessment Bot")
st.caption("Exeed National General Contracting LLC | Bloom Living Projects | Abu Dhabi")
st.markdown("---")

SYSTEM_PROMPT = """
You are ADOSH Risk Bot, an expert HSE Risk Assessment generator for
Exeed National General Contracting LLC (ENGC), Abu Dhabi, UAE.

You ONLY generate Risk Assessments. Nothing else.

STRICT RULES:
1. Always follow ADOSH-SF Version 4.0 (July 2024)
2. Always use Abu Dhabi Codes of Practice
3. Always reference MOHRE Resolution No. 44/2022 (midday ban 12:30-15:00)
4. Always follow Hierarchy of Controls: Elimination to PPE
5. Risk Matrix: Probability x Severity
   - 1-4 = LOW
   - 5-9 = MODERATE
   - 10-14 = HIGH
   - 15-25 = EXTREME
6. Generate minimum 8 rows
7. Always reduce residual risk to LOW or MODERATE

Return ONLY this exact JSON. No extra text. No markdown. No explanation:
{
  "activity": "activity name",
  "project": "project name",
  "rows": [
    {
      "sn": 1,
      "activity_element": "activity step",
      "hazards": "hazard description",
      "consequences": "risk and consequences",
      "who_harmed": "who is harmed and exactly how",
      "prob_initial": 4,
      "sev_initial": 4,
      "risk_initial": 16,
      "risk_level_initial": "HIGH",
      "controls": "1. Elimination\n2. Substitution\n3. Engineering\n4. Administrative\n5. PPE",
      "prob_residual": 2,
      "sev_residual": 3,
      "risk_residual": 6,
      "risk_level_residual": "MODERATE"
    }
  ],
  "legal_references": "ADOSH CoP references"
}
"""

def get_risk_color(level):
    colors = {
        "LOW": "00AF50",
        "MODERATE": "FFFF00",
        "HIGH": "FFC000",
        "EXTREME": "FF0000"
    }
    return colors.get(str(level).upper(), "FFFFFF")

def shade_cell(cell, hex_color):
    try:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), hex_color)
        tcPr.append(shd)
    except:
        pass

def read_pdf_file(uploaded_file):
    if not PDF_SUPPORT:
        return "PDF reading not available", 0
    try:
        pdf_bytes = uploaded_file.read()
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        extracted_text = ""
        total_pages = len(pdf_doc)
        for page_num in range(total_pages):
            page = pdf_doc[page_num]
            extracted_text += f"\n--- Page {page_num + 1} ---\n{page.get_text()}"
        pdf_doc.close()
        if len(extracted_text) > 8000:
            extracted_text = extracted_text[:8000] + "\n[Truncated...]"
        return extracted_text, total_pages
    except Exception as e:
        return f"ERROR: {str(e)}", 0

def read_docx_file(uploaded_file):
    try:
        docx_bytes = io.BytesIO(uploaded_file.read())
        doc = Document(docx_bytes)
        extracted_text = ""
        for para in doc.paragraphs:
            if para.text.strip():
                extracted_text += para.text + "\n"
        for table in doc.tables:
            for row in table.rows:
                row_text = " | ".join(
                    cell.text.strip() for cell in row.cells if cell.text.strip()
                )
                if row_text:
                    extracted_text += row_text + "\n"
        if len(extracted_text) > 8000:
            extracted_text = extracted_text[:8000] + "\n[Truncated...]"
        return extracted_text, len(doc.paragraphs)
    except Exception as e:
        return f"ERROR: {str(e)}", 0

def generate_docx(ra_data, project_name, topic):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("RISK ASSESSMENT")
    run.bold = True
    run.font.size = Pt(14)
    run.font.name = 'Calibri'

    info_table = doc.add_table(rows=4, cols=2)
    info_table.style = 'Table Grid'
    info_data = [
        ("Entity Name:", "Exeed National General Contracting LLC (ENGC)"),
        ("Project:", project_name),
        ("Activity:", topic),
        ("Date:", date.today().strftime('%d-%b-%Y')),
    ]
    for i, (label, value) in enumerate(info_data):
        info_table.rows[i].cells[0].text = label
        info_table.rows[i].cells[1].text = value
        shade_cell(info_table.rows[i].cells[0], "C0C0C0")
        for cell in info_table.rows[i].cells:
            for para in cell.paragraphs:
                for r in para.runs:
                    r.font.size = Pt(9)
                    r.font.name = 'Calibri'

    doc.add_paragraph()

    table = doc.add_table(rows=2, cols=14)
    table.style = 'Table Grid'

    col_widths = [
        Cm(0.9), Cm(1.7), Cm(2.8), Cm(3.2), Cm(2.1),
        Cm(1.0), Cm(1.1), Cm(1.1),
        Cm(1.7), Cm(5.5),
        Cm(1.0), Cm(1.1), Cm(1.3), Cm(2.2)
    ]

    hdr1 = table.rows[0]
    hdr1.cells[5].merge(hdr1.cells[7])
    hdr1.cells[10].merge(hdr1.cells[12])

    h1 = {
        0: "S/N", 1: "Activity Element", 2: "Hazards / Impact",
        3: "Risk & Potential Consequences", 4: "Who Might Be Harmed and How?",
        5: "Risk Classification", 8: "Initial Risk Level",
        9: "Controls", 10: "Revised Risk Classification",
        13: "Residual Risk Level"
    }
    for col_idx, text in h1.items():
        cell = hdr1.cells[col_idx]
        cell.text = text
        shade_cell(cell, "C0C0C0")
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in para.runs:
                r.bold = True
                r.font.size = Pt(8)
                r.font.name = 'Calibri'
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    hdr2 = table.rows[1]
    h2 = {5: "P", 6: "S", 7: "R\nPxS", 10: "P", 11: "S", 12: "R\nPxS"}
    for col_idx in range(14):
        cell = hdr2.cells[col_idx]
        cell.text = h2.get(col_idx, "")
        shade_cell(cell, "BFBFBF")
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in para.runs:
                r.bold = True
                r.font.size = Pt(8)
                r.font.name = 'Calibri'
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for item in ra_data.get("rows", []):
        row = table.add_row()
        cells = row.cells
        data_map = {
            0: str(item.get("sn", "")),
            1: str(item.get("activity_element", "")),
            2: str(item.get("hazards", "")),
            3: str(item.get("consequences", "")),
            4: str(item.get("who_harmed", "")),
            5: str(item.get("prob_initial", "")),
            6: str(item.get("sev_initial", "")),
            7: str(item.get("risk_initial", "")),
            8: str(item.get("risk_level_initial", "")),
            9: str(item.get("controls", "")),
            10: str(item.get("prob_residual", "")),
            11: str(item.get("sev_residual", "")),
            12: str(item.get("risk_residual", "")),
            13: str(item.get("risk_level_residual", "")),
        }
        for col_idx, text in data_map.items():
            cells[col_idx].text = text
            cells[col_idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for para in cells[col_idx].paragraphs:
                for r in para.runs:
                    r.font.size = Pt(8)
                    r.font.name = 'Calibri'

        shade_cell(cells[8], get_risk_color(item.get("risk_level_initial", "LOW")))
        shade_cell(cells[13], get_risk_color(item.get("risk_level_residual", "LOW")))

    doc.add_paragraph()
    sig = doc.add_table(rows=2, cols=3)
    sig.style = 'Table Grid'
    for i, label in enumerate(["Prepared By:", "Reviewed By:", "Approved By:"]):
        sig.rows[0].cells[i].text = label
        shade_cell(sig.rows[0].cells[i], "C0C0C0")
        for para in sig.rows[0].cells[i].paragraphs:
            for r in para.runs:
                r.bold = True
                r.font.size = Pt(9)
                r.font.name = 'Calibri'
    designations = ["HSE Engineer", "HSE Manager", "Project Manager"]
    for i, desig in enumerate(designations):
        sig.rows[1].cells[i].text = (
            f"Name: _______________\n"
            f"Designation: {desig}\n"
            f"Date: ____________\n"
            f"Signature: ________"
        )
        for para in sig.rows[1].cells[i].paragraphs:
            for r in para.runs:
                r.font.size = Pt(8)
                r.font.name = 'Calibri'

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def call_claude_api(prompt, project_name):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}]
    )
    response_text = message.content[0].text.strip()

    if "```json" in response_text:
        response_text = response_text.split("```json")[1].split("```")[0].strip()
    elif "```" in response_text:
        response_text = response_text.split("```")[1].split("```")[0].strip()

    start = response_text.find('{')
    end = response_text.rfind('}')
    if start != -1 and end != -1:
        response_text = response_text[start:end+1]

    return json.loads(response_text)

def show_results(ra_data, project_name, topic):
    st.success(f"✅ Generated {len(ra_data['rows'])} activities!")

    preview = []
    for row in ra_data["rows"]:
        preview.append({
            "S/N": row.get("sn", ""),
            "Activity": str(row.get("activity_element", ""))[:45],
            "Hazard": str(row.get("hazards", ""))[:35] + "...",
            "Initial Risk": f"{row.get('risk_level_initial','')} ({row.get('risk_initial','')})",
            "Residual Risk": f"{row.get('risk_level_residual','')} ({row.get('risk_residual','')})"
        })
    st.dataframe(preview, use_container_width=True)

    try:
        docx_buffer = generate_docx(ra_data, project_name, topic)
        filename = f"ENGC_RA_{topic.replace(' ','_')}_{date.today().strftime('%d%b%Y')}.docx"
        st.download_button(
            label="⬇️ Download Risk Assessment (DOCX)",
            data=docx_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        st.success("✅ DOCX ready — click the Download button above!")
    except Exception as e:
        st.error(f"❌ DOCX error: {str(e)}")
        st.download_button(
            label="⬇️ Download Raw JSON",
            data=json.dumps(ra_data, indent=2),
            file_name=f"RA_{topic}_{date.today()}.json",
            mime="application/json",
            use_container_width=True
        )

    if ra_data.get("legal_references"):
        with st.expander("📚 Legal References"):
            st.write(ra_data["legal_references"])

# ── SIDEBAR ──
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    project_name = st.selectbox("Select Project:", [
        "Bloom Living – Almeria (Residential Villas)",
        "Bloom Living – Cordoba Community Center",
        "Bloom Living – Toledo Community Center",
        "Bloom Living Service Center (Plot C13, Zayed City)"
    ])
    st.markdown("---")
    st.markdown("**📋 Quick Topics:**")
    quick_topics = [
        "Fire Prevention", "Working at Heights", "Excavation Works",
        "Crane Operations", "Electrical Works", "Confined Space Entry",
        "Heat Stress Management", "Night Shift Work", "Chemical Handling",
        "Scaffolding Erection", "Concrete Works", "Lifting Operations"
    ]
    for topic in quick_topics:
        if st.button(f"📄 {topic}", key=f"btn_{topic}", use_container_width=True):
            st.session_state.quick_input = topic

# ── MAIN TABS ──
tab1, tab2 = st.tabs(["✏️ Type a Topic", "📎 Upload PDF / DOCX"])

# ── TAB 1 ──
with tab1:
    st.info("Type any construction activity → Click Generate → Download DOCX")
    user_input = st.text_input(
        "🔍 Enter Activity / Topic:",
        value=st.session_state.get("quick_input", ""),
        placeholder="e.g. Scaffolding Erection, Crane Operations..."
    )
    if st.button("🚀 Generate Risk Assessment", type="primary",
                 use_container_width=True, key="gen_tab1"):
        if not user_input.strip():
            st.warning("Please enter a topic first.")
        else:
            with st.spinner(f"Generating RA for: {user_input}..."):
                try:
                    prompt = (
                        f"Generate a complete Risk Assessment for: {user_input}\n"
                        f"Project: {project_name}\n"
                        f"Location: Abu Dhabi, UAE\n"
                        f"Season: Summer. Always include heat stress controls."
                    )
                    ra_data = call_claude_api(prompt, project_name)
                    show_results(ra_data, project_name, user_input)
                except json.JSONDecodeError:
                    st.error("⚠️ Parse error — please click Generate again.")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

# ── TAB 2 ──
with tab2:
    st.info("Upload a PDF or DOCX → Bot reads it → Generates RA automatically")
    uploaded_file = st.file_uploader(
        "Choose file:",
        type=["pdf", "docx", "doc"],
        help="Method Statement, existing RA, or work procedure"
    )
    if uploaded_file:
        file_size = len(uploaded_file.getvalue()) / 1024
        st.success(f"✅ {uploaded_file.name} ({file_size:.1f} KB)")

        extra = st.text_area(
            "Additional instructions (optional):",
            placeholder="e.g. Focus on excavation only, add 10 rows minimum...",
            height=70
        )

        if st.button("🚀 Generate RA from File", type="primary",
                     use_container_width=True, key="gen_tab2"):
            with st.spinner("Reading file and generating RA..."):
                uploaded_file.seek(0)
                if uploaded_file.name.lower().endswith(".pdf"):
                    file_text, pages = read_pdf_file(uploaded_file)
                    label = f"PDF ({pages} pages)"
                else:
                    file_text, _ = read_docx_file(uploaded_file)
                    label = "DOCX"

                if "ERROR" in str(file_text):
                    st.error(f"Cannot read file: {file_text}")
                else:
                    st.success(f"✅ File read: {label}")
                    topic_name = uploaded_file.name.rsplit(".", 1)[0].replace("_", " ")
                    prompt = (
                        f"Analyze this document and generate a complete ADOSH-SF "
                        f"Risk Assessment for all activities described.\n"
                        f"Project: {project_name}\n"
                        f"Location: Abu Dhabi, UAE\n"
                        f"Season: Summer. Always include heat stress controls.\n"
                        f"Document content:\n{file_text}\n"
                        f"{'Instructions: ' + extra if extra else ''}"
                    )
                    try:
                        ra_data = call_claude_api(prompt, project_name)
                        show_results(ra_data, project_name, topic_name)
                    except json.JSONDecodeError:
                        st.error("⚠️ Parse error — please try again.")
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")

st.markdown("---")
st.caption(
    "⚠️ AI output must be reviewed by a competent HSE professional. "
    "| ADOSH-SF Version 4.0 | © ENGC 2026"
)
