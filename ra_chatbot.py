# Imports
import streamlit as st
import anthropic
import os
import json
import io
import base64
from datetime import date
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── NEW: File reading libraries ──
import fitz  # PyMuPDF — reads PDF

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ADOSH Risk Bot – ENGC",
    page_icon="🦺",
    layout="wide"
)

st.title("🦺 ADOSH Risk Assessment Bot")
st.caption("Exeed National General Contracting LLC | Bloom Living Projects | Abu Dhabi")
st.markdown("---")

# ─────────────────────────────────────────────
# SYSTEM PROMPT (This is your RA expert brain)
# ─────────────────────────────────────────────
SYSTEM_PROMPT = """
You are ADOSH Risk Bot, an expert HSE Risk Assessment generator for 
Exeed National General Contracting LLC (ENGC), Abu Dhabi, UAE.

You ONLY generate Risk Assessments. Nothing else.

STRICT RULES:
1. Always follow ADOSH-SF Version 4.0 (July 2024)
2. Always use Abu Dhabi Codes of Practice (CoP 2.0, 11.0, 15.0, 21.0, 34.0, 35.0, 36.0, 44.0, 53.1)
3. Always reference MOHRE Resolution No. 44/2022 (midday ban 12:30-15:00, June 15 – September 15)
4. Always follow Hierarchy of Controls: Elimination → Substitution → Engineering → Administrative → PPE
5. Risk Matrix: Probability (1-5) × Severity (1-5)
   - 1-4 = LOW (Green)
   - 5-9 = MODERATE (Yellow)  
   - 10-14 = HIGH (Orange)
   - 15-25 = EXTREME (Red)
6. Generate 8-12 activity rows minimum
7. Always reduce residual risk to LOW or MODERATE

OUTPUT FORMAT — Return ONLY this JSON structure, no other text:
{
  "activity": "Name of activity assessed",
  "project": "Project name",
  "rows": [
    {
      "sn": 1,
      "activity_element": "Specific activity step",
      "hazards": "Hazard identified",
      "consequences": "Risk and potential consequences",
      "who_harmed": "Who might be harmed and HOW specifically",
      "prob_initial": 4,
      "sev_initial": 4,
      "risk_initial": 16,
      "risk_level_initial": "HIGH",
      "controls": "1. Elimination measure\\n2. Substitution\\n3. Engineering control\\n4. Admin control (ADOSH CoP ref)\\n5. PPE required",
      "prob_residual": 2,
      "sev_residual": 3,
      "risk_residual": 6,
      "risk_level_residual": "MODERATE"
    }
  ],
  "legal_references": "ADOSH-SF CoP references used"
}

Return ONLY valid JSON. No markdown. No explanation. No preamble.
"""

# ─────────────────────────────────────────────
# RISK LEVEL COLOR HELPER
# ─────────────────────────────────────────────
def get_risk_color(level):
    colors = {
        "LOW": "00AF50",
        "MODERATE": "FFFF00", 
        "HIGH": "FFC000",
        "EXTREME": "FF0000"
    }
    return colors.get(level.upper(), "FFFFFF")

def shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

# ─────────────────────────────────────────────
# GENERATE DOCX FROM RA DATA
# ─────────────────────────────────────────────
def generate_docx(ra_data, project_name, topic):
    doc = Document()
    
    # Page setup - A4 Landscape
    section = doc.sections[0]
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)

    # ── HEADER INFO BLOCK ──
    header_table = doc.add_table(rows=2, cols=3)
    header_table.style = 'Table Grid'
    
    cells = header_table.rows[0].cells
    cells[0].text = "Entity Name: Exeed National General Contracting LLC (ENGC)"
    cells[1].text = f"Project: {project_name}"
    cells[2].text = f"Date: {date.today().strftime('%d-%b-%Y')}"
    
    cells = header_table.rows[1].cells
    cells[0].text = f"Activity: {topic}"
    cells[1].text = "Document Ref: ENGC/OSH-RA/BLM/001"
    cells[2].text = "Revision: 00"

    for row in header_table.rows:
        for cell in row.cells:
            shade_cell(cell, "C0C0C0")
            for para in cell.paragraphs:
                for run in para.runs:
                    run.bold = True
                    run.font.size = Pt(9)
                    run.font.name = 'Calibri'

    doc.add_paragraph()

    # ── MAIN RA TABLE ──
    # 14 columns
    col_widths_cm = [0.9, 1.7, 2.8, 3.2, 2.1, 1.0, 1.1, 1.1, 1.7, 5.8, 1.0, 1.1, 1.3, 2.4]
    
    table = doc.add_table(rows=3, cols=14)
    table.style = 'Table Grid'

    # ── HEADER ROW 1 ──
    row1 = table.rows[0]
    headers_row1 = [
        "S/N", "Activity Element", "Hazards / Impact",
        "Risk & Potential Consequences", "Who Might Be Harmed and How?",
        "Risk Classification", "", "",
        "Initial Risk Level", "Controls",
        "Revised Risk Classification", "", "",
        "Residual Risk Level"
    ]
    
    # Merge Risk Classification cols 5-7
    row1.cells[5].merge(row1.cells[7])
    # Merge Revised Risk Classification cols 10-12
    row1.cells[10].merge(row1.cells[12])
    
    col_labels = ["S/N", "Activity Element", "Hazards / Impact",
                  "Risk & Potential Consequences", "Who Might Be Harmed and How?",
                  "Risk Classification", "Initial Risk Level", "Controls",
                  "Revised Risk Classification", "Residual Risk Level"]
    
    for i, cell in enumerate(row1.cells):
        shade_cell(cell, "C0C0C0")
        
    # Set merged header text
    row1.cells[0].text = "S/N"
    row1.cells[1].text = "Activity Element"
    row1.cells[2].text = "Hazards / Impact"
    row1.cells[3].text = "Risk & Potential Consequences"
    row1.cells[4].text = "Who Might Be Harmed and How?"
    row1.cells[5].text = "Risk Classification"
    row1.cells[8].text = "Initial Risk Level"
    row1.cells[9].text = "Controls"
    row1.cells[10].text = "Revised Risk Classification"
    row1.cells[13].text = "Residual Risk Level"

    # ── HEADER ROW 2 — P | S | R subheaders ──
    row2 = table.rows[1]
    subheaders = ["", "", "", "", "", "P", "S", "R\nP×S", "", "", "P", "S", "R\nP×S", ""]
    
    for i, text in enumerate(subheaders):
        row2.cells[i].text = text
        shade_cell(row2.cells[i], "BFBFBF")

    # Format all header cells
    for row in [row1, row2]:
        for cell in row.cells:
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    run.bold = True
                    run.font.size = Pt(8)
                    run.font.name = 'Calibri'
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # ── DATA ROWS ──
    # Remove the empty 3rd row
    table._tbl.remove(table.rows[2]._tr)
    
    for item in ra_data.get("rows", []):
        row = table.add_row()
        cells = row.cells
        
        cells[0].text = str(item["sn"])
        cells[1].text = item["activity_element"]
        cells[2].text = item["hazards"]
        cells[3].text = item["consequences"]
        cells[4].text = item["who_harmed"]
        cells[5].text = str(item["prob_initial"])
        cells[6].text = str(item["sev_initial"])
        cells[7].text = str(item["risk_initial"])
        
        # Initial Risk Level with color
        cells[8].text = item["risk_level_initial"]
        shade_cell(cells[8], get_risk_color(item["risk_level_initial"]))
        
        cells[9].text = item["controls"]
        cells[10].text = str(item["prob_residual"])
        cells[11].text = str(item["sev_residual"])
        cells[12].text = str(item["risk_residual"])
        
        # Residual Risk Level with color
        cells[13].text = item["risk_level_residual"]
        shade_cell(cells[13], get_risk_color(item["risk_level_residual"]))
        
        # Format all cells
        for cell in cells:
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(8)
                    run.font.name = 'Calibri'

    # ── SIGNATURE BLOCK ──
    doc.add_paragraph()
    sig_table = doc.add_table(rows=2, cols=3)
    sig_table.style = 'Table Grid'
    
    sig_headers = ["Prepared By:", "Reviewed By:", "Approved By:"]
    sig_names = ["HSE Engineer", "HSE Manager", "Project Manager"]
    
    for i, (header, name) in enumerate(zip(sig_headers, sig_names)):
        sig_table.rows[0].cells[i].text = header
        sig_table.rows[1].cells[i].text = f"Name: _______________\nDesignation: {name}\nDate: ________________\nSignature: ____________"
        shade_cell(sig_table.rows[0].cells[i], "C0C0C0")
        for para in sig_table.rows[0].cells[i].paragraphs:
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(9)
                run.font.name = 'Calibri'

    # Save to buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ─────────────────────────────────────────────
# MAIN APP UI
# ─────────────────────────────────────────────

# Sidebar
with st.sidebar:
    st.image("https://via.placeholder.com/200x60?text=ENGC+Logo", width=200)
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
        if st.button(f"📄 {topic}", key=topic, use_container_width=True):
            st.session_state.quick_input = topic

st.markdown("### 📝 Generate Risk Assessment")
st.info("**How to use:** Type any construction activity below → Click Generate → Download your DOCX")

# Input
user_input = st.text_input(
    "🔍 Enter Activity / Topic:",
    value=st.session_state.get("quick_input", ""),
    placeholder="e.g. Scaffolding Erection, Crane Operations, Confined Space Entry..."
)

if st.button("🚀 Generate Risk Assessment", type="primary", use_container_width=True):
    if not user_input.strip():
        st.warning("Please enter an activity topic first.")
    else:
        with st.spinner(f"🔄 Generating ADOSH-SF compliant RA for: **{user_input}**..."):
            try:
                # Call Claude API
                client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
                
                message = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=4000,
                    system=SYSTEM_PROMPT,
                    messages=[
                        {
                            "role": "user",
                            "content": f"Generate a complete Risk Assessment for: {user_input}\nProject: {project_name}\nLocation: Abu Dhabi, UAE\nSeason: Summer (Heat stress must be included in controls)"
                        }
                    ]
                )
                
                # Parse JSON response
                import json
                response_text = message.content[0].text.strip()
                
                # Clean if needed
                if response_text.startswith("```"):
                    response_text = response_text.split("```")[1]
                    if response_text.startswith("json"):
                        response_text = response_text[4:]
                
                ra_data = json.loads(response_text)
                
                st.success(f"✅ Risk Assessment Generated: **{len(ra_data['rows'])} activities** identified")
                
                # Show preview table
                st.markdown("### 📊 Preview")
                preview_data = []
                for row in ra_data["rows"]:
                    preview_data.append({
                        "S/N": row["sn"],
                        "Activity": row["activity_element"],
                        "Hazard": row["hazards"][:50] + "...",
                        "Initial Risk": f"{row['risk_level_initial']} ({row['risk_initial']})",
                        "Residual Risk": f"{row['risk_level_residual']} ({row['risk_residual']})"
                    })
                st.dataframe(preview_data, use_container_width=True)
                
                # Generate and offer DOCX download
                docx_buffer = generate_docx(ra_data, project_name, user_input)
                
                filename = f"ENGC_RA_{user_input.replace(' ', '_')}_{date.today().strftime('%d%b%Y')}.docx"
                
                st.download_button(
                    label="⬇️ Download Risk Assessment (DOCX)",
                    data=docx_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                # Show legal references
                if ra_data.get("legal_references"):
                    with st.expander("📚 Legal References Used"):
                        st.write(ra_data["legal_references"])
                        
            except json.JSONDecodeError:
                st.error("⚠️ Error parsing response. Please try again.")
                st.code(response_text)
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

st.markdown("---")
st.caption("⚠️ AI output must be reviewed by a competent HSE professional before submission. | ADOSH-SF Version 4.0 | © ENGC 2026")