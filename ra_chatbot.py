import streamlit as st
import anthropic
import os
import json
import io
import re
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

SYSTEM_PROMPT = """You are ADOSH Risk Bot for ENGC Abu Dhabi UAE.
Generate Risk Assessments following ADOSH-SF Version 4.0.
Include MOHRE Resolution No. 44/2022 heat stress controls.
Hierarchy of Controls: Elimination, Substitution, Engineering, Administrative, PPE.
Risk = Probability x Severity. LOW=1-4, MODERATE=5-9, HIGH=10-14, EXTREME=15-25.

Generate EXACTLY 10 rows specific to the activity requested.
Analyze the activity and break it into 10 real construction steps.
Each row must be a DIFFERENT specific step of that exact activity.
Think like this for any activity:
- Excavation: site setup, marking, breaking, shoring, dewatering, soil removal, underground services, backfill, compaction, reinstatement
- Scaffolding: delivery, ground prep, base plates, standards, ledgers, transoms, boarding, handrails, ladders, inspection
- Electrical: isolation, permit to work, cable routing, pulling, termination, testing, energizing, earthing, panel work, inspection
- Concrete: formwork, reinforcement, delivery, pouring, compaction, curing, stripping, finishing, testing, waste
- Working at Heights: equipment check, edge protection, harness, anchors, platform, tool tethering, weather, rescue plan, supervision, descent
- Fire Prevention: site assessment, ignition sources, storage, detection, suppression, signage, evacuation, training, hot works, inspection
- Chemical Handling: MSDS review, storage, PPE selection, mixing, spillage, disposal, ventilation, emergency, first aid, inspection
- Lifting Operations: pre-lift plan, crane inspection, outrigger setup, rigging, slinging, trial lift, main lift, slewing, load landing, post-lift
- Confined Space: risk assessment, permit to work, atmospheric testing, ventilation, entry, monitoring, communication, rescue, exit, debrief
Apply same thinking to ANY other activity requested.
Row 9 always = Heat Stress (MOHRE Resolution 44/2022, midday ban 12:30-15:00).
Row 10 always = Emergency Response and First Aid.
Keep each text field under 50 words. Residual risk must be LOW or MODERATE.

OUTPUT RULES - CRITICAL:
1. Output ONLY a JSON object
2. Start with { and end with }
3. Zero text before or after the JSON
4. No markdown, no code blocks, no explanation

JSON format:
{
  "activity": "name",
  "project": "name",
  "rows": [
    {
      "sn": 1,
      "activity_element": "step",
      "hazards": "hazard",
      "consequences": "consequences",
      "who_harmed": "who and how",
      "prob_initial": 4,
      "sev_initial": 4,
      "risk_initial": 16,
      "risk_level_initial": "HIGH",
      "controls": "1. Control one\n2. Control two\n3. Control three\n4. Control four\n5. PPE required",
      "prob_residual": 2,
      "sev_residual": 3,
      "risk_residual": 6,
      "risk_level_residual": "MODERATE"
    }
  ],
  "legal_references": "ADOSH CoP references used"
}"""


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
    except Exception:
        pass


def extract_json(text):
    text = text.strip()
    if "```json" in text:
        text = text.split("```json")[1]
        if "```" in text:
            text = text.split("```")[0]
        text = text.strip()
    elif "```" in text:
        text = text.split("```")[1]
        if "```" in text:
            text = text.split("```")[0]
        text = text.strip()
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError("No JSON found")
    text = text[start:end + 1]
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    last = text.rfind('"}')
    if last > 0:
        truncated = text[:last + 2]
        ob = truncated.count('{') - truncated.count('}')
        ob2 = truncated.count('[') - truncated.count(']')
        if ob2 > 0:
            truncated += ']' * ob2
        if ob > 0:
            truncated += '}' * ob
        try:
            return json.loads(truncated)
        except json.JSONDecodeError:
            pass
    text = re.sub(r',\s*}', '}', text)
    text = re.sub(r',\s*]', ']', text)
    return json.loads(text)


def read_pdf_file(uploaded_file):
    if not PDF_SUPPORT:
        return "PDF not supported", 0
    try:
        pdf_bytes = uploaded_file.read()
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = ""
        pages = len(pdf_doc)
        for i in range(pages):
            text += f"\n--- Page {i+1} ---\n{pdf_doc[i].get_text()}"
        pdf_doc.close()
        return text[:8000], pages
    except Exception as e:
        return f"ERROR: {str(e)}", 0


def read_docx_file(uploaded_file):
    try:
        doc = Document(io.BytesIO(uploaded_file.read()))
        text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        return text[:8000], len(doc.paragraphs)
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

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("RISK ASSESSMENT")
    r.bold = True
    r.font.size = Pt(14)
    r.font.name = "Calibri"

    t = doc.add_table(rows=4, cols=2)
    t.style = "Table Grid"
    info_rows = [
        ("Entity Name:", "Exeed National General Contracting LLC (ENGC)"),
        ("Project:", project_name),
        ("Activity:", topic),
        ("Date:", date.today().strftime("%d-%b-%Y")),
    ]
    for i, (lbl, val) in enumerate(info_rows):
        t.rows[i].cells[0].text = lbl
        t.rows[i].cells[1].text = val
        shade_cell(t.rows[i].cells[0], "C0C0C0")
        for cell in t.rows[i].cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(9)
                    run.font.name = "Calibri"

    doc.add_paragraph()

    tbl = doc.add_table(rows=2, cols=14)
    tbl.style = "Table Grid"

    h1 = tbl.rows[0]
    h1.cells[5].merge(h1.cells[7])
    h1.cells[10].merge(h1.cells[12])

    header1 = {
        0: "S/N",
        1: "Activity Element",
        2: "Hazards / Impact",
        3: "Risk & Potential Consequences",
        4: "Who Might Be Harmed and How?",
        5: "Risk Classification",
        8: "Initial Risk Level",
        9: "Controls",
        10: "Revised Risk Classification",
        13: "Residual Risk Level"
    }
    for idx, txt in header1.items():
        c = h1.cells[idx]
        c.text = txt
        shade_cell(c, "C0C0C0")
        for para in c.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(8)
                run.font.name = "Calibri"
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    h2 = tbl.rows[1]
    sub = {5: "P", 6: "S", 7: "R\nPxS", 10: "P", 11: "S", 12: "R\nPxS"}
    for idx in range(14):
        c = h2.cells[idx]
        c.text = sub.get(idx, "")
        shade_cell(c, "BFBFBF")
        for para in c.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(8)
                run.font.name = "Calibri"
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for item in ra_data.get("rows", []):
        row = tbl.add_row()
        vals = {
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
        for idx, val in vals.items():
            row.cells[idx].text = val
            row.cells[idx].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for para in row.cells[idx].paragraphs:
                for run in para.runs:
                    run.font.size = Pt(8)
                    run.font.name = "Calibri"
        shade_cell(row.cells[8],
                   get_risk_color(item.get("risk_level_initial", "LOW")))
        shade_cell(row.cells[13],
                   get_risk_color(item.get("risk_level_residual", "LOW")))

    doc.add_paragraph()
    sig = doc.add_table(rows=2, cols=3)
    sig.style = "Table Grid"
    for i, lbl in enumerate(["Prepared By:", "Reviewed By:", "Approved By:"]):
        sig.rows[0].cells[i].text = lbl
        shade_cell(sig.rows[0].cells[i], "C0C0C0")
        for para in sig.rows[0].cells[i].paragraphs:
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(9)
                run.font.name = "Calibri"
    for i, d in enumerate(["HSE Engineer", "HSE Manager", "Project Manager"]):
        sig.rows[1].cells[i].text = (
            "Name: _______________\n"
            "Designation: " + d + "\n"
            "Date: ____________\n"
            "Signature: ________"
        )
        for para in sig.rows[1].cells[i].paragraphs:
            for run in para.runs:
                run.font.size = Pt(8)
                run.font.name = "Calibri"

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_prompt(activity, project):
    return (
        "Generate Risk Assessment for the following construction activity.\n"
        "Activity: " + activity + "\n"
        "Project: " + project + "\n"
        "Location: Abu Dhabi UAE\n"
        "Season: Summer\n"
        "INSTRUCTIONS:\n"
        "1. Break the activity into exactly 10 specific construction steps\n"
        "2. Every row must be unique and specific to the activity\n"
        "3. No generic rows - only steps that belong to this specific activity\n"
        "4. Row 9 must be Heat Stress - MOHRE Resolution 44/2022 midday ban\n"
        "5. Row 10 must be Emergency Response and First Aid\n"
        "6. Keep all text fields under 50 words each"
    )


def build_file_prompt(activity, project, file_text, extra):
    prompt = (
        "Analyze this document and generate a Risk Assessment.\n"
        "Activity: " + activity + "\n"
        "Project: " + project + "\n"
        "Location: Abu Dhabi UAE\n"
        "Season: Summer\n"
        "INSTRUCTIONS:\n"
        "1. Extract all activities from the document\n"
        "2. Generate 10 specific rows based on document content\n"
        "3. Row 9 must be Heat Stress\n"
        "4. Row 10 must be Emergency Response\n"
        "5. Keep all text fields under 50 words\n"
        "Document content:\n" + file_text
    )
    if extra:
        prompt += "\nAdditional instructions: " + extra
    return prompt


def call_api_and_show(prompt, project_name, topic):
    try:
        client = anthropic.Anthropic(
            api_key=os.environ.get("ANTHROPIC_API_KEY")
        )
        msg = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=8192,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = msg.content[0].text

        with st.expander("🔍 Debug - Raw API Response"):
            st.text(raw[:500])

        ra_data = extract_json(raw)

        if "rows" not in ra_data or len(ra_data["rows"]) == 0:
            st.error("No rows generated. Please try again.")
            return

        st.success("✅ Generated " + str(len(ra_data["rows"])) + " activity rows!")

        preview = []
        for r in ra_data["rows"]:
            preview.append({
                "S/N": r.get("sn", ""),
                "Activity": str(r.get("activity_element", ""))[:40],
                "Hazard": str(r.get("hazards", ""))[:30] + "...",
                "Initial": str(r.get("risk_level_initial", "")) + "(" + str(r.get("risk_initial", "")) + ")",
                "Residual": str(r.get("risk_level_residual", "")) + "(" + str(r.get("risk_residual", "")) + ")"
            })
        st.dataframe(preview, use_container_width=True)

        try:
            buf = generate_docx(ra_data, project_name, topic)
            fname = "ENGC_RA_" + topic.replace(" ", "_") + "_" + date.today().strftime("%d%b%Y") + ".docx"
            st.download_button(
                label="⬇️ Download Risk Assessment (DOCX)",
                data=buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            st.success("✅ Click the button above to download your DOCX!")
        except Exception as de:
            st.error("DOCX error: " + str(de))
            st.download_button(
                "⬇️ Download JSON (backup)",
                data=json.dumps(ra_data, indent=2),
                file_name="RA_" + topic + ".json",
                mime="application/json",
                use_container_width=True
            )

        if ra_data.get("legal_references"):
            with st.expander("📚 Legal References"):
                st.write(ra_data["legal_references"])

    except json.JSONDecodeError as je:
        st.error("JSON parse error: " + str(je))
        st.info("Please click Generate again — usually works on retry.")
    except Exception as e:
        st.error("Error: " + str(e))
        if "api_key" in str(e).lower() or "auth" in str(e).lower():
            st.warning("Check your API key in Streamlit Secrets!")


# ── SIDEBAR ──
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    project_name = st.selectbox("Select Project:", [
        "Bloom Living – Almeria (Residential Villas)",
        "Bloom Living – Cordoba Community Center",
        "Bloom Living – Toledo Community Center",
        "Bloom Living Service Center (Plot C13, Zayed City)",
    ])
    st.markdown("---")
    st.markdown("**📋 Quick Topics:**")
    for topic in [
        "Fire Prevention", "Working at Heights", "Excavation Works",
        "Crane Operations", "Electrical Works", "Confined Space Entry",
        "Heat Stress Management", "Night Shift Work", "Chemical Handling",
        "Scaffolding Erection", "Concrete Works", "Lifting Operations"
    ]:
        if st.button("📄 " + topic, key="q_" + topic,
                     use_container_width=True):
            st.session_state.quick_input = topic

# ── TABS ──
tab1, tab2 = st.tabs(["✏️ Type a Topic", "📎 Upload PDF / DOCX"])

with tab1:
    st.info("Type any construction activity → Click Generate → Download DOCX")
    user_input = st.text_input(
        "🔍 Enter Activity / Topic:",
        value=st.session_state.get("quick_input", ""),
        placeholder="e.g. Excavation Works, Scaffolding Erection..."
    )
    if st.button("🚀 Generate Risk Assessment", type="primary",
                 use_container_width=True, key="g1"):
        if not user_input.strip():
            st.warning("Please enter a topic first.")
        else:
            with st.spinner("⏳ Generating RA for: " + user_input + "..."):
                call_api_and_show(
                    build_prompt(user_input, project_name),
                    project_name,
                    user_input
                )

with tab2:
    st.info("Upload PDF or DOCX → Bot reads it → Generates RA")
    uf = st.file_uploader("Choose file:", type=["pdf", "docx", "doc"])
    if uf:
        st.success("✅ " + uf.name + " (" + str(round(len(uf.getvalue())/1024, 1)) + " KB)")
        extra = st.text_area("Extra instructions:", height=60)
        if st.button("🚀 Generate RA from File", type="primary",
                     use_container_width=True, key="g2"):
            with st.spinner("Reading file..."):
                uf.seek(0)
                if uf.name.lower().endswith(".pdf"):
                    txt, pg = read_pdf_file(uf)
                else:
                    txt, pg = read_docx_file(uf)
                if "ERROR" in str(txt):
                    st.error(txt)
                else:
                    tname = uf.name.rsplit(".", 1)[0].replace("_", " ")
                    call_api_and_show(
                        build_file_prompt(tname, project_name, txt, extra),
                        project_name,
                        tname
                    )

st.markdown("---")
st.caption(
    "⚠️ AI output must be reviewed by a competent HSE professional. "
    "| ADOSH-SF Version 4.0 | © ENGC 2026"
)
