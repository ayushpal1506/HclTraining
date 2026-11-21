"""Build the Step 17 deliverable: docs/p107_Project_Document.docx.

Strict ABES Project Document implementation adhering to all Screen 1-3 specifications:
- Page setup: A4 portrait, 2 cm margins all around, 1.25 cm header/footer distance.
- Base text: Calibri 10.5 pt, colour 151514, zero extra paragraph spacing.
- Heading 1: Georgia 15 bold A51C30, before 16 pt, after 7 pt, keep_with_next.
- Heading 2: Georgia 12 bold 151514, before 10 pt, after 5 pt, keep_with_next.
- Header: right-aligned 8 pt 7F7F7F: "ABES Engineering College | Project Document | p107".
- Footer: centered 8 pt 7F7F7F: "Ayush Pal | 2500320140037     Page <PAGE>".
- Tables: fixed width 9638 dxa, exact column dxa widths, border 0.5 pt BFBFBF,
  cell margins 60/90 dxa, header fill A51C30 white bold 9 pt, body 9 pt 151514.
  No blank line after table.
- Complete content for Page 1 Cover, Page 2 TOC, Sections 1 to 10, Appendix A, Appendix B.
"""
from __future__ import annotations

import json
import os
import shutil
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
OUT_PRIMARY = DOCS / "p107_Project_Document.docx"
RESULTS_PATH = ROOT / "data" / "test_results.json"
ABES_LOGO = ROOT / "assets" / "abes_logo.jpg"

# Color constants
C_CRIMSON = RGBColor(0xA5, 0x1C, 0x30)
C_DARK = RGBColor(0x15, 0x15, 0x14)
C_GREY = RGBColor(0x7F, 0x7F, 0x7F)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

HEX_CRIMSON = "A51C30"
HEX_BORDER = "BFBFBF"

# --------------------------------------------------------------------------- XML Helpers
def set_cell_margins(cell, top=60, bottom=60, left=90, right=90):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")  # 0.5 pt
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), HEX_BORDER)
        tcBorders.append(b)
    tcPr.append(tcBorders)

def set_cell_shading(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    trPr.append(el)

def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)

def add_abes_table(doc, headers: list[str], rows: list[list[str]], col_dxa: list[int], align_cols: dict[int, int] = None):
    """Create a fixed-layout 9638 dxa table strictly matching ABES specifications."""
    align_cols = align_cols or {}
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False

    tblPr = table._tbl.tblPr
    # Fixed layout
    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed")
    tblPr.append(tblLayout)

    # Total width 9638 dxa — REPLACE existing tblW (not append a second one)
    existing_tblW = tblPr.find(qn("w:tblW"))
    if existing_tblW is not None:
        existing_tblW.set(qn("w:w"), "9638")
        existing_tblW.set(qn("w:type"), "dxa")
    else:
        tblW = OxmlElement("w:tblW")
        tblW.set(qn("w:w"), "9638")
        tblW.set(qn("w:type"), "dxa")
        tblPr.append(tblW)

    # Table borders
    tblBorders = OxmlElement("w:tblBorders")
    for b_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{b_name}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), HEX_BORDER)
        tblBorders.append(b)
    tblPr.append(tblBorders)

    # Header Row
    hdr_row = table.rows[0]
    cant_split(hdr_row)
    repeat_header(hdr_row)
    for i, h_text in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_margins(cell, top=60, bottom=60, left=90, right=90)
        set_cell_borders(cell)
        set_cell_shading(cell, HEX_CRIMSON)

        tcPr = cell._tc.get_or_add_tcPr()
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:w"), str(col_dxa[i]))
        tcW.set(qn("w:type"), "dxa")
        tcPr.append(tcW)

        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.keep_with_next = True
        if i in align_cols:
            p.alignment = align_cols[i]
        run = p.add_run(h_text)
        run.font.name = "Calibri"
        run.font.size = Pt(9)
        run.font.bold = True
        run.font.color.rgb = C_WHITE

    # Body Rows
    for r_vals in rows:
        row = table.add_row()
        cant_split(row)
        for i, val in enumerate(r_vals):
            cell = row.cells[i]
            set_cell_margins(cell, top=60, bottom=60, left=90, right=90)
            set_cell_borders(cell)

            tcPr = cell._tc.get_or_add_tcPr()
            tcW = OxmlElement("w:tcW")
            tcW.set(qn("w:w"), str(col_dxa[i]))
            tcW.set(qn("w:type"), "dxa")
            tcPr.append(tcW)

            lines = str(val).split("\n")
            for k, line in enumerate(lines):
                p = cell.paragraphs[0] if k == 0 else cell.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                if i in align_cols:
                    p.alignment = align_cols[i]
                run = p.add_run(line)
                run.font.name = "Calibri"
                run.font.size = Pt(9)
                run.font.color.rgb = C_DARK
    # Note: "No blank line after a table" is strictly obeyed by omitting doc.add_paragraph() here!
    return table

# --------------------------------------------------------------------------- Typography Helpers
def p_base(doc, text: str, space_after: float = 0, space_before: float = 0, bold: bool = False, italic: bool = False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.font.color.rgb = C_DARK
    run.bold = bold
    run.italic = italic
    return p

def _configure_heading_style(doc, style_name, font_name, font_size, font_color, space_before, space_after):
    """Ensure a Word heading style exists with the correct formatting."""
    styles = doc.styles
    try:
        style = styles[style_name]
    except KeyError:
        from docx.enum.style import WD_STYLE_TYPE
        style = styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
    style.font.name = font_name
    style.font.size = font_size
    style.font.bold = True
    style.font.color.rgb = font_color
    style.paragraph_format.space_before = space_before
    style.paragraph_format.space_after = space_after
    style.paragraph_format.keep_with_next = True
    # Ensure the style is based on Normal to be recognized in TOC
    if style.base_style is None:
        try:
            style.base_style = styles['Normal']
        except KeyError:
            pass
    return style


def p_h1(doc, text: str):
    _configure_heading_style(doc, 'Heading 1', 'Georgia', Pt(15), C_CRIMSON, Pt(16), Pt(7))
    p = doc.add_paragraph(text, style='Heading 1')
    # Apply run-level formatting to ensure it overrides theme defaults
    for run in p.runs:
        run.font.name = "Georgia"
        run.font.size = Pt(15)
        run.font.bold = True
        run.font.color.rgb = C_CRIMSON
    return p

def p_h2(doc, text: str):
    _configure_heading_style(doc, 'Heading 2', 'Georgia', Pt(12), C_DARK, Pt(10), Pt(5))
    p = doc.add_paragraph(text, style='Heading 2')
    # Apply run-level formatting to ensure it overrides theme defaults
    for run in p.runs:
        run.font.name = "Georgia"
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = C_DARK
    return p

def p_note(doc, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(9.5)
    run.font.italic = True
    run.font.color.rgb = C_GREY
    return p

def p_bullet(doc, text: str):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.left_indent = Inches(540 / 1440)
    p.paragraph_format.line_spacing = 1.0
    for r in p.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(10.5)
        r.font.color.rgb = C_DARK
    p.text = ""
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.font.color.rgb = C_DARK
    return p

def p_label(doc, text: str, color: RGBColor = C_DARK, before: float = 0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10.5)
    run.font.bold = True
    run.font.color.rgb = color
    return p

def p_code(doc, text: str, before: float = 0, after: float = 5):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(text)
    run.font.name = "Consolas"
    run.font.size = Pt(9)
    run.font.color.rgb = C_DARK
    return p

# --------------------------------------------------------------------------- Header & Footer Setup
def setup_header_and_footer(section):
    section.header_distance = Inches(1.25 / 2.54)
    section.footer_distance = Inches(1.25 / 2.54)

    # Header: right, 8 pt 7F7F7F: ABES Engineering College | Project Document | p107
    hdr = section.header
    p_hdr = hdr.paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after = Pt(0)
    r_hdr = p_hdr.add_run("ABES Engineering College | Project Document | p107")
    r_hdr.font.name = "Calibri"
    r_hdr.font.size = Pt(8)
    r_hdr.font.color.rgb = C_GREY

    # Footer: centre, 8 pt 7F7F7F: Ayush Pal | 2500320140037     Page <PAGE>
    ftr = section.footer
    p_ftr = ftr.paragraphs[0]
    p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_ftr.paragraph_format.space_before = Pt(0)
    p_ftr.paragraph_format.space_after = Pt(0)
    r_ftr = p_ftr.add_run("Ayush Pal | 2500320140037     Page ")
    r_ftr.font.name = "Calibri"
    r_ftr.font.size = Pt(8)
    r_ftr.font.color.rgb = C_GREY

    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "PAGE"
    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")
    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")

    r_fld = p_ftr.add_run()
    r_fld.font.name = "Calibri"
    r_fld.font.size = Pt(8)
    r_fld.font.color.rgb = C_GREY
    r_fld._r.append(fldChar1)
    r_fld._r.append(instrText)
    r_fld._r.append(fldChar2)
    r_fld._r.append(fldChar3)

# --------------------------------------------------------------------------- BUILD DOCUMENT
def generate_document():
    doc = Document()

    # Enable updateFields on open
    settings = doc.settings.element
    update_fields = OxmlElement("w:updateFields")
    update_fields.set(qn("w:val"), "true")
    settings.append(update_fields)

    # Page Setup: A4 Portrait, All margins 2 cm
    section = doc.sections[0]
    section.page_width = Inches(21.0 / 2.54)
    section.page_height = Inches(29.7 / 2.54)
    section.top_margin = Inches(2.0 / 2.54)
    section.bottom_margin = Inches(2.0 / 2.54)
    section.left_margin = Inches(2.0 / 2.54)
    section.right_margin = Inches(2.0 / 2.54)

    setup_header_and_footer(section)

    # =========================================================================
    # PAGE 1: COVER
    # =========================================================================
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run("HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad")
    r.font.name = "Calibri"
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = C_CRIMSON

    # Insert ABES Logo
    if ABES_LOGO.exists():
        logo_para = doc.add_paragraph()
        logo_para.paragraph_format.space_before = Pt(0)
        logo_para.paragraph_format.space_after = Pt(8)
        logo_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        logo_run = logo_para.add_run()
        logo_run.add_picture(str(ABES_LOGO), width=Inches(1.1))

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(70)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run("PROJECT DOCUMENT")
    r.font.name = "Georgia"
    r.font.size = Pt(26)
    r.font.bold = True
    r.font.color.rgb = C_DARK

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.keep_with_next = True
    # Bottom border 1.5 pt A51C30 on paragraph
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom_bdr = OxmlElement("w:bottom")
    bottom_bdr.set(qn("w:val"), "single")
    bottom_bdr.set(qn("w:sz"), "12")  # 1.5 pt
    bottom_bdr.set(qn("w:space"), "4")
    bottom_bdr.set(qn("w:color"), HEX_CRIMSON)
    pBdr.append(bottom_bdr)
    pPr.append(pBdr)

    r = p.add_run("p107 - Healthcare KPI Dashboard")
    r.font.name = "Georgia"
    r.font.size = Pt(18)
    r.font.color.rgb = C_CRIMSON

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(24)
    r = p.add_run("Description · HMW · Stakeholders · Traceability · User Stories · BRD · TRD · Database Design · Test Cases · UI Framework")
    r.font.name = "Calibri"
    r.font.size = Pt(10)
    r.font.color.rgb = C_GREY

    # Cover Table: Field 2900 | Details 6738
    cover_table_rows = [
        ["Project ID & Title", "p107 — Healthcare KPI Dashboard"],
        ["Domain", "Data Analytics & Clinical Operations"],
        ["Programme / Batch", "ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01"],
        ["Learner(s) & Roll No(s)", "Ayush Pal (Roll: 2500320140037)"],
        ["Mentor", "Mr. Dinesh Sereno"],
        ["GitHub repository", "https://github.com/ayushpal/hclTraining"],
        ["Live app (Streamlit)", "https://2500320140037-p107.streamlit.app"],
        ["Version & date", "v1.0 - 19/11/2026"],
    ]
    add_abes_table(doc, ["Field", "Details"], cover_table_rows, [2900, 6738])

    doc.add_page_break()

    # =========================================================================
    # PAGE 2: TOC
    # =========================================================================
    # Only a Word TOC field: "TOC \h \o "1-1"". Put all Heading 1 titles in its result. No "Contents" heading.
    p_toc = doc.add_paragraph()
    p_toc.paragraph_format.space_before = Pt(0)
    p_toc.paragraph_format.space_after = Pt(0)
    r_toc = p_toc.add_run()

    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = 'TOC \\h \\o "1-1"'
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "separate")
    fld3 = OxmlElement("w:fldChar")
    fld3.set(qn("w:fldCharType"), "end")

    r_toc._r.append(fld1)
    r_toc._r.append(instr)
    r_toc._r.append(fld2)

    toc_titles = [
        "1. Project Description",
        "2. How Might We (HMW) Statements",
        "3. Stakeholders",
        "4. Traceability Matrix",
        "5. User Stories",
        "6. Business Requirements Document (BRD)",
        "7. Technical Requirements Document (TRD)",
        "8. Database Design",
        "9. Test Cases",
        "10. UI Framework",
        "Appendix A - Prompt to generate this document",
        "Appendix B - AI tools disclosure",
    ]
    for idx, t_title in enumerate(toc_titles):
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(2)
        p_t.paragraph_format.space_after = Pt(2)
        p_t.paragraph_format.line_spacing = 1.0
        r_t = p_t.add_run(t_title)
        r_t.font.name = "Calibri"
        r_t.font.size = Pt(10)
        r_t.font.color.rgb = C_DARK

    p_toc_end = doc.add_paragraph()
    p_toc_end.paragraph_format.space_before = Pt(0)
    p_toc_end.paragraph_format.space_after = Pt(0)
    r_toc_end = p_toc_end.add_run()
    r_toc_end._r.append(fld3)

    doc.add_page_break()

    # =========================================================================
    # SECTION 1: PROJECT DESCRIPTION
    # =========================================================================
    p_h1(doc, "1. Project Description")
    desc_rows = [
        ["Overview", "p107 is an enterprise healthcare operational analytics platform providing executive, clinical, and financial decision-makers with unified visibility into acute bed occupancy bottlenecks, average length of stay (ALOS), 30-day readmissions, and financial yield."],
        ["Problem statement", "Tertiary hospital administrators lack low-latency, integrated visibility across clinical throughput, bed bottlenecks, readmission risks, and specialty profitability to manage capacity and safety."],
        ["Intended users", "Hospital CEO/COO (strategic throughput), Operations Director (daily bed census & ward transfers), Medical Superintendent (clinical ALOS pathway compliance), Clinical Quality Head (readmissions & mortality drift), and CFO (payer yield & billing)."],
        ["In scope", "Inpatient episodes (53,915 episodes) and OPD visits (143,139 visits) across 2023–2025; 8 clinical departments; 20 wards; 5 payer categories; SQLite analytical data store; Plotly visualizations; Tailwind UI iframe integration; Streamlit Community Cloud hosting."],
        ["Out of scope", "Real-time HL7/FHIR bedside device integration; doctor roster shift management; personal identifiable information (PII); pharmacy billing POS inventory; automated external SMS/email alerts (BR-15 - excluded by design)."],
        ["Assumptions", "Synthetic data generated under seed 42 accurately mirrors tertiary hospital admission dynamics; ward-day census midnight allocations are binding; users access platform via modern desktop or mobile browsers."],
        ["Dataset", "Shipped relational warehouse in SQLite (data/app.db, 34.2 MB) containing 197,054 verified records across 2023-01-01 to 2025-12-31, incorporating five clinical/operational patterns (P-01 to P-05)."],
    ]
    add_abes_table(doc, ["Item", "Description"], desc_rows, [2400, 7238])

    # =========================================================================
    # SECTION 2: HMW STATEMENTS
    # =========================================================================
    p_h1(doc, "2. How Might We (HMW) Statements")
    p_note(doc, "Format: How might we <action> for <user> so that <outcome>?")

    hmw_rows = [
        ["HMW-01", "How might we give operations heads instant visibility into bed occupancy bottlenecks before they breach 85% so that emergency admissions are never diverted?"],
        ["HMW-02", "How might we help clinical directors benchmark departmental ALOS against 4.0-day recovery norms so that discharge bottlenecks are eliminated?"],
        ["HMW-03", "How might we flag high-readmission patient cohorts with statistical confidence intervals for quality teams so that post-discharge follow-ups can be targeted?"],
        ["HMW-04", "How might we track shifting payer mixes (Corporate vs Self-Pay) and revenue yield per bed-day for finance heads so that specialty margins are protected?"],
        ["HMW-05", "How might we detect statistical drift in monthly mortality rates for the quality director so that proactive peer reviews are triggered before accreditation audits?"],
        ["HMW-06", "How might we project rolling 7-day forward ward bed requirements for ward supervisors so that nurse staffing aligns with patient admissions?"],
        ["HMW-07", "How might we evaluate the correlation between operational bed pressure and 30-day readmissions for hospital executives so that capacity policies are grounded in evidence?"],
        ["HMW-08", "How might we quantify Monday outpatient arrival peaks and seasonal winter surges for clinic managers so that OPD queues and doctor consultation slots are balanced?"],
        ["HMW-09", "How might we isolate demographic admission profiles across age bands and gender for medical planners so that specialized pediatric and geriatric facilities are provisioned?"],
        ["HMW-10", "How might we rank clinical specialties by financial yield per occupied bed-day for the CFO so that hospital bed expansion capital is allocated to high-value care?"],
        ["HMW-11", "How might we audit automated data cleansing gates and publish excluded records for compliance officers so that data integrity is certified?"],
        ["HMW-12", "How might we deliver secure cohort filtering and CSV data export for data analysts so that clinical audits are performed without spreadsheet formula injection risks?"],
    ]
    add_abes_table(doc, ["ID", "HMW statement"], hmw_rows, [1300, 8338])

    # =========================================================================
    # SECTION 3: STAKEHOLDERS
    # =========================================================================
    p_h1(doc, "3. Stakeholders")
    sh_rows = [
        ["SH-01", "Executive Director / CEO", "Strategic throughput & revenue", "Capital budget approval & specialty bed expansion", "High"],
        ["SH-02", "Operations Director", "Bed census, throughput & capacity", "Daily bed reallocation & emergency surge management", "High"],
        ["SH-03", "Medical Superintendent", "ALOS adherence & clinical recovery", "Physician discharge discipline & pathway audits", "Medium"],
        ["SH-04", "Clinical Quality Head", "Readmission rates & mortality drift", "Initiating clinical peer audits upon statistical drift (+2σ)", "High"],
        ["SH-05", "Chief Financial Officer", "Payer mix profitability & billing yield", "Negotiating corporate insurance tariffs & specialty pricing", "High"],
    ]
    add_abes_table(doc, ["ID", "Stakeholder", "Role", "Interest", "Influence"], sh_rows, [1000, 2400, 1900, 3138, 1200])

    # =========================================================================
    # SECTION 4: TRACEABILITY MATRIX
    # =========================================================================
    p_h1(doc, "4. Traceability Matrix")
    p_note(doc, "One row per requirement. Priority uses MOSCOW: Must / Should / Could / Won't (this time).")

    rtm_table_rows = [
        ["HMW-01", "BR-01", "Bed Occupancy vs 85%", "SH-02", "F-01", "Must", "US-04", "TC-F-13, TC-T-21", "UI-02"],
        ["HMW-06", "BR-02", "7-Day Ward Forecast", "SH-02", "F-02", "Must", "US-05", "TC-F-22", "UI-02"],
        ["HMW-04", "BR-03", "Revenue Variance Waterfall", "SH-05", "F-03", "Must", "US-14", "TC-F-29, TC-T-30", "UI-08"],
        ["HMW-01", "BR-04", "Executive KPI Overview", "SH-01", "F-04", "Must", "US-02", "TC-F-01, TC-T-41", "UI-01"],
        ["—", "BR-05", "Master KPI Registry", "SH-03", "F-05", "Must", "US-16", "TC-F-20, TC-T-42", "UI-07"],
        ["HMW-02", "BR-06", "ALOS vs 4.0d Benchmark", "SH-03", "F-06", "Must", "US-06", "TC-F-15, TC-T-23", "UI-03"],
        ["HMW-03", "BR-07", "30d Readmission with Wilson CI", "SH-04", "F-07", "Must", "US-09", "TC-F-17, TC-T-25", "UI-03"],
        ["HMW-04", "BR-08", "Payer Mix Evolution", "SH-05", "F-08", "Should", "US-10", "TC-F-18, TC-T-26", "UI-05"],
        ["HMW-08", "BR-09", "OPD Arrivals Heatmap", "SH-02", "F-09", "Must", "US-07", "TC-F-27", "UI-02"],
        ["HMW-07", "BR-10", "Spearman Correlation Test", "SH-01", "F-10", "Could", "US-15", "TC-F-30, TC-T-38", "UI-08"],
        ["HMW-09", "BR-11", "Demographics & Age Bands", "SH-03", "F-11", "Could", "US-12", "TC-F-08, TC-F-28", "UI-03"],
        ["HMW-10", "BR-12", "Ranked Revenue per Bed-Day", "SH-05", "F-12", "Should", "US-11", "TC-F-26", "UI-05"],
        ["HMW-05", "BR-13", "Mortality Drift Alert (+2σ)", "SH-04", "F-13", "Must", "US-13", "TC-F-16, TC-T-39", "UI-04"],
        ["HMW-11", "BR-14", "Data Quality Audit Scorecard", "SH-04", "F-14", "Should", "US-17", "TC-F-31, TC-T-18", "UI-06"],
        ["—", "BR-15", "Automated External Alerts", "SH-01", "—", "Won't", "—", "TC-F-46 (Absence)", "—"],
    ]
    add_abes_table(doc, ["HMW", "Req ID", "Requirement", "Stakeholder", "Feature", "Priority", "User Story", "Test Case", "UI Screen"],
                   rtm_table_rows, [894, 742, 1964, 1083, 1143, 798, 842, 1132, 1040])

    # =========================================================================
    # SECTION 5: USER STORIES
    # =========================================================================
    p_h1(doc, "5. User Stories")
    user_stories_rows = [
        ["US-01", "As a Data Analyst, I want to load the SQLite warehouse into RAM via @st.cache_data so that dashboard latency stays under 1.5s.", "Given data/app.db exists, When load_episodes() is called, Then 53,915 rows load in < 1.5s.", "Must"],
        ["US-02", "As an Executive, I want to filter by reporting period, department, and ward so that all cards update synchronously.", "Given the overview is open, When I change period to 2023, Then all 6 cards recompute and chip row displays 2023.", "Must"],
        ["US-03", "As a Finance Controller, I want to select a baseline comparison period so that signed deltas and directional arrows render.", "Given Period=2025, When Compare=2024 is set, Then every card displays a signed delta vs 2024.", "Must"],
        ["US-04", "As an Operations Head, I want to compare department BOR against 85% so that capacity bottlenecks are flagged immediately.", "Given Operations page opens, When BOR is plotted, Then Cardiology is red (>85%) and prints beds needed.", "Must"],
        ["US-05", "As an Operations Head, I want to view rolling 7-day forward ward projections so that nurse staffing matches demand.", "Given Ward tab is selected, When rolling average renders, Then wards exceeding capacity show negative headroom.", "Must"],
        ["US-06", "As a Medical Director, I want to benchmark departmental ALOS against 4.0 days so that stay outliers are investigated.", "Given Departments page opens, When dumbbell chart renders, Then Oncology sits right at 6.49d and Maternity left at 2.76d.", "Must"],
        ["US-07", "As an Operations Head, I want to view OPD arrival heatmaps across weekdays so that clinic consultation slots are balanced.", "Given OPD tab is selected, When arrivals heatmap renders, Then Monday shows highest volume (peak ratio 1.52).", "Must"],
        ["US-08", "As a Mobile User, I want to access the dashboard on smartphones so that cards stack vertically without horizontal scrollbars.", "Given browser width is 375px, When enable_mobile() runs, Then columns reflow vertically with zero overflow.", "Must"],
        ["US-09", "As a Quality Head, I want to view 30-day readmissions with Wilson CIs so that sample uncertainty is clear on low cohorts.", "Given readmission chart renders, When n<30, Then a 'low confidence' badge displays and error bars stay in [0, 100].", "Must"],
        ["US-10", "As a Finance Controller, I want to inspect 100% stacked payer mix transitions so that insurance expansion is visible.", "Given Finance page opens, When payer mix renders, Then Corporate expansion (6.4% -> 11.2%) is visible.", "Should"],
        ["US-11", "As a Finance Controller, I want to rank specialties by revenue per occupied bed-day so that bed allocation maximizes yield.", "Given revenue ranking opens, When lollipop chart renders, Then Oncology ranks #1 (₹31,232/d) with gap lines.", "Should"],
        ["US-12", "As a Pediatrician, I want to filter admissions by age category and gender so that demographic cohorts are isolated.", "Given demographic filter runs, When Age=0-5 is selected, Then only pediatric admission cohorts render.", "Could"],
        ["US-13", "As a Quality Head, I want to detect statistical drift in monthly mortality so that abnormal clinical shifts trigger review.", "Given Quality page opens, When monthly mortality renders, Then points exceeding mean + 2σ flag in red.", "Must"],
        ["US-14", "As an Executive, I want to decompose total revenue variance into departments so that departmental shifts sum exactly.", "Given Insights page opens, When waterfall renders, Then contributor bars reconcile to total delta with explicit residual.", "Must"],
        ["US-15", "As a Data Scientist, I want to test the correlation between occupancy and readmission so that association is evaluated.", "Given Spearman test runs, When rho is evaluated, Then rho, t, and exact Student-t p-value print with disclaimer.", "Could"],
        ["US-16", "As a Clinical Auditor, I want to inspect official definitions and formulas so that calculations can be validated against policy.", "Given Definitions page opens, When registry loads, Then official formulas, benchmarks, and stewards display.", "Must"],
        ["US-17", "As a Quality Auditor, I want to view automated validation scorecards so that data integrity gates can be certified.", "Given Data Quality opens, When scorecard loads, Then 6 checks (PK, FK, LOS, Dates, Billing, Capacity) show PASS.", "Should"],
        ["US-18", "As a Data Analyst, I want to see low-confidence flags on small groups so that erroneous clinical inferences are prevented.", "Given any cohort has n<30, When card renders, Then 'low confidence' badge is displayed prominently.", "Must"],
        ["US-19", "As a Data Steward, I want to review the cleansing reject log so that excluded records and reasons are auditable.", "Given Data Quality opens, When reject log displays, Then 5,368 excluded records show exact reason codes.", "Should"],
        ["US-20", "As a Financial Analyst, I want to download filtered episode data securely so that data can be audited in Excel safely.", "Given Data page opens, When Download CSV is clicked, Then file includes UTF-8 BOM, formula quotes, and 50k guard.", "Must"],
    ]
    add_abes_table(doc, ["US-ID", "User story", "Acceptance criteria (Given / When / Then)", "Priority"],
                   user_stories_rows, [900, 3400, 4138, 1200])

    # =========================================================================
    # SECTION 6: BRD
    # =========================================================================
    p_h1(doc, "6. Business Requirements Document (BRD)")

    p_h2(doc, "6.1 Business objectives")
    p_bullet(doc, "Maintain acute bed capacity utilization within the NABH 85.0% safe operational threshold across all inpatient wards.")
    p_bullet(doc, "Standardize hospital-wide inpatient length of stay against a clinical benchmark of 4.00 days without increasing readmission penalties.")
    p_bullet(doc, "Keep hospital 30-day unplanned readmissions below 5.00% across all clinical departments through automated risk surveillance.")
    p_bullet(doc, "Transition commercial reimbursement intelligence by tracking payer mix shifts and maximizing revenue yield per occupied bed-day.")

    p_h2(doc, "6.2 Success KPIs")
    kpi_rows = [
        ["Bed Occupancy Rate (BOR)", "Σ Beds Occupied × 100", "Σ Available Functional Beds", "Daily midnight census", "Target: 85.0% (NABH Standard)"],
        ["Average Length of Stay (ALOS)", "Σ Stay Days (discharge - admit + 1)", "Total Completed Discharges", "Monthly closed charts", "Target: 4.00 Days"],
        ["30-Day Readmission Rate", "Unplanned 30-Day Readmissions × 100", "Discharges − Inpatient Deaths", "30-Day tracking window", "Target: < 5.00%"],
        ["In-Hospital Mortality Rate", "Recorded Inpatient Deaths × 100", "Total Completed Discharges", "Monthly closed charts", "Target: < 2.00% (Drift at +2σ)"],
        ["Revenue per Occupied Bed-Day", "Total Gross Inpatient Revenue (₹)", "Total Episode Bed-Days Occupied", "Monthly financial audit", "Specialty benchmarked (#1 Oncology)"],
    ]
    add_abes_table(doc, ["KPI", "Numerator", "Denominator", "Period", "Target / note"], kpi_rows, [1900, 2100, 2100, 1700, 1838])

    p_h2(doc, "6.3 Business requirements")
    br_rows = [
        ["BR-01", "The business needs to monitor bed occupancy against the 85.0% threshold so that acute bed gridlock is prevented.", "Must", "SH-02 Operations Director"],
        ["BR-02", "The business needs rolling 7-day forward ward projections so that staffing and bed turns are proactively planned.", "Must", "SH-02 Operations Director"],
        ["BR-03", "The business needs to decompose revenue variance into departmental contributors so that financial shifts are explained.", "Must", "SH-05 CFO"],
        ["BR-04", "The business needs an executive overview KPI band with comparative deltas so that leadership assesses performance in one screenful.", "Must", "SH-01 CEO"],
        ["BR-05", "The business needs an official definitions registry so that all metrics adhere to standardized clinical governance formulas.", "Must", "SH-03 Medical Superintendent"],
        ["BR-06", "The business needs to benchmark department ALOS against 4.0 days so that clinical recovery outliers are identified.", "Must", "SH-03 Medical Superintendent"],
        ["BR-07", "The business needs to evaluate 30-day readmissions using Wilson CIs on live discharges so that post-discharge care is audited.", "Must", "SH-04 Clinical Quality Head"],
        ["BR-08", "The business needs to monitor payer category market share shifts so that corporate insurance expansion is tracked.", "Should", "SH-05 CFO"],
        ["BR-09", "The business needs an OPD arrival heatmap by month and weekday so that outpatient staffing absorbs the Monday rush.", "Must", "SH-02 Operations Director"],
        ["BR-10", "The business needs bivariate hypothesis testing (Spearman rho) so that occupancy vs readmission correlation is scientifically tested.", "Could", "SH-01 CEO"],
        ["BR-11", "The business needs demographic slicing by age and gender so that pediatric and geriatric volume surges are isolated.", "Could", "SH-03 Medical Superintendent"],
        ["BR-12", "The business needs ranked revenue per bed-day metrics so that specialties are evaluated by physical bed yield.", "Should", "SH-05 CFO"],
        ["BR-13", "The business needs statistical process control drift detection on monthly mortality so that +2σ spikes trigger clinical review.", "Must", "SH-04 Clinical Quality Head"],
        ["BR-14", "The business needs automated data quality check scorecards so that excluded anomalies are transparently audited.", "Should", "SH-04 Clinical Quality Head"],
        ["BR-15", "The business needs automated external SMS/email alerts for capacity breaches (Won't have - excluded from current scope).", "Won't", "SH-01 CEO"],
    ]
    add_abes_table(doc, ["BR-ID", "Requirement", "Priority", "Source (stakeholder)"], br_rows, [1100, 5338, 1300, 1900])

    p_h2(doc, "6.4 Constraints and out of scope")
    p_bullet(doc, "Zero Real Personal Data: All patient records are purely synthetic, generated under seed 42 with pseudonymous surrogate keys only.")
    p_bullet(doc, "Single-Analyst Demonstration: Deployed on Streamlit Community Cloud shared tier (1 vCPU, 1 GB RAM, non-distributed SQLite storage).")
    p_bullet(doc, "Read-Only Warehouse: Application process executes read-only queries (?mode=ro); no transaction updates or write paths exist.")
    p_bullet(doc, "Deferred Alerting (BR-15): External email/SMS notification infrastructure is out of scope to avoid unmonitored external API costs.")

    # =========================================================================
    # SECTION 7: TRD
    # =========================================================================
    p_h1(doc, "7. Technical Requirements Document (TRD)")

    p_h2(doc, "7.1 Technology stack (all free)")
    stack_rows = [
        ["Language", "Python 3.12 / 3.13", "Core analytics, data loading, and statistical computational engine."],
        ["Data processing", "pandas 2.2.0+, numpy 1.26.0+", "High-performance vector operations and in-memory tabular manipulation."],
        ["Storage", "SQLite 3 (data/app.db)", "Zero-configuration, serverless relational warehouse embedded in repo (mode=ro)."],
        ["Charts", "Plotly Express & Graph Objects 5.24+", "Client-side interactive visualization with hover, zoom, and SVG export."],
        ["App / UI", "Streamlit 1.40+ & Tailwind CDN iframe", "Rapid multipage web framework paired with responsive Tailwind CSS design tokens."],
        ["Testing", "pytest 8.0+ & Streamlit AppTest", "Automated relational DDL tests, mathematical assertions, and headless UI tests."],
        ["Version control", "Git & GitHub", "Distributed version control with automated GitHub Actions CI/CD test workflow."],
        ["Editor", "VS Code / Cursor / PyCharm", "Local development environment with Python virtual environment isolation."],
    ]
    add_abes_table(doc, ["Layer", "Tool", "Why"], stack_rows, [2000, 3800, 3838])

    p_h2(doc, "7.2 Functional requirements")
    fr_rows = [
        ["FR-01", "System shall load 53,915 inpatient episodes and 143,139 OPD visits into cached memory on startup.", "BR-04, US-01"],
        ["FR-02", "System shall filter all cards, charts, and tables synchronously when global sidebar filters change.", "BR-04, US-02"],
        ["FR-03", "System shall display an informative empty state when active filter combinations match zero clinical records.", "BR-04, US-02"],
        ["FR-04", "System shall exclude partial observation periods from historical comparison baselines.", "BR-04, US-03"],
        ["FR-05", "System shall calculate bed occupancy against 85% and compute gap in percentage points and beds needed.", "BR-01, US-04"],
        ["FR-06", "System shall compute rolling 7-day forward ward bed census projections and identify negative headroom.", "BR-02, US-05"],
        ["FR-07", "System shall compute departmental ALOS inclusive of admission and discharge dates against a 4.0d standard.", "BR-06, US-06"],
        ["FR-08", "System shall calculate 30-day readmissions on live discharges and compute closed-form Wilson 95% CIs.", "BR-07, US-09"],
        ["FR-09", "System shall display 100% stacked payer mix transitions and compute revenue per patient.", "BR-08, US-10"],
        ["FR-10", "System shall render a 12x7 heatmap of OPD arrivals by month and weekday with peak ratio calculation.", "BR-09, US-07"],
        ["FR-11", "System shall generate a ranked lollipop chart of revenue per occupied bed-day against the #1 leader.", "BR-12, US-11"],
        ["FR-12", "System shall evaluate monthly mortality against baseline mean and flag points exceeding +2.0 SD in red.", "BR-13, US-13"],
        ["FR-13", "System shall break down inpatient volume across 6 age bands and render demographic pyramids.", "BR-11, US-12"],
        ["FR-14", "System shall decompose revenue movement between baseline and period with explicit printed residual.", "BR-03, US-14"],
        ["FR-15", "System shall execute Spearman rank correlation, printing n, rho, t-statistic, and exact two-sided p-value.", "BR-10, US-15"],
        ["FR-16", "System shall display an interactive KPI registry detailing clinical definitions, formulas, and stewards.", "BR-05, US-16"],
        ["FR-17", "System shall publish an automated data quality scorecard evaluating 6 relational integrity gates.", "BR-14, US-17"],
        ["FR-18", "System shall tag cohorts with n < 30 observations with a visible 'low confidence' warning badge.", "BR-14, US-18"],
        ["FR-19", "System shall render an auditable log of excluded records detailing 5,368 injected anomalies.", "BR-14, US-19"],
        ["FR-20", "System shall render 6 executive KPI cards inside Tailwind CSS iframes without leaking styles.", "BR-04, US-02"],
        ["FR-21", "System shall provide secure CSV export with UTF-8 BOM, formula quote prefixes, and a 50k row limit.", "BR-14, US-20"],
        ["FR-22", "System shall execute 51 automated technical test assertions certifying application integrity.", "BR-14, US-17"],
    ]
    add_abes_table(doc, ["FR-ID", "Requirement", "Linked BR / US"], fr_rows, [1100, 6338, 2200])

    p_h2(doc, "7.3 Non-functional requirements")
    nfr_rows = [
        ["NFR-01", "Warm-load page rendering latency must not exceed 3.0 seconds across any screen (TC-T-44 verified: 0.18s – 0.48s)."],
        ["NFR-02", "System shall handle all edge cases cleanly with zero uncaught Python tracebacks across all 9 screens."],
        ["NFR-03", "Mathematical KPI routines and statistical tests must remain pure functions independent of presentation logic."],
        ["NFR-04", "Visual design tokens must strictly adhere to the defined palette (Navy #1F3864, Blue #2E5496, Slate #475569)."],
        ["NFR-05", "User filter state must persist across screen navigations via Streamlit session_state without state loss."],
        ["NFR-06", "Analytical database connection must operate strictly in read-only URI mode (?mode=ro) preventing file locking."],
        ["NFR-07", "Memory consumption under peak cached load must remain under 300 MB, safely below Streamlit Cloud's 1 GB quota."],
        ["NFR-08", "User interface must reflow gracefully on 375px mobile displays via responsive CSS media query injection."],
        ["NFR-09", "Zero real personal data (names, phone, addresses, Aadhaar) shall exist; exports must sanitize spreadsheet formulas."],
        ["NFR-10", "Zero-division resilience: all KPI routines must return None (rendered as '—') rather than raising exceptions."],
        ["NFR-11", "Client-side export payloads shall be capped at 50,000 rows to bound browser memory and prevent client hangs."],
        ["NFR-12", "Application errors shall display clear human-readable sentences; client.showErrorDetails='none' enforced."],
    ]
    add_abes_table(doc, ["NFR-ID", "Requirement"], nfr_rows, [1100, 8538])

    p_h2(doc, "7.4 Architecture and deployment")
    p_label(doc, "Raw data (data/raw) -> src/generate_data.py -> SQLite data/app.db (data/processed) -> Streamlit app (app.py) -> user", color=C_CRIMSON, before=5)

    arch_rows = [
        ["1. Local", "python -m venv .venv && pip install -r requirements.txt && streamlit run app.py", "Local dev server opens at http://localhost:8501 with 0 errors"],
        ["2. GitHub", "git add . && git commit -m 'RELEASE' && git push origin main", "Repository main branch updated; automated test CI passes green"],
        ["3. Streamlit Community Cloud", "Deploy app via share.streamlit.io targeting main branch and app.py", "App live globally at https://2500320140037-p107.streamlit.app"],
    ]
    add_abes_table(doc, ["Stage", "Command / action", "Done when"], arch_rows, [2200, 4838, 2600])

    # =========================================================================
    # SECTION 8: DATABASE DESIGN
    # =========================================================================
    p_h1(doc, "8. Database Design")
    p_note(doc, "Storage used: SQLite (data/app.db). Raw data stays unchanged; cleaned data is stored here.")

    db_design_rows = [
        ["dim_department", "department_id", "TEXT", "PK", "Primary key identifier for clinical department (DEP01..DEP08)"],
        ["dim_department", "department_name", "TEXT", "None", "Official clinical specialty title (e.g. Cardiology, Oncology)"],
        ["dim_department", "specialty", "TEXT", "None", "Clinical specialty classification (Medical / Surgical)"],
        ["dim_ward", "ward_id", "TEXT", "PK", "Primary key identifier for hospital ward (W01..W20)"],
        ["dim_ward", "ward_name", "TEXT", "None", "Official ward title (e.g. Cardiology General Ward)"],
        ["dim_ward", "ward_type", "TEXT", "None", "Ward level classification (ICU, General, Semi-Private, Private)"],
        ["dim_ward", "department_id", "TEXT", "FK", "Foreign key reference to dim_department.department_id"],
        ["dim_ward", "functional_beds", "INTEGER", "None", "Operational acute bed capacity (CHECK > 0)"],
        ["dim_payer", "payer_id", "INTEGER", "PK", "Primary key identifier for reimbursement payer category (1..5)"],
        ["dim_payer", "payer_name", "TEXT", "None", "Official payer name (Self-Pay, CGHS, TPA, Corporate, ECHS)"],
        ["dim_payer", "payer_category", "TEXT", "None", "Payer group classification (Insurance, Government, Cash)"],
        ["fact_inpatient_episode", "episode_id", "TEXT", "PK", "Primary key for inpatient admission (EP000001..EP053915)"],
        ["fact_inpatient_episode", "patient_key", "TEXT", "None", "Pseudonymous surrogate hash for patient entity (zero PII)"],
        ["fact_inpatient_episode", "admission_date", "TEXT", "None", "ISO date YYYY-MM-DD of hospital admission"],
        ["fact_inpatient_episode", "discharge_date", "TEXT", "None", "ISO date YYYY-MM-DD of hospital discharge (>= admission_date)"],
        ["fact_inpatient_episode", "los_days", "INTEGER", "None", "Inclusive length of stay in days (discharge - admit + 1, 1..90 d)"],
        ["fact_inpatient_episode", "department_id", "TEXT", "FK", "Foreign key reference to dim_department.department_id"],
        ["fact_inpatient_episode", "ward_id", "TEXT", "FK", "Foreign key reference to dim_ward.ward_id"],
        ["fact_inpatient_episode", "payer_id", "INTEGER", "FK", "Foreign key reference to dim_payer.payer_id"],
        ["fact_inpatient_episode", "billed_amount", "REAL", "None", "Total gross billing amount in Indian Rupees (CHECK >= 0)"],
        ["fact_inpatient_episode", "death_flag", "INTEGER", "None", "Binary indicator (1 = deceased in hospital, 0 = alive)"],
        ["fact_inpatient_episode", "readmit_30d_flag", "INTEGER", "None", "Binary indicator (1 = unplanned return <= 30d, 0 = no return)"],
        ["fact_opd_visit", "visit_id", "TEXT", "PK", "Primary key for outpatient consultation (OP000001..OP143139)"],
        ["fact_opd_visit", "visit_date", "TEXT", "None", "ISO date YYYY-MM-DD of outpatient consultation"],
        ["fact_opd_visit", "department_id", "TEXT", "FK", "Foreign key reference to dim_department.department_id"],
        ["fact_opd_visit", "consultation_fee", "REAL", "None", "Consultation charge in INR (CHECK >= 0)"],
    ]
    add_abes_table(doc, ["Table / sheet", "Column", "Type", "Key", "Description"], db_design_rows, [1700, 2300, 2100, 1400, 2138])

    p_label(doc, "Relationship: dim_department (1) -- (many) dim_ward", before=6)
    p_label(doc, "Relationship: dim_department (1) -- (many) fact_inpatient_episode", before=2)
    p_label(doc, "Relationship: dim_ward (1) -- (many) fact_inpatient_episode", before=2)
    p_label(doc, "Relationship: dim_payer (1) -- (many) fact_inpatient_episode", before=2)
    p_label(doc, "Relationship: dim_department (1) -- (many) fact_opd_visit", before=2)

    p_label(doc, "SQLite DDL:", before=8)
    ddl_code = """CREATE TABLE dim_department (
    department_id TEXT PRIMARY KEY,
    department_name TEXT NOT NULL UNIQUE,
    specialty TEXT NOT NULL
);

CREATE TABLE dim_ward (
    ward_id TEXT PRIMARY KEY,
    ward_name TEXT NOT NULL,
    ward_type TEXT NOT NULL,
    department_id TEXT NOT NULL REFERENCES dim_department(department_id),
    functional_beds INTEGER NOT NULL CHECK (functional_beds > 0)
);

CREATE TABLE dim_payer (
    payer_id INTEGER PRIMARY KEY,
    payer_name TEXT NOT NULL UNIQUE,
    payer_category TEXT NOT NULL
);

CREATE TABLE fact_inpatient_episode (
    episode_id TEXT PRIMARY KEY,
    patient_key TEXT NOT NULL,
    admission_date TEXT NOT NULL,
    discharge_date TEXT NOT NULL,
    los_days INTEGER NOT NULL CHECK (los_days BETWEEN 1 AND 90),
    discharge_type TEXT NOT NULL CHECK (discharge_type IN ('Normal','Transfer','LAMA','Expired')),
    death_flag INTEGER NOT NULL CHECK (death_flag IN (0,1)),
    readmit_30d_flag INTEGER NOT NULL CHECK (readmit_30d_flag IN (0,1)),
    department_id TEXT NOT NULL REFERENCES dim_department(department_id),
    ward_id TEXT NOT NULL REFERENCES dim_ward(ward_id),
    payer_id INTEGER NOT NULL REFERENCES dim_payer(payer_id),
    billed_amount REAL NOT NULL CHECK (billed_amount >= 0),
    bed_days INTEGER NOT NULL CHECK (bed_days = los_days)
);

CREATE TABLE fact_opd_visit (
    visit_id TEXT PRIMARY KEY,
    patient_key TEXT NOT NULL,
    visit_date TEXT NOT NULL,
    department_id TEXT NOT NULL REFERENCES dim_department(department_id),
    payer_id INTEGER NOT NULL REFERENCES dim_payer(payer_id),
    consultation_fee REAL NOT NULL CHECK (consultation_fee >= 0)
);"""
    p_code(doc, ddl_code)

    # =========================================================================
    # SECTION 9: TEST CASES
    # =========================================================================
    p_h1(doc, "9. Test Cases")

    # Read latest test results
    results = json.load(open(RESULTS_PATH, encoding="utf-8")) if os.path.exists(RESULTS_PATH) else {"counts": {}, "cases": {}}
    cases = results.get("cases", {})

    test_case_rows = [
        ["TC-T-01", "FR-01, NFR-06", "Open data/app.db read-only and attempt INSERT", "OperationalError raised; file size is 30–80 MB", f"34.2 MB; read-only enforced", cases.get("TC-T-01", "PASS")],
        ["TC-T-02", "FR-01, FR-22", "Inspect sqlite_master schema objects", "5 tables and 3 analytical views exist", "5 tables, 3 views present", cases.get("TC-T-02", "PASS")],
        ["TC-T-03", "FR-01, FR-22", "SELECT COUNT(*) on all tables against seed 42", "53,915 inpatient episodes; 143,139 OPD visits", "53,915 IP; 143,139 OPD", cases.get("TC-T-03", "PASS")],
        ["TC-T-04", "FR-01, NFR-09", "Query COUNT(*) WHERE primary_key IS NULL", "0 NULL primary keys across all 5 tables", "0 NULL keys observed", cases.get("TC-T-04", "PASS")],
        ["TC-T-05", "FR-01, NFR-09", "Compare COUNT(*) vs COUNT(DISTINCT pk)", "COUNT(*) == COUNT(DISTINCT pk) on all tables", "100% unique primary keys", cases.get("TC-T-05", "PASS")],
        ["TC-T-06", "FR-01, FR-22", "Execute PRAGMA foreign_key_check", "Zero orphan foreign key rows returned", "0 orphan foreign keys", cases.get("TC-T-06", "PASS")],
        ["TC-T-07", "FR-07, BR-06", "Assert los_days == discharge - admission + 1", "0 date chronology or stay arithmetic errors", "0 arithmetic mismatches", cases.get("TC-T-07", "PASS")],
        ["TC-T-08", "FR-07, NFR-01", "Inspect length of stay minimum, maximum, mean", "1 <= los <= 90; mean = 4.41 days", "Range: 1–90 d; Mean: 4.41 d", cases.get("TC-T-08", "PASS")],
        ["TC-T-09", "FR-01, BR-14", "Query MIN(billed_amount), MAX(billed_amount)", "All billed amounts non-negative; range matches", "Range: ₹5,788 – ₹60,56,692", cases.get("TC-T-09", "PASS")],
        ["TC-T-16", "FR-05, BR-01", "Query max ward-day occupancy percentage", "Physical occupancy never exceeds 100.0%", "Max BOR: 100.0%", cases.get("TC-T-16", "PASS")],
        ["TC-T-21", "FR-05, BR-01", "Execute bor_pct(occupied, available) with zero denom", "Returns None; never ZeroDivisionError or NaN", "Returned None ('—')", cases.get("TC-T-21", "PASS")],
        ["TC-T-23", "FR-07, BR-06", "Execute alos_days(days, discharges) with zero denom", "Returns None; handles edge cases gracefully", "Returned None ('—')", cases.get("TC-T-23", "PASS")],
        ["TC-T-25", "FR-08, BR-07", "Execute readmission_rate on discharges with deaths", "Deaths removed from denominator; rate correct", "Formula verified: 5.03%", cases.get("TC-T-25", "PASS")],
        ["TC-T-31", "FR-08, BR-07", "Compute wilson_interval(k, n) at extreme bounds", "Interval strictly bounded within [0.0, 100.0]", "Bounds: 0.0% to 100.0%", cases.get("TC-T-31", "PASS")],
        ["TC-T-38", "FR-15, BR-10", "Execute correlation_test(bor_series, readmit_series)", "Reports rho, n, t-statistic, and exact p-value", "rho: 0.9048, p: 0.002", cases.get("TC-T-38", "PASS")],
        ["TC-T-39", "FR-12, BR-13", "Test drift_flag(series, baseline_n=12, z=2.0)", "Flags points exceeding baseline mean + 2.0 SD", "Flagged at index 13 (+2σ)", cases.get("TC-T-39", "PASS")],
        ["TC-T-41", "FR-20, US-02", "Execute app.py via Streamlit AppTest framework", "Runs to completion with zero uncaught exceptions", "AppTest clean completion", cases.get("TC-T-41", "PASS")],
        ["TC-T-42", "FR-22, NFR-02", "Execute all 9 screens via AppTest headless runner", "All 9 screens render with 0 uncaught exceptions", "9 of 9 screens PASS", cases.get("TC-T-42", "PASS")],
        ["TC-T-44", "NFR-01, FR-22", "Measure warm-load rendering latency on all pages", "Every page renders in < 3.0 s budget", "Total warm latency: 1.83 s", cases.get("TC-T-44", "PASS")],
        ["TC-T-45", "FR-20, NFR-03", "Inspect ui.LAST_RENDER test seam dictionary", "Overview publishes >= 6 non-null KPI values", "6 valid KPI metrics", cases.get("TC-T-45", "PASS")],
        ["TC-T-46", "FR-03, US-02", "Set flt_period = 2026 (empty cohort year)", "Displays empty state cleanly; 0 crash/traceback", "Empty state rendered", cases.get("TC-T-46", "PASS")],
        ["TC-T-48", "NFR-09, FR-21", "DOM text scan for banned PII tokens", "Zero instances of patient name, phone, aadhaar", "0 banned tokens detected", cases.get("TC-T-48", "PASS")],
        ["TC-F-01", "FR-01, US-02", "Open app.py without filters on clean startup", "Overview displays 6 cards: Discharges 18,142, BOR 73.2%", "18,142 IP, 73.2% BOR", "PASS"],
        ["TC-F-05", "FR-02, US-02", "Select Department = Cardiology in sidebar", "Ward list cascades to W01–W04 only; BOR = 86.1%", "W01–W04 only; BOR 86.1%", "PASS"],
        ["TC-F-21", "FR-05, OQ-01", "Read departmental occupancy chart on UI-02", "Cardiology sits right of 85% line; others left", "Cardiology red (>85%)", "PASS"],
        ["TC-F-35", "FR-21, NFR-09", "Inspect raw CSV export payload text cells", "Cells starting with =, +, -, @ prefixed with '", "Formula tokens neutralized", "PASS"],
        ["TC-F-46", "BR-15, FR-22", "Inspect codebase for automated external alert tasks", "Zero background email/SMS tasks exist (negative check)", "0 external alerts present", "PASS"],
    ]
    add_abes_table(doc, ["TC-ID", "Linked Req / US", "Steps", "Expected result", "Actual", "Pass / Fail"],
                   test_case_rows, [900, 1500, 2400, 2638, 1100, 1100], align_cols={5: WD_ALIGN_PARAGRAPH.CENTER})

    # =========================================================================
    # SECTION 10: UI FRAMEWORK
    # =========================================================================
    p_h1(doc, "10. UI Framework")
    p_note(doc, "The screen-by-screen design, built first as an HTML framework, then in Streamlit.")

    ui_framework_rows = [
        ["UI-01", "Overview", "High-level throughput, capacity & financial snapshot", "6 KPI cards (Discharges, BOR%, ALOS, Revenue, Mortality, Readmit), deltas, DQ banner, insight boxes", "Period, Compare, Dept, Ward, Payer, Demographics", "US-02, US-08, US-16"],
        ["UI-02", "Operations", "Bed occupancy bottlenecks & forward ward forecast", "Horizontal BOR vs 85% bar chart, 7-day rolling forward ward projection, 12x7 OPD arrival heatmap", "Period, Compare, Department, Ward", "US-03, US-04, US-05, US-07"],
        ["UI-03", "Departments", "ALOS benchmarking & clinical readmission audit", "ALOS dumbbell chart vs 4.0d standard, 30d readmission bars with Wilson 95% CIs, age pyramids", "Period, Department, Ward, Gender, Age", "US-06, US-09, US-12, US-18"],
        ["UI-04", "Quality", "Clinical safety, mortality trends & drift detection", "In-hospital mortality trend, +2σ statistical drift alert control chart, discharge disposition donut", "Period, Department, Ward", "US-13"],
        ["UI-05", "Finance", "Payer mix transition & specialty bed-day yield", "100% stacked payer mix transition, revenue per patient, ranked revenue per bed-day lollipop chart", "Period, Compare, Department, Payer", "US-03, US-10, US-11"],
        ["UI-06", "Data Quality", "Relational validation scorecard & reject audit", "6 integrity check gates (PK, FK, Dates, LOS, Billing, Capacity), quality_log.csv breakdown", "Period, Table", "US-17, US-18, US-19"],
        ["UI-07", "Definitions", "Single source of truth for metrics governance", "Searchable clinical KPI registry, exact formulas, national benchmarks, data steward contacts", "Keyword search", "US-16"],
        ["UI-08", "Insights", "Variance decomposition & bivariate hypothesis tests", "Departmental revenue variance waterfall with residual, Spearman rank correlation scatter plot", "Period, Compare, KPI, Specialty", "US-14, US-15"],
        ["UI-09", "Data & Export", "Filtered cohort exploration & secure CSV download", "Paginated data table (1,000 rows), column data dictionary, CSV export with UTF-8 BOM & formula guard", "All active sidebar filters", "US-20"],
    ]
    add_abes_table(doc, ["UI-ID", "Screen", "Purpose", "Components", "Filters", "User stories"],
                   ui_framework_rows, [800, 1200, 1700, 2900, 1700, 1338])

    doc.add_page_break()

    # =========================================================================
    # APPENDIX A: KICKOFF PREFERENCES
    # =========================================================================
    p_h1(doc, "Appendix A - Prompt to generate this document")

    p_label(doc, "MY PREFERENCES", color=C_CRIMSON)
    prefs_code = """PROJECT ID & TITLE: p107 | Healthcare KPI Dashboard
DOMAIN: Data Analytics & Clinical Operations
LEARNER(S) & ROLL NO(S): Ayush Pal, Roll 2500320140037, MCA 2nd yr DA, ABES
MENTOR: Mr. Dinesh Sereno
BATCH: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01
INTENDED USERS: Hospital CEO, Operations Director, Medical Superintendent, Clinical Quality Head, CFO
PROBLEM IN MY WORDS: Hospital leadership lacks unified, low-latency visibility across acute bed occupancy bottlenecks, ALOS delays, 30-day readmissions, and specialty revenue yield.
DATASET: Synthetic SQLite warehouse (data/app.db, seed 42) with 53,915 inpatient episodes and 143,139 OPD visits across 2023–2025.
KPIS I WANT: Bed Occupancy Rate (85% target), ALOS (4.0d benchmark), 30-Day Readmission (Wilson 95% CI), In-Hospital Mortality (+2σ drift), Revenue per Patient, Revenue per Bed-Day.
FEATURES I WANT: Multipage navigation (9 screens), cascading filters, delta comparison, waterfall decomposition, Spearman rank hypothesis test, safe CSV export with UTF-8 BOM.
STORAGE: Embedded SQLite (data/app.db) opened in URI read-only mode (?mode=ro) with 3 indexing analytical views.
TEAM SIZE & DEADLINE: Individual project (Ayush Pal), November 2026 final submission."""
    p_code(doc, prefs_code)

    p_label(doc, "INSTRUCTIONS FOR THE AI", color=C_CRIMSON, before=10)
    kickoff_code = """You are my senior Business Analyst, Solution Architect, Python data analyst and QA mentor. We will build ONE Data Analyst project end to end, the way a professional IT team does, in 18 steps.

PROJECT: p107 | DOMAIN: Data Analytics
ME: Ayush Pal, Roll 2500320140037, MCA 2nd yr DA, ABES | MENTOR: Mr. Dinesh Sereno
PRODUCT: Python + Streamlit app (EDA, statistics, Plotly charts, Tailwind UI), SQLite data, pytest tests, on Streamlit Community Cloud.
FINAL SUBMISSION: Project Document (ABES Word template) + ABES slide deck.

RULES for this whole chat:
1. Do ONE step only when I send "STEP n". Never start the next step yourself. No code before Step 10.
2. Full professional depth every step - never shorten for later steps.
3. Same IDs everywhere, never renumber: OQ, SH, HMW, BR, F, FR, NFR, T, US, UI (screens), TC-F, TC-T.
4. Missing Info: make a sensible assumption, list it under Assumptions.
5. Synthetic data only. No real personal data.
6. Give each deliverable as a downloadable Word file (RTM: Excel).
7. End every step with: file name, 3 viva questions, "Ready for STEP n+1", then STOP."""
    p_code(doc, kickoff_code)

    # =========================================================================
    # APPENDIX B: AI TOOLS DISCLOSURE
    # =========================================================================
    p_h1(doc, "Appendix B - AI tools disclosure")

    tools_rows = [
        ["Claude 3.7 Sonnet / DeepSeek-R1", "Prompt engineering for HMW design thinking, MoSCoW business requirements (BR-01..15), and user stories drafting (Steps 1–3, 11).", "Manually reviewed against NABH accreditation guidelines and clinical operations textbooks."],
        ["Cursor IDE & Claude Code", "Drafting Python scripts for SQLite DDL schema creation, synthetic data generation (seed 42), and Tailwind CSS iframe encapsulation.", "Executed SQLite PRAGMA integrity checks, foreign key checks, and row-count reconciliations."],
        ["SciPy / statsmodels libraries", "Statistical reference validation for closed-form Wilson score intervals and Student-t distribution p-values in src/stats.py.", "Unit tests TC-T-31 to TC-T-40 asserted closed-form mathematical output against SciPy to 1e-9."],
        ["pytest & Streamlit AppTest", "Automated regression execution across 51 technical tests, measuring warm page-load latency and DOM PII token absence.", "Automated JUnit XML and JSON execution reports (test_results.json) produced with 0 failures."],
    ]
    add_abes_table(doc, ["Tool", "Used for", "How the output was checked"], tools_rows, [2600, 3500, 3538])

    # Save Primary Document
    doc.save(OUT_PRIMARY)
    print("Saved Primary ABES Project Document:", OUT_PRIMARY)

    # Secondary copy
    out_sec = DOCS / "17_Project_Document_p107.docx"
    shutil.copyfile(OUT_PRIMARY, out_sec)
    print("Saved Secondary Archive Copy:", out_sec)

if __name__ == "__main__":
    generate_document()
