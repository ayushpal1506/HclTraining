"""Build and render Step 18 Presentation deck (PPTX + PDF + slide images).

Generates:
1. docs/HKD-2026_Presentation.pptx & submission_project_folder/2_Project_Presentation.pptx
2. docs/slide_images/slide_01.png to slide_11.png (high-res 1600x900 with all text and shapes)
3. docs/HKD-2026_Presentation.pdf & submission_project_folder/2_Project_Presentation.pdf
"""
import os
import shutil
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
SUB_DIR = ROOT / "submission_project_folder"
IMG_DIR = DOCS / "slide_images"
IMG_DIR.mkdir(parents=True, exist_ok=True)
SUB_DIR.mkdir(parents=True, exist_ok=True)

# Colors
C_RED = RGBColor(0xA5, 0x1C, 0x30)
C_BLACK = RGBColor(0x15, 0x15, 0x14)
C_DARK = RGBColor(0x3A, 0x3A, 0x38)
C_GREY = RGBColor(0x5E, 0x5E, 0x5B)
C_CARD = RGBColor(0xF4, 0xF1, 0xF1)
C_PINK = RGBColor(0xD9, 0xB5, 0xBC)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_BORDER = RGBColor(0xD9, 0xD9, 0xD9)
C_NUM_GREY = RGBColor(0x8A, 0x8A, 0x86)

HEX_RED = "A51C30"
HEX_BORDER = "D9D9D9"

# Pillow colors
RGB_RED = (165, 28, 48)
RGB_BLACK = (21, 21, 20)
RGB_DARK = (58, 58, 56)
RGB_GREY = (94, 94, 91)
RGB_CARD = (244, 241, 241)
RGB_PINK = (217, 181, 188)
RGB_WHITE = (255, 255, 255)
RGB_BORDER = (217, 217, 217)
RGB_NUM_GREY = (138, 138, 134)

# Load Fonts for Pillow
FONT_GEORGIA_36 = ImageFont.truetype("/System/Library/Fonts/Supplemental/Georgia.ttf", 36)
FONT_GEORGIA_28 = ImageFont.truetype("/System/Library/Fonts/Supplemental/Georgia.ttf", 26)
FONT_GEORGIA_44 = ImageFont.truetype("/System/Library/Fonts/Supplemental/Georgia.ttf", 44)
FONT_GEORGIA_40 = ImageFont.truetype("/System/Library/Fonts/Supplemental/Georgia.ttf", 38)
FONT_HELVETICA_BOLD_15 = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 15)
FONT_HELVETICA_13 = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 13)
FONT_HELVETICA_12 = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 12)
FONT_HELVETICA_11 = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 11)
FONT_HELVETICA_9 = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 9)
FONT_HELVETICA_BOLD_12 = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 12)


def set_cell_border(cell, color_hex=HEX_BORDER, width_emu=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    for border_name in ("lnL", "lnR", "lnT", "lnB"):
        existing = tcPr.find(f"{{http://schemas.openxmlformats.org/drawingml/2006/main}}{border_name}")
        if existing is not None:
            tcPr.remove(existing)
        xml = (
            f'<a:{border_name} {nsdecls("a")} w="{width_emu}" cmpd="s">'
            f'<a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>'
            f'</a:{border_name}>'
        )
        tcPr.append(parse_xml(xml))


def add_card(slide, x, y, w, h):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = C_CARD
    shape.line.color.rgb = C_CARD
    shape.adjustments[0] = 0.08
    return shape


def add_badge(slide, x, y, text):
    badge = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.42), Inches(0.42))
    badge.fill.solid()
    badge.fill.fore_color.rgb = C_RED
    badge.line.fill.background()
    tf = badge.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = str(text)
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Calibri"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    return badge


def add_badge_card(slide, x, y, w, h, badge_text, head_text, body_text):
    add_card(slide, x, y, w, h)
    add_badge(slide, x + 0.2, y + 0.2, badge_text)
    
    box_head = slide.shapes.add_textbox(Inches(x + 0.75), Inches(y + 0.17), Inches(w - 0.9), Inches(0.45))
    tf_h = box_head.text_frame
    tf_h.word_wrap = True
    tf_h.margin_left = tf_h.margin_right = tf_h.margin_top = tf_h.margin_bottom = 0
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = tf_h.paragraphs[0]
    p_h.text = head_text
    p_h.font.name = "Calibri"
    p_h.font.size = Pt(15)
    p_h.font.bold = True
    p_h.font.color.rgb = C_BLACK
    
    box_body = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    tf_b = box_body.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_right = tf_b.margin_top = tf_b.margin_bottom = 0
    tf_b.vertical_anchor = MSO_ANCHOR.TOP
    for idx, line in enumerate(body_text.split("\n")):
        p = tf_b.paragraphs[0] if idx == 0 else tf_b.add_paragraph()
        p.text = line
        p.font.name = "Calibri"
        p.font.size = Pt(13)
        p.font.color.rgb = C_DARK
    return box_head


def add_plain_card(slide, x, y, w, h, head_text, body_text):
    add_card(slide, x, y, w, h)
    box_head = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.17), Inches(w - 0.35), Inches(0.45))
    tf_h = box_head.text_frame
    tf_h.word_wrap = True
    tf_h.margin_left = tf_h.margin_right = tf_h.margin_top = tf_h.margin_bottom = 0
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = tf_h.paragraphs[0]
    p_h.text = head_text
    p_h.font.name = "Calibri"
    p_h.font.size = Pt(15)
    p_h.font.bold = True
    p_h.font.color.rgb = C_BLACK
    
    box_body = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    tf_b = box_body.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_right = tf_b.margin_top = tf_b.margin_bottom = 0
    tf_b.vertical_anchor = MSO_ANCHOR.TOP
    for idx, line in enumerate(body_text.split("\n")):
        p = tf_b.paragraphs[0] if idx == 0 else tf_b.add_paragraph()
        p.text = line
        p.font.name = "Calibri"
        p.font.size = Pt(13)
        p.font.color.rgb = C_DARK
    return box_head


def add_abes_table(slide, x, y, headers, rows, col_widths):
    num_rows = len(rows) + 1
    num_cols = len(headers)
    total_w = sum(col_widths)
    total_h = 0.5 * num_rows
    tbl_shape = slide.shapes.add_table(num_rows, num_cols, Inches(x), Inches(y), Inches(total_w), Inches(total_h))
    tbl = tbl_shape.table
    
    tblPr = tbl._tbl.tblPr
    style_id = tblPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId")
    if style_id is not None:
        tblPr.remove(style_id)
        
    for i, w in enumerate(col_widths):
        tbl.columns[i].width = Inches(w)
    for r in range(num_rows):
        tbl.rows[r].height = Inches(0.5)
        
    for c, h_text in enumerate(headers):
        cell = tbl.cell(0, c)
        cell.margin_left = cell.margin_right = Inches(0.1)
        cell.margin_top = cell.margin_bottom = Inches(0.05)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_RED
        set_cell_border(cell, HEX_BORDER, 6350)
        tf = cell.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.text = h_text
        p.font.name = "Calibri"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.LEFT
        
    for r_idx, r_data in enumerate(rows):
        for c_idx, cell_text in enumerate(r_data):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.margin_left = cell.margin_right = Inches(0.1)
            cell.margin_top = cell.margin_bottom = Inches(0.05)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_WHITE
            set_cell_border(cell, HEX_BORDER, 6350)
            tf = cell.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.text = cell_text
            p.font.name = "Calibri"
            p.font.size = Pt(13)
            p.font.bold = False
            p.font.color.rgb = C_BLACK
            p.alignment = PP_ALIGN.LEFT
    return tbl_shape


def build_dark_slide(slide, title_text, lines_list, pink_line=None):
    # Full-bleed solid dark rectangle so it is guaranteed visible across all viewers
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = C_BLACK
    bg_shape.line.fill.background()

    b_logo = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf_l = b_logo.text_frame
    p_l = tf_l.paragraphs[0]
    p_l.text = "[ABES logo]"
    p_l.font.name = "Calibri"
    p_l.font.size = Pt(10)
    p_l.font.italic = True
    p_l.font.color.rgb = C_PINK

    b_hdr = slide.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf_h = b_hdr.text_frame
    p_h = tf_h.paragraphs[0]
    p_h.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    p_h.alignment = PP_ALIGN.RIGHT
    p_h.font.name = "Calibri"
    p_h.font.size = Pt(11)
    p_h.font.color.rgb = C_PINK

    b_title = slide.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_t = b_title.text_frame
    p_t = tf_t.paragraphs[0]
    p_t.text = title_text
    p_t.font.name = "Cambria"
    p_t.font.size = Pt(36)
    p_t.font.bold = True
    p_t.font.color.rgb = C_WHITE

    b_lines = slide.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_lines = b_lines.text_frame
    for idx, line in enumerate(lines_list):
        p = tf_lines.paragraphs[0] if idx == 0 else tf_lines.add_paragraph()
        p.text = line
        p.font.name = "Calibri"
        p.font.size = Pt(15)
        p.font.color.rgb = C_CARD
        p.space_before = Pt(4)

    if pink_line:
        p_pink = tf_lines.add_paragraph()
        p_pink.text = pink_line
        p_pink.font.name = "Calibri"
        p_pink.font.size = Pt(15)
        p_pink.font.color.rgb = C_PINK
        p_pink.space_before = Pt(4)


def setup_content_slide(slide, title_text, slide_num):
    # Full-bleed solid white rectangle
    bg_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(5.625))
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = C_WHITE
    bg_shape.line.fill.background()

    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.75))
    tf = title_shape.text_frame
    p = tf.paragraphs[0]
    p.text = title_text
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Cambria"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = C_BLACK

    box_ftr = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(7.5), Inches(0.3))
    tf_f = box_ftr.text_frame
    p_f = tf_f.paragraphs[0]
    p_f.text = "HCLTech Career Shaper ABES Engineering College | HKD-2026 - Healthcare KPI Dashboard"
    p_f.font.name = "Calibri"
    p_f.font.size = Pt(9)
    p_f.font.color.rgb = C_GREY

    box_num = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.5), Inches(0.3))
    tf_n = box_num.text_frame
    p_n = tf_n.paragraphs[0]
    p_n.text = str(slide_num)
    p_n.alignment = PP_ALIGN.RIGHT
    p_n.font.name = "Calibri"
    p_n.font.size = Pt(9)
    p_n.font.color.rgb = C_NUM_GREY


def set_speaker_notes(slide, notes_text):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text


# --------------------------------------------------------------------------- BUILD PPTX
def build_all():
    print("=== Building PPTX Presentation ===")
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    # SLIDE 1
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    build_dark_slide(s1, "HKD-2026 - Healthcare KPI Dashboard", [
        "Ayush Pal | Roll No 2503201450035 / MCA 2nd Year Data Analyst",
        "Mentor: Prof. Kalpana Dwevedi"
    ], pink_line="19 November 2026")
    set_speaker_notes(s1, "Good morning mentors and evaluators. I am Ayush Pal, presenting project HKD-2026, the Healthcare KPI Dashboard.")

    # SLIDE 2
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    setup_content_slide(s2, "Problem and How Might We", 2)
    prob_text = ("Tertiary hospital administrators lack low-latency visibility into bed saturation and readmissions.\n\n"
                 "Without tracking, acute ICU occupancy breaches 85% unexpectedly and discharge bottlenecks delay emergency admissions.")
    add_badge_card(s2, 0.5, 1.25, 4.3, 3.7, "P", "Problem statement", prob_text)
    add_badge_card(s2, 5.2, 1.25, 4.3, 1.75, "1", "HMW-01: Bed Capacity",
                   "How might we give operations heads instant visibility into bed occupancy bottlenecks before they breach 85%?")
    add_badge_card(s2, 5.2, 3.2, 4.3, 1.75, "2", "HMW-02: Clinical ALOS",
                   "How might we help clinical directors benchmark departmental ALOS against 4.0-day norms so that discharge delays drop?")
    set_speaker_notes(s2, "We formulated HMW-01 for real-time bed visibility and HMW-02 for clinical ALOS benchmarks.")

    # SLIDE 3
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    setup_content_slide(s3, "Stakeholders and users", 3)
    add_badge_card(s3, 0.5, 1.35, 2.85, 3.5, "1", "Operations Director",
                   "Role: Daily bed census, ward allocation, throughput.\nNeed: Real-time bed visibility.\n\nInfluence: High")
    add_badge_card(s3, 3.55, 1.35, 2.85, 3.5, "2", "Chief Medical Officer",
                   "Role: Clinical protocols, ALOS adherence, quality.\nNeed: Specialty benchmarks.\n\nInfluence: High")
    add_badge_card(s3, 6.6, 1.35, 2.85, 3.5, "3", "Academic Evaluator",
                   "Role: Curriculum mentor & viva defense auditor.\nNeed: Traceability matrix.\n\nInfluence: Medium")
    set_speaker_notes(s3, "Our stakeholders span operations, clinical leadership, and academic curriculum mentors.")

    # SLIDE 4
    s4 = prs.slides.add_slide(prs.slide_layouts[6])
    setup_content_slide(s4, "Requirements and traceability", 4)
    rtm_headers = ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"]
    rtm_rows = [
        ["HMW-01", "Real-time bed occupancy tracking vs 85%", "Must", "US-04", "TC-F-13", "UI-02"],
        ["HMW-02", "Department ALOS benchmarking (4.0d)", "Must", "US-06", "TC-F-15", "UI-03"],
        ["HMW-03", "30d Readmission Wilson CI calculation", "Must", "US-09", "TC-F-17", "UI-03"],
        ["HMW-04", "Payer mix revenue yield & margins", "Should", "US-10", "TC-F-18", "UI-05"],
    ]
    add_abes_table(s4, 0.5, 1.3, rtm_headers, rtm_rows, [1.0, 3.4, 1.0, 1.2, 1.2, 1.2])
    b_trace = s4.shapes.add_textbox(Inches(0.5), Inches(4.3), Inches(9.0), Inches(0.35))
    p_tr = b_trace.text_frame.paragraphs[0]
    p_tr.text = "Full matrix: Project Document, section 4"
    p_tr.font.name = "Calibri"
    p_tr.font.size = Pt(12)
    p_tr.font.italic = True
    p_tr.font.color.rgb = C_GREY
    set_speaker_notes(s4, "Every requirement traces end-to-end from user problem to test case and screen.")

    # SLIDE 5
    s5 = prs.slides.add_slide(prs.slide_layouts[6])
    setup_content_slide(s5, "Data and preparation", 5)
    stats = [("53,915", "rows (admissions)"), ("21", "columns (features)"), ("2023-2025", "period (3 years)"), ("1,842", "issues fixed")]
    for idx, (val, lbl) in enumerate(stats):
        x = 0.5 + idx * 2.3
        b_val = s5.shapes.add_textbox(Inches(x), Inches(1.3), Inches(2.1), Inches(0.8))
        p_v = b_val.text_frame.paragraphs[0]
        p_v.text = val
        p_v.alignment = PP_ALIGN.CENTER
        p_v.font.name = "Cambria"
        p_v.font.size = Pt(30)
        p_v.font.bold = True
        p_v.font.color.rgb = C_RED

        b_lbl = s5.shapes.add_textbox(Inches(x), Inches(2.1), Inches(2.1), Inches(0.35))
        p_l = b_lbl.text_frame.paragraphs[0]
        p_l.text = lbl
        p_l.alignment = PP_ALIGN.CENTER
        p_l.font.name = "Calibri"
        p_l.font.size = Pt(12)
        p_l.font.color.rgb = C_GREY

    add_plain_card(s5, 0.5, 2.7, 4.3, 2.25, "Source",
                   "Dataset: Synthetic acute care EHR warehouse (seed 42).\nOne row: One completed inpatient episode with admission, ward, vitals, discharge, and billing.")
    add_plain_card(s5, 5.2, 2.7, 4.3, 2.25, "Cleaning",
                   "Top 3 fixes:\n1. 420 negative LOS resolved from timezone shifts.\n2. 312 future discharge dates clamped.\n3. 1,110 null vitals imputed with ward medians.")
    set_speaker_notes(s5, "We processed 53,915 rows and cleaned 1,842 data quality defects.")

    # SLIDE 6
    s6 = prs.slides.add_slide(prs.slide_layouts[6])
    setup_content_slide(s6, "Solution and technology", 6)
    steps = [
        ("Raw data", "CSV / xlsx"),
        ("Clean & KPIs", "Python, pandas"),
        ("Store", "SQLite app.db"),
        ("Dashboard", "Streamlit UI"),
        ("Deploy", "GitHub -> Cloud"),
    ]
    for idx, (st_top, st_bot) in enumerate(steps):
        x = 0.5 + idx * 1.85
        shape = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.5), Inches(1.6), Inches(1.2))
        shape.adjustments[0] = 0.1
        shape.fill.solid()
        shape.fill.fore_color.rgb = C_RED if idx == 3 else C_CARD
        shape.line.color.rgb = C_RED if idx == 3 else C_CARD
        tf = shape.text_frame
        p1 = tf.paragraphs[0]
        p1.text = st_top
        p1.alignment = PP_ALIGN.CENTER
        p1.font.bold = True
        p1.font.size = Pt(12)
        p1.font.color.rgb = C_WHITE if idx == 3 else C_BLACK
        p2 = tf.add_paragraph()
        p2.text = st_bot
        p2.alignment = PP_ALIGN.CENTER
        p2.font.size = Pt(11)
        p2.font.color.rgb = C_WHITE if idx == 3 else C_DARK

        if idx < 4:
            b_ar = s6.shapes.add_textbox(Inches(x + 1.6), Inches(1.5), Inches(0.25), Inches(1.2))
            p_a = b_ar.text_frame.paragraphs[0]
            p_a.text = ">"
            p_a.alignment = PP_ALIGN.CENTER
            p_a.font.size = Pt(18)
            p_a.font.bold = True
            p_a.font.color.rgb = C_RED

    b_tools = s6.shapes.add_textbox(Inches(0.5), Inches(3.15), Inches(9.0), Inches(0.5))
    p_t = b_tools.text_frame.paragraphs[0]
    p_t.text = "Tools: Python 3.12 · pandas · NumPy · SQLite · Streamlit · Plotly · pytest · GitHub"
    p_t.alignment = PP_ALIGN.CENTER
    p_t.font.name = "Calibri"
    p_t.font.size = Pt(13)
    p_t.font.color.rgb = C_DARK

    b_run = s6.shapes.add_textbox(Inches(0.5), Inches(3.75), Inches(9.0), Inches(0.5))
    p_r = b_run.text_frame.paragraphs[0]
    p_r.text = "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud"
    p_r.alignment = PP_ALIGN.CENTER
    p_r.font.name = "Calibri"
    p_r.font.size = Pt(13)
    p_r.font.italic = True
    p_r.font.color.rgb = C_GREY
    set_speaker_notes(s6, "Our architecture connects pandas ETL to an indexed SQLite warehouse and Streamlit UI.")

    # SLIDE 7
    s7 = prs.slides.add_slide(prs.slide_layouts[6])
    setup_content_slide(s7, "Dashboard", 7)
    frm = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7))
    frm.fill.solid()
    frm.fill.fore_color.rgb = C_CARD
    frm.line.color.rgb = RGBColor(0xBF, 0xBF, 0xBF)
    tf_f = frm.text_frame
    p_fr = tf_f.paragraphs[0]
    p_fr.text = "[Live Streamlit Dashboard UI-01 to UI-09]"
    p_fr.alignment = PP_ALIGN.CENTER
    p_fr.font.italic = True
    p_fr.font.size = Pt(14)
    p_fr.font.color.rgb = C_GREY

    screens = [
        ("1", "UI-01", "Executive Overview & KPI Cards"),
        ("2", "UI-02", "Bed Occupancy & Census Heatmap"),
        ("3", "UI-03", "ED Throughput & Door-to-Doctor"),
        ("4", "UI-04", "30-Day Readmission Risk Cohorts"),
    ]
    for idx, (num, code, desc) in enumerate(screens):
        y = 1.3 + idx * 0.9
        add_badge(s7, 6.7, y, num)
        b_sc = s7.shapes.add_textbox(Inches(7.25), Inches(y), Inches(2.25), Inches(0.42))
        p_s = b_sc.text_frame.paragraphs[0]
        p_s.text = f"{code}: {desc}"
        p_s.font.size = Pt(11)
        p_s.font.bold = True
        p_s.font.color.rgb = C_BLACK
    set_speaker_notes(s7, "The dashboard features 9 pages covering executive overview, bed census, and readmissions.")

    # SLIDE 8
    s8 = prs.slides.add_slide(prs.slide_layouts[6])
    setup_content_slide(s8, "Key findings", 8)
    findings = [
        ("93.8%", "ICU Bed Saturation", "Evidence: Exceeds 85% NABH safety threshold, creating acute admission diversions."),
        ("16.8%", "General Medicine Readmission", "Evidence: Driven by chronic heart failure and COPD patients over 65 years."),
        ("38.4m", "Average ED Wait Time", "Evidence: Evening surge (17:00-21:00) spikes wait time to 48m for ESI Level 2."),
    ]
    for idx, (num, mean, evid) in enumerate(findings):
        x = 0.5 + idx * 3.05
        add_card(s8, x, 1.3, 2.85, 3.6)
        b_n = s8.shapes.add_textbox(Inches(x + 0.2), Inches(1.5), Inches(2.45), Inches(1.0))
        p_n = b_n.text_frame.paragraphs[0]
        p_n.text = num
        p_n.alignment = PP_ALIGN.CENTER
        p_n.font.name = "Cambria"
        p_n.font.size = Pt(36)
        p_n.font.bold = True
        p_n.font.color.rgb = C_RED

        b_m = s8.shapes.add_textbox(Inches(x + 0.2), Inches(2.6), Inches(2.45), Inches(0.9))
        p_m = b_m.text_frame.paragraphs[0]
        p_m.text = mean
        p_m.alignment = PP_ALIGN.CENTER
        p_m.font.size = Pt(14)
        p_m.font.bold = True
        p_m.font.color.rgb = C_BLACK

        b_e = s8.shapes.add_textbox(Inches(x + 0.2), Inches(3.6), Inches(2.45), Inches(1.1))
        p_e = b_e.text_frame.paragraphs[0]
        p_e.text = evid
        p_e.alignment = PP_ALIGN.CENTER
        p_e.font.size = Pt(11)
        p_e.font.color.rgb = C_GREY
    set_speaker_notes(s8, "Our analysis identified critical findings in ICU saturation, readmissions, and ED arrival peaks.")

    # SLIDE 9
    s9 = prs.slides.add_slide(prs.slide_layouts[6])
    setup_content_slide(s9, "Testing and validation", 9)
    b_pt = s9.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(2.6), Inches(1.0))
    p_pt = b_pt.text_frame.paragraphs[0]
    p_pt.text = "70 / 71"
    p_pt.alignment = PP_ALIGN.CENTER
    p_pt.font.name = "Cambria"
    p_pt.font.size = Pt(44)
    p_pt.font.bold = True
    p_pt.font.color.rgb = C_RED

    b_stat = s9.shapes.add_textbox(Inches(0.5), Inches(2.4), Inches(2.6), Inches(0.9))
    tf_st = b_stat.text_frame
    for idx, tl in enumerate(["70 test cases passed", "pytest: 70 passed", "manual: 19 verified"]):
        p = tf_st.paragraphs[0] if idx == 0 else tf_st.add_paragraph()
        p.text = tl
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(13)
        p.font.color.rgb = C_GREY

    test_headers = ["Test", "Checks", "Result"]
    test_rows = [
        ["TC-F-01", "Executive KPI cards calculation precision", "PASS"],
        ["TC-F-13", "Ward bed occupancy vs 85% threshold", "PASS"],
        ["TC-T-44", "Page latency (< 3.0s, actual 1.83s)", "PASS"],
    ]
    add_abes_table(s9, 3.4, 1.35, test_headers, test_rows, [1.1, 3.9, 1.1])
    set_speaker_notes(s9, "Quality is verified by 70 automated pytest test cases validating data integrity and latency.")

    # SLIDE 10
    s10 = prs.slides.add_slide(prs.slide_layouts[6])
    setup_content_slide(s10, "Recommendation, limitations and next steps", 10)
    add_badge_card(s10, 0.5, 1.25, 2.85, 3.7, "R", "Recommendation",
                   "Reallocate 8 step-down telemetry beds to ICU to relieve the 93.8% saturation.\n\nAutomate 48-hour post-discharge telehealth follow-ups for high-risk cardiac cohorts.")
    add_badge_card(s10, 3.58, 1.25, 2.85, 3.7, "L", "Limitations",
                   "Dataset is synthetic (seed 42), though calibrated to clinical distributions.\n\nCensus relies on midnight snapshots rather than continuous streaming bed telemetry.")
    add_badge_card(s10, 6.65, 1.25, 2.85, 3.7, "N", "Next steps",
                   "Integrate HL7/FHIR telemetry streams for live real-time bed tracking.\n\nExpand predictive readmission models with clinical EHR lab values.")
    set_speaker_notes(s10, "We recommend bed reallocations and telehealth follow-ups, with HL7 telemetry as next steps.")

    # SLIDE 11
    s11 = prs.slides.add_slide(prs.slide_layouts[6])
    build_dark_slide(s11, "Live demo and Q&A", [
        "App: https://2503201450035-hkd-2026.streamlit.app",
        "Code: https://github.com/ayushpal1506/HclTraining",
        "Thank you",
    ])
    set_speaker_notes(s11, "Thank you for your time. The application is live and I invite questions from the panel.")

    out_pptx = DOCS / "HKD-2026_Presentation.pptx"
    prs.save(out_pptx)
    shutil.copy(out_pptx, SUB_DIR / "2_Project_Presentation.pptx")
    print(f"Saved PPTX to: {out_pptx} and {SUB_DIR / '2_Project_Presentation.pptx'}")


# --------------------------------------------------------------------------- PILLOW HIGH-RES RENDERER
def render_all_slides():
    print("=== Rendering High-Resolution Slide Images with Pillow ===")
    w, h = 1600, 900
    dpi = 160

    # Slide 1
    im1 = Image.new("RGB", (w, h), RGB_BLACK)
    d1 = ImageDraw.Draw(im1)
    d1.text((80, 56), "[ABES logo]", fill=RGB_PINK, font=FONT_HELVETICA_12)
    d1.text((700, 56), "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad", fill=RGB_PINK, font=FONT_HELVETICA_BOLD_12)
    d1.text((80, 248), "HKD-2026 - Healthcare KPI Dashboard", fill=RGB_WHITE, font=FONT_GEORGIA_36)
    d1.text((80, 480), "Ayush Pal | Roll No 2503201450035 / MCA 2nd Year Data Analyst\nMentor: Prof. Kalpana Dwevedi", fill=RGB_CARD, font=FONT_HELVETICA_BOLD_15, spacing=10)
    d1.text((80, 560), "19 November 2026", fill=RGB_PINK, font=FONT_HELVETICA_13)
    im1.save(IMG_DIR / "slide_01.png")

    # Helper for white slide header/footer
    def draw_content_base(d, title, num):
        d.rectangle([0, 0, w, h], fill=RGB_WHITE)
        d.text((800, 48), title, fill=RGB_BLACK, font=FONT_GEORGIA_28, anchor="mt")
        d.text((80, 832), "HCLTech Career Shaper ABES Engineering College | HKD-2026 - Healthcare KPI Dashboard", fill=RGB_GREY, font=FONT_HELVETICA_9)
        d.text((1480, 832), str(num), fill=RGB_NUM_GREY, font=FONT_HELVETICA_9)

    def draw_badge(d, x, y, text):
        d.ellipse([x, y, x + 67, y + 67], fill=RGB_RED)
        d.text((x + 33, y + 33), str(text), fill=RGB_WHITE, font=FONT_HELVETICA_BOLD_15, anchor="mm")

    def draw_card(d, x, y, cw, ch):
        d.rounded_rectangle([x, y, x + cw, y + ch], radius=14, fill=RGB_CARD, outline=RGB_CARD)

    # Slide 2
    im2 = Image.new("RGB", (w, h), RGB_WHITE)
    d2 = ImageDraw.Draw(im2)
    draw_content_base(d2, "Problem and How Might We", 2)
    draw_card(d2, 80, 200, 688, 592)
    draw_badge(d2, 112, 232, "P")
    d2.text((200, 240), "Problem statement", fill=RGB_BLACK, font=FONT_HELVETICA_BOLD_15)
    d2.text((112, 320), "Tertiary hospital administrators lack low-latency visibility into acute\nbed occupancy bottlenecks and readmission risks.\n\nWithout automated tracking, ICU beds breach 85% unexpectedly and\ndischarge delays compound, leading to emergency diversions.", fill=RGB_DARK, font=FONT_HELVETICA_13, spacing=8)
    
    draw_card(d2, 832, 200, 688, 280)
    draw_badge(d2, 864, 232, "1")
    d2.text((952, 240), "HMW-01: Bed Capacity", fill=RGB_BLACK, font=FONT_HELVETICA_BOLD_15)
    d2.text((864, 320), "How might we give operations heads instant visibility into bed occupancy\nbottlenecks before they breach 85% so emergency care is preserved?", fill=RGB_DARK, font=FONT_HELVETICA_13, spacing=6)

    draw_card(d2, 832, 512, 688, 280)
    draw_badge(d2, 864, 544, "2")
    d2.text((952, 552), "HMW-02: Clinical ALOS", fill=RGB_BLACK, font=FONT_HELVETICA_BOLD_15)
    d2.text((864, 632), "How might we help clinical directors benchmark departmental ALOS\nagainst 4.0-day recovery norms so discharge delays drop?", fill=RGB_DARK, font=FONT_HELVETICA_13, spacing=6)
    im2.save(IMG_DIR / "slide_02.png")

    # Slide 3
    im3 = Image.new("RGB", (w, h), RGB_WHITE)
    d3 = ImageDraw.Draw(im3)
    draw_content_base(d3, "Stakeholders and users", 3)
    sh_list = [
        ("1", "Operations Director", "Role: Daily bed census & ward management.\nNeed: Real-time bed bottleneck visibility.\n\nInfluence: High", 80),
        ("2", "Chief Medical Officer", "Role: Clinical protocols, ALOS adherence.\nNeed: Departmental benchmark drilldowns.\n\nInfluence: High", 568),
        ("3", "Academic Evaluator", "Role: Programme mentor & viva auditor.\nNeed: Full requirement traceability.\n\nInfluence: Medium", 1056),
    ]
    for b_t, h_t, b_y, x_pos in sh_list:
        draw_card(d3, x_pos, 216, 456, 560)
        draw_badge(d3, x_pos + 32, 248, b_t)
        d3.text((x_pos + 120, 260), h_t, fill=RGB_BLACK, font=FONT_HELVETICA_BOLD_15)
        d3.text((x_pos + 32, 340), b_y, fill=RGB_DARK, font=FONT_HELVETICA_13, spacing=8)
    im3.save(IMG_DIR / "slide_03.png")

    # Slide 4
    im4 = Image.new("RGB", (w, h), RGB_WHITE)
    d4 = ImageDraw.Draw(im4)
    draw_content_base(d4, "Requirements and traceability", 4)
    # Draw Table
    tx, ty = 80, 208
    cols = [160, 544, 160, 192, 192, 192]
    headers = ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"]
    # Header row
    cx = tx
    for idx, c_w in enumerate(cols):
        d4.rectangle([cx, ty, cx + c_w, ty + 64], fill=RGB_RED, outline=RGB_BORDER)
        d4.text((cx + 12, ty + 24), headers[idx], fill=RGB_WHITE, font=FONT_HELVETICA_BOLD_12)
        cx += c_w
    # Rows
    rows_d = [
        ["HMW-01", "Real-time bed occupancy tracking vs 85%", "Must", "US-04", "TC-F-13", "UI-02"],
        ["HMW-02", "Department ALOS benchmarking (4.0d)", "Must", "US-06", "TC-F-15", "UI-03"],
        ["HMW-03", "30d Readmission Wilson CI calculation", "Must", "US-09", "TC-F-17", "UI-03"],
        ["HMW-04", "Payer mix revenue yield & margins", "Should", "US-10", "TC-F-18", "UI-05"],
    ]
    ry = ty + 64
    for row in rows_d:
        cx = tx
        for idx, c_w in enumerate(cols):
            d4.rectangle([cx, ry, cx + c_w, ry + 64], fill=RGB_WHITE, outline=RGB_BORDER)
            d4.text((cx + 12, ry + 24), row[idx], fill=RGB_BLACK, font=FONT_HELVETICA_12)
            cx += c_w
        ry += 64
    d4.text((80, 680), "Full matrix: Project Document, section 4", fill=RGB_GREY, font=FONT_HELVETICA_13)
    im4.save(IMG_DIR / "slide_04.png")

    # Slide 5
    im5 = Image.new("RGB", (w, h), RGB_WHITE)
    d5 = ImageDraw.Draw(im5)
    draw_content_base(d5, "Data and preparation", 5)
    s_list = [("53,915", "rows (admissions)", 80), ("21", "columns (features)", 448), ("2023-2025", "period (3 years)", 816), ("1,842", "issues fixed", 1184)]
    for num, lbl, px in s_list:
        d5.text((px + 160, 208), num, fill=RGB_RED, font=FONT_GEORGIA_36, anchor="mm")
        d5.text((px + 160, 280), lbl, fill=RGB_GREY, font=FONT_HELVETICA_12, anchor="mm")
    draw_card(d5, 80, 432, 688, 360)
    d5.text((112, 456), "Source", fill=RGB_BLACK, font=FONT_HELVETICA_BOLD_15)
    d5.text((112, 512), "Dataset: Synthetic acute care EHR warehouse (seed 42).\nOne row: One completed inpatient admission episode with ward,\nvitals, length of stay, discharge status, and billing yield.", fill=RGB_DARK, font=FONT_HELVETICA_13, spacing=6)
    draw_card(d5, 832, 432, 688, 360)
    d5.text((864, 456), "Cleaning", fill=RGB_BLACK, font=FONT_HELVETICA_BOLD_15)
    d5.text((864, 512), "Top 3 cleaning fixes executed:\n1. 420 negative LOS resolved from timezone shifts.\n2. 312 future discharge timestamps corrected.\n3. 1,110 null vitals imputed with ward medians.", fill=RGB_DARK, font=FONT_HELVETICA_13, spacing=6)
    im5.save(IMG_DIR / "slide_05.png")

    # Slide 6
    im6 = Image.new("RGB", (w, h), RGB_WHITE)
    d6 = ImageDraw.Draw(im6)
    draw_content_base(d6, "Solution and technology", 6)
    wf_steps = [("Raw data", "CSV / xlsx"), ("Clean & KPIs", "Python, pandas"), ("Store", "SQLite app.db"), ("Dashboard", "Streamlit UI"), ("Deploy", "GitHub -> Cloud")]
    for idx, (top, bot) in enumerate(wf_steps):
        sx = 80 + idx * 296
        fill_b = RGB_RED if idx == 3 else RGB_CARD
        txt_b = RGB_WHITE if idx == 3 else RGB_BLACK
        txt_sub = RGB_WHITE if idx == 3 else RGB_DARK
        d6.rounded_rectangle([sx, 240, sx + 256, 432], radius=16, fill=fill_b)
        d6.text((sx + 128, 310), top, fill=txt_b, font=FONT_HELVETICA_BOLD_15, anchor="mm")
        d6.text((sx + 128, 355), bot, fill=txt_sub, font=FONT_HELVETICA_12, anchor="mm")
        if idx < 4:
            d6.text((sx + 276, 336), ">", fill=RGB_RED, font=FONT_HELVETICA_BOLD_15, anchor="mm")
    d6.text((800, 520), "Tools: Python 3.12 · pandas · NumPy · SQLite · Streamlit · Plotly · pytest · GitHub", fill=RGB_DARK, font=FONT_HELVETICA_BOLD_15, anchor="mm")
    d6.text((800, 580), "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud", fill=RGB_GREY, font=FONT_HELVETICA_13, anchor="mm")
    im6.save(IMG_DIR / "slide_06.png")

    # Slide 7
    im7 = Image.new("RGB", (w, h), RGB_WHITE)
    d7 = ImageDraw.Draw(im7)
    draw_content_base(d7, "Dashboard", 7)
    d7.rectangle([80, 200, 960, 792], fill=RGB_CARD, outline=RGB_BORDER, width=2)
    d7.text((520, 496), "[Live Multi-Page Streamlit Dashboard UI-01 to UI-09]", fill=RGB_GREY, font=FONT_HELVETICA_BOLD_15, anchor="mm")
    sc_items = [("1", "UI-01", "Executive Overview & 6 KPI Cards"), ("2", "UI-02", "Bed Occupancy & Census Heatmap"), ("3", "UI-03", "ED Door-to-Doctor Throughput"), ("4", "UI-04", "30-Day Readmission Risk Cohorts")]
    for idx, (b_n, c_code, desc) in enumerate(sc_items):
        sy = 220 + idx * 144
        draw_badge(d7, 1040, sy, b_n)
        d7.text((1130, sy + 18), f"{c_code}:", fill=RGB_RED, font=FONT_HELVETICA_BOLD_15)
        d7.text((1130, sy + 44), desc, fill=RGB_BLACK, font=FONT_HELVETICA_13)
    im7.save(IMG_DIR / "slide_07.png")

    # Slide 8
    im8 = Image.new("RGB", (w, h), RGB_WHITE)
    d8 = ImageDraw.Draw(im8)
    draw_content_base(d8, "Key findings", 8)
    findings_list = [
        ("93.8%", "ICU Bed Saturation", "Evidence: Exceeds 85% NABH safety threshold, creating acute admission diversions.", 80),
        ("16.8%", "General Medicine Readmission", "Evidence: Driven by chronic heart failure and COPD patients over 65 years.", 568),
        ("38.4m", "Average ED Wait Time", "Evidence: Evening surge (17:00-21:00) spikes wait time to 48m for ESI Level 2.", 1056),
    ]
    for num, mean, evid, px in findings_list:
        draw_card(d8, px, 208, 456, 576)
        d8.text((px + 228, 280), num, fill=RGB_RED, font=FONT_GEORGIA_44, anchor="mm")
        d8.text((px + 228, 380), mean, fill=RGB_BLACK, font=FONT_HELVETICA_BOLD_15, anchor="mm")
        d8.text((px + 32, 460), evid, fill=RGB_GREY, font=FONT_HELVETICA_13, spacing=6)
    im8.save(IMG_DIR / "slide_08.png")

    # Slide 9
    im9 = Image.new("RGB", (w, h), RGB_WHITE)
    d9 = ImageDraw.Draw(im9)
    draw_content_base(d9, "Testing and validation", 9)
    d9.text((288, 260), "70 / 71", fill=RGB_RED, font=FONT_GEORGIA_44, anchor="mm")
    d9.text((288, 340), "70 test cases passed\npytest: 70 passed (1 skip)\nmanual: 19 verified", fill=RGB_GREY, font=FONT_HELVETICA_13, anchor="mm", spacing=6)
    tx, ty = 544, 216
    t_cols = [176, 624, 176]
    t_headers = ["Test", "Checks", "Result"]
    cx = tx
    for idx, c_w in enumerate(t_cols):
        d9.rectangle([cx, ty, cx + c_w, ty + 64], fill=RGB_RED, outline=RGB_BORDER)
        d9.text((cx + 12, ty + 24), t_headers[idx], fill=RGB_WHITE, font=FONT_HELVETICA_BOLD_12)
        cx += c_w
    t_rows_data = [
        ["TC-F-01", "Executive KPI cards calculation precision", "PASS"],
        ["TC-F-13", "Ward bed occupancy vs 85% threshold", "PASS"],
        ["TC-T-44", "Page latency (< 3.0s, actual 1.83s)", "PASS"],
    ]
    ry = ty + 64
    for r in t_rows_data:
        cx = tx
        for idx, c_w in enumerate(t_cols):
            d9.rectangle([cx, ry, cx + c_w, ry + 64], fill=RGB_WHITE, outline=RGB_BORDER)
            d9.text((cx + 12, ry + 24), r[idx], fill=RGB_BLACK, font=FONT_HELVETICA_12)
            cx += c_w
        ry += 64
    im9.save(IMG_DIR / "slide_09.png")

    # Slide 10
    im10 = Image.new("RGB", (w, h), RGB_WHITE)
    d10 = ImageDraw.Draw(im10)
    draw_content_base(d10, "Recommendation, limitations and next steps", 10)
    rec_items = [
        ("R", "Recommendation", "Reallocate 8 step-down telemetry beds to ICU to relieve the 93.8% saturation.\n\nAutomate 48-hour post-discharge telehealth follow-ups for high-risk cardiac cohorts.", 80),
        ("L", "Limitations", "Dataset is synthetic (seed 42), though calibrated to clinical distributions.\n\nCensus relies on midnight snapshots rather than continuous streaming bed telemetry.", 568),
        ("N", "Next steps", "Integrate HL7/FHIR telemetry streams for live real-time bed tracking.\n\nExpand predictive readmission models with clinical EHR lab values.", 1056),
    ]
    for b_char, b_title, b_desc, px in rec_items:
        draw_card(d10, px, 200, 456, 592)
        draw_badge(d10, px + 32, 232, b_char)
        d10.text((px + 120, 244), b_title, fill=RGB_BLACK, font=FONT_HELVETICA_BOLD_15)
        d10.text((px + 32, 330), b_desc, fill=RGB_DARK, font=FONT_HELVETICA_13, spacing=8)
    im10.save(IMG_DIR / "slide_10.png")

    # Slide 11
    im11 = Image.new("RGB", (w, h), RGB_BLACK)
    d11 = ImageDraw.Draw(im11)
    d11.text((80, 56), "[ABES logo]", fill=RGB_PINK, font=FONT_HELVETICA_12)
    d11.text((700, 56), "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad", fill=RGB_PINK, font=FONT_HELVETICA_BOLD_12)
    d11.text((80, 248), "Live demo and Q&A", fill=RGB_WHITE, font=FONT_GEORGIA_36)
    d11.text((80, 440), "App: https://2503201450035-hkd-2026.streamlit.app\nCode: https://github.com/ayushpal1506/HclTraining\n\nThank you", fill=RGB_CARD, font=FONT_HELVETICA_BOLD_15, spacing=14)
    im11.save(IMG_DIR / "slide_11.png")

    print("All 11 slide images rendered with complete visual elements!")


def compile_presentation_pdf():
    print("=== Compiling Presentation PDF from Rendered Slides ===")
    imgs = sorted(IMG_DIR.glob("slide_*.png"))
    if not imgs:
        print("No slide images found!")
        return
    im_list = [Image.open(f).convert("RGB") for f in imgs]
    out_pdf = DOCS / "HKD-2026_Presentation.pdf"
    im_list[0].save(out_pdf, save_all=True, append_images=im_list[1:], resolution=150.0)
    shutil.copy(out_pdf, SUB_DIR / "2_Project_Presentation.pdf")
    print(f"Compiled PDF: {out_pdf} and {SUB_DIR / '2_Project_Presentation.pdf'}")


if __name__ == "__main__":
    build_all()
    render_all_slides()
    compile_presentation_pdf()
    print("=== Complete Presentation Suite Updated Successfully! ===")
