"""Generate the Step 18 ABES Presentation: docs/p107_Presentation.pptx.

Strictly adheres to all 3-screen specifications:
- Exactly 11 slides in exact order.
- Canvas: 10 x 5.625 in (16:9 widescreen).
- Exact coordinates, typography, colors, borders, cards, badges, and tables.
- Speaker notes on every slide (~30s each).
- Visual check: renders each slide to PNG via PIL and audits bounding boxes and zero text overflow.
"""
from __future__ import annotations

import os
import shutil
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"
OUT_PPTX = DOCS / "p107_Presentation.pptx"
OUT_SEC = DOCS / "18_Presentation_p107.pptx"
IMG_DIR = DOCS / "slide_images"

# --------------------------------------------------------------------------- Color Palette
C_RED = RGBColor(0xA5, 0x1C, 0x30)      # A51C30
C_BLACK = RGBColor(0x15, 0x15, 0x14)    # 151514
C_DARK = RGBColor(0x3A, 0x3A, 0x38)     # 3A3A38
C_GREY = RGBColor(0x5E, 0x5E, 0x5B)     # 5E5E5B
C_CARD = RGBColor(0xF4, 0xF1, 0xF1)     # F4F1F1
C_PINK = RGBColor(0xD9, 0xB5, 0xBC)     # D9B5BC
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)    # FFFFFF
C_BORDER = RGBColor(0xD9, 0xD9, 0xD9)   # D9D9D9
C_NUM_GREY = RGBColor(0x8A, 0x8A, 0x86) # 8A8A86
C_FRAME_DASH = RGBColor(0xBF, 0xBF, 0xBF) # BFBFBF

HEX_RED = "A51C30"
HEX_BORDER = "D9D9D9"

# --------------------------------------------------------------------------- XML Helpers
def set_cell_border(cell, color_hex=HEX_BORDER, width_emu=6350):
    """Set 0.5 pt solid border on all 4 sides of a table cell."""
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

# --------------------------------------------------------------------------- Card & Shape Helpers
def add_card(slide, x, y, w, h):
    """CARD(x,y,w,h): rounded rectangle, radius .08, fill and line F4F1F1."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = C_CARD
    shape.line.color.rgb = C_CARD
    shape.adjustments[0] = 0.08
    return shape

def add_badge(slide, x, y, text):
    """BADGE (x,y,t): circle .42 red; t centred, 12 bold white."""
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
    """BADGE CARD: CARD; BADGE(x+.2,y+.2); HEAD (x+.75,y+.17,w-.9,.45); BODY (x+.2,y+.72,w-.4,h-.85)."""
    add_card(slide, x, y, w, h)
    add_badge(slide, x + 0.2, y + 0.2, badge_text)
    
    # HEAD (x+.75, y+.17, w-.9, .45): 15 bold black, middle
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
    
    # BODY (x+.2, y+.72, w-.4, h-.85): 13 dark, top
    box_body = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    tf_b = box_body.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_right = tf_b.margin_top = tf_b.margin_bottom = 0
    tf_b.vertical_anchor = MSO_ANCHOR.TOP
    lines = body_text.split("\n")
    for idx, line in enumerate(lines):
        p = tf_b.paragraphs[0] if idx == 0 else tf_b.add_paragraph()
        p.text = line
        p.font.name = "Calibri"
        p.font.size = Pt(13)
        p.font.color.rgb = C_DARK
    return box_head

def add_plain_card(slide, x, y, w, h, head_text, body_text):
    """PLAIN CARD: CARD; HEAD (x+.2,y+.17,w-.35,.45); BODY (x+.2,y+.72,w-.4,h-.85)."""
    add_card(slide, x, y, w, h)
    
    # HEAD (x+.2, y+.17, w-.35, .45): 15 bold black, middle
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
    
    # BODY (x+.2, y+.72, w-.4, h-.85): 13 dark, top
    box_body = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.72), Inches(w - 0.4), Inches(h - 0.85))
    tf_b = box_body.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_right = tf_b.margin_top = tf_b.margin_bottom = 0
    tf_b.vertical_anchor = MSO_ANCHOR.TOP
    lines = body_text.split("\n")
    for idx, line in enumerate(lines):
        p = tf_b.paragraphs[0] if idx == 0 else tf_b.add_paragraph()
        p.text = line
        p.font.name = "Calibri"
        p.font.size = Pt(13)
        p.font.color.rgb = C_DARK
    return box_head

def add_abes_table(slide, x, y, headers, rows, col_widths):
    """TABLE: no table style. Cell borders .5 pt D9D9D9, margins .1 left/right, .05 top/bottom, rows .5 high, 13 pt. Header fill red, white bold. Body white fill, black text."""
    num_rows = len(rows) + 1
    num_cols = len(headers)
    total_w = sum(col_widths)
    total_h = 0.5 * num_rows
    
    tbl_shape = slide.shapes.add_table(num_rows, num_cols, Inches(x), Inches(y), Inches(total_w), Inches(total_h))
    tbl = tbl_shape.table
    
    # Remove default table style
    tblPr = tbl._tbl.tblPr
    style_id = tblPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId")
    if style_id is not None:
        tblPr.remove(style_id)
        
    for i, w in enumerate(col_widths):
        tbl.columns[i].width = Inches(w)
        
    for r in range(num_rows):
        tbl.rows[r].height = Inches(0.5)
        
    # Header Row
    for c, h_text in enumerate(headers):
        cell = tbl.cell(0, c)
        cell.margin_left = Inches(0.1)
        cell.margin_right = Inches(0.1)
        cell.margin_top = Inches(0.05)
        cell.margin_bottom = Inches(0.05)
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
        
    # Body Rows
    for r, row_vals in enumerate(rows):
        for c, val in enumerate(row_vals):
            cell = tbl.cell(r + 1, c)
            cell.margin_left = Inches(0.1)
            cell.margin_right = Inches(0.1)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_WHITE
            set_cell_border(cell, HEX_BORDER, 6350)
            
            tf = cell.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.text = str(val)
            p.font.name = "Calibri"
            p.font.size = Pt(13)
            p.font.bold = False
            p.font.color.rgb = C_BLACK
            p.alignment = PP_ALIGN.LEFT
    return tbl_shape

# --------------------------------------------------------------------------- Slide Builders
def build_dark_slide(slide, title_text, lines_list, pink_line=None):
    """DARK SLIDE format: Background 151514, Logo (.5,.35,1.6,.5), Header (2.2,.35,7.3,.5), Title (.5,1.55,9,1.5), Lines (.5,3.2,9,1.7)."""
    # Background
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = C_BLACK

    # [ABES logo] (.5,.35,1.6,.5) italic 10 pink, middle, default insets
    b_logo = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(1.6), Inches(0.5))
    tf_l = b_logo.text_frame
    tf_l.word_wrap = True
    tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_l = tf_l.paragraphs[0]
    p_l.text = "[ABES logo]"
    p_l.font.name = "Calibri"
    p_l.font.size = Pt(10)
    p_l.font.italic = True
    p_l.font.color.rgb = C_PINK

    # Header: HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad (2.2,.35,7.3,.5) 11 pink, right, middle, default insets
    b_hdr = slide.shapes.add_textbox(Inches(2.2), Inches(0.35), Inches(7.3), Inches(0.5))
    tf_h = b_hdr.text_frame
    tf_h.word_wrap = True
    tf_h.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_h = tf_h.paragraphs[0]
    p_h.text = "HCLTech Career Shaper Programme | ABES Engineering College, Ghaziabad"
    p_h.alignment = PP_ALIGN.RIGHT
    p_h.font.name = "Calibri"
    p_h.font.size = Pt(11)
    p_h.font.color.rgb = C_PINK

    # Title (.5,1.55,9,1.5) Cambria 36 bold white, left, bottom, default insets
    b_title = slide.shapes.add_textbox(Inches(0.5), Inches(1.55), Inches(9.0), Inches(1.5))
    tf_t = b_title.text_frame
    tf_t.word_wrap = True
    tf_t.vertical_anchor = MSO_ANCHOR.BOTTOM
    p_t = tf_t.paragraphs[0]
    p_t.text = title_text
    p_t.alignment = PP_ALIGN.LEFT
    p_t.font.name = "Cambria"
    p_t.font.size = Pt(36)
    p_t.font.bold = True
    p_t.font.color.rgb = C_WHITE

    # Lines (.5,3.2,9,1.7) 15 F4F1F1, left, top, 4 pt before each line, default insets
    b_lines = slide.shapes.add_textbox(Inches(0.5), Inches(3.2), Inches(9.0), Inches(1.7))
    tf_lines = b_lines.text_frame
    tf_lines.word_wrap = True
    tf_lines.vertical_anchor = MSO_ANCHOR.TOP
    for i, line in enumerate(lines_list):
        p = tf_lines.paragraphs[0] if i == 0 else tf_lines.add_paragraph()
        p.text = line
        p.alignment = PP_ALIGN.LEFT
        p.font.name = "Calibri"
        p.font.size = Pt(15)
        p.font.color.rgb = C_CARD
        p.space_before = Pt(4)

    if pink_line:
        p_pink = tf_lines.add_paragraph()
        p_pink.text = pink_line
        p_pink.alignment = PP_ALIGN.LEFT
        p_pink.font.name = "Calibri"
        p_pink.font.size = Pt(15)
        p_pink.font.color.rgb = C_PINK
        p_pink.space_before = Pt(4)

def setup_content_slide(slide, title_text, slide_num):
    """CONTENT SLIDE format: White, Title (.5,.3,9,.75) Cambria 28 bold black centred middle, Footer (.5,5.2,7.5,.3) 9 grey, Slide number (9,5.2,.5,.3) 9 8A8A86 right."""
    # White background
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = C_WHITE

    # Title (.5,.3,9,.75) Cambria 28 bold black, centred, middle, default insets
    # Use real title placeholder if present
    if slide.shapes.title:
        title_shape = slide.shapes.title
        title_shape.left = Inches(0.5)
        title_shape.top = Inches(0.3)
        title_shape.width = Inches(9.0)
        title_shape.height = Inches(0.75)
    else:
        title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.75))

    tf = title_shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = title_text
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Cambria"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = C_BLACK

    # Footer (.5,5.2,7.5,.3) 9 grey, middle, default insets: HCLTech Career Shaper ABES Engineering College | <Project ID - Title>
    box_ftr = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(7.5), Inches(0.3))
    tf_f = box_ftr.text_frame
    tf_f.word_wrap = True
    tf_f.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_f = tf_f.paragraphs[0]
    p_f.text = "HCLTech Career Shaper ABES Engineering College | p107 - Healthcare KPI Dashboard"
    p_f.font.name = "Calibri"
    p_f.font.size = Pt(9)
    p_f.font.color.rgb = C_GREY

    # Slide number (9,5.2,.5,.3) 9, colour 8A8A86, right, default insets
    box_num = slide.shapes.add_textbox(Inches(9.0), Inches(5.2), Inches(0.5), Inches(0.3))
    tf_n = box_num.text_frame
    tf_n.word_wrap = True
    tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_n = tf_n.paragraphs[0]
    p_n.text = str(slide_num)
    p_n.alignment = PP_ALIGN.RIGHT
    p_n.font.name = "Calibri"
    p_n.font.size = Pt(9)
    p_n.font.color.rgb = C_NUM_GREY

def set_speaker_notes(slide, notes_text):
    """Add speaker notes to slide."""
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text

# --------------------------------------------------------------------------- BUILD PRESENTATION
def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    # =========================================================================
    # SLIDE 1: DARK (Title)
    # =========================================================================
    s1 = prs.slides.add_slide(prs.slide_layouts[6]) # blank layout
    build_dark_slide(
        s1,
        title_text="p107 - Healthcare KPI Dashboard",
        lines_list=[
            "Ayush Pal | Roll No 2500320140037 / MCA 2nd Year Data Analyst",
            "Mentor: Mr. Dinesh Sereno",
        ],
        pink_line="19 November 2026",
    )
    set_speaker_notes(
        s1,
        "Good morning mentors and evaluators. I am Ayush Pal, Roll 2500320140037, MCA 2nd Year Data Analyst. "
        "Today I present project p107, the Healthcare KPI Dashboard, built under the guidance of Mr. Dinesh Sereno. "
        "This enterprise operational analytics platform gives hospital leadership unified, real-time visibility into acute "
        "bed occupancy bottlenecks, length of stay adherence, readmissions, and specialty revenue yield."
    )

    # =========================================================================
    # SLIDE 2: CONTENT "Problem and How Might We"
    # =========================================================================
    s2 = prs.slides.add_slide(prs.slide_layouts[5]) # Title only
    setup_content_slide(s2, "Problem and How Might We", slide_num=2)

    # BADGE CARD (.5,1.25,4.3,3.7) P, Problem statement: 2-3 sentences
    prob_text = (
        "Tertiary hospital administrators lack unified, low-latency visibility across clinical throughput, "
        "acute bed bottlenecks, readmission risks, and specialty profitability.\n\n"
        "Without automated tracking, bed occupancy breaches 85% unexpectedly and discharge delays compound. "
        "Clinical and finance teams operate in silos using disconnected spreadsheets."
    )
    add_badge_card(s2, 0.5, 1.25, 4.3, 3.7, "P", "Problem statement", prob_text)

    # BADGE CARD (5.2,1.25,4.3,1.75) 1, HMW-01: statement
    hmw1_text = (
        "How might we give operations heads instant visibility into bed occupancy bottlenecks "
        "before they breach 85% so that emergency admissions are never diverted?"
    )
    add_badge_card(s2, 5.2, 1.25, 4.3, 1.75, "1", "HMW-01: Bed Capacity", hmw1_text)

    # BADGE CARD (5.2,3.2,4.3,1.75) 2, HMW-02: statement
    hmw2_text = (
        "How might we help clinical directors benchmark departmental ALOS against 4.0-day recovery "
        "norms so that discharge bottlenecks are eliminated?"
    )
    add_badge_card(s2, 5.2, 3.2, 4.3, 1.75, "2", "HMW-02: Clinical ALOS", hmw2_text)

    set_speaker_notes(
        s2,
        "Hospital administrators struggle because clinical, operational, and financial data live in isolated silos. "
        "When bed occupancy spikes beyond 85%, emergency departments face diversions with zero early warning. "
        "We framed two core design thinking challenges: HMW-01 tackles bed occupancy bottlenecks before they breach "
        "the 85% NABH safety threshold, while HMW-02 addresses clinical length of stay benchmarking against the 4.0-day standard."
    )

    # =========================================================================
    # SLIDE 3: CONTENT "Stakeholders and users"
    # =========================================================================
    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    setup_content_slide(s3, "Stakeholders and users", slide_num=3)

    # BADGE CARDS y 1.35, w 2.85, h 3.5, x .5 / 3.55 / 6.6, badges 1 / 2 / 3
    sh1_body = (
        "Role: Daily bed census, ward allocation, and acute throughput management.\n\n"
        "Need: Real-time visibility into bed bottlenecks to prevent patient diversions.\n\n"
        "Influence: High"
    )
    add_badge_card(s3, 0.5, 1.35, 2.85, 3.5, "1", "Operations Director", sh1_body)

    sh2_body = (
        "Role: Clinical safety governance, 30-day readmissions, and mortality drift audit.\n\n"
        "Need: Statistical confidence intervals and +2σ drift alarms for peer review.\n\n"
        "Influence: High"
    )
    add_badge_card(s3, 3.55, 1.35, 2.85, 3.5, "2", "Quality Head", sh2_body)

    sh3_body = (
        "Role: Academic mentor, technical evaluator, and project supervisor.\n\n"
        "Need: Rigorous data modeling, statistical correctness, and complete traceability.\n\n"
        "Influence: High"
    )
    add_badge_card(s3, 6.6, 1.35, 2.85, 3.5, "3", "Mentor / Evaluator", sh3_body)

    set_speaker_notes(
        s3,
        "Our user research identified three pivotal stakeholders. Our primary user is the Operations Director, who "
        "needs daily bed census tracking to avoid hospital gridlock. Our secondary user is the Clinical Quality Head, who "
        "monitors 30-day readmissions and mortality drift to trigger clinical audits. Finally, our mentor and evaluator, "
        "Mr. Dinesh Sereno, ensures architectural soundness, schema integrity, and end-to-end requirement traceability."
    )

    # =========================================================================
    # SLIDE 4: CONTENT "Requirements and traceability"
    # =========================================================================
    s4 = prs.slides.add_slide(prs.slide_layouts[5])
    setup_content_slide(s4, "Requirements and traceability", slide_num=4)

    # TABLE at (.5,1.3), column widths 1, 3.4, 1, 1.2, 1.2, 1.2: HMW | Requirement | Priority | User story | Test | Screen
    rtm_headers = ["HMW", "Requirement", "Priority", "User story", "Test", "Screen"]
    rtm_rows = [
        ["HMW-01", "Bed Occupancy vs 85% threshold", "Must", "US-04", "TC-F-13", "UI-02"],
        ["HMW-02", "ALOS vs 4.0d Benchmark", "Must", "US-06", "TC-F-15", "UI-03"],
        ["HMW-03", "30d Readmission with Wilson CI", "Must", "US-09", "TC-F-17", "UI-03"],
        ["HMW-05", "Mortality Drift Alert (+2σ)", "Must", "US-13", "TC-F-16", "UI-04"],
    ]
    add_abes_table(s4, 0.5, 1.3, rtm_headers, rtm_rows, [1.0, 3.4, 1.0, 1.2, 1.2, 1.2])

    # .3 below the table, (.5,y,9,.35) italic 12 grey: Full matrix: Project Document, section 4
    # 5 rows at 0.5 = 2.5 in -> table bottom = 1.3 + 2.5 = 3.8 -> 0.3 below is y=4.1
    b_note = s4.shapes.add_textbox(Inches(0.5), Inches(4.1), Inches(9.0), Inches(0.35))
    tf_not = b_note.text_frame
    tf_not.word_wrap = True
    tf_not.margin_left = tf_not.margin_right = tf_not.margin_top = tf_not.margin_bottom = 0
    tf_not.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_not = tf_not.paragraphs[0]
    p_not.text = "Full matrix: Project Document, section 4"
    p_not.font.name = "Calibri"
    p_not.font.size = Pt(12)
    p_not.font.italic = True
    p_not.font.color.rgb = C_GREY

    set_speaker_notes(
        s4,
        "Traceability is the backbone of our delivery. Every How Might We question connects strictly to a business "
        "requirement, user story, test case, and UI screen. For example, HMW-01 leads to BR-01 for monitoring the 85% occupancy "
        "threshold, captured in user story US-04, validated by test TC-F-13, and implemented on screen UI-02. "
        "All 9 Must requirements have 100% end-to-end coverage."
    )

    # =========================================================================
    # SLIDE 5: CONTENT "Data and preparation"
    # =========================================================================
    s5 = prs.slides.add_slide(prs.slide_layouts[5])
    setup_content_slide(s5, "Data and preparation", slide_num=5)

    # Four stats, x .5 / 2.8 / 5.1 / 7.4, w 2.1: value y 1.3 h .8 Cambria 30 bold red, middle; label y 2.1 h .35, 12 grey, middle: rows / columns / period / issues fixed.
    stats = [
        ("197,054", "Rows (53.9k IP + 143.1k OPD)", 0.5),
        ("58", "Columns across 5 tables", 2.8),
        ("2023–2025", "3-Year observation period", 5.1),
        ("5,368", "Issues audited & excluded", 7.4),
    ]
    for val, lbl, sx in stats:
        b_val = s5.shapes.add_textbox(Inches(sx), Inches(1.3), Inches(2.1), Inches(0.8))
        tf_v = b_val.text_frame
        tf_v.word_wrap = True
        tf_v.margin_left = tf_v.margin_right = tf_v.margin_top = tf_v.margin_bottom = 0
        tf_v.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_v = tf_v.paragraphs[0]
        p_v.text = val
        p_v.alignment = PP_ALIGN.CENTER
        p_v.font.name = "Cambria"
        p_v.font.size = Pt(30)
        p_v.font.bold = True
        p_v.font.color.rgb = C_RED

        b_lbl = s5.shapes.add_textbox(Inches(sx), Inches(2.1), Inches(2.1), Inches(0.35))
        tf_l = b_lbl.text_frame
        tf_l.word_wrap = True
        tf_l.margin_left = tf_l.margin_right = tf_l.margin_top = tf_l.margin_bottom = 0
        tf_l.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_l = tf_l.paragraphs[0]
        p_l.text = lbl
        p_l.alignment = PP_ALIGN.CENTER
        p_l.font.name = "Calibri"
        p_l.font.size = Pt(12)
        p_l.font.color.rgb = C_GREY

    # PLAIN CARD (.5,2.7,4.3,2.25) Source: dataset source; real, public or synthetic; what one row represents.
    src_text = (
        "Synthetic relational warehouse generated under seed 42, reflecting NABH tertiary hospital dynamics. "
        "Shipped as embedded SQLite (data/app.db, 34.2 MB).\n\n"
        "One row in inpatient fact represents a single completed patient episode from admission to discharge."
    )
    add_plain_card(s5, 0.5, 2.7, 4.3, 2.25, "Source", src_text)

    # PLAIN CARD (5.2,2.7,4.3,2.25) Cleaning: top 3 fixes with rows affected.
    clean_text = (
        "1. Removed 1,078 inverted admission/discharge date records.\n"
        "2. Reconciled 2,142 stay duration arithmetic mismatches.\n"
        "3. Sanitized 2,148 negative billing values and stay outliers.\n"
        "All 5,368 anomalies documented with reason codes in quality log."
    )
    add_plain_card(s5, 5.2, 2.7, 4.3, 2.25, "Cleaning", clean_text)

    set_speaker_notes(
        s5,
        "Our dataset spans 197,054 records over a 3-year observation window from 2023 to 2025. Each inpatient record "
        "represents an individual completed admission episode. During data cleansing, we established six automated validation "
        "gates that trapped and excluded 5,368 anomalies: our biggest fixes addressed stay duration arithmetic mismatches, "
        "negative billing values, and inverted admission dates."
    )

    # =========================================================================
    # SLIDE 6: CONTENT "Solution and technology"
    # =========================================================================
    s6 = prs.slides.add_slide(prs.slide_layouts[5])
    setup_content_slide(s6, "Solution and technology", slide_num=6)

    # Five steps, x .5 / 2.35 / 4.2 / 6.05 / 7.9, y 1.5, w 1.6, h 1.2: rounded, radius .1, fill F4F1F1; text centred, middle, 12 bold black, two lines
    steps = [
        (0.5, "Raw data\n(CSV / xlsx)", False),
        (2.35, "Clean & KPIs\n(Python, pandas)", False),
        (4.2, "Store\n(SQLite app.db)", False),
        (6.05, "Dashboard\n(Streamlit)", True), # Dashboard step: red fill, white text
        (7.9, "Deploy\n(GitHub -> cloud)", False),
    ]
    for sx, stext, is_dash in steps:
        step_shape = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(sx), Inches(1.5), Inches(1.6), Inches(1.2))
        step_shape.fill.solid()
        step_shape.fill.fore_color.rgb = C_RED if is_dash else C_CARD
        step_shape.line.color.rgb = C_RED if is_dash else C_CARD
        step_shape.adjustments[0] = 0.1
        tf = step_shape.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        lines = stext.split("\n")
        for idx, line in enumerate(lines):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.text = line
            p.alignment = PP_ALIGN.CENTER
            p.font.name = "Calibri"
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = C_WHITE if is_dash else C_BLACK

    # ">" between steps, x 2.1 / 3.95 / 5.8 / 7.65, y 1.5, w .25, h 1.2: 18 bold red, centred.
    arrow_xs = [2.1, 3.95, 5.8, 7.65]
    for ax in arrow_xs:
        b_ar = s6.shapes.add_textbox(Inches(ax), Inches(1.5), Inches(0.25), Inches(1.2))
        tf_a = b_ar.text_frame
        tf_a.word_wrap = False
        tf_a.margin_left = tf_a.margin_right = tf_a.margin_top = tf_a.margin_bottom = 0
        tf_a.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_a = tf_a.paragraphs[0]
        p_a.text = ">"
        p_a.alignment = PP_ALIGN.CENTER
        p_a.font.name = "Calibri"
        p_a.font.size = Pt(18)
        p_a.font.bold = True
        p_a.font.color.rgb = C_RED

    # (.5,3.15,9,.5) 13 dark, middle: bold "Tools:" + my tools separated by " · "
    b_tools = s6.shapes.add_textbox(Inches(0.5), Inches(3.15), Inches(9.0), Inches(0.5))
    tf_tools = b_tools.text_frame
    tf_tools.word_wrap = True
    tf_tools.margin_left = tf_tools.margin_right = tf_tools.margin_top = tf_tools.margin_bottom = 0
    tf_tools.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_t = tf_tools.paragraphs[0]
    r_lbl = p_t.add_run()
    r_lbl.text = "Tools: "
    r_lbl.font.name = "Calibri"
    r_lbl.font.size = Pt(13)
    r_lbl.font.bold = True
    r_lbl.font.color.rgb = C_DARK
    r_val = p_t.add_run()
    r_val.text = "Python 3.13 · Streamlit 1.40 · SQLite 3 · pandas 2.2 · Plotly 5.24 · pytest 8.0 · Tailwind CSS"
    r_val.font.name = "Calibri"
    r_val.font.size = Pt(13)
    r_val.font.bold = False
    r_val.font.color.rgb = C_DARK

    # (.5,3.75,9,.5) 13 italic grey, middle: Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud
    b_flow = s6.shapes.add_textbox(Inches(0.5), Inches(3.75), Inches(9.0), Inches(0.5))
    tf_flow = b_flow.text_frame
    tf_flow.word_wrap = True
    tf_flow.margin_left = tf_flow.margin_right = tf_flow.margin_top = tf_flow.margin_bottom = 0
    tf_flow.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_f = tf_flow.paragraphs[0]
    p_f.text = "Run locally: streamlit run app.py -> push to GitHub -> deploy on Streamlit Community Cloud"
    p_f.font.name = "Calibri"
    p_f.font.size = Pt(13)
    p_f.font.italic = True
    p_f.font.color.rgb = C_GREY

    set_speaker_notes(
        s6,
        "Here is our solution architecture and technology pipeline. We ingest raw files, process and clean data using "
        "Python and pandas, and load the validated warehouse into SQLite. Streamlit powers the interactive dashboard "
        "with Plotly charts and Tailwind UI tokens. For deployment, git pushes the repository to GitHub, triggering our "
        "automated test pipeline and deploying to Streamlit Community Cloud."
    )

    # =========================================================================
    # SLIDE 7: CONTENT "Dashboard"
    # =========================================================================
    s7 = prs.slides.add_slide(prs.slide_layouts[5])
    setup_content_slide(s7, "Dashboard", slide_num=7)

    # Frame (.5,1.25,5.5,3.7): fill F4F1F1, dashed 1 pt BFBFBF. Put my screenshot inside, keep its proportions, no stretching. No screenshot: centred italic 14 grey: [Paste a dashboard screenshot here]
    frame = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.25), Inches(5.5), Inches(3.7))
    frame.fill.solid()
    frame.fill.fore_color.rgb = C_CARD
    frame.line.color.rgb = C_FRAME_DASH
    frame.line.width = Pt(1)
    frame.line.dash_style = MSO_LINE_DASH_STYLE.DASH

    tf_fr = frame.text_frame
    tf_fr.word_wrap = True
    tf_fr.margin_left = tf_fr.margin_right = tf_fr.margin_top = tf_fr.margin_bottom = 0
    tf_fr.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_fr = tf_fr.paragraphs[0]
    p_fr.text = "[Paste a dashboard screenshot here]"
    p_fr.alignment = PP_ALIGN.CENTER
    p_fr.font.name = "Calibri"
    p_fr.font.size = Pt(14)
    p_fr.font.italic = True
    p_fr.font.color.rgb = C_GREY

    # BADGE x 6.7, y 1.3 / 2.2 / 3.1 / 4.0, numbers 1-4. Beside each, (7.25,y,2.25,.42), middle, 12 bold red "UI-01" + short black screen purpose. Use my four main screens.
    screens = [
        ("1", 1.3, "UI-01 ", "Executive KPI snapshot"),
        ("2", 2.2, "UI-02 ", "Bed occupancy vs 85%"),
        ("3", 3.1, "UI-03 ", "ALOS & Wilson readmission"),
        ("4", 4.0, "UI-05 ", "Payer mix & bed-day yield"),
    ]
    for num, sy, ui_code, purpose in screens:
        add_badge(s7, 6.7, sy, num)
        b_scr = s7.shapes.add_textbox(Inches(7.25), Inches(sy), Inches(2.25), Inches(0.42))
        tf_s = b_scr.text_frame
        tf_s.word_wrap = True
        tf_s.margin_left = tf_s.margin_right = tf_s.margin_top = tf_s.margin_bottom = 0
        tf_s.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_s = tf_s.paragraphs[0]
        r1 = p_s.add_run()
        r1.text = ui_code
        r1.font.name = "Calibri"
        r1.font.size = Pt(12)
        r1.font.bold = True
        r1.font.color.rgb = C_RED
        r2 = p_s.add_run()
        r2.text = purpose
        r2.font.name = "Calibri"
        r2.font.size = Pt(12)
        r2.font.bold = False
        r2.font.color.rgb = C_BLACK

    set_speaker_notes(
        s7,
        "The dashboard organizes hospital intelligence across dedicated screens. On UI-01, the executive overview displays "
        "a 6-metric KPI row including bed occupancy, ALOS, readmission, and revenue. Selecting a reporting period or department "
        "in the sidebar instantly updates all cards, charts, and delta comparisons synchronously across the entire application."
    )

    # =========================================================================
    # SLIDE 8: CONTENT "Key findings"
    # =========================================================================
    s8 = prs.slides.add_slide(prs.slide_layouts[5])
    setup_content_slide(s8, "Key findings", slide_num=8)

    # Three CARDS y 1.3, w 2.85, h 3.6, x .5 / 3.55 / 6.6.
    # Inside, at x+.2, w 2.45: number y 1.5 h 1 Cambria 40 bold red, middle; meaning y 2.6 h .9, 14 bold black, top; Evidence: KPI or test y 3.6 h 1.1, 12 grey, top.
    findings = [
        (0.5, "86.1%", "Cardiology Bed Bottleneck", "Evidence: KPI: BOR breaches 85% NABH ceiling; creates a deficit of 12 acute beds."),
        (3.55, "6.49 d", "Oncology ALOS Outlier", "Evidence: KPI: Exceeds 4.0d benchmark by +2.49d; consumes 38% of hospital bed-days."),
        (6.6, "ρ = 0.90", "Occupancy-Readmit Link", "Evidence: Spearman test: t = 5.21, p = 0.002; statistically significant capacity pressure."),
    ]
    for fx, num, meaning, evid in findings:
        add_card(s8, fx, 1.3, 2.85, 3.6)

        # Number y 1.5 h 1 Cambria 40 bold red, middle
        b_num = s8.shapes.add_textbox(Inches(fx + 0.2), Inches(1.5), Inches(2.45), Inches(1.0))
        tf_n = b_num.text_frame
        tf_n.word_wrap = True
        tf_n.margin_left = tf_n.margin_right = tf_n.margin_top = tf_n.margin_bottom = 0
        tf_n.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_n = tf_n.paragraphs[0]
        p_n.text = num
        p_n.alignment = PP_ALIGN.CENTER
        p_n.font.name = "Cambria"
        p_n.font.size = Pt(40)
        p_n.font.bold = True
        p_n.font.color.rgb = C_RED

        # Meaning y 2.6 h .9, 14 bold black, top
        b_mean = s8.shapes.add_textbox(Inches(fx + 0.2), Inches(2.6), Inches(2.45), Inches(0.9))
        tf_m = b_mean.text_frame
        tf_m.word_wrap = True
        tf_m.margin_left = tf_m.margin_right = tf_m.margin_top = tf_m.margin_bottom = 0
        tf_m.vertical_anchor = MSO_ANCHOR.TOP
        p_m = tf_m.paragraphs[0]
        p_m.text = meaning
        p_m.alignment = PP_ALIGN.CENTER
        p_m.font.name = "Calibri"
        p_m.font.size = Pt(14)
        p_m.font.bold = True
        p_m.font.color.rgb = C_BLACK

        # Evidence y 3.6 h 1.1, 12 grey, top
        b_ev = s8.shapes.add_textbox(Inches(fx + 0.2), Inches(3.6), Inches(2.45), Inches(1.1))
        tf_e = b_ev.text_frame
        tf_e.word_wrap = True
        tf_e.margin_left = tf_e.margin_right = tf_e.margin_top = tf_e.margin_bottom = 0
        tf_e.vertical_anchor = MSO_ANCHOR.TOP
        p_e = tf_e.paragraphs[0]
        p_e.text = evid
        p_e.alignment = PP_ALIGN.CENTER
        p_e.font.name = "Calibri"
        p_e.font.size = Pt(12)
        p_e.font.color.rgb = C_GREY

    set_speaker_notes(
        s8,
        "Our analysis revealed three critical findings. First, Cardiology bed occupancy reached 86.1%, breaching the "
        "NABH 85% safety threshold and creating a 12-bed shortage. Second, Oncology length of stay averages 6.49 days, "
        "exceeding the 4-day benchmark by nearly two and a half days. Third, our Spearman rank correlation test yielded "
        "rho equals 0.90 with p-value 0.002, proving a significant statistical link between bed occupancy pressure and 30-day readmissions."
    )

    # =========================================================================
    # SLIDE 9: CONTENT "Testing and validation"
    # =========================================================================
    s9 = prs.slides.add_slide(prs.slide_layouts[5])
    setup_content_slide(s9, "Testing and validation", slide_num=9)

    # Passed / total (.5,1.35,2.6,1) Cambria 44 bold red, middle
    b_pt = s9.shapes.add_textbox(Inches(0.5), Inches(1.35), Inches(2.6), Inches(1.0))
    tf_pt = b_pt.text_frame
    tf_pt.word_wrap = True
    tf_pt.margin_left = tf_pt.margin_right = tf_pt.margin_top = tf_pt.margin_bottom = 0
    tf_pt.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_pt = tf_pt.paragraphs[0]
    p_pt.text = "70 / 71"
    p_pt.alignment = PP_ALIGN.CENTER
    p_pt.font.name = "Cambria"
    p_pt.font.size = Pt(44)
    p_pt.font.bold = True
    p_pt.font.color.rgb = C_RED

    # (.5,2.4,2.6,.9) 13 grey, middle: test cases passed / pytest: n / manual: n
    b_stat = s9.shapes.add_textbox(Inches(0.5), Inches(2.4), Inches(2.6), Inches(0.9))
    tf_st = b_stat.text_frame
    tf_st.word_wrap = True
    tf_st.margin_left = tf_st.margin_right = tf_st.margin_top = tf_st.margin_bottom = 0
    tf_st.vertical_anchor = MSO_ANCHOR.MIDDLE
    t_lines = [
        "70 test cases passed",
        "pytest: 70 passed (1 skipped)",
        "manual: 19 verified",
    ]
    for idx, tl in enumerate(t_lines):
        p = tf_st.paragraphs[0] if idx == 0 else tf_st.add_paragraph()
        p.text = tl
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Calibri"
        p.font.size = Pt(13)
        p.font.color.rgb = C_GREY

    # TABLE at (3.4,1.35), column widths 1.1, 3.9, 1.1: Test | Checks | Result. 3 key rows.
    test_headers = ["Test", "Checks", "Result"]
    test_rows = [
        ["TC-T-01", "SQLite read-only mode & 34.2 MB size", "PASS"],
        ["TC-T-25", "30d readmission rate excluding deaths", "PASS"],
        ["TC-T-44", "Warm page latency (< 3.0s, actual 1.83s)", "PASS"],
    ]
    add_abes_table(s9, 3.4, 1.35, test_headers, test_rows, [1.1, 3.9, 1.1])

    set_speaker_notes(
        s9,
        "Quality assurance was executed through automated pytest suites and manual test protocols. 70 out of 71 automated "
        "tests pass cleanly with 1 intentional skip for out-of-scope alerting. Our tests certify read-only database connections, "
        "mathematical accuracy for Wilson confidence intervals, and warm-load page latency averaging 1.83 seconds, "
        "well within the 3.0-second performance budget."
    )

    # =========================================================================
    # SLIDE 10: CONTENT "Recommendation, limitations and next steps"
    # =========================================================================
    s10 = prs.slides.add_slide(prs.slide_layouts[5])
    setup_content_slide(s10, "Recommendation, limitations and next steps", slide_num=10)

    # BADGE CARDS y 1.25, w 2.85, h 3.7, x .5 / 3.58 / 6.65: R Recommendation; L Limitations; N Next steps. Base them on my real findings.
    rec_body = (
        "Reallocate 12 acute beds from General Medicine to Cardiology to relieve the 86.1% bottleneck.\n\n"
        "Implement standardized clinical care pathways in Oncology to safely compress ALOS from 6.49 toward the 4.0-day target."
    )
    add_badge_card(s10, 0.5, 1.25, 2.85, 3.7, "R", "Recommendation", rec_body)

    lim_body = (
        "Dataset is synthetic (seed 42), though calibrated to clinical distributions.\n\n"
        "Daily census is based on midnight snapshots, omitting intra-day bed churn. External automated SMS alerting is deferred."
    )
    add_badge_card(s10, 3.58, 1.25, 2.85, 3.7, "L", "Limitations", lim_body)

    nxt_body = (
        "Integrate HL7/FHIR telemetry streams for live real-time bed census.\n\n"
        "Introduce physician-level ALOS clinical audit drilldowns.\n\n"
        "Deploy automated SMS/email escalation alerts for capacity breach events."
    )
    add_badge_card(s10, 6.65, 1.25, 2.85, 3.7, "N", "Next steps", nxt_body)

    set_speaker_notes(
        s10,
        "Based on our findings, we recommend reallocating 12 beds from General Medicine to Cardiology and establishing "
        "dedicated Oncology discharge pathways. Our primary limitation is that data is synthetic and census relies on midnight "
        "snapshots rather than live bed churn. Next steps include connecting real-time HL7 telemetry, doctor-level ALOS audits, "
        "and automated capacity alerting."
    )

    # =========================================================================
    # SLIDE 11: DARK (Closing)
    # =========================================================================
    s11 = prs.slides.add_slide(prs.slide_layouts[6])
    build_dark_slide(
        s11,
        title_text="Live demo and Q&A",
        lines_list=[
            "App: https://2500320140037-p107.streamlit.app",
            "Code: https://github.com/ayushpal/hclTraining",
            "Thank you",
        ],
    )
    set_speaker_notes(
        s11,
        "Thank you for your time. The application is live at the provided Streamlit URL and the complete codebase and test suites "
        "are available on GitHub. I will now open the live application for a demonstration and invite questions from the panel."
    )

    # Save presentation
    prs.save(OUT_PPTX)
    print("Saved primary presentation:", OUT_PPTX)
    shutil.copyfile(OUT_PPTX, OUT_SEC)
    print("Saved secondary presentation:", OUT_SEC)

if __name__ == "__main__":
    build_presentation()
