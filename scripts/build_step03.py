"""Build Step 3 deliverable: Needs Analysis / Business Requirements for p107."""
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x54, 0x96)
GREY = RGBColor(0x59, 0x59, 0x59)
HDR_FILL = "1F3864"
ALT_FILL = "EAF0F8"
MUST_FILL = "DDEBF7"
WONT_FILL = "F2F2F2"

PROJECT_ID = "p107"
PROJECT_TITLE = "Healthcare KPI Dashboard"
STUDENT = "Ayush Pal"
ROLL = "2500320140037"
MENTOR = "Mr. Dinesh Sereno"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "03_Needs_Analysis.docx")


def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def add_table(doc, headers, rows, widths, font_size=8.5, row_fill=None):
    total = float(sum(widths))
    widths = [w * 6.6 / total for w in widths]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.width = Inches(widths[i])
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(cell, HDR_FILL)
    cant_split(hdr)
    repeat_header(hdr)
    for r_idx, vals in enumerate(rows):
        row = table.add_row()
        cant_split(row)
        for i, v in enumerate(vals):
            cell = row.cells[i]
            cell.width = Inches(widths[i])
            cell.text = ""
            run = cell.paragraphs[0].add_run(str(v))
            run.font.size = Pt(font_size)
            if i == 0:
                run.bold = True
            if row_fill and r_idx in row_fill:
                shade(cell, row_fill[r_idx])
            elif r_idx % 2 == 1:
                shade(cell, ALT_FILL)
    doc.add_paragraph()
    return table


def para(doc, text, size=10.5, italic=False, bold=False, color=None,
         space_after=6, align=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.italic = italic
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color
    p.paragraph_format.space_after = Pt(space_after)
    if align:
        p.alignment = align
    return p


def bullet(doc, text, level=0):
    p = doc.add_paragraph(text, style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.25 + 0.25 * level)
    p.paragraph_format.space_after = Pt(3)
    for r in p.runs:
        r.font.size = Pt(10.5)
    return p


def h1(doc, text, page_break=False):
    if page_break:
        doc.add_page_break()
    h = doc.add_heading(text, level=1)
    for r in h.runs:
        r.font.color.rgb = NAVY
        r.font.size = Pt(15)


def h2(doc, text):
    h = doc.add_heading(text, level=2)
    for r in h.runs:
        r.font.color.rgb = BLUE
        r.font.size = Pt(12)


def footer(section):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{PROJECT_ID}  |  Step 3 — Needs Analysis  |  Page ")
    run.font.size = Pt(8)
    run.font.color.rgb = GREY
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    p._p.append(fld)


doc = Document()
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.12

sec = doc.sections[0]
sec.top_margin = Inches(0.75)
sec.bottom_margin = Inches(0.7)
sec.left_margin = Inches(0.8)
sec.right_margin = Inches(0.8)
footer(sec)

# ---------- cover ----------
para(doc, "ABES ENGINEERING COLLEGE", 11, bold=True, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para(doc, "Department of Computer Applications (MCA) — Data Analytics", 10, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=40)
para(doc, "STEP 3", 26, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "NEEDS ANALYSIS — BUSINESS REQUIREMENTS", 19, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
para(doc, PROJECT_TITLE, 17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Project ID: {PROJECT_ID}   |   Domain: Data Analytics", 11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=36)

meta = [("Student", STUDENT), ("Roll Number", ROLL), ("Mentor", MENTOR),
        ("Deliverable", "03_Needs_Analysis.docx (Step 3 of 18)"),
        ("Version / Date", "v1.0 — 19 November 2026"),
        ("Inputs", "01_Project_Understanding.docx, 02_HMW_Statements.docx"),
        ("Status", "Baseline — approved for Step 4 (RTM)")]
t = doc.add_table(rows=0, cols=2)
t.style = "Table Grid"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
for k, v in meta:
    row = t.add_row()
    row.cells[0].width = Inches(1.9)
    row.cells[1].width = Inches(4.3)
    row.cells[0].text = ""
    r0 = row.cells[0].paragraphs[0].add_run(k)
    r0.bold = True
    r0.font.size = Pt(10)
    shade(row.cells[0], "EAF0F8")
    row.cells[1].text = ""
    r1 = row.cells[1].paragraphs[0].add_run(v)
    r1.font.size = Pt(10)

para(doc, "", space_after=20)
para(doc, "This document states WHAT the business needs, never HOW it will be built. No technology, "
          "product or platform is named. All data referenced is synthetic (Step 1 §9, AS-07).",
     8.5, italic=True, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)

# ---------- control ----------
doc.add_page_break()
h1(doc, "Document Control")
h2(doc, "Revision history")
add_table(doc, ["Version", "Date", "Author", "Change description", "Status"],
          [["0.1", "19-Nov-2026", STUDENT, "First pass: one requirement per HMW", "Draft"],
           ["0.2", "19-Nov-2026", STUDENT, "Split HMW-01/04/06 needs; added Won't item; DQ requirement", "Draft"],
           ["1.0", "19-Nov-2026", STUDENT, "15 BRs finalised, 60% Must, gap analysis clean, OQ-12 closed", "Baseline"]],
          [0.8, 1.0, 1.1, 3.1, 0.8])

h2(doc, "Scope of this step")
para(doc, "Step 1 established the questions; Step 2 established the human problems. Step 3 establishes "
          "what the business must be able to do for those problems to be solved. The output is a set of "
          "15 business requirements, each traceable to exactly one source HMW (two HMWs are served by "
          "more than one BR, never by zero), each with a measurable acceptance test, and each mapped to "
          "the data it needs.")

h2(doc, "Rules applied")
add_table(doc, ["Rule", "How it was enforced"],
          [["N1 — WHAT, not HOW", "Banned vocabulary in requirement text: Streamlit, Python, SQL, Plotly, Tailwind, database, dashboard, chart, API, cloud, ML. Checked mechanically before saving."],
           ["N2 — Business voice", "Every requirement follows: \u201cThe business needs to <capability> so that <benefit>.\u201d"],
           ["N3 — Source HMW mandatory", "A requirement with no HMW is a wish; it was rejected or traced to an existing HMW."],
           ["N4 — Testable", "Each BR carries at least one acceptance measure with a number, a population and a method."],
           ["N5 — Atomic", "One capability per BR; \u201cand\u201d joining two decisions forces a split."],
           ["N6 — MoSCoW honest", "Must is reserved for requirements whose absence breaks a priority HMW or an SC; target 60% Must."]],
          [1.55, 5.05], font_size=8.5)

# ---------- 1 ----------
h1(doc, "1. From HMW to Need to Business Requirement", page_break=True)
para(doc, "For each HMW the analysis asked three questions: what must be true for this user to act? "
          "what information must exist at that moment? what must the business be able to do repeatedly, "
          "not just once? The answer to the third question becomes the BR.", space_after=4)
add_table(doc, ["HMW", "Need extracted (in plain business language)", "BR(s) satisfying it"],
          [["HMW-01", "Advance warning of a target miss, with the shortfall expressed in beds so escalation is concrete.", "BR-01, BR-15 (Won't)"],
           ["HMW-02", "Ward-and-shift granularity of the same occupancy picture, available before the roster is frozen.", "BR-02"],
           ["HMW-03", "Decomposition of a movement into its contributors, so a cause can be named and not guessed.", "BR-03"],
           ["HMW-04", "A predictable path from question to answer that a non-technical user can complete unaided.", "BR-04, BR-05"],
           ["HMW-05", "A comparison that is normalised, so peer challenge focuses on performance and not on method.", "BR-06"],
           ["HMW-06", "Separation of clinical signal from data noise, plus attribution to department and diagnosis.", "BR-07, BR-14"],
           ["HMW-07", "Earlier visibility of payer-mix movement with revenue per patient attached.", "BR-08"],
           ["HMW-08", "Demand pattern by weekday and season, quantified as a ratio that can be planned against.", "BR-09"],
           ["HMW-09", "Evidence, with stated assumptions, about whether occupancy pressure correlates with quality outcomes.", "BR-10"],
           ["HMW-10", "Visibility of the population actually served, by age band and gender, per department.", "BR-11"],
           ["HMW-11", "A normalised revenue measure per unit of capacity, ranked with the size of the gap.", "BR-12"],
           ["HMW-12", "Early detection of slow deterioration in outcome indicators, not a post-audit reconstruction.", "BR-13, BR-15 (Won't)"]],
          [0.7, 4.35, 1.55], font_size=8.2)

# ---------- 2 ----------
h1(doc, "2. Business Requirements Register", page_break=True)
para(doc, "Priority key: Must = release 1 cannot ship without it; Should = important, workaround exists; "
          "Could = desirable if capacity allows; Won't = consciously excluded from this release, with a "
          "reason recorded. The Must rows are shaded.", space_after=4)

br_rows = [
    ["BR-01",
     "The business needs to see, for every department, its bed occupancy against the 85% target for the current and next week together with the shortfall expressed in beds and bed-days, so that the operations manager can escalate a likely miss with a number before the 8 a.m. huddle.",
     "SH-02", "HMW-01", "Must",
     "All 8 departments appear for any chosen week; gap in beds = f(target, projected occupancy) shown to 0 decimals and reproduces a manual calculation \u00b10.1 pp on 3 sampled months; 0 departments missing."],
    ["BR-02",
     "The business needs ward-level occupancy projected across the next seven days, so that the nursing superintendent can move staff before a ward crosses 90% instead of after handover.",
     "SH-06", "HMW-02", "Must",
     "All wards listed for a 7-day horizon; every ward forecast above 90% is flagged; verified on 10 randomly chosen dates against manual calculation."],
    ["BR-03",
     "The business needs any movement in a KPI between two periods to be decomposed into the contributions of department, payer category and weekday, so that the director can name the cause in the board meeting rather than ask an analyst.",
     "SH-01", "HMW-03", "Must",
     "For 3 sampled KPI-period pairs, the top three contributors sum to the total movement within \u00b10.5 pp; every decomposition states the period and unit."],
    ["BR-04",
     "The business needs any operational question to be answerable in no more than three interactions (choose period, choose unit, read the answer), so that a first-time user can self-serve within 60 seconds without knowing the data model.",
     "SH-02, SH-10", "HMW-04", "Must",
     "8 users, 5 tasks each: median \u2264 60 s, p90 \u2264 90 s, \u2264 3 interactions, at least 7 of 8 succeed unaided (links SC-03)."],
    ["BR-05",
     "The business needs every published figure to carry its definition, denominator and data period in the same view, so that two people never quote different numbers for the same week.",
     "SH-02, SH-07", "HMW-04", "Must",
     "Checklist audit of every measure shown: 100% have definition + denominator + period; 0 exceptions accepted; the same definition text is reused from Step 1."],
    ["BR-06",
     "The business needs length-of-stay comparison between departments to be presented on a like-for-like basis with the case-mix context stated, so that department heads accept the benchmark and commit to an action.",
     "SH-03", "HMW-05", "Must",
     "ALOS value reproduces the Step 1 definition to \u00b10.01 day for 2 sampled months; the comparison states its normalisation basis; 3 HODs rate fairness \u2265 4/5."],
    ["BR-07",
     "The business needs the 30-day unplanned readmission rate reported by department and diagnosis group with the data-quality context of that period visible, so that the quality officer can separate a clinical spike from a registration or coding error before reporting it.",
     "SH-04", "HMW-06", "Must",
     "Recomputed rate matches a manual calculation on 2 random months (\u00b10.1 pp); the 30-day window and exclusion rule (AS-04) are stated on screen; every highlighted spike carries at least one data-quality indicator."],
    ["BR-08",
     "The business needs payer-mix share, its movement between periods and revenue per patient attached to it, so that finance can act within the same quarter instead of after the negotiation window closes.",
     "SH-05", "HMW-07", "Should",
     "Payer shares sum to 100% \u00b10.1 for every period; movement shown for all 5 payer categories; revenue per discharge reconciles to 20 sampled billing records within \u00b10.5%."],
    ["BR-09",
     "The business needs outpatient demand summarised by weekday and by season with a peak-to-average ratio, so that rosters and consultant slots are planned before the queue forms.",
     "SH-06, SH-02", "HMW-08", "Must",
     "Peak weekday identified for each of 12 months; peak-to-average ratio reproduced by manual calculation within \u00b12%; both average and peak values shown with units."],
    ["BR-10",
     "The business needs a statistically defensible statement about whether occupancy pressure is associated with longer stays or higher readmissions, with method and assumptions disclosed, so that leadership stops asserting risk without evidence.",
     "SH-01, SH-04", "HMW-09", "Could",
     "Analysis reports coefficient, p-value, n and method name; assumptions checked and stated; result reproducible from the documented seed; the association-not-causation caveat appears in the same view."],
    ["BR-11",
     "The business needs admissions and outpatient volume broken down by age band and gender for each department, so that service-line plans reflect the population actually served.",
     "SH-03", "HMW-10", "Could",
     "Department-level shares sum to 100% \u00b10.1 across age bands and across genders; 3 departments spot-checked against manual counts."],
    ["BR-12",
     "The business needs revenue per bed-day ranked across departments with the size of the gap to the best performer, so that capacity and service-line decisions are evidence-based rather than historical.",
     "SH-05", "HMW-11", "Should",
     "Values for all 8 departments; ranking stable across two adjacent months unless a \u2265 5% change is documented; gap to best reproduced within \u00b11%."],
    ["BR-13",
     "The business needs a three-month moving view of mortality and discharge-type mix by department with a drift indicator, so that a quality review is triggered before the audit rather than during it.",
     "SH-04", "HMW-12", "Must",
     "Drift flagged when a point falls outside \u00b12 standard deviations of the 12-month baseline; the rule is validated on a deliberately injected synthetic drift; 0 unflagged injected drifts."],
    ["BR-14",
     "The business needs the data-quality status of each period (missing payer, duplicate registration, orphan department, date coverage, negative length of stay, occupancy above 100%) shown alongside the KPI, so that a suspect number is challenged before it is used in a decision.",
     "SH-04, SH-08", "HMW-06", "Should",
     "All 6 checks execute and return pass/fail with counts for any selected period; a period with an injected defect returns fail on the expected check; results reconcile with the data-quality test suite."],
    ["BR-15",
     "The business needs threshold breaches pushed automatically to managers without anyone opening the tool, so that a saturation or mortality-drift event is actioned within minutes.",
     "SH-02, SH-04", "HMW-01, HMW-12", "Won't",
     "Not measured this release (excluded by Step 1 X7 and CN-02). If revisited: alert within 5 minutes of breach, false-positive rate \u2264 5%."],
]
fills = {}
for i, r in enumerate(br_rows):
    if r[4] == "Must":
        fills[i] = MUST_FILL
    elif r[4] == "Won't":
        fills[i] = WONT_FILL
add_table(doc, ["BR ID", "Business requirement (WHAT, not HOW)", "Stake-holder", "Source HMW", "Priority", "How we will measure it (acceptance test)"],
          br_rows, [0.55, 2.85, 0.65, 0.65, 0.55, 1.95], font_size=7.4, row_fill=fills)

# ---------- 3 ----------
h1(doc, "3. Prioritisation Check", page_break=True)
h2(doc, "3.1 MoSCoW distribution")
add_table(doc, ["Priority", "Count", "Share", "IDs"],
          [["Must", "9", "60.0%", "BR-01, 02, 03, 04, 05, 06, 07, 09, 13"],
           ["Should", "3", "20.0%", "BR-08, 12, 14"],
           ["Could", "2", "13.3%", "BR-10, 11"],
           ["Won't", "1", "6.7%", "BR-15"],
           ["Total", "15", "100%", "\u2014"]],
          [0.9, 0.6, 0.6, 4.5], font_size=8.5)
para(doc, "The Must share is exactly 60%, inside the 55–65% band a healthy backlog should occupy. If "
          "every requirement were Must, priority would carry no information and Step 14 (build plan) "
          "would have no sequencing signal.", size=9.5, italic=True)

h2(doc, "3.2 Why each Must is a Must")
add_table(doc, ["BR", "Breaks which priority HMW / SC if removed", "Consequence of removal"],
          [["BR-01", "HMW-01 (rank 1), SC-01, SC-03", "The primary daily decision has no advance signal; the app becomes retrospective."],
           ["BR-02", "HMW-02, OBJ-01", "Rooms are staffed after saturation; the ward view loses its action value."],
           ["BR-03", "HMW-03, OBJ-01", "Leadership can see a movement but still cannot name the cause."],
           ["BR-04", "HMW-04 (rank 2), SC-03", "Adoption fails; every other requirement becomes academic."],
           ["BR-05", "HMW-04, BR-06, SH-07 pain", "Numbers are distrusted; the review meeting returns to reconciliation."],
           ["BR-06", "HMW-05 (rank 3), SC-01", "HODs reject peer comparison, so no corrective action is committed."],
           ["BR-07", "HMW-06 (rank 4), SC-01", "Quality reporting stays a compliance artefact without signal separation."],
           ["BR-09", "HMW-08 (rank 5), OBJ-01", "Roster planning remains reactive; peaks are explained after the fact."],
           ["BR-13", "HMW-12, OBJ-01 (mortality KPI)", "One of the six core KPIs would have no trend or drift view."]],
          [0.55, 2.45, 3.6], font_size=8.2)

h2(doc, "3.3 What is deliberately not required (and why)")
add_table(doc, ["Excluded need", "Why it is excluded", "Where the pain is otherwise addressed"],
          [["Automatic push alerts (BR-15)", "Platform constraint CN-02 and out-of-scope X7; alerting without a scheduler would be theatre.", "Daily 8 a.m. routine plus the default view (BR-01); recorded as a future enhancement."],
           ["Patient-facing wait-time visibility (SH-09 P1/P3)", "Out of scope X5/X7 — no public patient channel in this release.", "Accepted gap; noted in §5.3."],
           ["Predictive readmission scoring", "Out of scope X3; would change the project into ML engineering and dilute 18-step delivery.", "Descriptive readmission drivers only (BR-07)."],
           ["Role-based access and audit logging", "Out of scope X5; free-tier single deployment.", "A single shared passphrase is the maximum (Step 16)."]],
          [1.6, 2.8, 2.2], font_size=8.2)

# ---------- 4 ----------
h1(doc, "4. Data Each Requirement Needs", page_break=True)
para(doc, "Field names are still provisional — T-IDs are issued in Step 8. The concepts below are the "
          "same six concepts proven feasible in Step 1 §11.1, so no new data source is being invented "
          "at Step 3.", space_after=4)
add_table(doc, ["BR", "Data concepts needed", "Fields required", "Grain / period", "Feasible from Step 1 §11?"],
          [["BR-01", "Bed capacity, Inpatient episode", "department_id, week, bed_days_available, bed_days_occupied, beds_functional", "Department-week", "Yes (OQ-01)"],
           ["BR-02", "Bed capacity, Ward dimension", "ward_id, department_id, date, bed_days_available, bed_days_occupied", "Ward-day", "Yes (OQ-01, extended to ward)"],
           ["BR-03", "All four facts + dimensions", "kpi_value by period, department_id, payer_id, weekday", "KPI-month", "Yes (OQ-02)"],
           ["BR-04", "None beyond others", "No new field; governs presentation and interaction", "Any", "Yes (n/a — design requirement)"],
           ["BR-05", "Definition registry", "kpi_code, definition_text, numerator, denominator, version", "KPI", "Yes (new: registry, not data)"],
           ["BR-06", "Inpatient episode, Department dim", "department_id, admission_date, discharge_date, diagnosis_group", "Department-month", "Yes (OQ-03)"],
           ["BR-07", "Inpatient episode (self-join), Department, Payer", "patient_key, admission_date, discharge_date, department_id, diagnosis_group, readmit_flag, payer_id", "Department-month", "Yes (OQ-04)"],
           ["BR-08", "Finance, Payer dim, Inpatient episode", "payer_id, billed_amount, discharge_date, department_id", "Payer-month", "Yes (OQ-05, gross billed per AS-05)"],
           ["BR-09", "OPD visit", "visit_date, department_id, visit_count", "Day \u00d7 department", "Yes (OQ-06)"],
           ["BR-10", "Derived from BR-01/06/07 outputs", "occupancy_rate, alos, readmission_rate per department-month", "Department-month", "Yes (OQ-07)"],
           ["BR-11", "Inpatient episode, OPD visit", "age_band, gender, department_id, admission_date, visit_date", "Department-month", "Yes (OQ-08)"],
           ["BR-12", "Finance, Bed capacity", "billed_amount, bed_days_occupied, department_id", "Department-month", "Yes (OQ-09)"],
           ["BR-13", "Inpatient episode", "discharge_date, discharge_type, death_flag, department_id", "Department-month", "Yes (OQ-10)"],
           ["BR-14", "All facts", "null checks on payer_id/age_band, duplicate patient_key, orphan department_id, min/max date, LOS < 0, BOR > 100", "Period", "Yes (new: quality checks over existing data)"],
           ["BR-15", "Not built", "\u2014", "\u2014", "N/A (Won't)"]],
          [0.5, 1.3, 3.0, 1.0, 0.9], font_size=7.4)
para(doc, "One genuinely new artefact appears at Step 3: a KPI definition registry (BR-05). It is not a "
          "new data source — it is a one-row-per-KPI reference table that keeps a single definition in "
          "use across every requirement. Step 8 will decide whether it becomes a table or a configuration "
          "block; the identifier T-xx is not issued yet.", size=9.5, italic=True)

# ---------- 5 ----------
h1(doc, "5. Gap Analysis", page_break=True)
h2(doc, "5.1 HMW with no BR (upward gaps)")
add_table(doc, ["Check", "Result", "Evidence"],
          [["HMWs with at least one BR", "12 of 12", "HMW-01\u2192BR-01/15, 02\u2192BR-02, 03\u2192BR-03, 04\u2192BR-04/05, 05\u2192BR-06, 06\u2192BR-07/14, 07\u2192BR-08, 08\u2192BR-09, 09\u2192BR-10, 10\u2192BR-11, 11\u2192BR-12, 12\u2192BR-13/15"],
           ["HMWs with zero BR", "0", "None"],
           ["HMWs mapped only to a Won't BR", "1 (partly)", "HMW-01 and HMW-12 both also have a Must/Should BR, so no HMW depends solely on BR-15."]],
          [1.7, 0.9, 4.0], font_size=8.2)

h2(doc, "5.2 BR with no HMW (downward gaps)")
add_table(doc, ["Check", "Result", "Note"],
          [["BRs with exactly one source HMW", "14 of 15", "Normal case"],
           ["BRs with two source HMWs", "1 of 15", "BR-15 (HMW-01, HMW-12) — one alerting need serving two questions"],
           ["BRs with zero source HMW", "0", "Gate PASS — no orphan requirement"]],
          [1.9, 0.9, 3.8], font_size=8.2)

h2(doc, "5.3 Accepted gaps carried forward (with owner)")
add_table(doc, ["Gap ID", "Gap", "Why accepted", "Owner / when closed"],
          [["GAP-01", "SH-09 patient-facing wait and bed visibility has no BR.", "Public patient channel is out of scope (X5/X7); benefit is indirect.", "SH-09 via SH-02; revisit after Step 16"],
           ["GAP-02", "SH-07 asked for an enterprise data dictionary; BR-05 covers published KPI definitions, not every raw field.", "A full dictionary is outside a single-app scope; the KPI registry removes the reported disagreement.", "SH-07; Step 8 will document field definitions in the database design"],
           ["GAP-03", "BR-15 alerting is excluded, so nobody is notified without opening the tool.", "CN-02 / X7; default view plus a 8 a.m. routine is the accepted mitigation.", "SH-10 sign-off at Step 16"],
           ["GAP-04", "HMW-09's statistical claim depends on synthetic seasonality being realistic.", "AS-07 already mandates seasonality; Step 10 must deliver it or BR-10 becomes decorative.", "Step 10 validation checklist"]],
          [0.62, 2.0, 2.5, 1.45], font_size=8.2)

h2(doc, "5.4 OQ coverage re-checked through BRs (nothing lost from Step 2)")
add_table(doc, ["OQ", "Served by BR(s)", "Covered?"],
          [["OQ-01", "BR-01, BR-02, BR-15", "Yes"],
           ["OQ-02", "BR-03, BR-04", "Yes"],
           ["OQ-03", "BR-06", "Yes"],
           ["OQ-04", "BR-07, BR-14", "Yes"],
           ["OQ-05", "BR-08", "Yes"],
           ["OQ-06", "BR-09", "Yes"],
           ["OQ-07", "BR-10", "Yes"],
           ["OQ-08", "BR-11", "Yes"],
           ["OQ-09", "BR-12", "Yes"],
           ["OQ-10", "BR-13", "Yes"]],
          [0.9, 2.6, 1.0], font_size=8.5)

# ---------- 6 ----------
h1(doc, "6. Decisions Taken and Assumptions Confirmed in Step 3")
add_table(doc, ["Item", "Decision", "Effect on earlier documents"],
          [["OQ-12 (open since Step 1)", "Finance views report gross billed revenue per discharge; net collection is out of scope for release 1.", "Step 1 §12.2 closes this item; AS-05 upgraded from assumption to confirmed decision for BR-08 and BR-12."],
           ["OQ-11 (open)", "Occupancy will be shown monthly by default; a daily/ward heatmap is deferred to the UI step.", "Still open; carried to Step 12."],
           ["OQ-13 (open)", "Passphrase gate on the public URL remains undecided.", "Still open; carried to Step 16."],
           ["AS-04 readmission rule", "Confirmed as the operative definition for BR-07.", "No wording change; the 30-day unplanned exclusion now has a consumer requirement."],
           ["New AR-01", "Priority share 60% Must is accepted as the release-1 commitment.", "Steps 11, 13, 14 will assume Must items are all built and tested."]],
          [1.4, 3.2, 2.0], font_size=8.2)

# ---------- 7 ----------
h1(doc, "7. Step 3 Sign-off")
bullet(doc, "15 business requirements issued (BR-01 … BR-15) from 12 HMW statements; 0 orphan HMW, 0 orphan BR.")
bullet(doc, "Priority mix: 9 Must (60%), 3 Should, 2 Could, 1 Won't — with a written reason for every Must and for the exclusion.")
bullet(doc, "Every BR is testable: 15 of 15 carry a numeric acceptance measure, population and method.")
bullet(doc, "Data need mapped for 14 of 15 BRs; BR-15 is not built, so it needs no data.")
bullet(doc, "OQ coverage preserved: 10 of 10 OQs still served, now through BRs.")
para(doc, "Prepared by: " + STUDENT + "   |   Roll: " + ROLL, size=10)
para(doc, "Reviewed by (mentor): " + MENTOR + "   |   Date: 19-Nov-2026", size=10)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print("saved:", OUT)
