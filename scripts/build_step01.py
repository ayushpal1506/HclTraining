"""Build Step 1 deliverable: Project Understanding for HKD-2026."""
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x54, 0x96)
GREY = RGBColor(0x59, 0x59, 0x59)
HDR_FILL = "1F3864"
ALT_FILL = "EAF0F8"

PROJECT_ID = "HKD-2026"
PROJECT_TITLE = "Healthcare KPI Dashboard"
STUDENT = "Ayush Pal"
ROLL = "2503201450035"
MENTOR = "Aplana Lodhi"
OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "docs", "01_Project_Understanding.docx")


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


def keep_together(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def add_table(doc, headers, rows, widths, font_size=8.5, first_col_bold=False):
    # normalise column widths so the table never spills past the 6.67in text column
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
        p = cell.paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(cell, HDR_FILL)
    keep_together(hdr)
    repeat_header(hdr)
    for r_idx, row_vals in enumerate(rows):
        row = table.add_row()
        keep_together(row)
        for i, val in enumerate(row_vals):
            cell = row.cells[i]
            cell.width = Inches(widths[i])
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(font_size)
            if (first_col_bold and i == 0) or i == 0:
                run.bold = True
            if r_idx % 2 == 1:
                shade(cell, ALT_FILL)
    doc.add_paragraph()
    return table


def para(doc, text, size=10.5, italic=False, bold=False, color=None,
         space_after=6, align=None, style=None):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
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
    return h


def h2(doc, text):
    h = doc.add_heading(text, level=2)
    for r in h.runs:
        r.font.color.rgb = BLUE
        r.font.size = Pt(12)
    return h


def add_page_number_footer(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{PROJECT_ID}  |  Step 1 — Project Understanding  |  Page ")
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

# ---------- base styles ----------
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
add_page_number_footer(sec)

# ---------- cover ----------
para(doc, "ABES ENGINEERING COLLEGE", 11, bold=True, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para(doc, "Department of Computer Applications (MCA) — Data Analytics", 10,
     color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=40)

para(doc, "STEP 1", 26, bold=True, color=NAVY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "PROJECT UNDERSTANDING", 20, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
para(doc, PROJECT_TITLE, 17, bold=True,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Project ID: {PROJECT_ID}   |   Domain: Data Analytics", 11,
     color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=36)

meta = [
    ("Student", STUDENT),
    ("Roll Number", ROLL),
    ("Programme", "MCA — 2nd Year, Data Analytics"),
    ("Institution", "ABES Engineering College"),
    ("Mentor", MENTOR),
    ("Deliverable", "01_Project_Understanding.docx (Step 1 of 18)"),
    ("Version / Date", "v1.0 — 19 November 2026"),
    ("Status", "Baseline — approved for Step 2"),
]
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

para(doc, "", space_after=24)
para(doc, "Document confidentiality: academic project. All data used in this project is "
          "synthetic; no real patient or hospital records are processed.",
     8.5, italic=True, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)

# ---------- document control ----------
doc.add_page_break()
h1(doc, "Document Control")

h2(doc, "Revision history")
add_table(doc,
          ["Version", "Date", "Author", "Change description", "Status"],
          [["0.1", "19-Nov-2026", STUDENT, "First draft of project understanding", "Draft"],
           ["1.0", "19-Nov-2026", STUDENT, "Reviewed against mentor scope; OQ-to-data feasibility added", "Baseline"]],
          [0.8, 1.0, 1.1, 3.1, 0.8])

h2(doc, "How to read this document")
para(doc, "This is the constitutional document of the project: every later artefact — "
          "How-Might-We statements, business requirements, BRD, PRD, database design, user stories, "
          "test cases and the final project report — must trace back to an identifier defined here. "
          "Identifiers are frozen once issued. Section 11 explicitly proves that each operational "
          "question can actually be answered by the data the app will hold, which is the quality gate "
          "for leaving Step 1.")

h2(doc, "Identifier scheme used in this project")
add_table(doc,
          ["Prefix", "Meaning", "Introduced in"],
          [["SH", "Stakeholder", "Step 1"],
           ["OQ", "Operational question", "Step 1"],
           ["OBJ", "SMART objective", "Step 1"],
           ["SC", "Success criterion", "Step 1"],
           ["AS / CN", "Assumption / Constraint", "Step 1"],
           ["HMW", "How-Might-We statement", "Step 2"],
           ["BR", "Business requirement", "Step 3"],
           ["F / FR / NFR", "Feature / functional / non-functional requirement", "Step 3 onwards"],
           ["T", "Database table", "Step 8"],
           ["US / UI", "User story / application screen", "Steps 11–12"],
           ["TC-F / TC-T", "Functional / technical test case", "Step 13"]],
          [1.2, 3.6, 1.4])

# ---------- 1 ----------
h1(doc, "1. Project Title and Context", page_break=True)
para(doc, "Project title: ", bold=True, space_after=0)
para(doc, "Healthcare KPI Dashboard — a self-service analytics application for hospital "
          "operations and quality KPIs.", space_after=8)

para(doc, "Project identifier: ", bold=True, space_after=0)
para(doc, f"{PROJECT_ID}. The identifier is used on every artefact, every screenshot and "
          "every test record so that a reviewer can move from a chart in the app back to the "
          "requirement, the test case and the report section that justify it.", space_after=8)

para(doc, "One-paragraph context. ", bold=True, space_after=0)
para(doc, "A mid-sized multi-speciality hospital generates a large volume of operational data every "
          "day: admissions and discharges, bed occupancy, outpatient footfall, payer mix, revenue and "
          "clinical outcome events. In practice this data sits in a hospital information system (HIS), "
          "in nursing registers, in the billing system and in ad-hoc Excel files maintained by "
          "department assistants. Each source uses slightly different definitions — one register counts "
          "a transfer as a new admission, another does not; occupancy is computed against sanctioned "
          "beds in one report and against functional beds in another. The result is that leadership "
          "meetings spend more time reconciling numbers than deciding what to do about bed shortages, "
          "long stays or rising readmissions. This project builds a single, opinionated, well-documented "
          "dashboard that fixes the definitions once, shows the trend and the driver behind each KPI, "
          "and can be operated by a non-technical hospital manager without a data analyst in the room. "
          "Because real patient data cannot be used for an academic submission, the project generates a "
          "statistically realistic synthetic dataset with the same grain as the real sources, and "
          "validates the app against that dataset.", space_after=8)

# ---------- 2 ----------
h1(doc, "2. Domain and Business Background")
para(doc, "Healthcare operations management is the discipline of running a hospital's finite capacity — "
          "beds, operation theatres, Nurses, doctors, diagnostic slots — efficiently while meeting "
          "clinical quality and financial targets. Analytics in this domain is judged by three things: "
          "are the definitions unambiguous, is the comparison fair (same period, same unit, same "
          "denominator), and does the number lead to an action. The dashboard therefore reports KPIs in "
          "pairs — value and benchmark, value and trend, value and driver — rather than as isolated "
          "figures.", space_after=8)
para(doc, "The six KPIs that anchor this project are standard, externally recognisable hospital "
          "indicators:", space_after=4)
bullet(doc, "Bed occupancy rate (BOR) — bed-days occupied ÷ bed-days available, per department and month.")
bullet(doc, "Average length of stay (ALOS) — total inpatient days ÷ discharges, per department.")
bullet(doc, "30-day readmission rate — unplanned readmissions within 30 days ÷ index discharges.")
bullet(doc, "Outpatient (OPD) footfall — visits per day, weekday pattern and department mix.")
bullet(doc, "Revenue per patient and payer mix — billed revenue ÷ discharges, split by payer category.")
bullet(doc, "Clinical outcome indicator — in-hospital mortality rate and discharge-type mix.")
para(doc, "The domain also supplies the vocabulary the whole project will use: admission, discharge, "
          "transfer, bed-day, index admission, readmission window, payer category, sanctioned vs "
          "functional bed. Defining these in Step 1 prevents rework in Steps 3–8 and prevents the "
          "classic viva question \u201cwhy does your occupancy cross 100%?\u201d — the answer must be a "
          "definition, not an apology.", space_after=6)

# ---------- 3 ----------
h1(doc, "3. Problem Statement")
para(doc, "Hospital leadership has no single, trustworthy, same-day view of operational KPIs — bed "
          "occupancy, length of stay, readmissions, OPD load and revenue per patient are scattered "
          "across registers, spreadsheets and HIS exports with inconsistent definitions — so capacity, "
          "staffing and quality decisions are taken on stale or contradictory numbers.", italic=True,
     bold=True, space_after=8)
para(doc, "Problem decomposition (what makes it a real problem, not just an inconvenience):", space_after=4)
add_table(doc,
          ["Dimension", "Current pain (as observed in a typical hospital)", "Consequence"],
          [["Data access", "Numbers arrive by WhatsApp/Excel after the morning meeting; month-end consolidation takes 2–4 days.",
            "Decisions are taken on last month's reality."],
           ["Definition conflict", "Occupancy computed against sanctioned beds in one report and functional beds in another; transfers counted twice.",
            "Two departments quote different numbers for the same week; trust erodes."],
           ["Effort", "One analyst spends ~45 minutes per ad-hoc question rebuilding a pivot table.",
            "Only 3–5 questions get answered per week; the rest are guessed."],
           ["Driver blindness", "Reports show the KPI value but not which department, payer or weekday caused the change.",
            "Corrective action is generic (\u201cimprove occupancy\u201d) rather than targeted."],
           ["Quality visibility", "Readmission and mortality figures are reviewed only at NABH audit time.",
            "Clinical quality becomes a compliance ritual, not a management lever."]],
          [1.25, 3.55, 1.9])

# ---------- 4 ----------
h1(doc, "4. Intended Audience / Users")
para(doc, "The app is designed for a small set of named roles, each with a different question, a "
          "different time budget and a different tolerance for detail. Usage frequency below is the "
          "assumption used to size the UI and the default landing view.", space_after=4)
add_table(doc,
          ["User ID", "User role", "What they open the app for", "How often", "Session length", "Primary screens"],
          [["U-01", "Operations Manager (primary persona)",
            "Morning capacity huddle: occupancy, admissions, discharges, OPD load for yesterday and the week",
            "Daily (2–3 times)", "10–15 min", "UI-01, UI-02, UI-03"],
           ["U-02", "Department Head (HOD)",
            "Monthly departmental review: ALOS, readmissions, revenue per bed versus peers",
            "Weekly / monthly", "15–20 min", "UI-02, UI-03, UI-05"],
           ["U-03", "Quality / NABH Accreditation Officer",
            "Audit preparation: mortality, readmission, discharge-type trend and evidence export",
            "Monthly + pre-audit", "20–30 min", "UI-04, UI-05"],
           ["U-04", "Finance / Revenue Cycle Manager",
            "Payer mix, revenue per patient, department contribution, month-on-month movement",
            "Monthly", "15 min", "UI-05"],
           ["U-05", "Hospital Director / Medical Superintendent",
            "Executive summary: six KPIs against target with one-line explanation of variance",
            "Weekly / monthly", "3–5 min", "UI-01"],
           ["U-06", "Nursing Superintendent",
            "Ward-level occupancy and discharge load to plan shift rosters",
            "Weekly", "10 min", "UI-02"],
           ["U-07", "Data analyst / faculty evaluator",
            "Deep dive, data quality checks, statistical tests, reproducibility",
            "Ad-hoc", "20–45 min", "UI-03, UI-06, UI-07"]],
          [0.62, 1.45, 2.05, 0.95, 0.78, 0.95], font_size=8)

# ---------- 5 ----------
h1(doc, "5. Stakeholders", page_break=True)
para(doc, "Influence (H/M/L) = ability to change the project's scope, definitions or acceptance. "
          "Interest (H/M/L) = how much the outcome affects their daily work. Where influence is high and "
          "interest is low (e.g. IT governance), the strategy is to keep them informed with short, "
          "decision-ready artefacts rather than long reviews.", space_after=4)
add_table(doc,
          ["SH ID", "Stakeholder", "Role", "Interest", "Influence", "Decision they make with the app"],
          [["SH-01", "Hospital Director / Medical Superintendent", "Executive sponsor", "M", "H",
            "Approve capacity expansion, new bed blocks or department budget changes based on KPI trend."],
           ["SH-02", "Operations Manager", "Primary user and process owner", "H", "H",
            "Move beds/staff between departments, escalate a bottleneck, set next week's occupancy target."],
           ["SH-03", "Department Heads (Cardiology, Orthopaedics, General Medicine, Paediatrics, Maternity, Oncology, Emergency, ENT)", "Clinical unit owners", "H", "M",
            "Accept or challenge their ALOS/readmission numbers and commit a corrective action."],
           ["SH-04", "Quality / NABH Accreditation Officer", "Compliance owner", "H", "M",
            "Decide whether a quality indicator breaches the internal threshold and needs a CAPA (corrective and preventive action)."],
           ["SH-05", "Finance / Revenue Cycle Manager", "Revenue owner", "H", "M",
            "Prioritise payer-mix improvement, pricing review or discharge-billing correction."],
           ["SH-06", "Nursing Superintendent", "Ward staffing owner", "M", "M",
            "Approve shift roster changes and ward-level escalation when occupancy crosses 90%."],
           ["SH-07", "IT / Data Governance Officer", "Data custodian (in real hospital)", "L", "M",
            "Approve data extraction, refresh schedule and access policy for the dashboard."],
           ["SH-08", "Registration / Front-desk Supervisor", "Data entry owner", "M", "L",
            "Fix registration errors (missing payer, wrong department) that distort KPIs."],
           ["SH-09", "Patient / Attendant", "Indirect beneficiary", "M", "L",
            "No direct decision — benefits from shorter waits and better bed allocation."],
           ["SH-10", "Faculty mentor and evaluation panel", "Academic approver", "H", "H",
            "Accept the deliverable, grade the viva, approve progression to the next step."]],
          [0.6, 1.5, 1.2, 0.6, 0.65, 2.15], font_size=8)
para(doc, "Engagement plan (summary): SH-02 and SH-10 are engaged every step; SH-01, SH-03, SH-04 and "
          "SH-05 review the BRD (Step 5) and user stories (Step 11); SH-07 reviews data governance "
          "statements in Steps 7–8; SH-08 receives a one-page data-quality checklist at Step 15.",
     size=9.5, italic=True)

# ---------- 6 ----------
h1(doc, "6. Operational Questions the App Must Answer")
para(doc, "These ten questions define the functional boundary of the app. Every screen, filter, chart "
          "and test case must serve at least one OQ; anything that serves none is out of scope. Each "
          "question is written so that the answer is a number or a ranked list — not an essay.",
     space_after=4)
add_table(doc,
          ["OQ ID", "Operational question", "Decision it supports", "KPI / measure used", "Asked"],
          [["OQ-01", "Which departments are above or below the 85% bed-occupancy target this month, and by how much?",
            "Where to move beds/staff; which HOD to escalate", "BOR vs target, gap in percentage points", "Daily / weekly"],
           ["OQ-02", "How are admissions, discharges and OPD visits trending week over week, and is any week an outlier?",
            "Staffing and OPD slot planning for the coming week", "Admissions, discharges, OPD visits; WoW % change", "Weekly"],
           ["OQ-03", "Which departments exceed the hospital ALOS benchmark of 4.0 days, and which patient segments drive the excess?",
            "Case-mix review; discharge-process fix", "ALOS by department, age group, payer", "Monthly"],
           ["OQ-04", "What is the 30-day readmission rate, and which departments and diagnoses contribute most?",
            "Targeted clinical audit and discharge-counselling effort", "Readmission rate, contribution %", "Monthly"],
           ["OQ-05", "How does revenue per patient vary by payer category, and which payer mix is improving or deteriorating?",
            "Payer negotiation, pricing review, collection focus", "Revenue per discharge, payer share", "Monthly"],
           ["OQ-06", "Which weekdays and months carry peak OPD footfall, and how far does peak exceed average?",
            "Roster planning, OPD token capacity, consultant scheduling", "OPD visits by weekday/month, peak-to-average ratio", "Monthly"],
           ["OQ-07", "Is higher occupancy associated with longer stays or higher readmissions, and how strong is that association?",
            "Decide whether occupancy pressure is a quality risk", "Pearson/Spearman correlation, scatter plot", "Quarterly"],
           ["OQ-08", "Which age groups and genders dominate admissions, and how does the mix differ by department?",
            "Service-line planning, paediatrics/geriatric capacity", "Admission share by age band and gender", "Quarterly"],
           ["OQ-09", "Which departments earn the least revenue per bed-day, and what is the size of the gap to the best performer?",
            "Resource reallocation and service-line profitability review", "Revenue per bed-day, ranking, gap", "Monthly"],
           ["OQ-10", "Is the in-hospital mortality rate and discharge-type mix stable, or is any department drifting?",
            "Trigger a clinical quality review / CAPA", "Mortality rate, discharge-type share, trend", "Monthly"]],
          [0.62, 2.75, 1.6, 1.35, 0.75], font_size=8)
para(doc, "Non-functional expectation attached to every OQ: the answer must be reachable in at most "
          "three interactions (choose period → choose department → read) and must state the denominator "
          "and the data period on screen, so the user never has to ask \u201cof what?\u201d", size=9.5,
     italic=True)

# ---------- 7 ----------
h1(doc, "7. Objectives (SMART)")
add_table(doc,
          ["OBJ ID", "Objective", "Specific", "Measurable target", "Achieved by"],
          [["OBJ-01", "Consolidate the six core hospital KPIs into one dashboard with published definitions.",
            "Six KPIs: BOR, ALOS, readmission, OPD footfall, revenue per patient, mortality.",
            "6 of 6 KPIs live, each with a written definition and denominator shown on screen.", "Step 15"],
           ["OBJ-02", "Build a reproducible synthetic dataset at the same grain as real hospital sources.",
            "3 calendar years (2023–2025), 8 departments, admissions + bed capacity + OPD + finance.",
            "\u2265 18,000 admission rows, \u2265 60,000 OPD rows, 0 orphan foreign keys, regeneration in < 30 s.", "Step 10"],
           ["OBJ-03", "Cut the time to answer a routine operational question.",
            "Replace the ~45-minute manual pivot with a filtered dashboard view.",
            "Any OQ answered in \u2264 60 seconds by a first-time user (usability test, 8 users).", "Step 15"],
           ["OBJ-04", "Prove correctness through a documented test suite.",
            "Functional tests for every FR plus technical tests for data, performance and access.",
            "\u2265 25 test cases, 100% pass on functional, \u2265 85% pytest line coverage.", "Step 15"],
           ["OBJ-05", "Make the dashboard publicly reachable and responsive.",
            "Deploy on Streamlit Community Cloud against the packaged SQLite database.",
            "Public URL live; cold start \u2264 10 s, filter interaction \u2264 2 s on a 20 Mbps connection.", "Step 16"],
           ["OBJ-06", "Deliver submission-ready academic artefacts.",
            "ABES project document + slide deck, fully traceable to IDs defined from Step 1 onward.",
            "0 broken ID references in the RTM; document and deck submitted on or before the deadline.", "Steps 17–18"]],
          [0.65, 1.85, 1.55, 2.05, 0.6], font_size=8)

# ---------- 8 ----------
h1(doc, "8. Scope", page_break=True)
h2(doc, "8.1 In scope")
add_table(doc,
          ["#", "In-scope capability", "Why it is included"],
          [["S1", "Read-only analytical dashboard over a pre-built SQLite dataset (no live HIS write-back).",
            "Delivers the OQs without depending on hospital IT access."],
           ["S2", "Six KPI families: occupancy, ALOS/volume, readmissions, OPD footfall, revenue & payer mix, mortality/discharge mix.",
            "Matches the six indicators leadership already tracks."],
           ["S3", "Filters: date range, department, ward/bed class, payer category, age band, gender.",
            "Enables drill-down and fair comparison (same period, same unit)."],
           ["S4", "Interactive Plotly visuals: trend line, bar comparison, stacked payer mix, scatter with correlation, heatmap of weekday × department.",
            "Answers \u201cwhat changed\u201d and \u201cwho drove it\u201d in one view."],
           ["S5", "Descriptive statistics and a documented correlation test (OQ-07), with assumptions checked.",
            "Adds the statistical depth expected of a DA project."],
           ["S6", "KPI definition panel and data-quality panel (row counts, missing values, date coverage).",
            "Prevents the two failure modes of hospital dashboards: ambiguous definitions and silent data gaps."],
           ["S7", "CSV download of the filtered view for offline review.",
            "Supports SH-04's audit evidence and SH-03's review pack."],
           ["S8", "Synthetic data generator script, SQLite build script and pytest suite.",
            "Makes the project reproducible and gradeable."]],
          [0.45, 3.4, 2.85], font_size=8.5)

h2(doc, "8.2 Out of scope")
add_table(doc,
          ["#", "Out-of-scope item", "Reason"],
          [["X1", "Live integration with a real HIS / EHR or HL7/FHIR feed.", "No hospital access; academic project; real data prohibited."],
           ["X2", "Patient-level clinical notes, prescriptions or any identifiable health information.", "Privacy and ethics; synthetic data only."],
           ["X3", "Predictive modelling (readmission risk scoring, bed-demand forecasting).", "Kept as a future enhancement; would dilute 18-step delivery."],
           ["X4", "Write operations: bed allocation, discharge entry, billing edits.", "This is a decision-support tool, not a transaction system."],
           ["X5", "Role-based authentication, SSO or audit logging.", "Free-tier deployment; a single shared passphrase gate is the maximum."],
           ["X6", "Multi-hospital / multi-tenant comparison.", "Single-hospital model keeps definitions and schema honest."],
           ["X7", "Mobile-native app, offline mode, scheduled email alerts.", "Out of platform capability for a Streamlit Cloud academic deployment."],
           ["X8", "Statutory reporting formats (GST, insurance claim files, government returns).", "Domain-specific compliance work unrelated to the six KPIs."]],
          [0.45, 3.4, 2.85], font_size=8.5)

# ---------- 9 ----------
h1(doc, "9. Assumptions and Constraints")
h2(doc, "9.1 Assumptions (AS)")
para(doc, "Each assumption is testable; if one proves false, the affected sections are flagged rather "
          "than silently changed.", size=9.5, italic=True, space_after=4)
add_table(doc,
          ["AS ID", "Assumption", "Impact if false"],
          [["AS-01", "The hospital is a single 300-bed, 8-department multi-speciality unit.", "Dataset size, department filters and benchmarks change."],
           ["AS-02", "One row of the admission fact table = one inpatient episode (admission to discharge, including transfers within the stay).", "Readmission and ALOS maths must be redefined."],
           ["AS-03", "Bed occupancy is computed as bed-days occupied ÷ bed-days available (functional beds, not sanctioned beds).", "Occupancy values shift; target 85% may be unreachable."],
           ["AS-04", "An unplanned readmission is any inpatient admission to the same department within 30 days of a discharge, excluding planned follow-up procedures.", "Readmission rate would inflate by ~2–3 percentage points."],
           ["AS-05", "Revenue means gross billed amount per discharge, before deductions and write-offs.", "Revenue per patient ranking may change by payer."],
           ["AS-06", "Data refreshes once per day (T+1); near-real-time is not required.", "Streaming architecture would be needed — out of scope."],
           ["AS-07", "Synthetic data must reproduce realistic seasonality (winter respiratory peaks, monsoon fever/OPD spikes) and weekday effects (Monday OPD peak, weekend dip).", "Charts look synthetic and correlation tests lose meaning."],
           ["AS-08", "Users have basic spreadsheet literacy and a browser; no SQL knowledge is assumed.", "UI would need an SQL console or analyst-only framing."],
           ["AS-09", "The dataset is small enough for SQLite and in-memory aggregation (< 200 MB).", "Would need DuckDB/Postgres and a different deployment path."],
           ["AS-10", "Benchmarks used (85% occupancy, 4.0-day ALOS, < 5% readmission, 1.5% mortality) are domain-typical internal targets, not regulatory claims.", "Success criteria would need re-baselining against hospital policy."],
           ["AS-11", "English interface with Indian numbering conventions (\u20b9, lakh/crore shown as compact values).", "Localisation work required."],
           ["AS-12", "Mentor's ABES template governs the final document structure and font/heading conventions.", "Step 17 layout would be redone."]],
          [0.62, 3.75, 2.3], font_size=8.5)

h2(doc, "9.2 Constraints (CN)")
add_table(doc,
          ["CN ID", "Constraint", "Type", "Mitigation"],
          [["CN-01", "Academic calendar: the 18 steps must finish within the semester; roughly one step every 2–3 days.", "Schedule", "Steps 1–9 are document-only and can run in parallel with dataset thinking."],
           ["CN-02", "Streamlit Community Cloud free tier: ~1 GB RAM, app sleeps after inactivity, single shared deployment.", "Platform", "Pre-aggregate in SQL views; cache with st.cache_data; keep the app cold-start light."],
           ["CN-03", "SQLite is single-writer and file-based; the deployed database is read-only.", "Technology", "Ship a pre-built .db file; all writes happen in the build script, not in the app."],
           ["CN-04", "No real patient data may be used or shown anywhere, including screenshots.", "Ethics / legal", "Synthetic generator with documented seed; screenshots carry a \u201csynthetic data\u201d watermark."],
           ["CN-05", "Python 3.11+ only; libraries limited to Streamlit, pandas, NumPy, Plotly, pytest, python-docx/openpyxl for artefacts.", "Technology", "No heavy ML stack; keeps deployment and viva explanation simple."],
           ["CN-06", "Tailwind UI must be applied through Streamlit's HTML/CSS injection without breaking accessibility or mobile layout.", "Design", "One shared CSS block; contrast checked; tested at 375 px and 1440 px widths."],
           ["CN-07", "University deadline and mentor review slots are fixed; late artefacts block the next step.", "Governance", "Each step ends with a signed-off file before the next begins."],
           ["CN-08", "Single developer (the student) — no separate tester, designer or DBA.", "Resourcing", "pytest automation plus a self-review checklist act as the second pair of eyes."]],
          [0.62, 3.25, 0.95, 1.9], font_size=8.5)

# ---------- 10 ----------
h1(doc, "10. Success Criteria (Measurable)", page_break=True)
para(doc, "A criterion is only listed if it can be checked by a third party without asking the author. "
          "Each one maps to an objective and will be re-used verbatim in Step 13 (test cases) and "
          "Step 15 (build/test report).", space_after=4)
add_table(doc,
          ["SC ID", "Success criterion", "Target / threshold", "How it is measured", "Links to"],
          [["SC-01", "KPI coverage", "6 of 6 KPIs visible with definition and denominator", "Screen-by-screen checklist against Section 6", "OBJ-01"],
           ["SC-02", "Data completeness", "0 missing mandatory fields; 0 orphan FKs; 100% date coverage 2023-01-01 → 2025-12-31", "SQL quality queries in the data-quality panel + pytest", "OBJ-02"],
           ["SC-03", "Answer speed", "\u2264 60 s to answer any OQ for a first-time user", "Moderated usability test, 8 participants, timed tasks", "OBJ-03"],
           ["SC-04", "Technical response time", "p95 filter response \u2264 2 s; cold start \u2264 10 s", "Browser timing over 20 runs, median and p95 reported", "OBJ-05"],
           ["SC-05", "Functional correctness", "100% of functional test cases pass", "pytest + manual UI test log, 0 open Sev-1/Sev-2 defects", "OBJ-04"],
           ["SC-06", "Code quality", "pytest line coverage \u2265 85%; 0 flake8 E9/F errors", "coverage report + flake8 run archived with the test log", "OBJ-04"],
           ["SC-07", "Usability", "Mean SUS \u2265 75 from 8–10 evaluators; no critical accessibility failure", "10-item SUS questionnaire + contrast/keyboard check", "OBJ-03, OBJ-06"],
           ["SC-08", "Traceability", "Every BR → FR → US → TC chain resolvable; 0 orphan IDs in the RTM", "RTM cross-check script + manual audit sample of 10 chains", "OBJ-06"],
           ["SC-09", "Deployment", "Public URL reachable; all 7 screens load; no unhandled exception in logs for 7 consecutive days", "Deployment log + uptime spot checks", "OBJ-05"],
           ["SC-10", "Academic acceptance", "Mentor sign-off on Steps 1–18 and on the final document and deck", "Signed review record in the project log", "OBJ-06"]],
          [0.62, 1.85, 1.75, 1.9, 0.55], font_size=8)
para(doc, "Stretch (nice-to-have, not acceptance-blocking): heatmap export to PNG, per-department PDF one-pager, "
          "and a what-if slider for occupancy target.", size=9.5, italic=True)

# ---------- 11 ----------
h1(doc, "11. Check: Can Every OQ Be Answered From Data?")
para(doc, "This is the Step 1 quality gate. For each operational question the table names the exact "
          "fields the app must hold and states whether the question is answerable with the planned "
          "dataset. \u201cYes\u201d means the computation is fully defined here; \u201cYes + AS-nn\u201d "
          "means it is answerable provided the named assumption holds. No OQ may remain \u201cNo\u201d at "
          "the end of this step — an unanswerable OQ is either deleted or its data is added to scope.",
     space_after=4)
add_table(doc,
          ["OQ ID", "Data fields required", "Source concept (provisional)", "Answerable?", "Computation note"],
          [["OQ-01", "department_id, month, bed_days_available, bed_days_occupied", "Bed-capacity fact + admission fact", "Yes + AS-03",
            "BOR = \u03a3 bed_days_occupied \u00f7 \u03a3 bed_days_available; gap = BOR \u2212 85%."],
           ["OQ-02", "admission_date, discharge_date, visit_date, department_id", "Admission fact, OPD fact", "Yes",
            "Count per ISO week; WoW % = (current \u2212 prior) \u00f7 prior; outlier = |z| > 2."],
           ["OQ-03", "admission_date, discharge_date, department_id, age_band, payer_id", "Admission fact", "Yes",
            "ALOS = \u03a3 (discharge \u2212 admission) \u00f7 discharges; benchmark 4.0 days (AS-10)."],
           ["OQ-04", "patient_key (pseudonymous), admission_date, discharge_date, department_id, diagnosis_group, readmit_flag", "Admission fact (self-join)", "Yes + AS-04",
            "Rate = unplanned readmissions \u2264 30 days \u00f7 index discharges; contribution % per department."],
           ["OQ-05", "payer_id, billed_amount, discharge_date, department_id", "Finance fact + dimension payer", "Yes + AS-05",
            "Revenue per discharge = \u03a3 billed_amount \u00f7 discharges; payer share = \u03a3 amount by payer \u00f7 total."],
           ["OQ-06", "visit_date, department_id, weekday (derived), visit_count", "OPD fact", "Yes",
            "Average visits per weekday; peak-to-average ratio = max weekday avg \u00f7 overall avg."],
           ["OQ-07", "occupancy_rate (monthly, per dept), alos, readmission_rate", "Aggregated from OQ-01/03/04", "Yes + AS-02",
            "Spearman + Pearson on monthly dept-level points (n = 8 \u00d7 36 = 288); report r, p, n and caveat that association \u2260 causation."],
           ["OQ-08", "age_band, gender, department_id, admission_date", "Admission fact", "Yes",
            "Share = admissions in band \u00f7 total admissions; chi-square test of independence by department (optional depth)."],
           ["OQ-09", "department_id, billed_amount, bed_days_occupied", "Finance fact + bed-capacity fact", "Yes + AS-03, AS-05",
            "Revenue per bed-day = \u03a3 billed_amount \u00f7 \u03a3 bed_days_occupied; gap = best \u2212 department."],
           ["OQ-10", "discharge_date, discharge_type, death_flag, department_id", "Admission fact", "Yes",
            "Mortality rate = deaths \u00f7 discharges; control-chart style 3-month moving average to show drift."]],
          [0.6, 2.0, 1.45, 0.85, 2.15], font_size=7.8)

h2(doc, "11.1 Minimum data the app must hold (provisional)")
para(doc, "This is not the database design (that is Step 8, where T-IDs are issued). It exists only to "
          "prove the OQs are answerable. Grain and field names below are indicative and will be frozen in "
          "Step 8.", size=9.5, italic=True, space_after=4)
add_table(doc,
          ["Source concept", "Grain (one row =)", "Key fields", "Supports"],
          [["Inpatient episode", "one admission/discharge episode", "patient_key, department_id, ward_id, admission_date, discharge_date, discharge_type, death_flag, diagnosis_group, payer_id", "OQ-02/03/04/08/10"],
           ["Bed capacity", "one department-month (optionally ward)", "department_id, month, beds_functional, bed_days_available, bed_days_occupied", "OQ-01/09"],
           ["OPD visit", "one outpatient visit", "visit_date, department_id, age_band, gender, payer_id", "OQ-02/06/08"],
           ["Finance", "one discharge billing record", "episode_ref, billed_amount, payer_id, department_id", "OQ-05/09"],
           ["Payer dimension", "one payer category", "payer_id, payer_name (Cash, Insurance, TPA, Government scheme, Corporate)", "OQ-05"],
           ["Department / ward dimension", "one department or ward", "department_id, department_name, ward_id, ward_type", "all"]],
          [1.15, 1.35, 3.1, 1.1], font_size=8)

para(doc, "Gate result: 10 of 10 operational questions are answerable from the planned data. Five depend "
          "on explicitly stated assumptions (AS-02, AS-03, AS-04, AS-05, AS-10), all of which are recorded "
          "in Section 9. No OQ requires data that is out of scope, and no OQ needs to be dropped. Verdict: "
          "PASS — proceed to Step 2.", bold=True)

# ---------- 12 ----------
h1(doc, "12. Risks and Open Questions Carried Forward")
h2(doc, "12.1 Top risks")
add_table(doc,
          ["Risk ID", "Risk", "L", "I", "Response"],
          [["R-01", "Occupancy cross 100% because bed-days available use sanctioned beds while occupied uses functional beds.", "M", "H", "Freeze one definition (AS-03) and validate with a SQL check \u2264 100% before Step 10 is closed."],
           ["R-02", "Synthetic data too uniform, so correlation and trend visuals look artificial.", "M", "M", "Inject seasonality and weekday effects (AS-07); test with OQ-07 and OQ-06 before Step 15."],
           ["R-03", "Scope creep toward prediction or live HIS integration.", "H", "M", "Out-of-scope list X1–X8 is contractual; changes need a written change note."],
           ["R-04", "Streamlit Cloud memory limit during heavy groupby operations.", "L", "M", "Pre-aggregate into SQL views and cache; load < 200 MB (AS-09)."],
           ["R-05", "Definition drift between the doc and the app as steps progress.", "M", "H", "KPI definitions panel reads the same text block used in the BRD; RTM cross-check at Step 4 and 15."]],
          [0.7, 3.0, 0.4, 0.4, 2.15], font_size=8)

h2(doc, "12.2 Open questions (to be closed by the named step)")
add_table(doc,
          ["OQ-Ref", "Open question", "Owner", "Close by"],
          [["OQ-11", "Should bed occupancy be shown monthly only, or also as a daily heatmap?", "SH-02", "Step 12 (UI)"],
           ["OQ-12", "Is revenue reported gross billed or net collected for the finance view?", "SH-05", "Step 3 (BR)"],
           ["OQ-13", "Does the mentor expect a passphrase gate on the public URL?", "SH-10", "Step 16 (deploy)"]],
          [0.7, 3.55, 1.05, 1.4], font_size=8.5)

# ---------- 13 ----------
h1(doc, "13. Step 1 Sign-off")
para(doc, "What this document commits the project to:", space_after=4)
bullet(doc, "10 operational questions (OQ-01 … OQ-10), each proven answerable from planned data.")
bullet(doc, "6 SMART objectives (OBJ-01 … OBJ-06) and 10 measurable success criteria (SC-01 … SC-10).")
bullet(doc, "10 stakeholders (SH-01 … SH-10), 12 assumptions (AS-01 … AS-12), 8 constraints (CN-01 … CN-08).")
bullet(doc, "An explicit in-scope (S1–S8) and out-of-scope (X1–X8) boundary.")
bullet(doc, "Open questions OQ-11 … OQ-13 carried forward with named owners.")
para(doc, "Nothing in later steps may contradict this document without a stated change note: "
          "\u201cStep n changes <ID> from <old> to <new> because <reason>\u201d.", italic=True, space_after=10)
para(doc, "Prepared by: " + STUDENT + "   |   Roll: " + ROLL, size=10)
para(doc, "Reviewed by (mentor): " + MENTOR + "   |   Date: 19-Nov-2026", size=10)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print("saved:", OUT)
