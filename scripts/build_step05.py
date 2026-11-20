"""Build Step 5 deliverable: Business Requirements Document (BRD) for HKD-2026."""
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

PROJECT_ID = "HKD-2026"
PROJECT_TITLE = "Healthcare KPI Dashboard"
STUDENT = "Ayush Pal"
ROLL = "2503201450035"
MENTOR = "Aplana Lodhi"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "05_BRD.docx")


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
    run = p.add_run(f"{PROJECT_ID}  |  Step 5 — Business Requirements Document  |  Page ")
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

# ---------------- cover ----------------
para(doc, "ABES ENGINEERING COLLEGE", 11, bold=True, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para(doc, "Department of Computer Applications (MCA) — Data Analytics", 10, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=40)
para(doc, "STEP 5", 26, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "BUSINESS REQUIREMENTS DOCUMENT (BRD)", 18, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
para(doc, PROJECT_TITLE, 17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Project ID: {PROJECT_ID}   |   Domain: Data Analytics", 11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=36)

meta = [("Document", "Business Requirements Document (BRD)"),
        ("Student", STUDENT), ("Roll Number", ROLL),
        ("Mentor / reviewer", MENTOR),
        ("Version / Date", "v1.0 — 19 November 2026"),
        ("Baseline inputs", "Steps 1–4 (Project Understanding, HMW, Needs Analysis, RTM)"),
        ("Status", "Baseline — approved for Step 6 (PRD)"),
        ("Classification", "Academic; synthetic data only")]
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

# ---------------- 1 doc control ----------------
doc.add_page_break()
h1(doc, "1. Document Control")
h2(doc, "1.1 Revision history")
add_table(doc, ["Version", "Date", "Author", "Change description", "Status"],
          [["0.1", "19-Nov-2026", STUDENT, "Initial consolidation of Steps 1–4", "Draft"],
           ["0.2", "19-Nov-2026", STUDENT, "Added business rules BRL-01…BRL-14 and dependency register", "Draft"],
           ["1.0", "19-Nov-2026", STUDENT, "Baseline issued for Step 6; 15 BRs, 14 business rules, RACI agreed", "Approved"]],
          [0.8, 1.0, 1.1, 3.1, 0.8])

h2(doc, "1.2 Reviewers and approval")
add_table(doc, ["Role", "Name / ID", "Review scope", "Outcome"],
          [["Mentor (approver)", MENTOR + " (SH-10)", "Scope, requirement quality, traceability, academic fit", "Approved v1.0"],
           ["Process owner", "Operations Manager (SH-02)", "Daily-use realism, KPI relevance, wording clarity", "Reviewed, no change"],
           ["Quality reviewer", "Quality / NABH Officer (SH-04)", "Readmission, mortality and data-quality rules", "Reviewed, BRL-04 amended"],
           ["Finance reviewer", "Finance Manager (SH-05)", "Revenue definition (gross billed), payer categories", "Reviewed, BRL-05 added"],
           ["Data governance", "IT / Data Governance (SH-07)", "Definitions, refresh expectation, field ownership", "Reviewed, no change"]],
          [1.35, 1.85, 2.4, 1.0], font_size=8.2)

h2(doc, "1.3 Change control")
para(doc, "This document is baselined at v1.0. Any change after this point follows a fixed routine:")
bullet(doc, "A new requirement is added at the end of the register as the next free BR number (BR-16, BR-17 …). Existing BR numbers are never reused or renumbered.")
bullet(doc, "Each change carries a note in the form: \u201cAdded BR-nn because <reason>; source HMW-xx; priority <MoSCoW>; RTM updated on <date>\u201d.")
bullet(doc, "The RTM file (04_Traceability_Matrix.xlsx) must be updated in the same sitting, otherwise the BRD and the matrix disagree and traceability is broken.")
bullet(doc, "A requirement is never deleted. It is marked Deferred with the reason and the stakeholder agreement recorded.")
para(doc, "Instruction to the project team: if a new requirement is introduced during any later step, add it at "
          "the end and explicitly state \u201cRTM update required\u201d.", italic=True)

# ---------------- 2 exec summary ----------------
h1(doc, "2. Executive Summary", page_break=True)
para(doc, "The Healthcare KPI Dashboard (HKD-2026) is an analytical decision-support product for a 300-bed, "
          "eight-department multi-speciality hospital. It gives hospital leadership one consistent, "
          "self-service view of six operational indicators — bed occupancy, length of stay, 30-day "
          "readmissions, outpatient footfall, revenue per patient by payer, and mortality with discharge-type "
          "mix — each shown against a target, a trend and a driver.")
para(doc, "The business problem is not a shortage of data but a shortage of agreement. Six sources describe the "
          "same week differently, a routine question costs about 45 minutes of manual work, and the numbers "
          "that decide bed allocation and staffing are reviewed after the decision has already been taken. "
          "The product replaces that reconciliation ritual with a fixed set of definitions, a reproducible "
          "calculation, and a view that names the department, payer and weekday responsible for a movement.")
para(doc, "Scope for release 1 is deliberately narrow: 15 business requirements, of which 9 are Must; ten "
          "operational questions all proven answerable from the planned data; and one consciously excluded "
          "requirement (automatic alerting, BR-15). Delivery is incremental across 18 documented steps, with "
          "traceability from every operational question to the test case that proves it works.")
para(doc, "Expected benefit, stated conservatively: a routine operational question answered in under a minute "
          "instead of 45 minutes; a single agreed occupancy figure, ending the two-number argument in review "
          "meetings; and early visibility of a department drifting on readmissions or mortality before an "
          "audit finds it. Success is measured by ten criteria (SC-01 … SC-10), not by opinion.", space_after=8)

# ---------------- 3 problem ----------------
h1(doc, "3. Business Problem and Opportunity")
h2(doc, "3.1 Problem statement (unchanged from Step 1)")
para(doc, "Hospital leadership has no single, trustworthy, same-day view of operational KPIs — bed occupancy, "
          "length of stay, readmissions, OPD load and revenue per patient are scattered across registers, "
          "spreadsheets and HIS exports with inconsistent definitions — so capacity, staffing and quality "
          "decisions are taken on stale or contradictory numbers.", italic=True, bold=True)
h2(doc, "3.2 Evidence of the problem")
add_table(doc, ["Symptom", "Observed impact", "Affected stakeholder"],
          [["Month-end consolidation takes 2–4 days", "Decisions run on last month's reality; a bad week is discovered three weeks late", "SH-01, SH-02"],
           ["Occupancy computed against different bed bases", "Two departments quote different numbers for the same week; trust collapses", "SH-01, SH-03"],
           ["~45 minutes per ad-hoc question", "Only 3–5 questions answered per week; the rest are guessed", "SH-02"],
           ["Reports show value but not driver", "Corrective action is generic (\u201cimprove occupancy\u201d) instead of targeted", "SH-01, SH-03"],
           ["Quality indicators reviewed only at audit", "Drift becomes a finding rather than a managed signal", "SH-04"],
           ["Registration errors discovered at billing", "Revenue KPI distorted weeks later, with no feedback loop", "SH-05, SH-08"]],
          [1.85, 3.3, 1.35], font_size=8.4)
h2(doc, "3.3 Opportunity")
para(doc, "Three opportunities justify the investment: (i) the six KPIs already exist in the data — the gap is "
          "definition and presentation, not data collection; (ii) a single normalised view removes the "
          "reconciliation meeting entirely, returning roughly two analyst-days per month; (iii) once "
          "definitions are frozen, later additions (forecasting, alerting, payer analytics) become extensions "
          "rather than rewrites.")
h2(doc, "3.4 Business value (qualitative, with the measure that will prove it)")
add_table(doc, ["Value driver", "Mechanism", "Proved by"],
          [["Faster decisions", "Question → answer in one screen instead of a spreadsheet rebuild", "SC-03: median \u2264 60 s"],
           ["One version of truth", "Single definition registry reused in every view", "SC-01: 6/6 KPIs with definition and denominator"],
           ["Targeted action", "Movement decomposed into dept / payer / weekday contributions", "BR-03 acceptance test"],
           ["Earlier risk detection", "Drift indicator on readmission and mortality", "BR-07, BR-13 acceptance tests"],
           ["Trust in numbers", "Data-quality status shown beside the KPI", "BR-14 acceptance test"]],
          [1.5, 3.1, 1.9], font_size=8.4)

# ---------------- 4 objectives ----------------
h1(doc, "4. Objectives and Success Metrics", page_break=True)
h2(doc, "4.1 Business objectives (from Step 1, unchanged)")
add_table(doc, ["OBJ", "Objective", "Target"],
          [["OBJ-01", "Consolidate six core KPIs with published definitions", "6 of 6 live with denominator shown"],
           ["OBJ-02", "Reproducible synthetic dataset at real hospital grain", "\u2265 18,000 admissions, \u2265 60,000 OPD rows, 0 orphan keys"],
           ["OBJ-03", "Cut time to answer a routine question", "\u2264 60 s by a first-time user"],
           ["OBJ-04", "Prove correctness with a documented test suite", "\u2265 25 cases, 100% functional pass, \u2265 85% coverage"],
           ["OBJ-05", "Public, responsive deployment", "Cold start \u2264 10 s, interaction \u2264 2 s"],
           ["OBJ-06", "Submission-ready artefacts, fully traceable", "0 broken IDs in the RTM"]],
          [0.75, 3.2, 2.6], font_size=8.4)

h2(doc, "4.2 KPI targets the product must support")
add_table(doc, ["KPI", "Definition used (see BRL)", "Internal target", "Source OQ", "Delivers"],
          [["Bed occupancy rate (BOR)", "Bed-days occupied \u00f7 bed-days available (functional beds), monthly and weekly", "85% (\u00b15 pp tolerance band)", "OQ-01", "BR-01"],
           ["Average length of stay (ALOS)", "Inpatient days \u00f7 discharges", "4.0 days", "OQ-03", "BR-06"],
           ["30-day readmission rate", "Unplanned readmissions \u2264 30 days \u00f7 index discharges", "< 5.0%", "OQ-04", "BR-07"],
           ["OPD footfall peak ratio", "Busiest weekday average \u00f7 overall weekday average", "\u2264 1.30 (planning assumption)", "OQ-06", "BR-09"],
           ["Revenue per patient", "Gross billed \u00f7 discharges", "Department-specific; reported, not targeted", "OQ-05", "BR-08"],
           ["Revenue per bed-day", "Gross billed \u00f7 bed-days occupied", "Baseline year 2023; improvement tracked", "OQ-09", "BR-12"],
           ["In-hospital mortality", "Deaths \u00f7 discharges", "< 1.5%", "OQ-10", "BR-13"]],
          [1.35, 2.35, 1.35, 0.7, 0.6], font_size=7.8)
para(doc, "These targets are internal planning benchmarks (AS-10), not regulatory claims. Every KPI shown must also "
          "display its denominator and, where the denominator is below 30, a low-confidence marker (BRL-10).",
     size=9.5, italic=True)

h2(doc, "4.3 Product success criteria (from Step 1, unchanged)")
add_table(doc, ["SC", "Criterion", "Threshold"],
          [["SC-01", "KPI coverage", "6/6 with definition and denominator"],
           ["SC-02", "Data completeness", "0 missing mandatory fields, 0 orphan FKs, full date coverage"],
           ["SC-03", "Answer speed", "\u2264 60 s, first-time user"],
           ["SC-04", "Technical response", "p95 \u2264 2 s, cold start \u2264 10 s"],
           ["SC-05", "Functional correctness", "100% functional tests pass"],
           ["SC-06", "Code quality", "\u2265 85% coverage, 0 flake8 E9/F errors"],
           ["SC-07", "Usability", "Mean SUS \u2265 75"],
           ["SC-08", "Traceability", "0 orphan IDs in the RTM"],
           ["SC-09", "Deployment", "Public URL, 7 days without unhandled exception"],
           ["SC-10", "Academic acceptance", "Mentor sign-off, Steps 1–18"]],
          [0.65, 2.1, 3.7], font_size=8.4)

# ---------------- 5 RACI ----------------
h1(doc, "5. Stakeholders and RACI", page_break=True)
para(doc, "R = responsible (does the work), A = accountable (single owner of the outcome), C = consulted "
          "(two-way input required), I = informed (one-way update). Every activity has exactly one A.",
     space_after=4)
add_table(doc, ["Activity", "SH-01\nDirector", "SH-02\nOps Mgr", "SH-03\nHODs", "SH-04\nQuality",
                "SH-05\nFinance", "SH-06\nNursing", "SH-07\nIT Gov", "SH-08\nRegistr.", "SH-09\nPatient", "SH-10\nMentor"],
          [["A1 Business requirements", "C", "R", "C", "C", "C", "I", "I", "I", "I", "A"],
           ["A2 KPI definition sign-off", "A", "R", "C", "C", "C", "I", "C", "I", "\u2013", "C"],
           ["A3 Data rules & quality", "I", "A", "I", "R", "C", "I", "C", "C", "\u2013", "I"],
           ["A4 Design / UI review", "I", "R", "C", "C", "I", "C", "I", "I", "\u2013", "A"],
           ["A5 Usability & UAT", "I", "R", "C", "C", "C", "C", "I", "I", "\u2013", "A"],
           ["A6 Scope change approval", "C", "C", "I", "I", "I", "I", "I", "\u2013", "\u2013", "A"],
           ["A7 Final document & deck", "C", "C", "I", "I", "I", "I", "I", "\u2013", "\u2013", "A"]],
          [1.7] + [0.52] * 10, font_size=7.6)
para(doc, "The student (author) is R on every activity in practice; the table records who the work is done for, "
          "not who types it. SH-09 is informed only — the patient does not participate in decisions in this "
          "release (GAP-01, Step 3 §5.3).", size=9.5, italic=True)

h2(doc, "5.1 Stakeholder engagement summary")
add_table(doc, ["Stakeholder", "Interest", "Influence", "How they are engaged"],
          [["SH-01 Hospital Director", "M", "H", "Monthly review of the executive view; receives BRD and final deck"],
           ["SH-02 Operations Manager", "H", "H", "Primary user; reviews every artefact; daily UAT participant"],
           ["SH-03 Department Heads", "H", "M", "Review ALOS and readmission views; fairness check on comparisons"],
           ["SH-04 Quality / NABH", "H", "M", "Owns readmission/mortality rules; provides audit questions"],
           ["SH-05 Finance", "H", "M", "Owns revenue definitions and payer categories"],
           ["SH-06 Nursing Superintendent", "M", "M", "Validates ward-level occupancy and roster use case"],
           ["SH-07 IT / Data Governance", "L", "M", "Reviews definitions, refresh expectation and data dictionary"],
           ["SH-08 Registration Supervisor", "M", "L", "Receives data-quality feedback checklist"],
           ["SH-09 Patient / Attendant", "M", "L", "Informed only; benefit measured indirectly through wait metrics"],
           ["SH-10 Mentor / Panel", "H", "H", "Approves each step; final academic sign-off"]],
          [1.9, 0.7, 0.8, 3.2], font_size=8.4)

# ---------------- 6 scope ----------------
h1(doc, "6. Scope", page_break=True)
h2(doc, "6.1 In scope")
add_table(doc, ["ID", "In-scope capability", "Acceptance evidence"],
          [["S1", "Read-only analytical view over a pre-built database; no write-back to any hospital system", "No editing control exists anywhere in the product"],
           ["S2", "Six KPI families: occupancy, ALOS/volume, readmissions, OPD footfall, revenue & payer mix, mortality/discharge mix", "6/6 KPIs with definition and denominator (SC-01)"],
           ["S3", "Filters: date range, department, ward/bed class, payer, age band, gender", "Every filter changes the KPI and its denominator consistently"],
           ["S4", "Trend, comparison, payer mix, scatter with correlation, weekday heatmap visuals", "Each visual carries title, unit, period, source"],
           ["S5", "Descriptive statistics plus a documented correlation test with assumptions", "BR-10 acceptance test"],
           ["S6", "KPI definition panel and data-quality panel", "BR-05 and BR-14 acceptance tests"],
           ["S7", "Download of the filtered view for offline review", "File matches the on-screen period and filters"],
           ["S8", "Synthetic data generator, database build scripts and an automated test suite", "Regeneration reproducible from a documented seed"]],
          [0.45, 3.5, 2.6], font_size=8.4)

h2(doc, "6.2 Out of scope")
add_table(doc, ["ID", "Excluded", "Reason"],
          [["X1", "Live HIS/EHR integration, HL7/FHIR feed", "No hospital access; real data prohibited"],
           ["X2", "Patient-level clinical notes or any identifiable health information", "Privacy and ethics"],
           ["X3", "Predictive models (readmission risk, bed forecasting)", "Dilutes 18-step delivery; future enhancement"],
           ["X4", "Write operations: bed allocation, discharge entry, billing edits", "Decision-support product, not a transaction system"],
           ["X5", "Role-based access, SSO, audit logging", "Free-tier deployment; single passphrase maximum"],
           ["X6", "Multi-hospital / multi-tenant comparison", "Single-hospital model keeps definitions honest"],
           ["X7", "Mobile-native app, offline mode, scheduled email or push alerts", "Platform capability; BR-15 explicitly excluded"],
           ["X8", "Statutory reporting formats (GST, insurance files, government returns)", "Out of the six-KPI remit"]],
          [0.45, 3.5, 2.6], font_size=8.4)

# ---------------- 7 requirements ----------------
h1(doc, "7. Business Requirements", page_break=True)
para(doc, "Fifteen requirements, identifiers unchanged from Step 3 and Step 4. Priority: 9 Must, 3 Should, "
          "2 Could, 1 Won't. Acceptance criteria are the same tests carried in the RTM, abbreviated here for "
          "reading; the RTM remains the controlling copy.", space_after=4)

brs = [
    ("BR-01", "The business needs to see, for every department, its bed occupancy against the 85% target for the current and next week, with the shortfall expressed in beds and bed-days, so the operations manager can escalate a likely miss with a number before the 8 a.m. huddle.",
     "Escalation is currently verbal and late; a quantified gap converts an opinion into a staffing or bed decision.", "Must", "HMW-01",
     "8/8 departments for any chosen week; gap in beds to 0 dp; reproduces manual calculation \u00b10.1 pp on 3 sampled months."),
    ("BR-02", "The business needs ward-level occupancy projected across the next seven days, so the nursing superintendent can move staff before a ward crosses 90% instead of after handover.",
     "The roster is frozen a week ahead; without a forward view, saturation is managed by crisis transfer.", "Must", "HMW-02",
     "All wards for a 7-day horizon; every ward above 90% flagged; verified on 10 random dates."),
    ("BR-03", "The business needs any KPI movement between two periods decomposed into department, payer and weekday contributions, so the director can name the cause in the board meeting.",
     "Movement without attribution produces generic action; attribution converts a review into a decision.", "Must", "HMW-03",
     "Top-three contributors sum to total movement within \u00b10.5 pp on 3 sampled KPI-period pairs."),
    ("BR-04", "The business needs any operational question to be answerable in no more than three interactions (choose period, choose unit, read), so a first-time user can self-serve within 60 seconds.",
     "Adoption depends on effort; a five-step path returns the user to Excel.", "Must", "HMW-04",
     "8 users \u00d7 5 tasks: median \u2264 60 s, p90 \u2264 90 s, \u2264 3 interactions, \u2265 7/8 unaided success."),
    ("BR-05", "The business needs every published figure to carry its definition, denominator and data period in the same view, so two people never quote different numbers for the same week.",
     "This is the root fix for the definition conflict that made the whole project necessary.", "Must", "HMW-04",
     "100% of displayed measures carry definition + denominator + period; 0 exceptions."),
    ("BR-06", "The business needs length-of-stay comparison between departments on a like-for-like basis with case-mix context stated, so department heads accept the benchmark and commit to action.",
     "An unfair comparison is rejected by HODs and kills the monthly review before it starts.", "Must", "HMW-05",
     "ALOS reproduces Step 1 definition within \u00b10.01 day; normalisation basis stated; 3 HODs rate fairness \u2265 4/5."),
    ("BR-07", "The business needs the 30-day unplanned readmission rate by department and diagnosis group with the period's data-quality context visible, so the quality officer can separate clinical signal from coding error.",
     "Reporting a coding artefact as a clinical failure damages credibility and triggers unnecessary audits.", "Must", "HMW-06",
     "Matches manual recalculation on 2 random months (\u00b10.1 pp); window and exclusion rule stated; every spike shows \u2265 1 data-quality indicator."),
    ("BR-08", "The business needs payer-mix share, its movement between periods and revenue per patient, so finance can act within the same quarter rather than after the negotiation window closes.",
     "Payer mix moves slowly but a missed window costs a full negotiation cycle.", "Should", "HMW-07",
     "Shares sum to 100% \u00b10.1; movement shown for all 5 payer categories; reconciles to 20 sampled billing records within \u00b10.5%."),
    ("BR-09", "The business needs outpatient demand summarised by weekday and season with a peak-to-average ratio, so rosters and consultant slots are planned before the queue forms.",
     "Staffing is the largest controllable cost and the main driver of patient wait.", "Must", "HMW-08",
     "Peak weekday identified for 12 months; ratio reproduces manual calculation within \u00b12%."),
    ("BR-10", "The business needs a statistically defensible statement about whether occupancy pressure is associated with longer stays or higher readmissions, with method and assumptions disclosed.",
     "Leadership currently asserts that a full hospital is a risky hospital; the claim should be tested, not assumed.", "Could", "HMW-09",
     "Reports coefficient, p-value, n, method; assumptions stated; reproducible from the documented seed."),
    ("BR-11", "The business needs admissions and outpatient volume broken down by age band and gender for each department, so service-line plans reflect the population actually served.",
     "Service-line plans are currently built from anecdote about \u201cwho comes here\u201d.", "Could", "HMW-10",
     "Shares sum to 100% \u00b10.1 by age band and by gender; 3 departments spot-checked."),
    ("BR-12", "The business needs revenue per bed-day ranked across departments with the gap to the best performer, so capacity and service-line decisions are evidence-based rather than historical.",
     "Capacity that earns least per bed-day is invisible in raw revenue totals.", "Should", "HMW-11",
     "All 8 departments; ranking stable unless a \u2265 5% change is documented; gap reproduced within \u00b11%."),
    ("BR-13", "The business needs a three-month moving view of mortality and discharge-type mix by department with a drift indicator, so a quality review starts before the audit.",
     "Mortality drift found at audit is retrospective and reputational rather than corrective.", "Must", "HMW-12",
     "Flags points outside \u00b12 SD of a 12-month baseline; validated on injected drift; 0 missed injected drifts."),
    ("BR-14", "The business needs the data-quality status of each period (missing payer, duplicate registration, orphan department, date coverage, negative LOS, occupancy above 100%) shown beside the KPI.",
     "A suspect number must be challenged before it reaches a decision, not after.", "Should", "HMW-06",
     "6 checks return pass/fail with counts; injected defect fails the expected check; reconciles with the test suite."),
    ("BR-15", "The business needs threshold breaches pushed automatically to managers without anyone opening the tool, so a saturation or mortality-drift event is actioned within minutes.",
     "A daily routine still depends on someone logging in; push closes that gap — but the platform cannot support it in release 1.", "Won't", "HMW-01, HMW-12",
     "Not measured this release (X7, CN-02). Revisit criteria: alert \u2264 5 min, false positives \u2264 5%."),
]
rows = [[b[0], b[1], b[2], b[3], b[4], b[5]] for b in brs]
fills = {i: (MUST_FILL if b[3] == "Must" else WONT_FILL if b[3] == "Won't" else None)
         for i, b in enumerate(brs)}
fills = {k: v for k, v in fills.items() if v}
add_table(doc, ["ID", "Requirement (WHAT the business needs)", "Rationale", "Pri.", "Source HMW", "Acceptance criteria (abridged)"],
          rows, [0.5, 2.2, 1.5, 0.45, 0.65, 1.95], font_size=7.2, row_fill=fills)

h2(doc, "7.1 Requirement summary")
add_table(doc, ["Priority", "Count", "IDs"],
          [["Must", "9", "BR-01, 02, 03, 04, 05, 06, 07, 09, 13"],
           ["Should", "3", "BR-08, 12, 14"],
           ["Could", "2", "BR-10, 11"],
           ["Won't", "1", "BR-15"]],
          [0.9, 0.6, 4.9], font_size=8.5)
para(doc, "Traceability is held in 04_Traceability_Matrix.xlsx: 16 rows for 15 requirements because BR-15 links "
          "to two HMWs. This BRD does not restate traceability — it references it, so there is one controlling "
          "copy.", size=9.5, italic=True)

h2(doc, "7.2 Business-level quality expectations (formalised as NFRs in Step 6)")
add_table(doc, ["Expectation", "Business-level statement", "Linked SC"],
          [["Answer speed", "A routine answer is available inside a minute without analyst help", "SC-03, SC-04"],
           ["Consistency", "The same definition is used everywhere; the number never disagrees with itself", "SC-01, SC-08"],
           ["Comparability", "No ranking is shown unless it is normalised to a common unit", "BR-06, BR-12"],
           ["Transparency", "Every figure can be traced to period, denominator and data quality", "BR-05, BR-14"],
           ["Availability", "Reachable during working hours through a normal browser", "SC-09"],
           ["Usability", "Usable by non-technical managers without training", "SC-07"]],
          [1.25, 3.9, 1.45], font_size=8.4)

# ---------------- 8 business rules ----------------
h1(doc, "8. Business Rules", page_break=True)
para(doc, "Business rules are the calculation and interpretation laws of the domain. They are binding on every "
          "requirement and on the database design. Changing one is a change request, not an implementation "
          "detail.", space_after=4)
add_table(doc, ["Rule", "Statement", "Rationale", "Consumed by"],
          [["BRL-01", "Bed occupancy rate = bed-days occupied \u00f7 bed-days available \u00d7 100, computed on functional beds, monthly and weekly.",
            "Removes the sanctioned-vs-functional bed ambiguity; AS-03.", "BR-01, BR-02, BR-12"],
           ["BRL-02", "A bed-day is counted for each midnight a patient occupies a bed; an admission and discharge on the same calendar day counts as one bed-day.",
            "Prevents day-care cases from disappearing and keeps BOR \u2264 100%.", "BR-01, BR-02"],
           ["BRL-03", "Length of stay = discharge date \u2212 admission date in days; a same-day discharge counts as 1 day. ALOS = total inpatient days \u00f7 discharges.",
            "Removes the zero-day division anomaly and the 0 vs 1 inconsistency.", "BR-06"],
           ["BRL-04", "An unplanned readmission is any inpatient admission to the same department within 30 days of an earlier discharge, excluding admissions whose diagnosis group is on the planned-procedure list. The earlier discharge is the index discharge.",
            "Confirms AS-04 after quality review; prevents planned follow-up inflating the rate.", "BR-07"],
           ["BRL-05", "Revenue means gross billed amount per discharge, before deductions, write-offs, refunds and package discounts.",
            "Confirms OQ-12 decision; keeps revenue comparable across payers.", "BR-08, BR-12"],
           ["BRL-06", "Payer categories are a closed list of five: Cash/Self-pay, Private insurance, TPA, Government scheme, Corporate. Any other value is a data defect.",
            "A closed list makes payer mix sum to 100% and prevents long-tail noise.", "BR-08, BR-14"],
           ["BRL-07", "Outpatient footfall counts visits, not unique patients; a same-day revisit to a different department is a separate visit.",
            "Staffing decisions depend on visits, not people.", "BR-09"],
           ["BRL-08", "A week is ISO Monday–Sunday; a month is a calendar month. Comparisons use complete periods only; a partial current period is labelled \u201cin progress\u201d and excluded from ranking.",
            "Stops partial-month distortion in trend and ranking.", "BR-01, BR-03"],
           ["BRL-09", "Display rounding: percentages 1 dp, days 1 dp, currency to the nearest rupee with compact lakh/crore only on axis labels. Units are never mixed inside one visual.",
            "Consistency prevents apparent contradictions caused by rounding.", "All requirements"],
           ["BRL-10", "Any rate must show its denominator; if the denominator is below 30, the rate is marked low-confidence.",
            "Small denominators produce volatile percentages that look like events.", "BR-06, BR-07, BR-13"],
           ["BRL-11", "In-hospital mortality = deaths during the admission \u00f7 discharges, where discharge type = Expired counts exactly once. Transfers-out are not deaths.",
            "Prevents double counting and conflates no outcome with a transfer.", "BR-13"],
           ["BRL-12", "Rejected records are logged, never silently dropped: negative length of stay, discharge before admission, age outside 0–120, unknown department or payer.",
            "Silent dropping is the most common cause of \u201cmissing\u201d patients in hospital reports.", "BR-14"],
           ["BRL-13", "No cross-department ranking is displayed without normalisation (per bed-day or per discharge) and a stated period.",
            "Raw totals reward size, not performance.", "BR-06, BR-12"],
           ["BRL-14", "A KPI may not be displayed unless its definition exists in the definition registry with numerator, denominator and version.",
            "Implements BR-05 as a governance rule rather than a style choice.", "BR-05"]],
          [0.65, 3.0, 2.0, 1.0], font_size=7.4)

# ---------------- 9 data ----------------
h1(doc, "9. Data Requirements", page_break=True)
para(doc, "All data is synthetic, generated to the same grain as the six source concepts proven feasible in "
          "Step 1 §11.1. Physical table names (T-xx) are issued in Step 8; nothing here should be read as a "
          "schema.", space_after=4)
add_table(doc, ["Concept", "What it holds", "Source", "Refresh", "Owner", "Volume (target)"],
          [["Inpatient episode", "One row per admission-to-discharge episode: department, ward, dates, discharge type, diagnosis group, payer, patient key", "Synthetic generator (seeded)", "Static build, T+1 semantics", "SH-07 (design), student (build)", "\u2265 18,000 rows"],
           ["Bed capacity", "Department/ward-month bed base and bed-days available/occupied", "Synthetic generator derived from episode volume", "Monthly snapshot", "SH-06, SH-02", "8 depts \u00d7 36 months"],
           ["OPD visit", "One row per outpatient visit: date, department, age band, gender, payer", "Synthetic generator with weekday/season effects", "Daily grain", "SH-02", "\u2265 60,000 rows"],
           ["Finance", "One row per discharge billing: gross billed amount, payer, department, episode reference", "Derived from episodes with tariff logic", "Monthly", "SH-05", "1:1 with episodes"],
           ["Payer dimension", "Closed list of five payer categories", "Reference data", "On change only", "SH-05", "5 rows"],
           ["Department / ward dimension", "8 departments, wards and bed classes", "Reference data", "On change only", "SH-02", "8 depts, ~20 wards"],
           ["KPI definition registry", "One row per KPI: code, definition, numerator, denominator, version (BR-05)", "Authored, not generated", "On definition change", "SH-04", "6 core + 4 derived"],
           ["Data-quality log", "Rejected/excluded records with reason code (BRL-12)", "Build pipeline output", "Every rebuild", "SH-07", "Full audit of rejects"]],
          [1.2, 2.1, 1.35, 0.9, 0.75, 0.85], font_size=7.4)
para(doc, "Refresh expectation: the deployed product reads a static, versioned database built by the generator; "
          "there is no live feed (AS-06). The period filter therefore carries the latest available complete month, "
          "and a partial month is labelled per BRL-08.", size=9.5, italic=True)

# ---------------- 10 assumptions ----------------
h1(doc, "10. Assumptions, Constraints and Dependencies", page_break=True)
h2(doc, "10.1 Assumptions still in force")
add_table(doc, ["AS", "Assumption", "Status at v1.0"],
          [["AS-01", "Single 300-bed, eight-department hospital", "Confirmed"],
           ["AS-02", "One episode row = one admission to discharge including internal transfers", "Confirmed"],
           ["AS-03", "Occupancy on functional beds (BRL-01)", "Confirmed"],
           ["AS-04", "30-day unplanned readmission definition (BRL-04)", "Confirmed, wording refined"],
           ["AS-05", "Gross billed revenue (BRL-05)", "Upgraded to decision (OQ-12 closed)"],
           ["AS-06", "Daily refresh, not real time", "Confirmed"],
           ["AS-07", "Synthetic data reproduces seasonality and weekday effects", "Confirmed; validation in Step 10"],
           ["AS-08", "Users have spreadsheet literacy, no SQL", "Confirmed"],
           ["AS-09", "Dataset under 200 MB, SQLite-scale", "Confirmed"],
           ["AS-10", "Benchmarks are internal targets, not regulatory", "Confirmed"],
           ["AS-11", "English UI, Indian numbering conventions", "Confirmed"],
           ["AS-12", "ABES template governs the final document", "Confirmed"]],
          [0.7, 4.2, 1.7], font_size=8.4)

h2(doc, "10.2 Constraints")
add_table(doc, ["CN", "Constraint", "Impact on this BRD"],
          [["CN-01", "Semester timeline, one step every 2–3 days", "Baseline must be stable; late requirement changes will slip the deadline"],
           ["CN-02", "Free-tier hosting: ~1 GB RAM, app sleeps when idle", "No scheduler, so BR-15 cannot be satisfied"],
           ["CN-03", "Read-only deployed database", "No write-back requirements may be added"],
           ["CN-04", "No real patient data", "All evidence is synthetic; no hospital benchmark claims"],
           ["CN-05", "Python 3.11+ stack, limited libraries", "Rules expressed as computable definitions, not ML"],
           ["CN-06", "Tailwind-styled UI via CSS injection, must stay accessible", "Design must hold at 375 px and 1440 px"],
           ["CN-07", "Fixed mentor review slots", "Each baseline signed before the next step"],
           ["CN-08", "Single developer", "Test automation substitutes for a second reviewer"]],
          [0.7, 3.1, 2.8], font_size=8.4)

h2(doc, "10.3 Dependencies")
add_table(doc, ["DEP", "Dependency", "Needed by", "If it fails"],
          [["DEP-01", "Mentor approval of this BRD", "Step 6", "PRD cannot be baselined"],
           ["DEP-02", "Synthetic generator produces seasonality (AS-07)", "Step 10", "BR-09 and BR-10 become decorative"],
           ["DEP-03", "Free-tier deployment stays available", "Step 16", "Fallback: local demo with recorded walkthrough"],
           ["DEP-04", "KPI definition registry authored before build", "Step 15", "BR-05 cannot pass its acceptance test"],
           ["DEP-05", "8–10 usability participants available", "Step 15", "SC-03 and SC-07 unverifiable"],
           ["DEP-06", "ABES document and slide templates issued", "Steps 17–18", "Final formatting rework"]],
          [0.7, 3.0, 1.1, 1.8], font_size=8.4)

# ---------------- 11 risks ----------------
h1(doc, "11. Risks and Mitigation")
add_table(doc, ["R", "Risk", "L", "I", "Mitigation", "Owner"],
          [["R-01", "Occupancy exceeds 100% because bed base and occupied bed-days use different bed definitions", "M", "H", "BRL-01/BRL-02 frozen; automated check rejects BOR > 100 as a defect; validated before Step 10 closes", "Student"],
           ["R-02", "Readmission rate inflates because planned follow-ups are counted", "H", "H", "BRL-04 planned-procedure exclusion list; 30-day window computed on discharge date; reviewed by SH-04", "SH-04"],
           ["R-03", "Synthetic data too uniform, so trend and correlation views look artificial", "M", "M", "Seasonality and weekday effects mandated in the generator spec; tested with OQ-06 and OQ-07", "Student"],
           ["R-04", "Definition drift between BRD and product as steps progress", "M", "H", "Definition registry is the single source (BRL-14); RTM re-check at Steps 6, 12 and 15", "Student"],
           ["R-05", "Scope creep toward prediction or live integration", "H", "M", "X1–X8 contractual; changes require the §1.3 change note and mentor approval", "SH-10"],
           ["R-06", "Free-tier memory limit hit during aggregation", "L", "M", "Pre-aggregated views, caching, dataset under 200 MB (AS-09)", "Student"],
           ["R-07", "Small denominators create false alarms in department comparisons", "M", "M", "BRL-10 low-confidence marker below 30; drift rule uses 12-month baseline", "Student"],
           ["R-08", "Usability target missed because the first screen is too dense", "M", "M", "Progressive disclosure: summary first, drill-down second; 8-user test before deploy", "SH-02, SH-10"]],
          [0.55, 2.55, 0.4, 0.4, 2.1, 0.85], font_size=7.8)

# ---------------- 12 glossary ----------------
h1(doc, "12. Glossary", page_break=True)
add_table(doc, ["Term", "Meaning in this project"],
          [["Admission (episode)", "An inpatient stay from admission to discharge, including any internal ward transfer; the unit of the inpatient fact table."],
           ["ALOS", "Average length of stay; total inpatient days \u00f7 discharges (BRL-03)."],
           ["Bed-day", "One midnight of occupancy (BRL-02); the denominator unit for BOR."],
           ["BOR", "Bed occupancy rate; bed-days occupied \u00f7 bed-days available (BRL-01)."],
           ["Discharge type", "Outcome recorded at discharge: routine, transfer, LAMA (left against medical advice), Expired."],
           ["Functional bed", "A bed actually staffed and available for use; the base for occupancy in this project."],
           ["Index discharge", "The earlier discharge used as the time origin when testing for a readmission (BRL-04)."],
           ["LAMA", "Left against medical advice; excluded from planned readmission logic but counted in volume."],
           ["Low-confidence rate", "A rate computed on a denominator below 30, marked per BRL-10."],
           ["Must / Should / Could / Won't", "MoSCoW priority; Must items are the release-1 commitment."],
           ["OPD", "Outpatient department; footfall counts visits, not unique patients (BRL-07)."],
           ["Payer category", "One of five closed categories used for revenue and mix analysis (BRL-06)."],
           ["Readmission (30-day)", "Unplanned inpatient admission to the same department within 30 days of an index discharge (BRL-04)."],
           ["TBD", "Test Basis Document, produced at Step 7."],
           ["RTM", "Requirements Traceability Matrix; the controlling copy holds OQ → HMW → BR → later artefacts."],
           ["RACI", "Responsible, Accountable, Consulted, Informed."],
           ["Synthetic data", "Artificially generated records that mimic real distributions and seasonality; contains no real person."],
           ["T+1", "Data for day D is available on day D+1; the refresh expectation for this release."]],
          [1.5, 5.1], font_size=8.4)

# ---------------- 13 signoff ----------------
h1(doc, "13. Sign-off")
para(doc, "By signing, the reviewer confirms that the scope in §6, the requirements in §7 and the rules in §8 are "
          "complete and correct for release 1, and that any addition from this point follows the change control "
          "routine in §1.3 and triggers an RTM update.")
add_table(doc, ["Role", "Name / ID", "Signature", "Date"],
          [["Author (student)", STUDENT + " (" + ROLL + ")", "", "19-Nov-2026"],
           ["Mentor / approver", MENTOR + " (SH-10)", "", ""],
           ["Process owner", "Operations Manager (SH-02)", "", ""],
           ["Quality reviewer", "Quality / NABH Officer (SH-04)", "", ""]],
          [1.5, 2.6, 1.5, 1.0], font_size=8.6)
para(doc, "Next step: 06_PRD.docx — the product requirements that translate these business requirements into "
          "features, functional requirements and non-functional requirements, without introducing a new BR.",
     size=9.5, italic=True)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print("saved:", OUT)
