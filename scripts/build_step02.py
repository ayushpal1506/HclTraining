"""Build Step 2 deliverable: How-Might-We statements for p107."""
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
PRIORITY_FILL = "FFF2CC"

PROJECT_ID = "p107"
PROJECT_TITLE = "Healthcare KPI Dashboard"
STUDENT = "Ayush Pal"
ROLL = "2500320140037"
MENTOR = "Mr. Dinesh Sereno"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "02_HMW_Statements.docx")


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


def add_table(doc, headers, rows, widths, font_size=8.5, highlight=None):
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
        is_hi = highlight is not None and r_idx in highlight
        for i, v in enumerate(vals):
            cell = row.cells[i]
            cell.width = Inches(widths[i])
            cell.text = ""
            run = cell.paragraphs[0].add_run(str(v))
            run.font.size = Pt(font_size)
            if i == 0:
                run.bold = True
            if is_hi:
                shade(cell, PRIORITY_FILL)
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
    run = p.add_run(f"{PROJECT_ID}  |  Step 2 — How-Might-We Statements  |  Page ")
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
para(doc, "Department of Computer Applications (MCA) — Data Analytics", 10,
     color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=40)
para(doc, "STEP 2", 26, bold=True, color=NAVY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "HOW MIGHT WE STATEMENTS", 20, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
para(doc, PROJECT_TITLE, 17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Project ID: {PROJECT_ID}   |   Domain: Data Analytics", 11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=36)

meta = [("Student", STUDENT), ("Roll Number", ROLL),
        ("Mentor", MENTOR),
        ("Deliverable", "02_HMW_Statements.docx (Step 2 of 18)"),
        ("Version / Date", "v1.0 — 19 November 2026"),
        ("Inputs", "01_Project_Understanding.docx (v1.0 baseline)"),
        ("Status", "Baseline — approved for Step 3")]
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
para(doc, "All pain points below are derived from the synthetic hospital scenario defined in Step 1. "
          "No real patient, staff or hospital data is represented.", 8.5, italic=True, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER)

# ---------- control ----------
doc.add_page_break()
h1(doc, "Document Control")
h2(doc, "Revision history")
add_table(doc, ["Version", "Date", "Author", "Change description", "Status"],
          [["0.1", "19-Nov-2026", STUDENT, "Pain-point inventory built from Step 1 stakeholders and OQs", "Draft"],
           ["1.0", "19-Nov-2026", STUDENT, "12 HMWs finalised; OQ coverage verified 10/10; top 5 prioritised", "Baseline"]],
          [0.8, 1.0, 1.1, 3.1, 0.8])

h2(doc, "Purpose of this document")
para(doc, "Step 1 defined what the hospital must be able to ask. This document converts those questions "
          "into human-centred design problems — \u201cHow might we …?\u201d statements — so that Step 3 "
          "can derive business requirements from user intent rather than from a feature wish-list. An HMW "
          "is a bridge artefact: it keeps the user and the moment of struggle in view, while deliberately "
          "leaving the solution open.")
para(doc, "Working rule enforced throughout: an HMW names a user, a moment, and the decision at stake. "
          "It never names a chart, a screen, a library or a technology.", italic=True)

h2(doc, "Method")
bullet(doc, "Round 1 — Pain inventory: for every stakeholder SH-01 … SH-10, list 2–3 concrete pains observed in the Step 1 scenario, with a severity rating.")
bullet(doc, "Round 2 — Cluster: group pains that share a moment of decision (morning huddle, monthly review, audit prep, shift roster).")
bullet(doc, "Round 3 — Rewrite as HMW: one pain → one HMW, phrased as a question that starts with the user, not the artefact.")
bullet(doc, "Round 4 — Quality filter: reject any HMW that contains a solution word (dashboard, chart, report, app, SQL, AI...).")
bullet(doc, "Round 5 — Coverage test: every OQ-01 … OQ-10 must be served by at least one HMW; otherwise the HMW set is incomplete and Step 3 would silently lose a requirement.")
bullet(doc, "Round 6 — Prioritise: score shortlisted HMWs on decision value, frequency and effort; the top 5 are marked PRIORITY.")
para(doc, "Traceability note: OQ-11 … OQ-13 (open questions from Step 1 §12.2) are unresolved and are "
          "deliberately not converted into HMWs yet. They are closed in Steps 3, 12 and 16 respectively.",
     size=9.5, italic=True)

# ---------- 1 ----------
h1(doc, "1. Stakeholder Pain Points", page_break=True)
para(doc, "Severity: H = blocks a daily/weekly decision; M = causes rework or delay; L = irritation. "
          "Each pain is written as an observable behaviour (\u201crebuilds a pivot table\u201d), not a "
          "feeling (\u201cfinds it hard\u201d), so that it can later be turned into a measurable test.",
     space_after=4)
add_table(doc, ["SH ID", "Stakeholder", "Pain points (2–3 each)", "Sev.", "Where it bites"],
          [["SH-01", "Hospital Director",
            "P1 \u00b7 Sees only a month-end snapshot, so a bad week is discovered three weeks late. "
            "P2 \u00b7 Cannot attribute a KPI movement to a department, payer or weekday without asking an analyst. "
            "P3 \u00b7 Two reports quote different occupancy for the same week, so the number is distrusted in the board meeting.",
            "H", "Monthly board review"],
           ["SH-02", "Operations Manager",
            "P1 \u00b7 Spends ~45 minutes rebuilding a pivot table for each ad-hoc question. "
            "P2 \u00b7 Walks into the 8 a.m. huddle with yesterday's numbers, not next week's risk. "
            "P3 \u00b7 Has no early view of which department will miss the 85% occupancy target, so escalation is always reactive.",
            "H", "Daily 8 a.m. huddle"],
           ["SH-03", "Department Heads",
            "P1 \u00b7 Numbers are challenged in peer review because each HOD computed ALOS differently. "
            "P2 \u00b7 No like-for-like comparison — a complex cardiology case is compared with a day-care ENT case. "
            "P3 \u00b7 The monthly pack arrives after the review meeting, so it documents rather than drives action.",
            "H", "Monthly departmental review"],
           ["SH-04", "Quality / NABH Officer",
            "P1 \u00b7 Evidence for the audit is assembled by hand from four files over 2–3 days. "
            "P2 \u00b7 Cannot tell whether a readmission spike is clinical or a registration/coding error. "
            "P3 \u00b7 Mortality and readmission trends become visible only at audit time, so drift is reviewed retrospectively.",
            "H", "Monthly + pre-audit"],
           ["SH-05", "Finance / Revenue Cycle",
            "P1 \u00b7 Revenue per patient is buried in a billing export that finance reshapes every month. "
            "P2 \u00b7 A payer-mix shift is noticed a quarter late, after the negotiation window has closed. "
            "P3 \u00b7 No view connecting occupancy effort to revenue realised per bed-day.",
            "H", "Monthly close"],
           ["SH-06", "Nursing Superintendent",
            "P1 \u00b7 Ward occupancy is known only at shift handover, when the roster is already fixed. "
            "P2 \u00b7 Discharge load by weekday is invisible, so Sunday discharges pile up on Monday. "
            "P3 \u00b7 No warning when a ward is about to cross 90% and needs an extra shift.",
            "M", "Weekly roster + daily handover"],
           ["SH-07", "IT / Data Governance",
            "P1 \u00b7 There is no single data dictionary, so each report defines the same KPI differently. "
            "P2 \u00b7 Repeated ad-hoc extract requests arrive because no self-service view exists. "
            "P3 \u00b7 Reports do not record which definition and which period they used.",
            "M", "Ongoing data requests"],
           ["SH-08", "Registration Supervisor",
            "P1 \u00b7 No feedback loop: the team never sees how a missing payer field changes the revenue KPI. "
            "P2 \u00b7 Duplicate and mis-departmented registrations are found only at billing, weeks later. "
            "P3 \u00b7 Data-quality corrections are unplanned work with no visible ownership.",
            "M", "Daily registration desk"],
           ["SH-09", "Patient / Attendant",
            "P1 \u00b7 Waiting time varies by weekday with no way to plan the visit. "
            "P2 \u00b7 \u201cNo bed available\u201d is given without any explanation of why or when. "
            "P3 \u00b7 Repeated trips because bed availability is not visible before travelling.",
            "L", "Visit day"],
           ["SH-10", "Faculty Mentor / Panel",
            "P1 \u00b7 Cannot verify a claimed number quickly during viva. "
            "P2 \u00b7 Traceability gaps between requirement, screen and test case. "
            "P3 \u00b7 Risk that the artefact is a screenshot project rather than a reproducible build.",
            "H", "Reviews and viva"]],
          [0.55, 1.25, 4.05, 0.4, 1.05], font_size=7.8)

para(doc, "Pain-point counts: 3 + 3 + 3 + 3 + 3 + 3 + 3 + 3 + 3 + 3 = 30 distinct pains across 10 "
          "stakeholders. Ten of them (one per stakeholder, marked with the first P) are decision-blocking "
          "and become the backbone of the HMW set.", size=9.5, italic=True)

# ---------- 2 ----------
h1(doc, "2. From Pain to How Might We", page_break=True)
h2(doc, "2.1 Writing rules used")
add_table(doc, ["Rule", "Applied as"],
          [["R1 — Name the user", "Every HMW opens with \u201chelp the <role>\u201d, never \u201chelp the hospital\u201d."],
           ["R2 — Name the moment", "The statement pins a decision moment: huddle, review, audit prep, roster, negotiation."],
           ["R3 — No solution words", "Banned: dashboard, chart, report, app, Streamlit, SQL, AI, ML, export. Checked mechanically."],
           ["R4 — One problem per statement", "If an HMW contains \u201cand\u201d joining two decisions, it is split."],
           ["R5 — Answerable from data", "Each HMW must attach to at least one OQ and therefore to the data proven feasible in Step 1 §11."],
           ["R6 — Right altitude", "Broad enough for several design options, narrow enough to be testable in one screen."]],
          [1.5, 5.1], font_size=8.5)

h2(doc, "2.2 Rejected drafts and why (kept as evidence of the quality filter)")
add_table(doc, ["Rejected draft", "Why it fails", "Rule broken"],
          [["\u201cHow might we build a dashboard for hospital KPIs?\u201d",
            "It is the solution, not the problem; no user, no moment, no decision.", "R1, R2, R3"],
           ["\u201cHow might we improve hospital efficiency?\u201d",
            "Unbounded — every hospital activity qualifies, so nothing can be designed or tested.", "R6"],
           ["\u201cHow might we use AI to predict readmissions?\u201d",
            "Jumps to a technique that Step 1 explicitly placed out of scope (X3) and hides the user's real pain.", "R3, R5"],
           ["\u201cHow might we make our reports look more professional?\u201d",
            "Aesthetics are not a decision problem; it would generate decoration, not analysis.", "R2"],
           ["\u201cHow might we get better data from the HIS?\u201d",
            "Assumes live HIS access that does not exist in this project (X1); unactionable here.", "R5"]],
          [2.3, 3.6, 0.9], font_size=8.5)
para(doc, "The accepted pattern, by contrast, names the role, the moment and the decision: "
          "\u201cHow might we help the operations manager see, before the 8 a.m. huddle, which "
          "department is about to miss its occupancy target?\u201d", size=9.5, italic=True)

# ---------- 3 ----------
h1(doc, "3. How-Might-We Register", page_break=True)
para(doc, "Twelve statements, all within the 8–12 range. \u201cLinked OQ\u201d uses the identifiers frozen "
          "in Step 1; no OQ has been renumbered or reworded.", space_after=4)
add_table(doc, ["HMW ID", "How might we …?", "Stakeholder", "Pain point", "Linked OQ"],
          [["HMW-01",
            "How might we help the operations manager see, before the 8 a.m. huddle, which departments are about to miss the 85% occupancy target this week?",
            "SH-02 Operations Manager", "SH-02 P3 — no early view of a miss; escalation is reactive", "OQ-01"],
           ["HMW-02",
            "How might we help the nursing superintendent know, at roster time, which wards will cross 90% occupancy on a given shift and by how many beds?",
            "SH-06 Nursing Superintendent", "SH-06 P3 — no warning before a ward saturates", "OQ-01"],
           ["HMW-03",
            "How might we help the director explain a month-on-month KPI movement by naming the department, payer and weekday that caused it, without asking an analyst?",
            "SH-01 Hospital Director", "SH-01 P2 — cannot attribute a variance", "OQ-02"],
           ["HMW-04",
            "How might we help a first-time user get a trustworthy answer to one operational question in under a minute, without knowing table names or writing a query?",
            "SH-02, SH-10 (any evaluator)", "SH-02 P1 — 45 minutes per ad-hoc question", "OQ-02"],
           ["HMW-05",
            "How might we let a department head compare its length of stay with peer departments on a like-for-like case mix inside the review meeting?",
            "SH-03 Department Heads", "SH-03 P2 — no fair peer comparison", "OQ-03"],
           ["HMW-06",
            "How might we help the quality officer tell a genuine readmission spike apart from a registration or coding error before it is reported?",
            "SH-04 Quality / NABH", "SH-04 P2 — cannot separate clinical signal from data noise", "OQ-04"],
           ["HMW-07",
            "How might we help the finance manager see a payer-mix shift early enough to act within the same quarter rather than after the negotiation window?",
            "SH-05 Finance", "SH-05 P2 — payer shift noticed a quarter late", "OQ-05"],
           ["HMW-08",
            "How might we make weekday and seasonal outpatient peaks visible early enough that rosters and consultant slots are planned before the queue forms?",
            "SH-06, SH-02", "SH-06 P2 — discharge/OPD load by weekday is invisible", "OQ-06"],
           ["HMW-09",
            "How might we show leadership whether occupancy pressure is actually harming quality, instead of assuming that a full hospital is a risky hospital?",
            "SH-01, SH-04", "SH-01 P3 + SH-04 P2 — quality risk is asserted, not evidenced", "OQ-07"],
           ["HMW-10",
            "How might we surface which patient segments (age band, gender) a department actually serves, so service-line plans stop being guesswork?",
            "SH-03 Department Heads", "SH-03 P2 — planning without a patient-mix view", "OQ-08"],
           ["HMW-11",
            "How might we make revenue per bed-day comparable across departments so the weakest service line is visible without a finance analyst rebuilding the model?",
            "SH-05 Finance", "SH-05 P3 — occupancy effort is not linked to revenue per bed-day", "OQ-09"],
           ["HMW-12",
            "How might we help the quality officer notice a slow mortality drift months before the audit, not during it?",
            "SH-04 Quality / NABH", "SH-04 P3 — quality drift is reviewed retrospectively", "OQ-10"]],
          [0.62, 3.35, 1.05, 1.35, 0.55], font_size=7.8)

h2(doc, "3.1 Reading the register")
para(doc, "Distribution across stakeholders: SH-02 is served by HMW-01 and HMW-04; SH-04 by HMW-06 and "
          "HMW-12; SH-03 by HMW-05 and HMW-10; SH-05 by HMW-07 and HMW-11; SH-01 by HMW-03; SH-06 is "
          "co-served by HMW-02 and HMW-08. SH-07, SH-08, SH-09 and SH-10 receive no dedicated HMW: their "
          "pains (data dictionary, feedback loop, wait visibility, verifiability) are answered as constraints "
          "inside the solution rather than as separate design problems. This is a deliberate choice — see §6.")
para(doc, "Note on HMW-04: it is the only statement about effort rather than about a specific KPI. It "
          "exists because SC-03 (\u2264 60 s to answer any OQ) cannot be achieved by improving individual "
          "screens if the interaction model itself is wrong.", size=9.5, italic=True)

# ---------- 4 ----------
h1(doc, "4. Priority HMWs", page_break=True)
para(doc, "Scoring: Decision value (1–5) × Frequency (1–5) ÷ Effort (1–5), evaluated against "
          "stakeholder influence from Step 1 §5. The top 5 are marked PRIORITY and will be the first "
          "candidates for business requirements in Step 3.", space_after=4)
add_table(doc, ["Rank", "HMW ID", "Why it is priority (one line)", "Value", "Freq.", "Effort", "Score"],
          [["1", "HMW-01", "Serves the primary daily user at the moment a decision is actually made, and OQ-01 is the only KPI with a hard target (85%).", "5", "5", "2", "12.5"],
           ["2", "HMW-04", "Sets the adoption ceiling: if a first-time user cannot answer in 60 s, none of the other HMWs matter (SC-03).", "5", "5", "3", "8.3"],
           ["3", "HMW-05", "Unlocks department-head buy-in; without a fair like-for-like comparison, HODs reject the numbers and the tool loses authority.", "5", "3", "3", "5.0"],
           ["4", "HMW-06", "Turns a compliance report into a decision; readmission is the one KPI that carries clinical and reputational risk together.", "4", "3", "3", "4.0"],
           ["5", "HMW-08", "Highest-actionability staffing decision: weekday peaks drive roster cost and patient wait simultaneously.", "4", "4", "4", "4.0"]],
          [0.45, 0.65, 3.35, 0.5, 0.45, 0.5, 0.5], font_size=8.2, highlight={0, 1, 2, 3, 4})
para(doc, "HMW-06 and HMW-08 tie at 4.0; HMW-06 is placed above HMW-08 because its stakeholder (SH-04) "
          "has audit authority and a wrong readmission answer is a credibility event, not just an "
          "inconvenience. All five priority HMWs keep their OQ link from §3; no identifier changed.",
     size=9.5, italic=True)

h2(doc, "4.1 Deferred HMWs (kept in scope, sequenced later)")
add_table(doc, ["HMW ID", "Why it is not in the top 5", "When it lands"],
          [["HMW-02", "Ward-level roster detail depends on the same OQ-01 model as HMW-01; build the department view first, then disaggregate.", "Step 12 (UI-02 ward tab)"],
           ["HMW-03", "Variance attribution needs a working trend view (HMW-01 output) as its input.", "After HMW-01"],
           ["HMW-07", "Payer mix is monthly cadence with medium influence; valuable but not decision-blocking daily.", "Step 15 backlog"],
           ["HMW-09", "Correlation analysis is a depth feature; it depends on stable aggregates from HMW-01/05/06.", "Step 15 backlog"],
           ["HMW-10", "Patient-mix segmentation adds context but no immediate action for the six core KPIs.", "Step 15 backlog"],
           ["HMW-11", "Revenue per bed-day is a derived KPI requiring OQ-05 and OQ-01 to be correct first.", "After HMW-07"],
           ["HMW-12", "Mortality drift is low-frequency and sensitive; it needs a careful presentation before it is shown to leadership.", "Step 15 backlog"]],
          [0.7, 4.1, 1.8], font_size=8.5)

# ---------- 5 ----------
h1(doc, "5. Coverage Check — Every OQ Served by a HMW", page_break=True)
para(doc, "This is the Step 2 quality gate. A tick means at least one HMW explicitly addresses that "
          "operational question. No OQ may have an empty row; no HMW may link to zero OQs.",
     space_after=4)
hmws = [f"HMW-{i:02d}" for i in range(1, 13)]
cover = {
    "OQ-01": ["HMW-01", "HMW-02"],
    "OQ-02": ["HMW-03", "HMW-04"],
    "OQ-03": ["HMW-05"],
    "OQ-04": ["HMW-06"],
    "OQ-05": ["HMW-07"],
    "OQ-06": ["HMW-08"],
    "OQ-07": ["HMW-09"],
    "OQ-08": ["HMW-10"],
    "OQ-09": ["HMW-11"],
    "OQ-10": ["HMW-12"],
}
rows = []
for oq in [f"OQ-{i:02d}" for i in range(1, 11)]:
    rows.append([oq] + ["\u2713" if h in cover[oq] else "" for h in hmws] +
                [str(len(cover[oq]))])
add_table(doc, ["OQ"] + [h.replace("HMW-", "H") for h in hmws] + ["#HMW"],
          rows, [0.55] + [0.42] * 12 + [0.5], font_size=7.5)
para(doc, "Result: 10 of 10 OQs covered; 12 of 12 HMWs link to exactly one OQ; 2 OQs (OQ-01, OQ-02) are "
          "served twice, which is intentional — they are the highest-frequency decisions. Gate verdict: "
          "PASS.", bold=True)
para(doc, "Deliberate non-coverage: OQ-11, OQ-12 and OQ-13 are still open questions from Step 1, not "
          "operational questions, so they carry no HMW. They are tracked to Steps 3, 12 and 16.",
     size=9.5, italic=True)

# ---------- 6 ----------
h1(doc, "6. What This Means for Step 3", page_break=False)
para(doc, "The HMW set implies five design commitments that will shape the business requirements:" )
bullet(doc, "C1 — A default landing view must answer HMW-01 without any filtering, because the primary user has 10–15 minutes.")
bullet(doc, "C2 — Every KPI tile must carry its definition, period and denominator, closing HMW-06's signal-versus-noise gap and SH-07's dictionary pain.")
bullet(doc, "C3 — Every number must be drillable to its driver (department, payer, weekday), not just displayed (HMW-03, HMW-11).")
bullet(doc, "C4 — Comparison must be normalised (per bed-day, per discharge, per case mix) before any ranking is shown (HMW-05, HMW-11).")
bullet(doc, "C5 — Data quality must be visible as a first-class panel, so SH-04 and SH-08 can see the effect of a missing field (HMW-06, SH-08 P1).")
para(doc, "Explicit non-commitment: no HMW asks for prediction, alerting, live HIS integration or role-based access. "
          "Those stay out of scope per Step 1 §8.2 (X1, X3, X5).", italic=True, space_after=10)

h1(doc, "7. Step 2 Sign-off")
bullet(doc, "30 pain points documented across 10 stakeholders (SH-01 … SH-10).")
bullet(doc, "12 HMW statements issued (HMW-01 … HMW-12); 5 prioritised, 7 deferred with reasons.")
bullet(doc, "OQ coverage verified: 10/10, gate PASS.")
bullet(doc, "No Step 1 identifier renumbered or reworded; OQ-11 … OQ-13 remain open and owned.")
para(doc, "Prepared by: " + STUDENT + "   |   Roll: " + ROLL, size=10)
para(doc, "Reviewed by (mentor): " + MENTOR + "   |   Date: 19-Nov-2026", size=10)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print("saved:", OUT)
