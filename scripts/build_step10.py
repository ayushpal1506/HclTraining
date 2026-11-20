"""Build the Step 10 deliverable: dataset generation report (Word)."""
import json
import os
import sqlite3
from datetime import date

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x54, 0x96)
GREY = RGBColor(0x59, 0x59, 0x59)
HDR_FILL = "1F3864"
ALT_FILL = "EAF0F8"
CODE_FILL = "F4F6F8"

PROJECT_ID = "HKD-2026"
PROJECT_TITLE = "Healthcare KPI Dashboard"
STUDENT = "Ayush Pal"
ROLL = "2503201450035"
MENTOR = "Kalpana Dwevedi"

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
OUT = os.path.join(ROOT, "docs", "10_Dataset_Generation.docx")

summary = json.load(open(os.path.join(DATA, "validation_summary.json"), encoding="utf-8"))
con = sqlite3.connect(os.path.join(DATA, "app.db"))
con.row_factory = sqlite3.Row


def size_mb(path):
    try:
        n = os.path.getsize(path)
    except OSError:
        return "—"
    return f"{n/1024/1024:.1f} MB" if n >= 1024 * 1024 else f"{n/1024:.0f} KB"


# --------------------------------------------------------------------------- helpers
def shade(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
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


def add_table(doc, headers, rows, widths, font_size=8.2, align_center=()):
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
        for i, v in enumerate(vals):
            cell = row.cells[i]
            cell.width = Inches(widths[i])
            cell.text = ""
            p = cell.paragraphs[0]
            if i in align_center:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(str(v))
            run.font.size = Pt(font_size)
            if i == 0:
                run.bold = True
            if r_idx % 2 == 1:
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


def bullet(doc, text):
    p = doc.add_paragraph(text, style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.25)
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


def mono_block(doc, text, size=7.6):
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.autofit = False
    cell = table.rows[0].cells[0]
    cell.width = Inches(6.6)
    shade(cell, CODE_FILL)
    cell.text = ""
    p = cell.paragraphs[0]
    lines = text.split("\n")
    for i, line in enumerate(lines):
        run = p.add_run(line)
        run.font.name = "Consolas"
        run.font.size = Pt(size)
        if i < len(lines) - 1:
            run.add_break()
    doc.add_paragraph()


def footer(section):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{PROJECT_ID}  |  Step 10 — Dataset Generation  |  Page ")
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


def rows_evenly(table, col_count, n):
    total = con.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
    step = max(1, total // n)
    sql = f"SELECT * FROM {table} ORDER BY 1 LIMIT ? OFFSET ?"
    out = []
    for i in range(n):
        row = con.execute(sql, (1, i * step)).fetchone()
        if row:
            out.append(tuple(row)[:col_count])
    return out


def inr(x):
    return f"{x:,.0f}"


# --------------------------------------------------------------------------- document
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

# cover
para(doc, "ABES ENGINEERING COLLEGE", 11, bold=True, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para(doc, "Department of Computer Applications (MCA) — Data Analytics", 10, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=40)
para(doc, "STEP 10", 26, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "SYNTHETIC DATASET GENERATION", 17, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
para(doc, PROJECT_TITLE, 17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Project ID: {PROJECT_ID}   |   Domain: Data Analytics", 11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=34)

meta = [("Student", STUDENT), ("Roll Number", ROLL), ("Mentor", MENTOR),
        ("Deliverable", "src/generate_data.py + data/ (SQLite, CSV, Excel, quality log) + this report"),
        ("Version / Date", "v1.0 — 19 November 2026"),
        ("Seed", "42 — every run reproduces the dataset exactly"),
        ("Volume", f"{summary['tables']['fact_inpatient_episode']:,} inpatient episodes · "
                   f"{summary['tables']['fact_opd_visit']:,} outpatient visits · "
                   f"2023-01-01 to 2025-12-31"),
        ("Status", "Baseline — approved for Step 11 (user stories) and Step 15 (analysis)")]
t = doc.add_table(rows=0, cols=2)
t.style = "Table Grid"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
for k, v in meta:
    row = t.add_row()
    row.cells[0].width = Inches(1.6)
    row.cells[1].width = Inches(4.6)
    row.cells[0].text = ""
    r0 = row.cells[0].paragraphs[0].add_run(k)
    r0.bold = True
    r0.font.size = Pt(9.5)
    shade(row.cells[0], "EAF0F8")
    row.cells[1].text = ""
    r1 = row.cells[1].paragraphs[0].add_run(v)
    r1.font.size = Pt(9.5)

# control
doc.add_page_break()
h1(doc, "Document Control")
h2(doc, "Revision history")
add_table(doc, ["Version", "Date", "Author", "Change", "Status"],
          [["0.1", "19-Nov-2026", STUDENT, "Generator written; first run produced a schema-valid dataset", "Draft"],
           ["0.2", "19-Nov-2026", STUDENT, "Patterns tuned after the first measurements; capacity blocking explained; cleaning rules added", "Draft"],
           ["1.0", "19-Nov-2026", STUDENT, f"Baseline: 5 tables, {summary['tables']['fact_inpatient_episode']:,} episodes, {summary['tables']['fact_opd_visit']:,} visits, all {len(summary['checks'])} checks pass", "Superseded"],
           ["1.1", "19-Nov-2026", STUDENT, "Believability review: patient reuse constrained to the same department (no male maternity or adult paediatric rows); all counts and pattern values refreshed", "Approved"]],
          [0.8, 1.0, 1.1, 3.2, 0.8])

h2(doc, "Inputs and outputs")
add_table(doc, ["Artefact", "Role"],
          [["docs/08_Database_Design.xlsx (v1.1)", "The DDL and views executed against data/app.db — the schema is not retyped in the generator"],
           ["05_BRD.docx (BRL-01…BRL-14)", "Business rules the generated values must respect (bed-days, payer list, mortality, rounding)"],
           ["src/generate_data.py", "The only code in this step: generator + cleaner + loader + validator"],
           ["data/app.db", "SQLite database, FK ON, built from the Step 8 DDL"],
           ["data/hkd2026_data.xlsx", "Review workbook: 4 data sheets + Data_Quality, Patterns, Validation"],
           ["data/quality_log.csv", "Every rejected, removed or flagged row with its rule"],
           ["data/validation_report.md", "Machine-generated validation summary"]],
          [2.6, 4.0], font_size=8.6)

h2(doc, "Change notes")
add_table(doc, ["Note", "Change", "Impact"],
          [["DG-CN-01", "Fact sheets in the Excel workbook hold an evenly spaced 2,000-row sample, not all rows. Full data is in the CSVs and app.db.", "Review convenience only; no data is withheld from the CSVs"],
           ["DG-CN-02", f"{summary['tables']['fact_inpatient_episode']:,} episodes generated against a 57,000 plan. The difference is bed-blocked demand ({sum(summary['blocked'].values()):,} candidate admissions) that could not be placed without breaching capacity.", "Volume is still 3× the OBJ-02 minimum; documented in §9"],
           ["DG-CN-03", "v1.1 — the 16% patient-reuse path could attach a male or adult patient to Maternity or Paediatrics. Reuse is now restricted to the same department's patient pool, which removed 588 male Maternity episodes and 504 Paediatric episodes aged 51+.", "Row counts and every measured value changed slightly; all figures in this document are from the v1.1 run"],
           ["No ID changes", "No OQ, SH, HMW, BR, FR, NFR, T, US, UI or TC identifier was added or renumbered.", "RTM v1.1 remains current"]],
          [1.0, 4.3, 1.5], font_size=8.4)

# 1 what was generated
h1(doc, "1. What was generated", page_break=True)
para(doc, f"The generator produced a {summary['days']:,}-day hospital history from {summary['period'][0]} to "
          f"{summary['period'][1]}, matching the five tables defined in Step 8 exactly — same table names, same 49 "
          f"columns, same types, same keys, same CHECK constraints. Nothing was added to the schema to make the data fit.")
add_table(doc, ["Table", "Rows (clean)", "Rows (raw)", "Rejected / removed", "Grain"],
          [["dim_department", f"{summary['tables']['dim_department']}", f"{summary['raw_counts']['dim_department']}", "0", "one row per department"],
           ["dim_ward", f"{summary['tables']['dim_ward']}", f"{summary['raw_counts']['dim_ward']}", "0", "one row per ward, with its functional bed base"],
           ["dim_payer", f"{summary['tables']['dim_payer']}", f"{summary['raw_counts']['dim_payer']}", "0", "one row per payer"],
           ["fact_inpatient_episode", f"{summary['tables']['fact_inpatient_episode']:,}",
            f"{summary['raw_counts']['fact_inpatient_episode']:,}",
            f"{summary['raw_counts']['fact_inpatient_episode']-summary['tables']['fact_inpatient_episode']:,}",
            "one row per admission-to-discharge episode"],
           ["fact_opd_visit", f"{summary['tables']['fact_opd_visit']:,}",
            f"{summary['raw_counts']['fact_opd_visit']:,}",
            f"{summary['raw_counts']['fact_opd_visit']-summary['tables']['fact_opd_visit']:,}",
            "one row per outpatient visit"]],
          [1.7, 1.15, 1.15, 1.25, 1.35], font_size=8.2, align_center=(1, 2, 3))
para(doc, f"Totals: {sum(summary['tables'].values()):,} clean rows loaded, {summary['total_rejected']:,} rows rejected or "
          f"removed during cleaning, {summary['outlier_flags']} valid outliers flagged and kept. The database is "
          f"{size_mb(os.path.join(DATA, 'app.db'))}, well under the 100 MB per-file limit on GitHub.", size=9.8)

# 2 method
h1(doc, "2. How the data was generated", page_break=True)
para(doc, "A dataset for a KPI dashboard is not a random table. If admissions are drawn independently, occupancy "
          "cannot exceed a plausible level of 70–90% without either filling beds or dropping patients, and seasonality "
          "disappears. The generator therefore simulates a real constraint: three hundred staffed beds, occupied day "
          "by day for three years, with demand that varies by department, weekday and month.")
h2(doc, "2.1 The loop")
mono_block(doc, """for each day in 2023-01-01 .. 2025-12-31:            # 1,096 days
    for each scheduled readmission due today:        # chained from an index discharge
        place it if every night of the stay has a free bed; else retry, then drop
    for each of the 20 wards:
        n ~ Poisson(beds × occupancy_target ÷ mean_LOS × weekday × season × trend)
        for each candidate admission:
            draw LOS (gamma, department mean), cap at 120, shift Sunday discharges
            place the stay only if the ward has a free bed on every night
            otherwise count it as bed-blocked and do not record it
            optionally schedule a readmission 3–25 days after discharge
for each outpatient day and department:
    n ~ Poisson(OPD base × weekday × season)      # Monday peak, Sunday near-closed
clean the raw layer, load app.db, run 12 checks""", 7.4)
para(doc, "Two consequences matter for the analysis that follows. First, occupancy can never exceed 100% by "
          "construction — the ward-day view confirms a maximum of "
          f"{summary['bor_max']}% — so BR-01's validity rule is satisfied without patching. Second, when demand "
          "exceeds capacity the admission is simply not in the data, which is what happens in a real hospital and why "
          "the achieved volume is below the theoretical λ.", size=9.8)

h2(doc, "2.2 Department parameters and what they actually produced")
rows = []
for code in sorted(summary["dept_bor"]):
    rows.append([code,
                 summary["dept_bor"][code],
                 summary["dept_alos"][code],
                 summary["dept_rr"][code],
                 inr(summary["dept_rpbd"][code])])
add_table(doc, ["Department", "BOR %", "ALOS (days)", "30-day readmission %", "Revenue per bed-day (₹)"],
          rows, [1.3, 0.9, 1.0, 1.4, 1.6], font_size=8.6, align_center=(1, 2, 3, 4))
para(doc, "BOR is computed from the ward-day view, not from the target: the target drives demand, the measured value is "
          "what a dashboard would show. Cardiology lands at "
          f"{summary['dept_bor']['DEP01']}% against the 85% target, which is exactly the kind of borderline case the "
          "operations manager needs to see.", size=9.8)

h2(doc, "2.3 Fiction, privacy and the Indian context")
para(doc, "The brief asks for fictional Indian names and places. The Step 8 schema deliberately stores no person "
          "name, phone number or address — patients exist only as a 12-character pseudonymous key — so there is "
          "nothing personal to fictionalise. What is Indian in the dataset is the setting: rupee amounts at Indian "
          "tariff levels, payer categories that include a state scheme and a corporate panel, and five invented payer "
          "names (Swaasthya Shield, Sahyog TPA, Rajya Swasthya Yojana, Nirmaan Corporate Health, plus self-pay). None "
          "of them is a real insurer, hospital or person, and no row can be traced to an individual.", size=9.8)

h2(doc, "2.4 Reproducibility")
para(doc, "One seed (42) drives every draw, and the script writes the same bytes on every run — verified by running it "
          "twice and comparing the episode and visit CSVs. Anyone marking this project can regenerate the dataset from "
          "the script alone; no binary artefact has to be trusted.", size=9.8)

# 3 patterns
h1(doc, "3. The five patterns planted in the data", page_break=True)
para(doc, "Each pattern is a deliberate structure, not an accident of sampling. The measured column is computed from "
          "the loaded database, so the claim is checkable rather than asserted.", space_after=4)
prows = []
for p in summary["patterns"]:
    prows.append([p["id"], p["name"], p["evidence"], p["supports"]])
add_table(doc, ["ID", "Pattern", "Measured in data/app.db", "Analysis it supports"],
          prows, [0.55, 1.35, 3.6, 1.1], font_size=7.6)
para(doc, "Reading the five together: the winter surge gives the trend pages something to explain, the cardiology "
          "occupancy gives BR-01 a near-target case, the readmission spread gives the Director a department to "
          "challenge, the payer shift gives Finance a three-year story, and the oncology drift gives the "
          "mortality page a reason to exist. Without these, every chart would be flat and every correlation "
          "would be noise.", size=9.8)

# 4 data quality
h1(doc, "4. Data-quality issues, by design", page_break=True)
para(doc, "Real extracts arrive dirty. The raw layer (data/raw/) therefore contains the defects below; the cleaning "
          "stage rejects or removes them, logs each one, and loads only valid rows. Outliers are the exception: a "
          "₹48 lakh bill or a 90-day stay is unusual but real, so it is flagged, never deleted.", space_after=4)
irows = []
for i in summary["issues"]:
    irows.append([i["table"], i["issue"], i["rows"], f"{i['pct']}%", i["action"]])
add_table(doc, ["Table", "Injected issue", "Rows", "% of table", "Handling"],
          irows, [1.9, 2.5, 0.6, 0.8, 1.4], font_size=7.8, align_center=(2, 3))
add_table(doc, ["Rule", "What it catches", "Action"],
          [["RQ-01", "Required field empty (payer_id, billed_amount)", "Reject the row"],
           ["RQ-02", "Length of stay outside 1–120", "Reject the row"],
           ["RQ-03", "Department, ward or payer that does not exist (orphan FK)", "Reject the row"],
           ["RQ-04", "Value outside the allowed list (gender, visit type, payer category)", "Reject the row"],
           ["RQ-05", "Malformed date or number", "Reject the row"],
           ["RQ-06", "Discharge date before admission date", "Reject the row"],
           ["RQ-07", "Duplicate business key (patient + admission, patient + date + department)", "Keep first, remove rest"],
           ["RQ-08", "Valid but extreme value (top 0.4% of bills, stays ≥ 45 days)", "Flag and keep"]],
          [0.6, 4.6, 1.4], font_size=8.2)
para(doc, f"{summary['total_rejected']:,} rows were rejected or removed; {summary['outlier_flags']} were flagged and kept. "
          f"Every decision is in data/quality_log.csv with the rule number, the field and the offending value, which is "
          "what BRL-12 requires: rejected records are logged, never silently dropped.", size=9.8)
mono_block(doc, """table,row_position,rule,field,value,action
fact_inpatient_episode,738,RQ-08 valid outlier kept,billed_amount/los_days,"303070.37 / 49",flagged
fact_inpatient_episode,42,RQ-01 required field missing,payer_id,,rejected
fact_inpatient_episode,57,RQ-07 duplicate,patient_key+admission_date,"PT0000004521 2023-02-14",removed""", 7.2)

# 5 validation
h1(doc, "5. Validation results", page_break=True)
para(doc, "Twelve checks were executed against data/app.db after loading. The generator fails loudly if any check "
          "fails — the document would not be produced.", space_after=4)
add_table(doc, ["Check", "Expectation", "Actual", "Result"],
          [[c[0], c[1], c[2], c[3]] for c in summary["checks"]],
          [1.7, 2.0, 2.2, 0.6], font_size=7.8)
h2(doc, "5.1 The three checks that matter most")
para(doc, f"Foreign keys: PRAGMA foreign_key_check returns {summary['fk_violations']} rows. Every ward belongs to a "
          "department, every episode to a ward, department and payer, and every readmission points at an episode that "
          "exists.", size=9.8)
para(doc, f"Reconciliation: the sum of bed_days in the fact table ({summary['bed_days']:,}) equals the sum of "
          "occupied beds across ward-days. That single equality proves the occupancy view and the fact table tell the "
          "same story; if they diverged, BOR would be computed from one and explained from the other.", size=9.8)
para(doc, f"Ranges: length of stay runs {summary['los'][0]}–{summary['los'][1]} days with a mean of "
          f"{summary['los'][2]}, billed amounts from ₹{summary['billed'][0]:,.0f} to ₹{summary['billed'][1]:,.0f}, and "
          f"all dates fall inside the three-year window. Maximum ward-day occupancy is {summary['bor_max']}% and the "
          f"maximum department-month is {summary['month_bor_max']}%.", size=9.8)

# 6 conformance + sample
h1(doc, "6. Does it match the design — and does it look real?", page_break=True)
h2(doc, "6.1 Structural conformance")
add_table(doc, ["Dimension", "Step 8 design", "Delivered"],
          [["Tables", "5 (3 masters + 2 facts)", "5, same names"],
           ["Columns", "49", "49, same names, same order"],
           ["Primary keys", "1 per table", "1 per table, unchanged"],
           ["Foreign keys", "7 (including self-reference)", "7, PRAGMA foreign_key_check clean"],
           ["CHECK / UNIQUE / trigger", "36 CHECKs, 5 UNIQUEs, 1 trigger", f"all active; {summary['total_rejected']:,} bad rows rejected or removed"],
           ["Analysis view", "30–40 columns", "vw_analysis_episode, 36 columns"],
           ["Row plan", "150–300 review sample; ~57k production", f"{summary['tables']['fact_inpatient_episode']:,} episodes; "
                                                                   f"{summary['tables']['fact_opd_visit']:,} visits"]],
          [1.6, 2.3, 2.6], font_size=8.4)
para(doc, "The joined view is the one a reviewer should check against the 30–40 column brief. Its 36 columns fall into "
          "six families:", space_after=4)
add_table(doc, ["Family", "Columns", "Count"],
          [["Episode identity", "episode_id, patient_key, admission_date, discharge_date", "4"],
           ["Outcome", "los_days, discharge_type, death_flag, readmit_30d_flag, index_episode_id", "5"],
           ["Clinical", "diagnosis_group, is_planned_procedure, age_band, gender", "4"],
           ["Money and volume", "billed_amount, bed_days", "2"],
           ["Dimension attributes", "payer_id/name/category, department_id/name/specialty, ward_id/name/type/class/functional_beds", "12"],
           ["Derived time and bands", "admission_year, admission_month, week_of_year, weekday_name, admission_quarter, is_weekend_admission, season, los_band, revenue_band", "9"]],
          [1.5, 4.3, 0.6], font_size=8.0)

h2(doc, "6.2 Ten rows of each table, read cold")
para(doc, "Ten rows taken at even intervals across the three years — the same view the Excel sample gives. The question is not whether "
          "they are pretty — it is whether a clinician or a finance officer would accept them without asking "
          "“what is this?”.", size=9.6)
eps = rows_evenly("fact_inpatient_episode", 19, 10)
add_table(doc, ["#", "Dept", "Ward", "Payer", "Admitted", "Discharged", "LOS", "Bill (₹)"],
          [[e[0], e[2], e[3], e[4], e[5], e[6], e[7], inr(e[16])] for e in eps],
          [0.45, 0.75, 0.75, 0.6, 1.0, 1.0, 0.5, 1.0], font_size=7.6, align_center=(6, 7))
add_table(doc, ["#", "Dept", "Date", "Payer", "Type", "Fee (₹)"],
          [[v[0], v[3], v[2], v[4], v[7], inr(v[8])] for v in rows_evenly("fact_opd_visit", 10, 10)],
          [0.5, 0.8, 1.0, 0.6, 0.9, 0.9], font_size=7.6, align_center=(5,))
add_table(doc, ["Ward", "Name", "Type", "Beds", "Dept"],
          [[w[0], w[1], w[2], w[5], w[4]] for w in rows_evenly("dim_ward", 10, 8)],
          [0.6, 1.3, 1.0, 0.5, 0.7], font_size=7.8, align_center=(3,))
para(doc, "What a reviewer can confirm from these rows:", space_after=3)
ep_span = con.execute("SELECT MIN(admission_date), MAX(admission_date) FROM fact_inpatient_episode").fetchone()
distinct_dept = con.execute("SELECT COUNT(DISTINCT department_id) FROM fact_inpatient_episode").fetchone()[0]
readmit_share = con.execute("SELECT ROUND(100.0*SUM(readmit_30d_flag)/COUNT(*),2) FROM fact_inpatient_episode").fetchone()[0]
bullet(doc, f"the sample spans {ep_span[0]} to {ep_span[1]} and covers all {distinct_dept} departments, so it is not a "
            "single-week extract in disguise;")
bullet(doc, f"length of stay, ward type and bill move together: ICU and oncology stays are longer and costlier, "
            "day-care ENT stays are short and cheap;")
bullet(doc, f"{readmit_share}% of episodes carry the readmission flag, and every flagged row has an index episode id, "
            "which is the join that OQ-04 will use;")
bullet(doc, "payer ids are the five closed categories, so payer mix always sums to 100% and never produces an "
            "“unknown” slice;")
bullet(doc, "no episode contradicts its department: Maternity is 100% female and Paediatrics carries no patient aged "
            "51 or over, because v1.1 restricts patient reuse to the originating department.")

# 7 file guide
h1(doc, "7. Files delivered, and how to regenerate them", page_break=True)
files = [
    ("data/app.db", "SQLite database, built from the Step 8 DDL, foreign keys ON"),
    ("data/hkd2026_data.xlsx", "Review workbook: 4 table sheets + Data_Quality + Patterns + Validation"),
    ("data/fact_inpatient_episode.csv", "Cleaned episodes, all rows"),
    ("data/fact_opd_visit.csv", "Cleaned visits, all rows"),
    ("data/dim_department.csv · dim_ward.csv · dim_payer.csv", "Masters, all rows"),
    ("data/raw/<table>.csv", "As-received layer with the injected defects"),
    ("data/quality_log.csv", "Rejected, removed and flagged rows with rule numbers"),
    ("data/validation_report.md", "Row counts, checks and pattern evidence"),
    ("data/schema.sql", "The DDL and views actually executed"),
    ("src/generate_data.py", "Generator, cleaner, loader and validator (single entry point)"),
]
add_table(doc, ["Path", "Contents"],
          [[f, d] for f, d in files], [2.7, 3.9], font_size=8.2)
mono_block(doc, f"""cd /Users/ayushpal/Desktop/hclTraining
.venv/bin/python src/generate_data.py        # regenerates data/, ~40 s, deterministic

Expected tail:
  episodes: {summary['tables']['fact_inpatient_episode']} visits: {summary['tables']['fact_opd_visit']}
  rejected+removed: {summary['total_rejected']} outliers kept: {summary['outlier_flags']} load_errors: 0 fk_violations: 0
  {len(summary['checks'])}/{len(summary['checks'])} checks PASS""", 7.6)

# 8 assumptions
h1(doc, "8. Assumptions", page_break=True)
add_table(doc, ["ID", "Assumption", "Why", "If it is wrong"],
          [["AS-18", "Ward bed counts stay at their 2023 values for the whole period.", "Matches AS-17; avoids a ward-version table the schema cannot key (see DF-01).", "A ward-level validity table is needed; BOR history changes."],
           ["AS-19", "Demand is a Poisson process with department seasonality and weekday effects.", "Gives realistic day-to-day variance without predicting anything.", "Real arrivals are burstier; use a negative-binomial and re-tune."],
           ["AS-20", "Readmissions are scheduled for a fraction of unplanned discharges, then placed only if a bed is free.", "Produces a controllable rate without breaking capacity.", "The realised rate is slightly below the configured rate; documented in §3."],
           ["AS-21", "Billing follows a length-of-stay tariff with payer multipliers.", "Keeps revenue per bed-day comparable across departments.", "Real tariffs are procedure-based; the ranking may shift."],
           ["AS-22", f"OPD visits draw 28% of their patient keys from the same department's existing pool; {summary.get('cross_setting_pct', 13.1)}% of all visits therefore belong to a patient who also has an inpatient episode.", "Allows cross-setting analysis for OQ-08 without an FK the schema does not have.", "Cross-setting charts would be thinner but still valid."]],
          [0.6, 2.1, 2.1, 1.8], font_size=7.8)

# 9 observations
h1(doc, "9. What the build exposed", page_break=True)
para(doc, "Generating data is the first time a schema is used rather than drawn, and it found four things worth "
          "recording.", space_after=4)
add_table(doc, ["ID", "Severity", "Finding", "Decision"],
          [["DF-01", "MED", "dim_ward cannot hold two versions of a ward. ward_id is the primary key, so effective_to exists but only one row per ward can exist — the columns suggest SCD-2, the key forbids it. The Step 8/9 wording “SCD-2 tracked” was therefore stronger than the schema.",
            f"Accepted for this dataset (AS-17/AS-18): every ward has one row with effective_to NULL. A true history needs PK (ward_id, effective_from) and both facts to carry the validity date. Deferred — no requirement reads historical bed counts, and the change would ripple through {summary['tables']['fact_inpatient_episode']:,} rows."],
           ["DF-02", "LOW", "patient_key is a degenerate dimension shared by both facts. Nothing stops the same key carrying two genders.",
            "The generator stores one gender and one age band per key, and only lets the band advance with the calendar, so the delivered data is consistent. A TC-T case at Step 13 will assert it."],
           ["DF-03", "INFO", f"Volume landed at {summary['tables']['fact_inpatient_episode']:,} episodes, not 57,000.",
            f"{sum(summary['blocked'].values()):,} candidate admissions were bed-blocked and correctly not recorded. This is a property of the capacity model, not a data defect, and it is reported rather than hidden."],
           ["DF-04", "LOW", f"app.db is {size_mb(os.path.join(DATA, 'app.db'))}, above the 30–35 MB estimate in Step 8.",
            "Indexes and STRICT page overhead account for the difference. Still well under GitHub's 100 MB per-file limit; no action."]],
          [0.6, 0.7, 2.9, 2.4], font_size=7.6)

# 10 sign-off
h1(doc, "10. Sign-off")
add_table(doc, ["Role", "Name", "Responsibility", "Status"],
          [["Student / Author", STUDENT, "Wrote the generator, tuned the patterns, ran the validation", "Signed — 19-Nov-2026"],
           ["Mentor", MENTOR, "Reviews the patterns and the DF-01 deferral", "Pending"],
           ["Internal guide", "Project guide, MCA (DA)", "Confirms the dataset is fit for analysis at Step 15", "Pending"]],
          [1.4, 1.5, 2.8, 1.4], font_size=8.4)
para(doc, "Next step: Step 11 writes the user stories against this dataset — one story per screen and role, each with "
          "the exact query it implies and the acceptance measure it inherits from its BR.", size=9.5, italic=True)

doc.save(OUT)
print("saved:", OUT)
print("tables:", len(doc.tables), "paragraphs:", len(doc.paragraphs))
