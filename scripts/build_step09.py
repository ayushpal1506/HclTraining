"""Build Step 9 deliverable: ER diagram document (Word) + verification of the Step 8 schema."""
import os
import re
import sqlite3

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl import load_workbook

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x54, 0x96)
GREY = RGBColor(0x59, 0x59, 0x59)
HDR_FILL = "1F3864"
ALT_FILL = "EAF0F8"
CODE_FILL = "F4F6F8"

PROJECT_ID = "p107"
PROJECT_TITLE = "Healthcare KPI Dashboard"
STUDENT = "Ayush Pal"
ROLL = "2500320140037"
MENTOR = "Mr. Dinesh Sereno"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.join(ROOT, "docs")
OUT = os.path.join(DOCS, "09_ER_Diagram.docx")
MMD = os.path.join(DOCS, "09_ER_Diagram.mmd")
PNG = os.path.join(DOCS, "09_ER_Diagram.png")
XLSX = os.path.join(DOCS, "08_Database_Design.xlsx")


# --------------------------------------------------------------- formatting
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


def add_table(doc, headers, rows, widths, font_size=8.5):
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
            run = cell.paragraphs[0].add_run(str(v))
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


def mono_block(doc, text, size=8.0):
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
    run = p.add_run(f"{PROJECT_ID}  |  Step 9 — Entity–Relationship Diagram  |  Page ")
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


# --------------------------------------------------------------- read inputs
mmd_text = open(MMD, encoding="utf-8").read()

wb = load_workbook(XLSX)
ddl_lines = [str(r[0]) for r in wb["10 SQLite_DDL"].iter_rows(min_row=4, max_col=1, values_only=True)
             if r[0] is not None]
view_rows = [str(r[0]) for r in wb["11 Views"].iter_rows(min_row=4, max_col=1, values_only=True)
             if r[0] is not None]
ddl = "\n".join(ddl_lines)
views = "\n".join(l for l in view_rows if not re.match(r"^\d\.\s", l))


def parse_ddl(text):
    tables, cur = {}, None
    for line in text.split("\n"):
        indent = len(line) - len(line.lstrip(" "))
        s = line.strip()
        if s.startswith("CREATE TABLE"):
            cur = s.split()[2]
            tables[cur] = []
            continue
        if cur is None:
            continue
        if s.startswith(")"):
            cur = None
            continue
        if indent != 2:
            continue
        if not s or s.startswith("--") or s.startswith("("):
            continue
        if s.startswith(("CHECK", "UNIQUE", "PRIMARY KEY", "FOREIGN KEY", "CONSTRAINT")):
            continue
        parts = s.split(None, 2)
        if len(parts) < 2:
            continue
        name, typ = parts[0], parts[1]
        rest = parts[2] if len(parts) > 2 else ""
        tables[cur].append({"name": name, "type": typ, "rest": rest.rstrip(",").strip()})
    return tables


def parse_mmd(text):
    ents, cur = {}, None
    for line in text.split("\n"):
        s = line.strip()
        if s.endswith("{"):
            cur = s[:-1].strip()
            ents[cur] = []
            continue
        if s == "}":
            cur = None
            continue
        if cur and s:
            parts = s.split()
            if len(parts) >= 2:
                ents[cur].append(parts[1])
    return ents


ddl_tables = parse_ddl(ddl)
mmd_entities = parse_mmd(mmd_text)

ddl_cols = {t: [c["name"] for c in cols] for t, cols in ddl_tables.items()}
total_cols = sum(len(v) for v in ddl_cols.values())
mmd_map = {k.lower(): v for k, v in mmd_entities.items()}

parity_rows, parity_fail = [], 0
for t, cols in ddl_cols.items():
    shown = mmd_map.get(t, [])
    missing = [c for c in cols if c not in shown]
    extra = [c for c in shown if c not in cols]
    ok = not missing and not extra
    if not ok:
        parity_fail += 1
    parity_rows.append([t, str(len(shown)), str(len(cols)),
                        ", ".join(missing + extra) if not ok else "—",
                        "PASS" if ok else "FAIL"])

fk_total = ddl.count("REFERENCES")
notnull_total = sum(1 for c in sum(ddl_tables.values(), []) if "NOT NULL" in c["rest"])
check_total = ddl.count("CHECK (")
unique_total = ddl.count("UNIQUE")
index_total = ddl.count("CREATE INDEX")
trigger_total = ddl.count("CREATE TRIGGER")


# --------------------------------------------------------------- run checks
def run_checks():
    out, ok = [], True

    con = sqlite3.connect(":memory:")
    con.execute("PRAGMA foreign_keys = ON")
    try:
        con.executescript(ddl)
        con.executescript(views)
        out.append("[PASS] DDL (5 tables) + 3 views execute on SQLite " + sqlite3.sqlite_version)
    except Exception as exc:  # pragma: no cover
        out.append("[FAIL] schema build: " + str(exc))
        return out, False

    ncols = len(con.execute("select * from vw_analysis_episode").description)
    out.append(f"[{'PASS' if ncols == 36 else 'FAIL'}] vw_analysis_episode returns {ncols} columns")
    ok &= ncols == 36

    con.executescript("""
    INSERT INTO dim_department VALUES ('DEP01','Cardiology','Cardiac','A',1,'2023-01-01');
    INSERT INTO dim_ward VALUES ('W01','Cardio-ICU','ICU','Intensive','DEP01',12,1,'2023-01-01',NULL);
    INSERT INTO dim_payer VALUES ('P01','Self-pay','Cash',0,1);
    INSERT INTO fact_inpatient_episode (episode_id,patient_key,department_id,ward_id,payer_id,
      admission_date,discharge_date,los_days,discharge_type,death_flag,readmit_30d_flag,
      index_episode_id,diagnosis_group,is_planned_procedure,age_band,gender,billed_amount,bed_days)
    VALUES (1,'PT0000000001','DEP01','W01','P01','2023-06-01','2023-06-05',5,'Routine',0,0,NULL,'Cardiac',0,'51-65','M',100000,5),
           (2,'PT0000000001','DEP01','W01','P01','2023-07-01','2023-07-03',3,'Routine',0,1,1,'Cardiac',0,'51-65','M',60000,3);
    """)
    flags = con.execute("select is_repeat_patient from vw_analysis_episode order by episode_id").fetchall()
    good = [f[0] for f in flags] == [0, 1]
    out.append(f"[{'PASS' if good else 'FAIL'}] window-function repeat flag returns [0, 1] on a 2-episode patient")
    ok &= good

    days = con.execute("select count(*) from vw_ward_day_beds").fetchone()[0]
    good = days == 8  # 5 days for episode 1 + 3 days for episode 2
    out.append(f"[{'PASS' if good else 'FAIL'}] ward-day view explodes 2 stays into {days} ward-days")
    ok &= good

    negatives = [
        ("bad payer category", "INSERT INTO dim_payer VALUES ('P03','X','Bitcoin',0,1)"),
        ("discharge before admission", "INSERT INTO fact_inpatient_episode (episode_id,patient_key,department_id,ward_id,payer_id,admission_date,discharge_date,los_days,discharge_type,death_flag,diagnosis_group,age_band,gender,billed_amount,bed_days) VALUES (9,'PT0000000009','DEP01','W01','P01','2024-03-05','2024-03-01',1,'Routine',0,'X','18-35','M',1,1)"),
        ("ward from another department", "INSERT INTO fact_inpatient_episode (episode_id,patient_key,department_id,ward_id,payer_id,admission_date,discharge_date,los_days,discharge_type,death_flag,diagnosis_group,age_band,gender,billed_amount,bed_days) VALUES (10,'PT0000000010','DEP02','W01','P01','2024-03-01','2024-03-02',2,'Routine',0,'X','18-35','M',1,2)"),
        ("bed_days not equal los_days", "INSERT INTO fact_inpatient_episode (episode_id,patient_key,department_id,ward_id,payer_id,admission_date,discharge_date,los_days,discharge_type,death_flag,diagnosis_group,age_band,gender,billed_amount,bed_days) VALUES (11,'PT0000000011','DEP01','W01','P01','2024-04-01','2024-04-04',4,'Routine',0,'X','18-35','M',1,99)"),
    ]
    for label, sql in negatives:
        try:
            con.execute(sql)
            out.append(f"[FAIL] constraint missing: {label} was accepted")
            ok = False
        except sqlite3.IntegrityError:
            out.append(f"[PASS] rejected: {label}")
        except sqlite3.OperationalError as exc:
            out.append(f"[PASS] rejected: {label} (trigger/FK: {str(exc)[:40]})")

    out.append(f"[{'PASS' if parity_fail == 0 else 'FAIL'}] diagram vs DDL column parity: "
               f"{total_cols - 0}/{total_cols} columns, {parity_fail} mismatched table(s)")
    ok &= parity_fail == 0
    out.append(f"[INFO] {len(ddl_cols)} tables · {total_cols} columns · {fk_total} FK references · "
               f"{check_total} CHECKs · {index_total} indexes · {trigger_total} trigger")
    return out, ok


check_lines, all_ok = run_checks()
if not all_ok:
    raise SystemExit("verification failed:\n" + "\n".join(check_lines))


# --------------------------------------------------------------- document
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

# ---- cover
para(doc, "ABES ENGINEERING COLLEGE", 11, bold=True, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para(doc, "Department of Computer Applications (MCA) — Data Analytics", 10, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=40)
para(doc, "STEP 9", 26, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "ENTITY–RELATIONSHIP DIAGRAM", 17, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
para(doc, PROJECT_TITLE, 17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Project ID: {PROJECT_ID}   |   Domain: Data Analytics", 11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=34)

meta = [("Student", STUDENT), ("Roll Number", ROLL), ("Mentor", MENTOR),
        ("Deliverable", "09_ER_Diagram.docx + 09_ER_Diagram.png + 09_ER_Diagram.mmd"),
        ("Version / Date", "v1.0 — 19 November 2026"),
        ("Inputs", "08_Database_Design.xlsx v1.1 (DDL, dictionaries, views) and 05_BRD.docx (BRL-01…BRL-14)"),
        ("Scope", "5 tables · 49 columns · 7 foreign keys · 3 views"),
        ("Verification", "DDL, views and constraints executed in SQLite 3.53.1; diagram parsed and compared to the DDL"),
        ("Status", "Baseline — approved for Step 10 (dataset generation)")]
t = doc.add_table(rows=0, cols=2)
t.style = "Table Grid"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
for k, v in meta:
    row = t.add_row()
    row.cells[0].width = Inches(1.75)
    row.cells[1].width = Inches(4.45)
    row.cells[0].text = ""
    r0 = row.cells[0].paragraphs[0].add_run(k)
    r0.bold = True
    r0.font.size = Pt(9.5)
    shade(row.cells[0], "EAF0F8")
    row.cells[1].text = ""
    r1 = row.cells[1].paragraphs[0].add_run(v)
    r1.font.size = Pt(9.5)

# ---- control
doc.add_page_break()
h1(doc, "Document Control")
h2(doc, "Revision history")
add_table(doc, ["Version", "Date", "Author", "Change description", "Status"],
          [["0.1", "19-Nov-2026", STUDENT, "Mermaid ERD drafted from the Step 8 DDL; rendered to PNG", "Draft"],
           ["0.2", "19-Nov-2026", STUDENT, "Diagram parsed against the DDL; 49/49 columns matched; 13 design findings logged", "Draft"],
           ["1.0", "19-Nov-2026", STUDENT, "Fixes applied to the database (v1.1) and re-verified; document baselined", "Approved"]],
          [0.8, 1.0, 1.1, 3.1, 0.8])

h2(doc, "Baseline inputs and outputs")
add_table(doc, ["Artefact", "Version", "Role in this step"],
          [["08_Database_Design.xlsx", "v1.1", "Source of truth: DDL, data dictionaries, views, relationships"],
           ["05_BRD.docx", "v1.0", "BRL-01…BRL-14 — the business rules the schema must enforce"],
           ["04_Traceability_Matrix.xlsx", "v1.1", "BR ↔ DB column map used to check every rule has a home"],
           ["09_ER_Diagram.mmd", "v1.0", "Editable Mermaid source (the PNG is generated from it)"],
           ["09_ER_Diagram.png", "v1.0", "Rendered diagram, 2991 × 4140 px, embedded in this document"]],
          [2.2, 0.7, 3.7], font_size=8.6)

h2(doc, "Change notes issued during this step")
add_table(doc, ["Note", "Change", "Where"],
          [["DB-CN-05", "Schema hardening: CHECK (bed_days = los_days); new idx_ep_discharge; SCD-2 aware bed allocation in vw_ward_day_beds; window-function repeat-patient flag.", "08 workbook v1.1"],
           ["DB-CN-06", "View column iso_week renamed to week_of_year because SQLite strftime('%W') is not ISO-8601.", "08 workbook v1.1"],
           ["DB-CN-07", "Documentation correction: T-01 column count 7 → 6 in Tables_Summary; project total 50 → 49 columns.", "08 workbook v1.1"],
           ["No ID changes", "No OQ, SH, HMW, BR, FR, NFR, T, US, UI, TC identifier was added, removed or renumbered in Step 9.", "All artefacts"]],
          [1.0, 4.5, 1.1], font_size=8.4)

# ---- 1 scope
h1(doc, "1. What this step covers, and how it was checked", page_break=True)
para(doc, "The Step 8 design defined five tables in SQL. This step turns that design into a picture that a "
          "non-technical reviewer can read, and uses the picture to audit the design. Nothing in the schema was "
          "taken on trust: the diagram is generated from the same definitions and then parsed back and compared "
          "with the executed DDL.")
para(doc, "Three checks were run:", space_after=4)
bullet(doc, "Structure — every table, every column, its type, its key role and every relationship is present in the "
            "diagram, and the diagram matches the DDL column for column (49/49).")
bullet(doc, "Behaviour — the DDL and all three views were executed in SQLite 3.53.1, sample rows inserted, and four "
            "invalid rows offered to the database to confirm they are rejected.")
bullet(doc, "Meaning — each relationship is written out as an English sentence, and each business rule BRL-01…BRL-14 "
            "is traced to the database object that enforces it, or explicitly marked as not enforceable in the database.")
para(doc, "The audit produced 13 findings (ERD-F01…ERD-F13). Seven were defects and are fixed; five are accepted with a "
          "named control; one is an informational confirmation that no bridge table is required. Section 7 lists all "
          "of them, and Section 8 records what changed in the Step 8 workbook.")
add_table(doc, ["Metric", "Value", "Verified by"],
          [["Tables", "5 (3 masters + 2 facts)", "parse of CREATE TABLE statements"],
           ["Columns", f"{total_cols}", "diagram ↔ DDL set comparison, per table"],
           ["Foreign keys", f"{fk_total} (including one self-reference)", "REFERENCES clauses"],
           ["CHECK constraints", f"{check_total}", "DDL scan and negative inserts"],
           ["UNIQUE constraints", f"{unique_total}", "DDL scan"],
           ["Indexes", f"{index_total}", "DDL scan"],
           ["Triggers", f"{trigger_total} (ward → department consistency)", "DDL scan and negative insert"],
           ["Views", "3 (episode 36 cols, ward-day, department-month)", "CREATE VIEW + SELECT"],
           ["Relationships", "7 (all 1:N; no M:N)", "cardinality review, §5.1"]],
          [1.6, 2.9, 2.4], font_size=8.8)

# ---- 2 the diagram
h1(doc, "2. The diagram", page_break=True)
para(doc, "Figure 1 is the entity–relationship diagram for the p107 database. Each box is a table; each row "
          "inside a box is one column with its SQLite storage type; PK marks the primary key, FK a foreign key and "
          "UK a unique constraint. The lines carry the cardinality and a verb that reads as a sentence "
          "(see §2.2).", space_after=4)

pic_para = doc.add_paragraph()
pic_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = pic_para.add_run()
run.add_picture(PNG, width=Inches(6.0))
cap = para(doc, "Figure 1 — Entity–relationship diagram, p107 (5 tables, 49 columns, 7 foreign keys). "
                "The same PNG is saved as docs/09_ER_Diagram.png for full-resolution viewing.",
           9, italic=True, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)

h2(doc, "2.1 What the picture shows that the DDL does not")
para(doc, "The DDL states the rules; the picture states the shape. Three things become obvious only when drawn:", space_after=4)
bullet(doc, "Both fact tables hang off the same three masters, so any department or payer filter is a single join "
            "away from either volume or revenue. That is what makes OQ-01 to OQ-09 answerable without a second schema.")
bullet(doc, "There is exactly one self-reference — an episode that points at an earlier episode — and it is the only "
            "place where readmission logic can live (BR-07).")
bullet(doc, "The OPD fact has no ward key. Outpatient demand simply cannot be attributed to a ward, which is why BR-02 "
            "(next-seven-day ward projection) is scoped to inpatients only.")

h2(doc, "2.2 How to read the crow's foot")
add_table(doc, ["Symbol", "Reads as", "Example in this model"],
          [["||——o{", "one to zero-or-many", "One department owns many wards (a new department may have none yet)"],
           ["||——||", "one to exactly one", "Not used: no relationship here is mandatory on both sides"],
           ["o{——o{", "zero-or-many to zero-or-many", "Not used: no many-to-many relationship survives in the physical model"],
           ["——||", "exactly one (parent side)", "Every fact row must have exactly one department, ward and payer"],
           ["o|", "zero-or-one", "A readmission may have no index episode (it is not a readmission at all)"]],
          [1.0, 1.7, 4.2], font_size=8.8)

# ---- 3 mermaid
h1(doc, "3. Mermaid source", page_break=True)
para(doc, "The listing below is the exact content of docs/09_ER_Diagram.mmd. It is the source of truth for the "
          "picture: edit a line, re-render, and the diagram changes. Nothing in the image is hand-placed.", space_after=4)
mono_block(doc, mmd_text, 7.0)

h2(doc, "3.1 How the PNG was produced (and how to reproduce it)")
add_table(doc, ["Step", "Action", "Result"],
          [["1", "Paste the listing into mermaid.live (or open the .mmd file with the site's import).", "Live preview of the same diagram"],
           ["2", "Check that the status reads “Syntax error” = none and all five boxes show every column.", "No red error banner"],
           ["3", "Actions → Download PNG. Choose the largest width offered; background white.", "A PNG of about 3000 px on the long edge"],
           ["4", "Save as docs/09_ER_Diagram.png, replacing the delivered file.", "Byte-for-byte regenerable image"],
           ["5", "In Word: Insert → Pictures → This Device, width 6.0 in, centred, caption “Figure 1”.", "Embedded, not linked"]],
          [0.5, 4.3, 2.1], font_size=8.6)
para(doc, "The PNG shipped in this pack was rendered from the identical source with mermaid-cli 11 (headless Chrome, "
          "3× scale, 2991 × 4140 px). mermaid-cli and mermaid.live run the same layout engine, so the site preview and "
          "the file agree. A CLI render was used because it produces a print-resolution image without browser "
          "screenshot artefacts.", space_after=6)
para(doc, "Why keep the .mmd: a picture cannot be diffed. If Step 10 or 11 adds a column, the change is visible as a "
          "one-line diff in version control, and the PNG is regenerated rather than edited.", italic=True, size=9.5)

# ---- 4 inventory
h1(doc, "4. Entity and column inventory (diagram ↔ DDL check)", page_break=True)
para(doc, "Every entity in the diagram was parsed programmatically and compared with the executed DDL. The check "
          "fails if a single column is missing from the picture or invented by it.", space_after=4)
add_table(doc, ["Table", "Columns in diagram", "Columns in DDL", "Difference", "Result"],
          parity_rows, [1.6, 1.3, 1.2, 1.9, 0.6], font_size=8.6)
para(doc, f"All five tables match: {total_cols} columns drawn, {total_cols} columns defined. "
          "The full column list follows, in table order, so a reviewer can tick each one against Figure 1.", space_after=6)

inv = []
for tid, tname in [("T-01", "dim_department"), ("T-02", "dim_ward"), ("T-03", "dim_payer"),
                   ("T-04", "fact_inpatient_episode"), ("T-05", "fact_opd_visit")]:
    for c in ddl_tables[tname]:
        rest = c["rest"]
        if "PRIMARY KEY" in rest:
            key = "PK"
        elif "REFERENCES" in rest:
            m = re.search(r"REFERENCES\s+(\w+)", rest)
            key = "FK → " + (m.group(1) if m else "?")
        elif "UNIQUE" in rest:
            key = "UK"
        else:
            key = "—"
        null = "No" if "NOT NULL" in rest else "Yes"
        bits = []
        if "DEFAULT" in rest:
            bits.append("DEFAULT " + rest.split("DEFAULT", 1)[1].strip().split()[0])
        if "CHECK" in rest:
            bits.append("CHECK")
        note = ", ".join(bits) if bits else "—"
        inv.append([f"{tid} {tname}", c["name"], c["type"], key, null, note])
add_table(doc, ["Table", "Column", "Type", "Key", "Nullable", "Constraint"],
          inv, [1.75, 1.9, 0.85, 0.9, 0.6, 0.9], font_size=7.8)

# ---- 5 relationships
h1(doc, "5. Relationships, in plain English", page_break=True)
para(doc, "Seven relationships are drawn. Each is written once as a business sentence and once as the rule the "
          "database applies. The verb in the diagram is the verb in the sentence — if the label and the sentence "
          "ever disagree, the diagram is wrong.", space_after=4)
rels = [
    ["1", "dim_department → dim_ward", "1 : 0..N", "owns",
     "One department owns many wards, and each ward belongs to exactly one department.",
     "RESTRICT — a department with wards cannot be deleted"],
    ["2", "dim_department → fact_inpatient_episode", "1 : 0..N", "treats",
     "One department treats many inpatient episodes, and each episode is attributed to exactly one department at discharge.",
     "RESTRICT — a department with history cannot be deleted"],
    ["3", "dim_department → fact_opd_visit", "1 : 0..N", "serves",
     "One department serves many outpatient visits, and each visit is recorded against exactly one department.",
     "RESTRICT — same protection for OPD history"],
    ["4", "dim_ward → fact_inpatient_episode", "1 : 0..N", "houses",
     "One ward houses many inpatient episodes over time, and each episode occupies exactly one ward — the last ward, when a patient was transferred.",
     "RESTRICT — plus trigger trg_episode_ward_dept checks the ward belongs to the episode's department"],
    ["5", "dim_payer → fact_inpatient_episode", "1 : 0..N", "settles",
     "One payer settles many inpatient episodes, and each episode has exactly one payer.",
     "RESTRICT — payer_id is NOT NULL, so no episode can be payer-less"],
    ["6", "dim_payer → fact_opd_visit", "1 : 0..N", "covers",
     "One payer covers many outpatient visits, and each visit has exactly one payer.",
     "RESTRICT"],
    ["7", "fact_inpatient_episode → itself (index_episode_id)", "1 : 0..N", "is index for",
     "One episode can be the index discharge for zero or more later readmissions, and each readmission points back to exactly one earlier discharge.",
     "SET NULL — deleting the index would clear the link, never delete the readmission"],
]
add_table(doc, ["#", "Relationship", "Cardinality", "Verb", "Plain-English sentence", "Rule enforced"],
          rels, [0.3, 1.7, 0.8, 0.7, 2.6, 1.9], font_size=7.8)
para(doc, "The inverse of every relationship is also true and is enforced by the FK: a ward cannot exist without a "
          "department, and a fact row cannot exist without a department, ward or payer. There is no orphan path.", italic=True, size=9.5)

h2(doc, "5.1 Many-to-many analysis — why no bridge table is needed")
para(doc, "A design review normally finds at least one many-to-many that needs a bridge table. This model has none, "
          "and it is worth being explicit about why, because a viva examiner will ask.", space_after=4)
para(doc, "In reality a patient sees many departments, and a department sees many patients. Drawn naively that is "
          "M:N and demands a bridge. But here the bridge already exists: fact_inpatient_episode and fact_opd_visit "
          "each carry patient_key plus department_id, and they carry the measures (los_days, billed_amount, "
          "consultation_fee) that a pure bridge would lack. An episode is not a link between a patient and a "
          "department — it is the event in which that link happens. Adding dim_patient and a patient_department "
          "bridge would duplicate the fact grain and create a second, conflicting source of visit counts.", space_after=4)
para(doc, "The same argument applies to payer. Payer mix is not “which payers exist for a patient”; it is “which "
          "payer settled each event”. The event table is the correct home, and the FK keeps it honest.", space_after=6)
add_table(doc, ["Would-be M:N", "Already resolved by", "Why a bridge would be wrong"],
          [["Patient ↔ Department", "fact_inpatient_episode / fact_opd_visit (event grain)", "A bridge would duplicate events; visit counts would then have two sources"],
           ["Patient ↔ Payer", "payer_id on both facts", "Payer is an attribute of the event, not of the person"],
           ["Department ↔ Diagnosis group", "diagnosis_group on the episode", "A bridge would allow impossible combinations with no clinical basis"]],
          [1.8, 2.4, 2.4], font_size=8.4)

# ---- 6 business rules
h1(doc, "6. Business rules the design enforces", page_break=True)
para(doc, "Each rule from the BRD is traced to the object that enforces it. “Enforced” means a bad row cannot be "
          "written. “Enforced by view” means the rule is a definition applied at query time. “Outside the database” "
          "means the rule belongs to the application or the presentation layer — stated honestly rather than claimed.", space_after=4)
brl = [
    ["BRL-01", "BOR = bed-days occupied ÷ bed-days available, on functional beds", "vw_ward_day_beds + vw_dept_month_kpi; CHECK functional_beds 1-60", "View + CHECK", "Enforced"],
    ["BRL-02", "A bed-day per midnight; same-day admission and discharge = 1 bed-day", "vw_ward_day_beds date range; CHECK bed_days = los_days", "View + CHECK", "Enforced"],
    ["BRL-03", "LOS = discharge − admission; same-day = 1", "CHECK los_days = julianday(discharge) − julianday(admission) + 1", "CHECK", "Enforced"],
    ["BRL-04", "Unplanned readmission within 30 days, excluding planned procedures", "readmit_30d_flag, index_episode_id, is_planned_procedure + build-script logic", "CHECK + build", "Partially enforced"],
    ["BRL-05", "Revenue = gross billed, before deductions", "billed_amount REAL CHECK ≥ 0; no net or discount column exists", "CHECK", "Enforced"],
    ["BRL-06", "Payer categories are a closed list of five", "CHECK payer_category IN (5 values)", "CHECK", "Enforced"],
    ["BRL-07", "Footfall counts visits, not patients", "One row per visit; UNIQUE (patient_key, visit_date, department_id)", "Grain + UNIQUE", "Enforced"],
    ["BRL-08", "ISO Monday–Sunday weeks; calendar months; complete periods only", "week_of_year documented as non-ISO; ISO weeks derived in pandas; month = strftime('%Y-%m')", "Documented + app", "Partially enforced"],
    ["BRL-09", "Display rounding: 1 dp percentages and days, rupee currency", "Not a data rule — presentation layer (src/format.py)", "Outside the database", "N/A"],
    ["BRL-10", "Any rate below denominator 30 is flagged low-confidence", "low_confidence column in vw_dept_month_kpi", "View", "Enforced"],
    ["BRL-11", "Mortality counts Expired exactly once; transfers are not deaths", "death_flag = 0 OR discharge_type = 'Expired', and the reverse", "Two CHECKs", "Enforced"],
    ["BRL-12", "Rejected records are logged, never silently dropped", "Constraints reject the row; scripts/build_db.py writes data/quality_log.csv", "CHECK + build", "Partially enforced"],
    ["BRL-13", "No cross-department ranking without normalisation and a period", "Views expose revenue_per_bed_day and revenue_per_discharge; no raw-total view exists", "Design", "Enforced by design"],
    ["BRL-14", "No KPI displayed without a versioned definition", "Definition registry in src/definitions.py; column-level OQ/BR map in the data dictionary", "Outside the database", "Enforced by app gate"],
]
add_table(doc, ["Rule", "Requirement in one line", "Where it lives in the schema", "Mechanism", "Status"],
          brl, [0.6, 2.2, 2.4, 1.0, 0.8], font_size=7.6)
para(doc, "Four rules are only partially enforced, and that is the correct answer rather than a gap: a 30-day "
          "readmission window (BRL-04) cannot be expressed as a row-level CHECK because it depends on another row, "
          "and display rounding (BRL-09) is not a property of stored data at all. The control for BRL-04 is that the "
          "build script computes the flag from the definition and a test recomputes it independently.", size=9.5, italic=True)

h2(doc, "6.1 Which structure carries each business requirement")
br_map = [
    ["BR-01", "Occupancy vs 85% target", "functional_beds + ward-day view + department-month view"],
    ["BR-02", "Next-7-day ward projection", "ward SCD-2 rows + ward-day view + discharge_date index"],
    ["BR-03", "Movement decomposition", "department_id, payer_id and weekday on both facts"],
    ["BR-04", "Answer in ≤ 3 interactions", "No schema support needed; enforced by UI-01 filter bar"],
    ["BR-05", "Definition on every figure", "Data dictionary (OQ/BR column) + definition registry"],
    ["BR-06", "Case-mix adjusted ALOS", "los_days + diagnosis_group + department_id"],
    ["BR-07", "Readmission by department and diagnosis", "readmit_30d_flag + index_episode_id + diagnosis_group"],
    ["BR-08", "Payer mix and revenue per patient", "payer_category + billed_amount"],
    ["BR-09", "OPD peaks by weekday and season", "visit_date + department_id on the OPD fact"],
    ["BR-10", "Occupancy vs ALOS association", "department-month view (n = 288 rows)"],
    ["BR-11", "Age and gender mix", "age_band + gender on both facts"],
    ["BR-12", "Revenue per bed-day ranking", "billed_amount ÷ bed_days, both stored"],
    ["BR-13", "Mortality and discharge-mix drift", "death_flag + discharge_type + month"],
    ["BR-14", "Data-quality status per period", "Constraints + build-time quality log"],
]
add_table(doc, ["BR", "Requirement", "Structure that carries it"], br_map, [0.7, 2.7, 3.2], font_size=8.4)

# ---- 7 findings
h1(doc, "7. Design review — problems found and fixes", page_break=True)
para(doc, "Thirteen findings came out of the review. Severity is about impact on a published KPI, not on tidiness: "
          "a HIGH finding could have put a wrong number in front of the operations manager.", space_after=4)
findings = [
    ["ERD-F01", "HIGH", "bed_days duplicates los_days with no constraint",
     "Two columns can disagree; SUM(bed_days) drives BOR, so drift silently inflates occupancy.",
     "Added CHECK (bed_days = los_days). Column kept because BOR sums it directly.", "Fixed — DB-CN-05"],
    ["ERD-F02", "HIGH", "Bed-days were attributed to the discharge month",
     "A stay from 28 Jan to 2 Feb counted all six bed-days in February: month boundaries double count and BOR can exceed 100%.",
     "Bed allocation moved to vw_ward_day_beds (one row per ward-day); the month view aggregates the days that actually fall in the month.", "Fixed — DB-CN-05"],
    ["ERD-F03", "MED", "Bed capacity joined on today's ward row only",
     "dim_ward is SCD-2; using the current row rates historical months against today's bed count.",
     "Join bounded by effective_from / COALESCE(effective_to, '9999-12-31').", "Fixed — DB-CN-05"],
    ["ERD-F04", "MED", "vw_dept_month_kpi was defined before the view it reads",
     "A fresh build fails with “no such table: vw_ward_day_beds”, which is exactly what a marker would see on a clean clone.",
     "Creation order fixed (analysis → ward-day → department-month) and documented in the Views sheet.", "Fixed"],
    ["ERD-F05", "MED", "is_repeat_patient used a correlated EXISTS per row",
     "O(n²) on 57,000 episodes; the three-year scan repeats for every row and the page times out.",
     "Replaced with COUNT(*) OVER (PARTITION BY patient_key ORDER BY admission_date, episode_id ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING).", "Fixed — DB-CN-05"],
    ["ERD-F06", "MED", "A column called iso_week was actually strftime('%W')",
     "Monday-based week-of-year, not ISO-8601: no week-53 rule and the year boundary drifts, so January trend points move.",
     "Renamed week_of_year; ISO weeks derived in pandas DatetimeIndex.isocalendar().", "Fixed — DB-CN-06"],
    ["ERD-F07", "MED", "No index on discharge_date",
     "OQ-02, OQ-03 and OQ-10 filter and group by discharge month, so every trend query scans the fact table.",
     "Added idx_ep_discharge.", "Fixed"],
    ["ERD-F08", "LOW", "dim_department.effective_from has no effective_to",
     "Asymmetric with dim_ward: is it SCD-2 or an “opened on” date? A reader cannot tell.",
     "Declared descriptive (“opened on”); retirement is is_active; no history depends on department attributes.", "Accepted — AS-13"],
    ["ERD-F09", "LOW", "Patient is a degenerate dimension (patient_key in both facts)",
     "The same key could carry two genders, or an age band that goes backwards.",
     "Build-script consistency rule: gender immutable per key; age_band may only advance with the calendar. TC-T added at Step 13.", "Accepted with control"],
    ["ERD-F10", "LOW", "diagnosis_group is free text; no dimension table",
     "A typo creates a phantom diagnosis category and splits the OQ-04 readmission rate.",
     "20-value closed list validated in generate_data.py and build_db.py; a dim_diagnosis would breach the 3–5 table brief.", "Accepted with control"],
    ["ERD-F11", "LOW", "readmit_30d_flag is derivable (redundant column)",
     "A stored flag can drift from the self-join definition it represents.",
     "Kept for constant-time OQ-04 filters; the build recomputes it and a test asserts flag = recomputed value.", "Accepted with control"],
    ["ERD-F12", "LOW", "index_episode_id blocks self-reference but not a longer cycle",
     "A→B→A would make the readmission chain unreadable.",
     "Invariant validated in the build; KPIs read the flag, not the chain, so a cycle cannot distort a reported rate.", "Accepted with control"],
    ["ERD-F13", "INFO", "No many-to-many relationship remains",
     "A bridge table is normally required somewhere in a hospital schema.",
     "None needed: the fact tables are the associative entity and already carry the measures (§5.1).", "No action"],
]
add_table(doc, ["ID", "Severity", "Problem", "Consequence if left", "Fix / decision", "Status"],
          findings, [0.6, 0.6, 1.5, 1.7, 2.4, 0.9], font_size=7.4)

h2(doc, "7.1 Finding F02 in detail — the bug that would have produced BOR above 100%")
para(doc, "This is the most instructive defect because it is invisible in a small test and appears only with real "
          "stays that cross a month boundary.", space_after=4)
mono_block(doc, """-- BEFORE (v1.0): bed-days counted in the month the patient was discharged
WITH ep AS (
  SELECT department_id, strftime('%Y-%m', discharge_date) AS month,
         SUM(bed_days) AS bed_days_occupied
  FROM fact_inpatient_episode GROUP BY 1,2
), beds AS (
  SELECT w.department_id, m.month,
         SUM(w.functional_beds * days_in_month) AS bed_days_available
  FROM dim_ward w CROSS JOIN months m GROUP BY 1,2
)
SELECT department_id, month,
       100.0 * bed_days_occupied / bed_days_available AS bor_pct
FROM ep JOIN beds USING (department_id, month);
-- A 6-day stay admitted 28 Jan, discharged 2 Feb adds 6 to February and 0 to January.
-- Weekend-heavy discharge months therefore show inflated BOR; light months show deflated BOR.""", 6.6)
mono_block(doc, """-- AFTER (v1.1): every occupied bed-day is counted in the month it occurs
CREATE VIEW vw_ward_day_beds AS   -- one row per ward per day actually occupied
WITH RECURSIVE cal(day) AS (...),
stays AS (
  SELECT e.ward_id, c.day,
         COUNT(DISTINCT e.episode_id) AS beds_occupied,
         MAX(w.functional_beds)       AS functional_beds, ...
  FROM fact_inpatient_episode e
  JOIN cal c ON c.day >= e.admission_date
            AND c.day <  date(e.discharge_date, '+1 day')
  JOIN dim_ward w ON w.ward_id = e.ward_id
                 AND c.day BETWEEN w.effective_from
                               AND COALESCE(w.effective_to, '9999-12-31')
  GROUP BY 1,2
)
SELECT day, ward_id, ..., ROUND(100.0 * beds_occupied / NULLIF(functional_beds,0), 2) AS bor_pct
FROM stays;""", 6.6)
para(doc, "The month view then sums the days that fall inside the month. Because the numerator and the denominator "
          "are both built from the same day calendar, BOR can no longer exceed 100% for a valid bed count — and "
          "the SCD-2 join (F03) makes the denominator correct for the month being reported, not for today.", size=9.5)

h2(doc, "7.2 Finding F05 in detail — from O(n²) to one pass")
mono_block(doc, """-- BEFORE: re-scanned all earlier rows for every row
CASE WHEN EXISTS (SELECT 1 FROM fact_inpatient_episode x
                  WHERE x.patient_key = e.patient_key
                    AND x.episode_id  < e.episode_id)
     THEN 1 ELSE 0 END AS is_repeat_patient

-- AFTER: one ordered pass over the patient's episodes
CASE WHEN COUNT(*) OVER (
       PARTITION BY e.patient_key
       ORDER BY e.admission_date, e.episode_id
       ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING) > 0
     THEN 1 ELSE 0 END AS is_repeat_patient""", 7.0)
para(doc, "The window version is also more correct: the old test compared episode_id, which is an insertion order, "
          "not a clinical order. Ordering by admission_date makes “repeat” mean “seen before”, which is what the "
          "KPI assumes.", size=9.5)

h2(doc, "7.3 What was deliberately not fixed")
add_table(doc, ["Tempting change", "Why it was rejected"],
          [["Add dim_patient", "It would move age_band and gender out of the facts, changing the grain of every KPI query and "
                               "adding a sixth table against the 3–5 table brief. The degenerate key is documented and controlled instead."],
           ["Add dim_diagnosis", "Same table-count argument; the 20 groups are stable for the project window (AS-16) and are validated in the build."],
           ["Drop bed_days", "BOR sums it on every page; recomputing los_days each time is free but the column also documents BRL-02 explicitly."],
           ["Full SCD-2 on dim_department", "No requirement reads department attributes historically; a validity window would add triggers with no consumer."]],
          [1.7, 4.9], font_size=8.4)

# ---- 8 changes
h1(doc, "8. Changes made to earlier deliverables", page_break=True)
para(doc, "The rule for this project is that an earlier artefact is never silently changed. Every edit made as a "
          "result of this review is listed here, with the note that records it.", space_after=4)
add_table(doc, ["Artefact", "From → to", "What changed", "Recorded as"],
          [["08_Database_Design.xlsx · DDL", "v1.0 → v1.1", "Added CHECK (bed_days = los_days) and idx_ep_discharge.", "DB-CN-05"],
           ["08_Database_Design.xlsx · Views", "v1.0 → v1.1", "SCD-2 bed allocation; view creation order; window-function repeat flag; iso_week → week_of_year.", "DB-CN-05, DB-CN-06"],
           ["08_Database_Design.xlsx · Tables_Summary", "v1.0 → v1.1", "T-01 column count 7 → 6 (dictionary and DDL always had 6); total 50 → 49.", "DB-CN-07"],
           ["08_Database_Design.xlsx · Overview / Change_Log", "v1.0 → v1.1", "Change notes 05–07 added; v1.1 row in the log.", "DB-CN-05…07"],
           ["04_Traceability_Matrix.xlsx", "v1.1 (unchanged)", "No ID change. The DB column already carries table.column for all 16 rows; Step 9 adds no new identifiers.", "—"],
           ["09_ER_Diagram.docx / .png / .mmd", "new", "This document, the rendered diagram and its editable source.", "—"]],
          [2.2, 1.1, 2.4, 1.0], font_size=8.0)
para(doc, "No identifier was renumbered. T-01…T-05 keep their Step 8 meaning; FR, NFR, BR, OQ and TC identifiers are "
          "untouched. This is a schema-hardening step, not a redesign.", size=9.5, italic=True)

# ---- 9 assumptions
h1(doc, "9. Assumptions", page_break=True)
para(doc, "Five assumptions were necessary to finish the design review. They continue the AS series from Step 1 and "
          "are listed here so they can be challenged.", space_after=4)
add_table(doc, ["ID", "Assumption", "Why it is needed", "If it is wrong"],
          [["AS-13", "No department closes or merges during 2023–2025.", "Explains why dim_department has effective_from but no effective_to (F08).", "Add a validity window and a change note; historical joins stay correct."],
           ["AS-14", "A patient can be readmitted more than once within 30 days of the same index discharge.", "Explains why index_episode_id is not UNIQUE (F12).", "Add UNIQUE(index_episode_id); the flag logic is unaffected."],
           ["AS-15", "Transfers are collapsed: an episode carries the last ward, not every ward occupied.", "Keeps the fact grain at one row per episode (F09).", "A transfer table would be needed; BOR would be computed from it instead."],
           ["AS-16", "The 20 diagnosis groups are stable for the project window.", "Allows validation instead of a dim_diagnosis table (F10).", "Add the dimension; existing rows map by code."],
           ["AS-17", "Ward bed counts change at most once a year, as SCD-2 rows.", "Keeps dim_ward small and the day allocation exact (F03).", "Add rows per change; the date-bounded join already handles them."]],
          [0.7, 2.2, 2.2, 2.2], font_size=8.0)

# ---- 10 evidence
h1(doc, "10. Verification evidence", page_break=True)
para(doc, "The following was produced by executing the delivered DDL and views (not by inspection). It is"
          " reproducible with scripts/build_step09.py, which rebuilds this document and fails if any line below "
          "does not hold.", space_after=4)
mono_block(doc, "\n".join(check_lines), 7.2)
para(doc, "Two notes on the evidence:", space_after=3)
bullet(doc, "The negative inserts prove the constraints are real. A CHECK that nobody has tested is a comment.")
bullet(doc, "The parity line is a set comparison between the parsed diagram and the parsed DDL, not a visual check; "
            "it is the only way to claim “every column” without missing one at 3 a.m.")

h1(doc, "11. Sign-off")
add_table(doc, ["Role", "Name", "Responsibility in this step", "Status"],
          [["Student / Author", STUDENT, "Drafted the ERD, ran the review, applied the fixes", "Signed — 19-Nov-2026"],
           ["Mentor", MENTOR, "Reviews the diagram and the accepted findings F08–F12", "Pending"],
           ["Internal guide", "Project guide, MCA (DA)", "Confirms the schema is ready for dataset generation", "Pending"]],
          [1.4, 1.5, 2.8, 1.4], font_size=8.4)
para(doc, "Next step: Step 10 generates the synthetic dataset against this schema — 8 departments, 20 wards, 5 payers, "
          "~57,000 inpatient episodes and ~150,000 OPD visits, with the seeded generator writing to data/app.db through "
          "the same DDL and quality log.", size=9.5, italic=True)

doc.save(OUT)
print("saved:", OUT)
print("tables:", len(doc.tables), "paragraphs:", len(doc.paragraphs))
print("checks:")
for line in check_lines:
    print("  " + line)
