"""Build Step 7 deliverable: Technical Requirements Document (TRD) for p107."""
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
CODE_FILL = "F4F6F8"

PROJECT_ID = "p107"
PROJECT_TITLE = "Healthcare KPI Dashboard"
STUDENT = "Ayush Pal"
ROLL = "2500320140037"
MENTOR = "Mr. Dinesh Sereno"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "07_TRD.docx")


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


def mono_block(doc, text, size=8.0):
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.autofit = False
    cell = table.rows[0].cells[0]
    cell.width = Inches(6.6)
    shade(cell, CODE_FILL)
    cell.text = ""
    lines = text.split("\n")
    p = cell.paragraphs[0]
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
    run = p.add_run(f"{PROJECT_ID}  |  Step 7 — Technical Requirements Document  |  Page ")
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
para(doc, "STEP 7", 26, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "TECHNICAL REQUIREMENTS DOCUMENT (TRD)", 17, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18)
para(doc, PROJECT_TITLE, 17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Project ID: {PROJECT_ID}   |   Domain: Data Analytics", 11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=34)

meta = [("Student", STUDENT), ("Roll Number", ROLL), ("Mentor", MENTOR),
        ("Deliverable", "07_TRD.docx (Step 7 of 18)"),
        ("Version / Date", "v1.0 — 19 November 2026"),
        ("Inputs", "Steps 1–5 (Understanding, HMW, Needs, RTM, BRD)"),
        ("Stack", "Python 3.12 · pandas · NumPy · SciPy/statsmodels · Plotly · Streamlit (multipage) · SQLite · pytest + AppTest · GitHub · Streamlit Community Cloud"),
        ("Status", "Baseline — approved for Step 8 (database design)")]
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

# ---------------- control ----------------
doc.add_page_break()
h1(doc, "Document Control")
h2(doc, "Revision history")
add_table(doc, ["Version", "Date", "Author", "Change description", "Status"],
          [["0.1", "19-Nov-2026", STUDENT, "Architecture, folder structure and FR set drafted", "Draft"],
           ["0.2", "19-Nov-2026", STUDENT, "Statistical method added per OQ; NFRs issued; pins verified on PyPI", "Draft"],
           ["1.0", "19-Nov-2026", STUDENT, "Baseline: 22 FRs, 12 NFRs, 10 technical risks", "Approved"]],
          [0.8, 1.0, 1.1, 3.1, 0.8])

h2(doc, "Baseline inputs")
add_table(doc, ["Artefact", "Version", "Used for"],
          [["01_Project_Understanding.docx", "1.0", "OQ-01…OQ-10, stakeholders, objectives, SC"],
           ["02_HMW_Statements.docx", "1.0", "HMW context behind each module"],
           ["03_Needs_Analysis.docx", "1.0", "BR-01…BR-15, priorities, acceptance measures"],
           ["04_Traceability_Matrix.xlsx", "1.0", "OQ → HMW → BR rows (16 links)"],
           ["05_BRD.docx", "1.0", "Business rules BRL-01…BRL-14, data requirements"]],
          [2.3, 0.7, 3.6], font_size=8.6)

h2(doc, "Change note — Step 6 not delivered")
para(doc, "Step 6 (PRD) was not submitted before this step. STEP 7 was requested directly, so this TRD "
          "issues the functional (FR-xx) and non-functional (NFR-xx) identifiers that the PRD would normally "
          "have introduced. Recorded as change note CN-STEP6-01.", bold=False)
add_table(doc, ["Item", "Effect", "Action required"],
          [["FR-01…FR-22 issued here", "Functional layer exists from Step 7 onward; no FR is invented later.", "If a PRD is produced later, it must reference these FR IDs, not create new ones."],
           ["NFR-01…NFR-12 issued here", "Quality attributes are baselined now so Step 13 can write TC-T cases.", "No action."],
           ["RTM 'Feature' column", "Still empty: it expects F-xx from the PRD. FR-xx are technical, not features.", "Fill Feature with F-xx if Step 6 is delivered; otherwise map BR → FR using §3.4."],
           ["RTM update required", "Yes: add FR IDs to a new column or a mapping sheet when the next RTM revision is made.", "Mentor to confirm at Step 8."]],
          [1.5, 3.0, 2.1], font_size=8.4)

# ---------------- 1 architecture ----------------
h1(doc, "1. Architecture", page_break=True)
para(doc, "The product is a read-only analytical application: a static SQLite file built offline, queried by a "
          "Streamlit multipage front end that performs all aggregation in pandas and all statistics in "
          "SciPy/statsmodels. There is no write path, no live feed and no user-generated content.", space_after=6)

h2(doc, "1.1 End-to-end flow (Mermaid source)")
mono_block(doc, """flowchart LR
  subgraph Build time
    G1[scripts/generate_data.py<br/>seeded synthetic generator]
    G2[scripts/build_db.py<br/>schema + load + quality log]
    DB[(data/app.db<br/>SQLite, read-only)]
    G1 --> G2 --> DB
  end
  subgraph Request time
    L1[src/data_access.py<br/>SQL views + st.cache_data]
    L2[src/validate.py<br/>schema, nulls, ranges, referential checks]
    GATE{Validation gate}
    K1[src/kpi.py<br/>BOR, ALOS, readmission, revenue, mortality]
    K2[src/stats.py<br/>SciPy / statsmodels]
    K3[src/charts.py<br/>Plotly figures]
    UI[pages/ UI-01..UI-07<br/>Streamlit multipage<br/>Tailwind header + KPI cards via iframe]
    EX[src/export.py<br/>filtered CSV download]
    DB --> L1 --> L2 --> GATE
    GATE -- fail --> DQ[Data-quality panel<br/>BR-14: block or badge]
    GATE -- pass --> K1 --> K2 --> K3 --> UI --> EX
  end
  subgraph Quality
    T1[pytest + Streamlit AppTest] --> T2[GitHub Actions]
  end""", 7.2)
para(doc, "Note: the diagram is stored as Mermaid source so it can be rendered in the repository README; Word "
          "cannot render Mermaid natively.", size=9, italic=True)

h2(doc, "1.2 Component responsibilities")
add_table(doc, ["Layer", "Responsibility", "Module", "Must not"],
          [["Persistence", "Hold the versioned synthetic dataset as one SQLite file", "data/app.db", "Be written by the app at runtime"],
           ["Access", "Open read-only, run parameterised queries, cache results", "src/data_access.py", "Contain business rules or display logic"],
           ["Validation", "Check schema, nulls, ranges, referential integrity, BOR \u2264 100", "src/validate.py", "Silently drop bad records (BRL-12)"],
           ["Analysis", "Compute the six KPIs and their drivers from validated frames", "src/kpi.py", "Format currency or colours"],
           ["Statistics", "Run tests with assumption checks and return a result object", "src/stats.py", "Render text"],
           ["Visualisation", "Build Plotly figures from prepared frames only", "src/charts.py", "Query the database"],
           ["Presentation", "Compose pages, filters, KPIs, tables, downloads", "app.py, pages/*.py", "Re-implement a KPI formula"],
           ["Governance", "Definition registry and data-quality checks", "src/definitions.py, src/quality.py", "Allow an undefined KPI to render (BRL-14)"],
           ["Build-time", "Generate data, build the database, write the reject log", "scripts/*.py", "Ship in the app runtime path"],
           ["Test", "Unit tests on KPI maths, integration on queries, AppTest on pages", "tests/*.py", "Touch the production database"]],
          [0.95, 2.55, 1.6, 1.5], font_size=7.8)

h2(doc, "1.3 Design principles")
bullet(doc, "Single source of truth: every KPI formula exists once, in src/kpi.py; pages only call it (enables NFR-10).")
bullet(doc, "Definitions are data, not styling: the registry drives labels, tooltips and the block rule in BRL-14.")
bullet(doc, "Fail closed: if validation fails, the page shows the defect instead of a plausible-looking number.")
bullet(doc, "Stateless pages: filters live in URL query parameters so a view can be shared and reproduced.")
bullet(doc, "Offline builds, deterministic seeds: the same commit regenerates the same database (NFR-06).")

# ---------------- 2 folder structure ----------------
h1(doc, "2. Folder Structure", page_break=True)
mono_block(doc, """hclTraining/
├─ app.py                     # entry point: home / executive summary (UI-01)
├─ pages/
│  ├─ 1_Operations.py         # UI-02 occupancy & volume
│  ├─ 2_Departments.py        # UI-03 ALOS, readmission, patient mix
│  ├─ 3_Quality.py            # UI-04 mortality, drift, discharge mix
│  ├─ 4_Finance.py            # UI-05 payer mix, revenue per patient / bed-day
│  ├─ 5_Data_Quality.py       # UI-06 quality panel + reject log
│  └─ 6_Definitions.py        # UI-07 KPI registry & methodology
├─ src/
│  ├─ __init__.py
│  ├─ config.py               # paths, constants, target values (BRL-01, AS-10)
│  ├─ data_access.py          # read-only connection, SQL, st.cache_data wrappers
│  ├─ validate.py             # schema/null/range/referential checks
│  ├─ kpi.py                  # the six KPI families + decomposition
│  ├─ stats.py                # SciPy/statsmodels tests with assumption checks
│  ├─ charts.py               # Plotly figure builders
│  ├─ definitions.py          # KPI definition registry loader (BR-05)
│  ├─ quality.py              # 6 data-quality checks (BR-14)
│  ├─ export.py               # CSV export with formula-injection guard
│  └─ ui.py                   # Tailwind header/KPI cards (iframe), layout helpers
├─ data/
│  ├─ app.db                  # SQLite, read-only at runtime
│  └─ app.db.sha256           # integrity hash verified at start-up
├─ scripts/
│  ├─ generate_data.py        # seeded synthetic generator (Step 10)
│  └─ build_db.py             # schema + load + reject log
├─ tests/
│  ├─ conftest.py
│  ├─ test_kpi.py             # formula-level tests (1e-6 tolerance)
│  ├─ test_stats.py
│  ├─ test_validate.py
│  ├─ test_quality.py
│  └─ test_pages_apptest.py   # Streamlit AppTest smoke + interaction tests
├─ .streamlit/
│  ├─ config.toml             # theme, server settings
│  └─ secrets.toml            # NOT committed; only .example in repo
├─ .github/workflows/tests.yml
├─ requirements.txt
├─ requirements-dev.txt
├─ README.md
└─ docs/                      # this project's 18 deliverables""", 7.0)
para(doc, "Why this shape: pages are thin, src/ is testable without Streamlit, scripts/ never runs in the deployed "
          "app, and docs/ keeps the academic deliverables separate from the runtime code.", size=9.5, italic=True)

h2(doc, "2.1 File-to-requirement map")
add_table(doc, ["Path", "Primary FRs", "Primary BRs"],
          [["app.py", "FR-20, FR-21", "BR-04"],
           ["pages/1_Operations.py", "FR-05, FR-06, FR-14", "BR-01, BR-02, BR-03"],
           ["pages/2_Departments.py", "FR-07, FR-08, FR-13", "BR-06, BR-07, BR-11"],
           ["pages/3_Quality.py", "FR-12", "BR-13"],
           ["pages/4_Finance.py", "FR-09, FR-11", "BR-08, BR-12"],
           ["pages/5_Data_Quality.py", "FR-17, FR-19", "BR-14"],
           ["pages/6_Definitions.py", "FR-16, FR-15", "BR-05, BR-10"],
           ["src/kpi.py", "FR-05…FR-14", "BR-01…BR-13"],
           ["src/stats.py", "FR-15", "BR-10"]],
          [1.9, 2.0, 2.7], font_size=8.4)

# ---------------- 3 FRs ----------------
h1(doc, "3. Functional Requirements", page_break=True)
para(doc, "Twenty-two functional requirements, each assigned to exactly one owning module and traceable to at "
          "least one BR. User stories are issued at Step 11; the 'US' column names the persona the story will "
          "be written for, so nothing is invented twice.", space_after=4)

fr_rows = [
    ("FR-01", "Load the fact and dimension datasets from data/app.db over a read-only connection and expose them as typed DataFrames.", "M1 Access", "BR-01…BR-14", "SH-02"),
    ("FR-02", "Apply global filters (date range, department, ward/bed class, payer, age band, gender) consistently to every KPI, denominator and drill-down on all pages.", "M2 Filter", "BR-01, BR-04", "SH-02"),
    ("FR-03", "Show a clear empty state with the active filter summary when a filter combination returns no rows, without raising an exception.", "M2 Filter", "BR-04", "SH-02"),
    ("FR-04", "Label partial periods as \u201cin progress\u201d and exclude them from ranking and trend comparisons (BRL-08).", "M2 Filter", "BR-01, BR-09", "SH-05"),
    ("FR-05", "Compute bed occupancy per department for week and month, the gap to the 85% target in percentage points and in beds, and display both.", "M3 KPI", "BR-01", "SH-02"),
    ("FR-06", "Show ward-level occupancy for the next seven days using recorded scheduled discharges and confirmed admissions only (deterministic projection; no forecasting model, per X3).", "M3 KPI", "BR-02", "SH-06"),
    ("FR-07", "Compute ALOS by department with the case-mix normalisation basis stated beside the value.", "M3 KPI", "BR-06", "SH-03"),
    ("FR-08", "Compute the 30-day unplanned readmission rate by department and diagnosis group with the window and exclusion rule stated (BRL-04).", "M3 KPI", "BR-07", "SH-04"),
    ("FR-09", "Show payer mix share, its movement between periods, and revenue per patient for each of the five payer categories (BRL-05, BRL-06).", "M3 KPI", "BR-08", "SH-05"),
    ("FR-10", "Summarise OPD visits by weekday and month with peak-to-average ratio and per-department breakdown.", "M3 KPI", "BR-09", "SH-06"),
    ("FR-11", "Rank departments by revenue per bed-day and display the absolute and relative gap to the best performer (BRL-13).", "M3 KPI", "BR-12", "SH-05"),
    ("FR-12", "Show mortality rate and discharge-type mix by department with a three-month moving average and a drift flag at \u00b12 SD of the 12-month baseline (BR-13).", "M3 KPI", "BR-13", "SH-04"),
    ("FR-13", "Break admissions and OPD visits down by age band and gender for each department, as counts and shares.", "M3 KPI", "BR-11", "SH-03"),
    ("FR-14", "Decompose a KPI movement between two periods into department, payer and weekday contributions that reconcile to the total within \u00b10.5 pp.", "M3 KPI", "BR-03", "SH-01"),
    ("FR-15", "Run the documented statistical test for the selected OQ, display method, n, statistic, p-value, confidence interval and an association-not-causation caveat.", "M4 Stats", "BR-10", "SH-01, SH-04"),
    ("FR-16", "Render KPI title, definition, numerator, denominator and version from the definition registry; refuse to render a KPI that has no registry entry (BRL-14).", "M6 Governance", "BR-05", "SH-07"),
    ("FR-17", "Run the six data-quality checks for the selected period and show pass/fail with counts.", "M6 Governance", "BR-14", "SH-04, SH-08"),
    ("FR-18", "Mark any rate whose denominator is below 30 as low-confidence (BRL-10).", "M6 Governance", "BR-06, BR-07, BR-13", "SH-04"),
    ("FR-19", "Expose the rejected-record log with reason codes so a data defect can be traced to its source field (BRL-12).", "M6 Governance", "BR-14", "SH-08"),
    ("FR-20", "Render the page header and KPI cards through Tailwind in an st.iframe component, with values passed in server-side (no client fetch).", "M7 UI", "BR-04", "SH-01"),
    ("FR-21", "Download the currently filtered dataset as CSV, guarded against spreadsheet formula injection.", "M8 Export", "BR-14 (evidence), S7", "SH-04"),
    ("FR-22", "Serve a responsive layout at 375 px, 768 px and 1440 px without horizontal scrolling or clipped labels.", "M7 UI", "BR-04", "SH-06"),
]
add_table(doc, ["FR", "Functional requirement", "Module", "Linked BR", "US persona"],
          [[a, b, c, d, e] for a, b, c, d, e in fr_rows],
          [0.45, 3.8, 0.85, 0.95, 0.6], font_size=7.4)

h2(doc, "3.1 Modules")
add_table(doc, ["Module", "Owns", "Key files", "Depends on"],
          [["M1 Access", "Connection, queries, caching", "src/data_access.py", "SQLite file"],
           ["M2 Filter", "Filter state, URL sync, period logic", "src/ui.py, pages/*.py", "M1"],
           ["M3 KPI", "All KPI mathematics and decomposition", "src/kpi.py", "M1, M2"],
           ["M4 Stats", "Statistical tests and assumption checks", "src/stats.py", "M3"],
           ["M5 Charts", "Plotly figure construction", "src/charts.py", "M3, M4"],
           ["M6 Governance", "Definitions, quality checks, low-confidence rule", "src/definitions.py, src/quality.py", "M1"],
           ["M7 UI", "Layout, Tailwind header/cards, theme", "src/ui.py, .streamlit/config.toml", "M5, M6"],
           ["M8 Export", "CSV export and injection guard", "src/export.py", "M2, M3"]],
          [0.95, 2.2, 2.0, 1.45], font_size=8.2)

h2(doc, "3.2 Functional coverage of business requirements")
add_table(doc, ["BR", "Covered by", "Gap"],
          [["BR-01", "FR-02, FR-04, FR-05", "\u2014"],
           ["BR-02", "FR-06", "\u2014"],
           ["BR-03", "FR-14", "\u2014"],
           ["BR-04", "FR-02, FR-03, FR-20, FR-22", "\u2014"],
           ["BR-05", "FR-16", "\u2014"],
           ["BR-06", "FR-07, FR-18", "\u2014"],
           ["BR-07", "FR-08, FR-18", "\u2014"],
           ["BR-08", "FR-09", "\u2014"],
           ["BR-09", "FR-04, FR-10", "\u2014"],
           ["BR-10", "FR-15", "Depth feature; Could priority"],
           ["BR-11", "FR-13", "Could priority"],
           ["BR-12", "FR-11", "\u2014"],
           ["BR-13", "FR-12, FR-18", "\u2014"],
           ["BR-14", "FR-17, FR-19, FR-21", "\u2014"],
           ["BR-15", "\u2014", "Intentional: Won't (X7). No FR may exist for it."]],
          [0.7, 3.4, 2.5], font_size=8.2)
para(doc, "14 of 15 BRs have at least one FR. BR-15 has none by design; if an FR appears for it later, the "
          "scope decision in Step 1 §8.2 has been overturned and must be recorded as a change.", size=9.5,
     italic=True)

# ---------------- 4 data layer ----------------
h1(doc, "4. Data Layer", page_break=True)
h2(doc, "4.1 Connection and access rules")
bullet(doc, "One SQLite file at data/app.db, opened read-only: sqlite3.connect('file:data/app.db?mode=ro&immutable=1', uri=True).")
bullet(doc, "All SQL uses bound parameters (:start_date, :dept_id); no f-string SQL, no user input concatenation.")
bullet(doc, "Aggregation happens in SQL views for heavy groupings and in pandas for KPI composition; the app never loads a fact table without a period predicate.")
bullet(doc, "Every query returns a DataFrame with explicit dtypes set after read (category for department/payer, datetime64 for dates).")
bullet(doc, "The database hash in data/app.db.sha256 is verified at start-up; a mismatch blocks the app with an explicit error (protects against a stale or edited file).")

h2(doc, "4.2 Caching strategy (st.cache_data)")
add_table(doc, ["Cached function", "Returns", "Key", "TTL", "Invalidated by"],
          [["load_dimensions()", "3 small dimension frames", "none (static)", "24 h", "DB hash change"],
           ["load_period_bounds()", "min/max dates, available months", "none", "24 h", "DB hash change"],
           ["kpi_occupancy(start, end, depts, wards)", "aggregated frame", "all filter args", "1 h", "filter change"],
           ["kpi_flow(start, end, depts)", "admissions/discharges/OPD aggregates", "all filter args", "1 h", "filter change"],
           ["stats_result(oq_code, start, end, depts)", "result object (method, n, stat, p, ci)", "all filter args", "1 h", "filter change"],
           ["quality_report(start, end)", "6-check pass/fail + counts", "period", "1 h", "DB hash change"]],
          [1.85, 1.85, 1.2, 0.55, 1.15], font_size=7.8)
para(doc, "Rule: cache the aggregation, not the figure. Figures are cheap to build and expensive to keep stale. "
          "Cache keys must include every argument that changes the result; a missing key is a correctness bug, "
          "not a performance issue.", size=9.5, italic=True)

h2(doc, "4.3 Validation contract (fail-closed)")
add_table(doc, ["Check", "Rule", "On failure"],
          [["Schema", "Required tables and columns exist; dtypes match expectation", "Block page; show missing column names"],
           ["Referential integrity", "Every department_id, ward_id, payer_id resolves", "Block; list orphan counts (BR-14)"],
           ["Date coverage", "Contiguous dates within the declared range; no future dates", "Warn; mark affected periods incomplete"],
           ["Range", "LOS \u2265 0, age 0–120, BOR \u2264 100, billed_amount \u2265 0", "Reject row, log reason code (BRL-12)"],
           ["Null policy", "Mandatory: department, dates, discharge type, payer. Optional: comorbidity flags", "Block if mandatory; badge if optional null > 2%"],
           ["Duplicates", "One row per episode_id; no duplicate patient-day registrations", "Block; show duplicate count"],
           ["Integrity hash", "SHA-256 of app.db matches the committed hash", "Block with instructions to rebuild"]],
          [1.1, 3.3, 2.2], font_size=7.8)

h2(doc, "4.4 Query patterns")
mono_block(doc, """# src/data_access.py (shape, not final code)
SQL_OCCUPANCY = '''
SELECT d.department_name, c.month,
       c.bed_days_occupied, c.bed_days_available,
       ROUND(100.0 * c.bed_days_occupied / NULLIF(c.bed_days_available, 0), 4) AS bor_pct
FROM fact_bed_capacity c
JOIN dim_department d USING (department_id)
WHERE c.month BETWEEN :start AND :end
  AND (:dept_count = 0 OR c.department_id IN (SELECT value FROM json_each(:dept_json)))
'''
@st.cache_data(ttl=3600, show_spinner=False)
def kpi_occupancy(start, end, dept_ids):
    with connect_ro() as con:
        return pd.read_sql(SQL_OCCUPANCY, con, params={...},
                           parse_dates=['month'])""", 7.6)

# ---------------- 5 statistics ----------------
h1(doc, "5. Statistical Method per Operational Question", page_break=True)
para(doc, "Method choice is fixed here so that Step 13 can test it and the viva cannot ambush the analysis. "
          "Every test reports n, the assumption check performed, and the fallback actually used. Alpha = 0.05; "
          "where more than two groups are compared, Bonferroni correction applies.", space_after=4)
add_table(doc, ["OQ", "Question type", "Primary method / test", "Assumption check \u2192 fallback", "Output shown"],
          [["OQ-01", "Ratio vs target", "Ratio estimator with 1,000-sample bootstrap 95% CI on monthly BOR", "Independence of bed-days (documented, not tested)", "BOR, gap in pp and beds, CI"],
           ["OQ-02", "Time series / trend", "ISO-week aggregation, WoW % change, robust z (MAD) for outliers", "None required", "Trend line, WoW table, flagged outlier weeks"],
           ["OQ-03", "Group comparison", "One-way ANOVA on log(LOS)", "Shapiro-Wilk on residuals; if p < 0.05 \u2192 Kruskal-Wallis + pairwise Mann-Whitney (Bonferroni)", "Box plot, medians, p-value, method used"],
           ["OQ-04", "Proportion + association", "Wilson 95% CI; chi-square test of homogeneity across departments", "Expected cell count \u2265 5; else Fisher exact for 2\u00d72", "Rate, CI, contribution %, p-value"],
           ["OQ-05", "Distribution by group", "Kruskal-Wallis on revenue per discharge; median + IQR", "None (rank-based)", "Payer bars, medians, p-value, share table"],
           ["OQ-06", "Seasonality / weekday", "One-way ANOVA on daily visits; peak-to-average ratio", "Shapiro-Wilk \u2192 Kruskal-Wallis; seasonal_decompose (period 7) for pattern", "Heatmap, ratio, p-value"],
           ["OQ-07", "Association (two continuous)", "Spearman \u03c1 (primary) + Pearson r (secondary)", "Linearity/monotonicity by scatter inspection; report both, never only the stronger", "Scatter + fitted line, r, \u03c1, p, n, caveat"],
           ["OQ-08", "Categorical association", "Chi-square test of independence (age band \u00d7 department)", "Expected count \u2265 5; Cram\u00e9r's V as effect size", "Stacked bars, p-value, V"],
           ["OQ-09", "Ranking with uncertainty", "Revenue per bed-day, rank with bootstrap 95% CI per department", "Independence of bed-days; overlapping CIs flagged as \u201cnot distinguishable\u201d", "Ranked bar + CI, gap to best"],
           ["OQ-10", "Rate over time", "Poisson GLM with month trend; p-chart (3\u03c3) for drift", "Overdispersion check (Pearson \u03c7\u00b2/df); if > 1.5 \u2192 negative binomial", "Moving average, control limits, flag count"]],
          [0.5, 1.15, 2.05, 1.85, 1.35], font_size=7.2)
para(doc, "Interpretation guardrails enforced in the UI: no p-value without n; no ranking without a CI or a "
          "normalisation statement; no correlation presented as causation (label FR-15); a test on fewer than "
          "30 observations is labelled indicative only (BRL-10).", size=9.5, italic=True)

h2(doc, "5.1 Why not a t-test everywhere")
para(doc, "Hospital KPI distributions are skewed (a few very long stays, a few very large bills) and "
          "denominators vary per department. A t-test on raw LOS or revenue would violate normality and "
          "equal-variance assumptions; the table therefore defaults to rank-based methods and only uses "
          "parametric tests after a documented check.")

# ---------------- 6 UI ----------------
h1(doc, "6. UI and Presentation Notes", page_break=True)
h2(doc, "6.1 Tailwind approach (and its constraint)")
para(doc, "Tailwind is loaded from the CDN inside an isolated iframe built with st.components.v1.html. The "
          "iframe renders the page header and the KPI card grid only. Values are interpolated server-side into "
          "the HTML string, so there is no client-side fetch and no API surface.")
add_table(doc, ["Constraint", "Consequence", "Mitigation"],
          [["An iframe is a separate document", "Tailwind classes do not affect Streamlit widgets", "Keep Tailwind to header and cards; style the rest with .streamlit/config.toml theme"],
           ["Iframe height is not auto", "Content can clip or scroll", "Compute height from card count (e.g. rows \u00d7 132 + 96 px) and pass explicitly"],
           ["CDN availability", "Offline or blocked CDN leaves cards unstyled", "Include a minimal inline CSS fallback for the card skeleton"],
           ["No Streamlit state inside iframe", "Cards cannot hold widgets", "Cards are display-only; all interactivity stays in Streamlit widgets"],
           ["Print/PDF export", "Iframe may not print reliably", "Provide the on-screen table as the print view; CSV is the portable artefact"]],
          [1.7, 2.2, 2.7], font_size=7.8)

h2(doc, "6.2 Design tokens")
add_table(doc, ["Token", "Value", "Use"],
          [["Primary", "#1F3864", "Header background, active nav"],
           ["Accent", "#2E5496", "Links, selected tab"],
           ["Success / warn / danger", "#1E7B4F / #B7791F / #B42318", "KPI status: on target / watch / breach"],
           ["Surface", "#FFFFFF on #F4F6F8", "Cards on page background"],
           ["Type scale", "28 / 20 / 14 / 12 px", "Page title / section / body / caption"],
           ["Spacing", "4 px grid, card padding 20 px, gap 16 px", "Consistent rhythm"],
           ["Radius", "12 px cards, 8 px controls", "Soft, modern"],
           ["Font", "Inter (fallback system-ui)", "Loaded with the same CDN block, with a system fallback"]],
          [1.5, 2.1, 3.0], font_size=8.2)

h2(doc, "6.3 Responsive layout")
add_table(doc, ["Breakpoint", "Grid", "Behaviour"],
          [["\u2265 1440 px", "4 cards per row, 12-col content", "Full table columns, side legend"],
           ["1024–1439 px", "4 cards per row", "Table columns drop to 6"],
           ["768–1023 px", "2 cards per row", "Charts stack; filter bar collapses to two rows"],
           ["375–767 px", "1 card per row", "Charts min-height 280 px; tables become cards; no horizontal scroll"]],
          [1.4, 1.6, 3.6], font_size=8.2)
para(doc, "Accessibility: contrast \u2265 4.5:1 for text, never encode meaning by colour alone (breach also shows an "
          "icon and label), keyboard-reachable filters, visible focus ring, reduced-motion respected, and every "
          "chart carries a text summary beneath it.", size=9.5, italic=True)

# ---------------- 7 quality attributes ----------------
h1(doc, "7. Performance, Error Handling, Security", page_break=True)
h2(doc, "7.1 Non-functional requirements")
add_table(doc, ["NFR", "Requirement", "Target", "How verified"],
          [["NFR-01", "Filter response time", "p95 \u2264 2 s on 20 runs", "Browser timing log during UAT"],
           ["NFR-02", "Cold start", "\u2264 10 s on a sleeping app", "5 stop/start cycles, median"],
           ["NFR-03", "Memory ceiling", "\u2264 1 GB RSS; dataset \u2264 200 MB", "Local profiling + Cloud metrics"],
           ["NFR-04", "Availability", "Loads whenever Streamlit Cloud is awake; sleep message is explicit", "Daily spot check for 7 days"],
           ["NFR-05", "Correctness", "KPI values match fixtures to 1e-6", "pytest test_kpi.py"],
           ["NFR-06", "Reproducibility", "Same commit + seed \u2192 identical database hash", "Rebuild twice, compare SHA-256"],
           ["NFR-07", "Usability", "Median \u2264 60 s; SUS \u2265 75", "8-user moderated test"],
           ["NFR-08", "Accessibility", "WCAG 2.1 AA contrast; keyboard path complete; 375 px usable", "Checklist + keyboard-only run"],
           ["NFR-09", "Security", "No PII; read-only DB; secrets out of git", "Code review + .gitignore audit"],
           ["NFR-10", "Maintainability", "\u2265 85% coverage; no KPI formula outside src/kpi.py", "coverage report + grep audit"],
           ["NFR-11", "Access control", "Optional passphrase, constant-time comparison, 3 attempts", "Manual test when enabled"],
           ["NFR-12", "Observability", "Structured log line per page render; user-safe error codes", "Log inspection after UAT"]],
          [0.65, 1.6, 2.1, 2.25], font_size=7.8)

h2(doc, "7.2 Performance techniques")
add_table(doc, ["Technique", "Where", "Effect"],
          [["Pre-aggregated SQL views (monthly, weekly, ward-day)", "scripts/build_db.py", "Removes the largest group-bys from request time"],
           ["st.cache_data with correct keys", "src/data_access.py", "Repeat views become near-instant"],
           ["Column pruning + category dtypes", "src/data_access.py", "Cuts dataframe memory by 40–60%"],
           ["Point limits on scatter/heatmap (\u2264 5,000 marks)", "src/charts.py", "Keeps Plotly rendering under 300 ms"],
           ["Stats computed once per filter, not per widget", "src/stats.py", "Avoids re-running tests on every rerun"],
           ["Lazy imports for SciPy/statsmodels", "src/stats.py", "Faster cold start; heavy import only on the statistics page"]],
          [2.0, 1.7, 2.9], font_size=8.0)

h2(doc, "7.3 Error handling")
add_table(doc, ["Failure", "Detection", "User experience", "Log"],
          [["Database missing or unreadable", "Connection open fails", "Blocking error card with rebuild instructions", "ERROR + stack (server log only)"],
           ["Hash mismatch", "SHA-256 compare at start-up", "Blocking error naming the expected hash", "ERROR"],
           ["Validation failure", "src/validate.py", "Quality panel lists failed checks; KPI cards show a defect badge", "WARN with check id"],
           ["Empty filter result", "Row count = 0", "Empty state + active filter summary (FR-03)", "INFO"],
           ["Statistical test failure (singular matrix, n<3)", "Exception in src/stats.py", "Method card shows \u201cinsufficient data\u201d, page keeps working", "WARN"],
           ["CSV export error", "Write attempt", "Toast with retry; no partial file", "ERROR"],
           ["Unhandled exception", "Top-level handler in app.py", "Friendly panel with correlation id; no traceback to user", "ERROR + correlation id"]],
          [1.6, 1.45, 2.35, 1.2], font_size=7.8)

h2(doc, "7.4 Security")
bullet(doc, "No real patient data anywhere; patient keys are synthetic and non-reversible (CN-04).")
bullet(doc, "Database opened with mode=ro&immutable=1: the app cannot write even if compromised.")
bullet(doc, "Secrets only in Streamlit Cloud secrets or a local .streamlit/secrets.toml excluded by .gitignore; repository contains secrets.toml.example only.")
bullet(doc, "Passphrase compared with hmac.compare_digest; attempt counter in session state; no credential is logged (NFR-11).")
bullet(doc, "CSV export prefixes cells beginning with =, +, -, @ with an apostrophe (formula-injection guard, FR-21).")
bullet(doc, "Tailwind/KPI HTML is built from numeric values and registry text, never from raw user input; any future text input must be escaped.")
bullet(doc, "Dependencies pinned exactly (see §8) and reviewed before any bump; no post-install scripts, no unpinned CDN beyond Tailwind's versioned build.")

# ---------------- 8 requirements ----------------
h1(doc, "8. requirements.txt (Pinned)", page_break=True)
para(doc, "Versions were verified against PyPI on 19-Nov-2026 and are compatible with Python 3.12 (the runtime "
          "on Streamlit Community Cloud). Runtime pins go in requirements.txt; test-only tools go in "
          "requirements-dev.txt so the deployed image stays small.", space_after=4)
mono_block(doc, """# requirements.txt  — runtime only, exact pins
streamlit==1.67.0
pandas==3.1.1
numpy==2.5.5
scipy==1.18.1
statsmodels==0.15.0
plotly==7.1.0""", 9)
mono_block(doc, """# requirements-dev.txt  — local and CI only
-r requirements.txt
pytest==9.2.0
pytest-cov==7.1.0
flake8==7.3.0""", 9)
add_table(doc, ["Package", "Why it is required", "Removal impact"],
          [["streamlit", "Multipage app shell, widgets, caching, AppTest", "No product"],
           ["pandas", "Aggregation, joins, pivots, dtype control", "No analysis layer"],
           ["numpy", "Vector maths, bootstrap sampling, MAD", "Slower/verbose computations"],
           ["scipy", "Shapiro-Wilk, Mann-Whitney, Kruskal-Wallis, chi-square, Pearson/Spearman", "OQ-03…OQ-09 statistics become manual"],
           ["statsmodels", "Poisson GLM, seasonal decomposition, OLS with robust SE", "OQ-06 and OQ-10 lose the trend model"],
           ["plotly", "Interactive charts (hover, zoom, export)", "Static tables only"],
           ["pytest / pytest-cov", "Test execution and coverage gate (NFR-10)", "No automated proof of correctness"],
           ["flake8", "Static check for E9/F errors (SC-06)", "Style/correctness lint lost"]],
          [1.35, 3.55, 1.7], font_size=8.0)
para(doc, "Pin policy: no range specifiers, no latest-latest. A version bump is a change request with a test run "
          "and a note in the change log; the database hash and test results must be regenerated after any bump.",
     size=9.5, italic=True)

# ---------------- 9 deployment ----------------
h1(doc, "9. Deployment Requirements", page_break=True)
add_table(doc, ["Item", "Requirement"],
          [["Repository", "GitHub repository hclTraining, branch main, app at repository root (app.py)"],
           ["Platform", "Streamlit Community Cloud, free tier, Python 3.12 runtime"],
           ["Build command", "pip install -r requirements.txt (platform default)"],
           ["Entry point", "app.py (multipage via pages/ directory)"],
           ["Data file", "data/app.db committed to the repo (must stay under GitHub's 100 MB file limit; target \u2264 60 MB)"],
           ["Config", ".streamlit/config.toml committed (theme, server.headless); .streamlit/secrets.toml never committed"],
           ["Secrets", "APP_PASSPHRASE (optional) and LOG_LEVEL stored in the Cloud secrets editor"],
           ["Environment", "No environment variables required beyond secrets; no local file paths in code (use pathlib relative to repo root)"],
           ["Region / resources", "Default region; ~1 GB RAM, shared CPU; no always-on guarantee"],
           ["Cold start", "\u2264 10 s after sleep; loading placeholder shown"],
           ["Verification after deploy", "7 pages load, filters work, CSV downloads, logs clean for 7 days (SC-09)"],
           ["Rollback", "Revert to the previous commit on main; Cloud redeploys within ~2 minutes"],
           ["Monitoring", "Cloud logs reviewed after each deploy; error correlation ids searched weekly"],
           ["Accessibility of evidence", "Public URL plus a recorded walkthrough kept for the viva (in case the app is asleep)"]],
          [1.45, 5.15], font_size=8.2)

h2(doc, "9.1 Build and release sequence")
add_table(doc, ["Step", "Command / action", "Gate"],
          [["1", "python scripts/generate_data.py --seed 42", "Row counts \u2265 OBJ-02 targets"],
           ["2", "python scripts/build_db.py", "Quality log has 0 unexplained rejects; BOR \u2264 100"],
           ["3", "pytest --cov=src --cov-fail-under=85", "All tests pass, coverage \u2265 85%"],
           ["4", "flake8 src app.py pages --select=E9,F", "0 findings"],
           ["5", "Streamlit AppTest smoke on all 7 pages", "0 exceptions on default filters"],
           ["6", "git push origin main", "CI green"],
           ["7", "Streamlit Cloud redeploy", "All pages load; cold start \u2264 10 s"],
           ["8", "7-day log review", "0 unhandled exceptions (SC-09)"]],
          [0.5, 3.7, 2.4], font_size=8.2)

# ---------------- 10 risks ----------------
h1(doc, "10. Technical Risks")
add_table(doc, ["TR", "Risk", "L", "I", "Mitigation / trigger"],
          [["TR-01", "Iframe height mismatch clips KPI cards on mobile", "M", "M", "Height function + 375 px test in UAT; fallback to st.markdown cards if clipping persists"],
           ["TR-02", "Tailwind CDN blocked, cards render unstyled", "L", "M", "Inline skeleton CSS fallback; visual check with CDN blocked"],
           ["TR-03", "SQLite file exceeds GitHub's 100 MB limit", "L", "H", "Target \u2264 60 MB; check size in CI; compress by dropping unused columns"],
           ["TR-04", "Streamlit Cloud memory spike during bootstrap resampling", "L", "M", "Cap resamples at 1,000, operate on aggregated frames, profile before release"],
           ["TR-05", "pandas 3.x behaviour change (copy-on-write, string dtype)", "M", "M", "Explicit dtypes and .loc assignment; tests catch silent regressions"],
           ["TR-06", "Statistical result changes when data is regenerated with a different seed", "M", "M", "Seed stored in DB metadata and displayed in the About page"],
           ["TR-07", "Definition drift between registry and kpi.py", "M", "H", "One registry row tested against one function output (FR-16 test)"],
           ["TR-08", "Cache returns stale results after a filter key is missed", "M", "H", "Cache-key unit test; cache cleared on DB hash change"],
           ["TR-09", "Free-tier app sleeps during the viva", "H", "M", "Warm the app 10 minutes before; keep a recorded walkthrough"],
           ["TR-10", "Dependency bump breaks deployment", "L", "H", "Exact pins; bump only as a change request; rollback commit ready"]],
          [0.55, 2.6, 0.35, 0.35, 2.75], font_size=7.6)

h1(doc, "11. Step 7 Sign-off")
bullet(doc, "22 functional requirements (FR-01…FR-22) across 8 modules; 14 of 15 BRs covered; BR-15 intentionally uncovered.")
bullet(doc, "12 non-functional requirements (NFR-01…NFR-12) with measurable verification methods.")
bullet(doc, "Architecture, folder structure, data contract, statistical method per OQ, UI constraints, performance, error handling, security, pinned dependencies and deployment gates all fixed.")
bullet(doc, "Change note CN-STEP6-01 records that FR/NFR IDs were issued at Step 7 because Step 6 was not delivered; RTM update required.")
para(doc, "Prepared by: " + STUDENT + "   |   Roll: " + ROLL, size=10)
para(doc, "Reviewed by (mentor): " + MENTOR + "   |   Date: 19-Nov-2026", size=10)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print("saved:", OUT)
