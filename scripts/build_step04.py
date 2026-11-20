"""Build Step 4 deliverable: Requirements Traceability Matrix (Excel) for HKD-2026."""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "04_Traceability_Matrix.xlsx")

NAVY = "1F3864"
BLUE = "2E5496"
LIGHT = "EAF0F8"
MUST = "DDEBF7"
WONT = "F2F2F2"
WHITE = "FFFFFF"

thin = Side(style="thin", color="B4C6E7")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---------------- source data (from Steps 1-3, IDs frozen) ----------------
# (hmw, oq, br, requirement text, sh, priority)
LINKS = [
    ("HMW-01", "OQ-01", "BR-01",
     "The business needs to see each department's bed occupancy against the 85% target for the current and next week, with the shortfall expressed in beds and bed-days, so the operations manager can escalate a likely miss before the 8 a.m. huddle.",
     "SH-02", "Must"),
    ("HMW-01", "OQ-01", "BR-15",
     "The business needs threshold breaches pushed automatically to managers without anyone opening the tool, so a saturation or mortality-drift event is actioned within minutes.",
     "SH-02, SH-04", "Won't"),
    ("HMW-02", "OQ-01", "BR-02",
     "The business needs ward-level occupancy projected across the next seven days, so the nursing superintendent can move staff before a ward crosses 90% instead of after handover.",
     "SH-06", "Must"),
    ("HMW-03", "OQ-02", "BR-03",
     "The business needs any KPI movement between two periods decomposed into department, payer and weekday contributions, so the director can name the cause in the board meeting.",
     "SH-01", "Must"),
    ("HMW-04", "OQ-02", "BR-04",
     "The business needs any operational question answerable in no more than three interactions, so a first-time user can self-serve within 60 seconds without knowing the data model.",
     "SH-02, SH-10", "Must"),
    ("HMW-04", "OQ-02", "BR-05",
     "The business needs every published figure to carry its definition, denominator and data period in the same view, so two people never quote different numbers for the same week.",
     "SH-02, SH-07", "Must"),
    ("HMW-05", "OQ-03", "BR-06",
     "The business needs length-of-stay comparison between departments on a like-for-like basis with case-mix context stated, so department heads accept the benchmark and commit to action.",
     "SH-03", "Must"),
    ("HMW-06", "OQ-04", "BR-07",
     "The business needs the 30-day unplanned readmission rate by department and diagnosis group with the period's data-quality context visible, so the quality officer can separate clinical signal from coding error.",
     "SH-04", "Must"),
    ("HMW-06", "OQ-04", "BR-14",
     "The business needs the data-quality status of each period (missing payer, duplicate registration, orphan department, date coverage, negative length of stay, occupancy above 100%) shown beside the KPI.",
     "SH-04, SH-08", "Should"),
    ("HMW-07", "OQ-05", "BR-08",
     "The business needs payer-mix share, its movement between periods and revenue per patient, so finance can act within the same quarter rather than after the negotiation window closes.",
     "SH-05", "Should"),
    ("HMW-08", "OQ-06", "BR-09",
     "The business needs outpatient demand summarised by weekday and season with a peak-to-average ratio, so rosters and consultant slots are planned before the queue forms.",
     "SH-06, SH-02", "Must"),
    ("HMW-09", "OQ-07", "BR-10",
     "The business needs a statistically defensible statement about whether occupancy pressure is associated with longer stays or higher readmissions, with method and assumptions disclosed.",
     "SH-01, SH-04", "Could"),
    ("HMW-10", "OQ-08", "BR-11",
     "The business needs admissions and outpatient volume broken down by age band and gender for each department, so service-line plans reflect the population actually served.",
     "SH-03", "Could"),
    ("HMW-11", "OQ-09", "BR-12",
     "The business needs revenue per bed-day ranked across departments with the gap to the best performer, so capacity and service-line decisions are evidence-based rather than historical.",
     "SH-05", "Should"),
    ("HMW-12", "OQ-10", "BR-13",
     "The business needs a three-month moving view of mortality and discharge-type mix by department with a drift indicator, so a quality review starts before the audit.",
     "SH-04", "Must"),
    ("HMW-12", "OQ-10", "BR-15",
     "The business needs threshold breaches pushed automatically to managers without anyone opening the tool, so a saturation or mortality-drift event is actioned within minutes.",
     "SH-02, SH-04", "Won't"),
]

HMW_TEXT = {
    "HMW-01": "help the operations manager see, before the 8 a.m. huddle, which departments will miss the 85% occupancy target",
    "HMW-02": "help the nursing superintendent know, at roster time, which wards will cross 90% occupancy",
    "HMW-03": "help the director explain a month-on-month movement by naming the department, payer and weekday responsible",
    "HMW-04": "help a first-time user get a trustworthy answer in under a minute without writing a query",
    "HMW-05": "let a department head compare its length of stay with peers on a like-for-like case mix",
    "HMW-06": "help the quality officer tell a genuine readmission spike from a registration or coding error",
    "HMW-07": "help the finance manager see a payer-mix shift early enough to act within the quarter",
    "HMW-08": "make weekday and seasonal outpatient peaks visible before rosters are fixed",
    "HMW-09": "show leadership whether occupancy pressure actually harms quality",
    "HMW-10": "surface which patient segments (age, gender) each department serves",
    "HMW-11": "make revenue per bed-day comparable across departments",
    "HMW-12": "help the quality officer notice slow mortality drift before the audit",
}

OQS = {f"OQ-{i:02d}": t for i, t in enumerate([
    "Which departments miss the 85% occupancy target?",
    "How are admissions, discharges and OPD trending week over week?",
    "Which departments exceed the 4.0-day ALOS benchmark?",
    "What is the 30-day readmission rate and who drives it?",
    "How does revenue per patient vary by payer category?",
    "Which weekdays and months carry peak OPD load?",
    "Is occupancy associated with longer stays or readmissions?",
    "Which age groups and genders dominate admissions?",
    "Which departments earn least revenue per bed-day?",
    "Is mortality and discharge-mix stable by department?"], start=1)}

BR_META = {}
for hmw, oq, br, txt, sh, pri in LINKS:
    if br not in BR_META:
        BR_META[br] = {"text": txt, "sh": sh, "pri": pri, "hmws": [], "oqs": []}
    BR_META[br]["hmws"].append(hmw)
    if oq not in BR_META[br]["oqs"]:
        BR_META[br]["oqs"].append(oq)

PENDING = {
    "Feature": "Step 6 (F-xx)",
    "DB table.column": "Step 8 (T-xx)",
    "User Story": "Step 11 (US-xx)",
    "UI Screen": "Step 12 (UI-xx)",
    "Test Case": "Step 13 (TC-F / TC-T)",
}

wb = Workbook()

# ================= RTM SHEET =================
ws = wb.active
ws.title = "RTM"
ws.sheet_view.showGridLines = False

headers = ["OQ", "HMW", "Req ID (BR)", "Requirement (WHAT the business needs)",
           "Stakeholder (SH)", "Priority", "Status",
           "Feature", "DB table.column", "User Story", "UI Screen",
           "Test Case (TC-F / TC-T)"]

ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
c = ws.cell(row=1, column=1)
c.value = "HKD-2026  \u00b7  Healthcare KPI Dashboard  \u2014  Requirements Traceability Matrix (RTM)"
c.font = Font(name="Calibri", size=14, bold=True, color=WHITE)
c.fill = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(vertical="center", horizontal="left", indent=1)
ws.row_dimensions[1].height = 26

ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
c = ws.cell(row=2, column=1)
c.value = ("v1.0  \u00b7  19-Nov-2026  \u00b7  one row per HMW\u2013BR link  \u00b7  "
           "columns H\u2013L stay empty until Steps 6, 8, 11, 12 and 13  \u00b7  "
           "Status = Identified for every row at this baseline")
c.font = Font(name="Calibri", size=9, italic=True, color="404040")
c.fill = PatternFill("solid", fgColor=LIGHT)
c.alignment = Alignment(vertical="center", horizontal="left", indent=1)
ws.row_dimensions[2].height = 18

HDR_ROW = 3
FIRST = 4
for i, h in enumerate(headers, start=1):
    cell = ws.cell(row=HDR_ROW, column=i, value=h)
    cell.font = Font(name="Calibri", size=9, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border
ws.row_dimensions[HDR_ROW].height = 30

for r, (hmw, oq, br, txt, sh, pri) in enumerate(LINKS, start=FIRST):
    values = [oq, hmw, br, txt, sh, pri, "Identified", "", "", "", "", ""]
    for ci, v in enumerate(values, start=1):
        cell = ws.cell(row=r, column=ci, value=v)
        cell.font = Font(name="Calibri", size=9)
        cell.border = border
        cell.alignment = Alignment(vertical="top", wrap_text=True,
                                   horizontal="center" if ci in (1, 2, 3, 6, 7) else "left")
    if pri == "Must":
        for ci in range(1, len(headers) + 1):
            ws.cell(row=r, column=ci).fill = PatternFill("solid", fgColor=MUST)
    elif pri == "Won't":
        for ci in range(1, len(headers) + 1):
            ws.cell(row=r, column=ci).fill = PatternFill("solid", fgColor=WONT)
    ws.row_dimensions[r].height = 54

LAST = FIRST + len(LINKS) - 1

widths = {"A": 8, "B": 9, "C": 10, "D": 64, "E": 13, "F": 9, "G": 11,
          "H": 11, "I": 20, "J": 11, "K": 10, "L": 15}
for col, w in widths.items():
    ws.column_dimensions[col].width = w

ws.freeze_panes = "D4"
ws.auto_filter.ref = f"A{HDR_ROW}:L{LAST}"

# data validation
dv_pri = DataValidation(type="list", formula1='"Must,Should,Could,Won\'t"', allow_blank=False,
                        showErrorMessage=True, errorTitle="Invalid priority",
                        error="Use Must, Should, Could or Won't.")
dv_status = DataValidation(type="list",
                           formula1='"Identified,In Design,In Build,In Test,Tested,Deferred"',
                           allow_blank=False, showErrorMessage=True,
                           errorTitle="Invalid status",
                           error="Use one of the controlled status values.")
dv_pending = DataValidation(type="list", formula1='"-,TBD,N/A"', allow_blank=True)
ws.add_data_validation(dv_pri)
ws.add_data_validation(dv_status)
ws.add_data_validation(dv_pending)
dv_pri.add(f"F{FIRST}:F{LAST}")
dv_status.add(f"G{FIRST}:G{LAST}")
dv_pending.add(f"H{FIRST}:L{LAST}")

# conditional formatting on status
ws.conditional_formatting.add(f"G{FIRST}:G{LAST}",
    FormulaRule(formula=[f'$G{FIRST}="Identified"'],
                fill=PatternFill("solid", fgColor="FCE4D6")))
ws.conditional_formatting.add(f"G{FIRST}:G{LAST}",
    FormulaRule(formula=[f'$G{FIRST}="Tested"'],
                fill=PatternFill("solid", fgColor="C6EFCE")))

# print setup
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = f"{HDR_ROW}:{HDR_ROW}"

# ================= COVERAGE SHEET =================
cs = wb.create_sheet("Coverage Check")
cs.sheet_view.showGridLines = False
cs.merge_cells("A1:F1")
c = cs["A1"]
c.value = "Coverage Check  \u00b7  HKD-2026  \u00b7  v1.0 (19-Nov-2026)"
c.font = Font(name="Calibri", size=13, bold=True, color=WHITE)
c.fill = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(vertical="center", indent=1)
cs.row_dimensions[1].height = 24

cs.merge_cells("A2:F2")
c = cs["A2"]
c.value = ("Gate for Step 4: every HMW must have at least one BR, every BR must have a source HMW, "
           "and every OQ must be reachable. Counts below are the live formulas; the verdict line is the v1.0 snapshot.")
c.font = Font(name="Calibri", size=9, italic=True, color="404040")
c.alignment = Alignment(wrap_text=True, vertical="center")
cs.row_dimensions[2].height = 28

summary = [
    ("Metric", "Value", "Check", "Verdict"),
    ("RTM rows (HMW\u2013BR links)", len(LINKS), "\u2014", "Baseline"),
    ("Unique HMWs", 12, "= 12 in Step 2", "OK"),
    ("Unique BRs", 15, "= 15 in Step 3", "OK"),
    ("Unique OQs", 10, "= 10 in Step 1", "OK"),
    ("HMWs with no BR", 0, "must be 0", "PASS"),
    ("BRs with no HMW", 0, "must be 0", "PASS"),
    ("OQs not covered", 0, "must be 0", "PASS"),
    ("Rows still 'Identified'", len(LINKS), "13 more steps of detail to add", "Expected"),
]
r0 = 4
for i, row in enumerate(summary):
    for j, v in enumerate(row, start=1):
        cell = cs.cell(row=r0 + i, column=j, value=v)
        cell.border = border
        if i == 0:
            cell.font = Font(name="Calibri", size=9, bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=BLUE)
            cell.alignment = Alignment(horizontal="center")
        else:
            cell.font = Font(name="Calibri", size=9, bold=(j == 4))
            cell.alignment = Alignment(horizontal="center" if j > 1 else "left")

cs.merge_cells(start_row=r0 + len(summary), start_column=1,
               end_row=r0 + len(summary), end_column=6)
note = cs.cell(row=r0 + len(summary), column=1)
note.value = ("Reading the counts: the RTM holds 16 rows for 15 unique BRs because BR-15 is linked to two "
              "HMWs (HMW-01 and HMW-12) and therefore occupies two rows. Priority counts by row are "
              "Must 9, Should 3, Could 2, Won't 2 rows (= 1 unique BR).")
note.font = Font(size=8.5, italic=True)
note.alignment = Alignment(wrap_text=True, vertical="center")
cs.row_dimensions[r0 + len(summary)].height = 30

# 1. HMW without BR
r = r0 + len(summary) + 2
cs.cell(row=r, column=1, value="1. HMW \u2192 BR check (live)").font = Font(bold=True, size=11, color=NAVY)
r += 1
hdrs = ["HMW", "Statement (abridged)", "Linked BR", "# Links", "Linked OQ", "Verdict"]
for j, h in enumerate(hdrs, start=1):
    cell = cs.cell(row=r, column=j, value=h)
    cell.font = Font(size=9, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.border = border
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
r += 1
for hmw in HMW_TEXT:
    links = [l for l in LINKS if l[0] == hmw]
    brs = ", ".join(sorted({l[2] for l in links}, key=lambda x: int(x.split("-")[1])))
    oqs = ", ".join(sorted({l[1] for l in links}))
    vals = [hmw, HMW_TEXT[hmw], brs, len(links), oqs,
            f'=IF(COUNTIF(RTM!$B:$B,"{hmw}")=0,"GAP: no BR","OK")']
    for j, v in enumerate(vals, start=1):
        cell = cs.cell(row=r, column=j, value=v)
        cell.font = Font(size=8.5)
        cell.border = border
        cell.alignment = Alignment(vertical="top", wrap_text=(j == 2),
                                   horizontal="center" if j in (1, 4, 5, 6) else "left")
    cs.row_dimensions[r].height = 30
    r += 1

# 2. BR without HMW
r += 1
cs.cell(row=r, column=1, value="2. BR \u2192 HMW check (live)").font = Font(bold=True, size=11, color=NAVY)
r += 1
hdrs = ["BR", "Priority", "Source HMW", "# Links", "Linked OQ", "Verdict"]
for j, h in enumerate(hdrs, start=1):
    cell = cs.cell(row=r, column=j, value=h)
    cell.font = Font(size=9, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.border = border
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
r += 1
for br in sorted(BR_META, key=lambda x: int(x.split("-")[1])):
    m = BR_META[br]
    vals = [br, m["pri"], ", ".join(sorted(set(m["hmws"]))), len(m["hmws"]),
            ", ".join(m["oqs"]),
            f'=IF(COUNTIF(RTM!$C:$C,"{br}")=0,"GAP: no HMW","OK")']
    for j, v in enumerate(vals, start=1):
        cell = cs.cell(row=r, column=j, value=v)
        cell.font = Font(size=8.5)
        cell.border = border
        cell.alignment = Alignment(horizontal="center" if j in (1, 2, 4, 5, 6) else "left")
    r += 1

# 3. OQ coverage
r += 1
cs.cell(row=r, column=1, value="3. OQ coverage check (live)").font = Font(bold=True, size=11, color=NAVY)
r += 1
hdrs = ["OQ", "Operational question", "Served by BR", "# Links", "HMW behind it", "Verdict"]
for j, h in enumerate(hdrs, start=1):
    cell = cs.cell(row=r, column=j, value=h)
    cell.font = Font(size=9, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.border = border
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
r += 1
for oq in OQS:
    links = [l for l in LINKS if l[1] == oq]
    brs = ", ".join(sorted({l[2] for l in links}, key=lambda x: int(x.split("-")[1])))
    hmws = ", ".join(sorted({l[0] for l in links}))
    vals = [oq, OQS[oq], brs, len(links), hmws,
            f'=IF(COUNTIF(RTM!$A:$A,"{oq}")=0,"GAP: uncovered","Covered")']
    for j, v in enumerate(vals, start=1):
        cell = cs.cell(row=r, column=j, value=v)
        cell.font = Font(size=8.5)
        cell.border = border
        cell.alignment = Alignment(vertical="top", wrap_text=(j == 2),
                                   horizontal="center" if j in (1, 4, 5, 6) else "left")
    cs.row_dimensions[r].height = 26
    r += 1

# 4. forward columns
r += 1
cs.cell(row=r, column=1, value="4. Forward columns still to be filled").font = Font(bold=True, size=11, color=NAVY)
r += 1
hdrs = ["RTM column", "Populated in", "ID prefix", "Rows filled (v1.0)", "Target", "Status"]
for j, h in enumerate(hdrs, start=1):
    cell = cs.cell(row=r, column=j, value=h)
    cell.font = Font(size=9, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.border = border
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
r += 1
for col, when in PENDING.items():
    vals = [col, when, "-" if "(" not in when else when.split("(")[1].rstrip(")"),
            0, len(LINKS), "Pending"]
    for j, v in enumerate(vals, start=1):
        cell = cs.cell(row=r, column=j, value=v)
        cell.font = Font(size=8.5)
        cell.border = border
        cell.alignment = Alignment(horizontal="center" if j > 1 else "left")
    r += 1

for col, w in {"A": 16, "B": 62, "C": 24, "D": 10, "E": 18, "F": 20}.items():
    cs.column_dimensions[col].width = w
cs.freeze_panes = "A4"

# ================= HOW TO UPDATE =================
hs = wb.create_sheet("How to Update")
hs.sheet_view.showGridLines = False
hs.merge_cells("A1:C1")
c = hs["A1"]
c.value = "How this matrix is maintained (updated in Steps 6, 8, 11, 12 and 13)"
c.font = Font(size=13, bold=True, color=WHITE)
c.fill = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(indent=1)
hs.row_dimensions[1].height = 24

guide = [
    ("Column", "Filled when", "Rule"),
    ("OQ \u2192 Status", "Step 4 (this version)", "Never changed. Status moves Identified \u2192 In Design \u2192 In Build \u2192 In Test \u2192 Tested, or Deferred with a reason in the change log."),
    ("Feature", "Step 6 (PRD)", "Enter the F-xx ID that delivers the BR. One BR may map to one or more features; add a row only if the BR genuinely splits."),
    ("DB table.column", "Step 8 (database design)", "Enter T-xx.column (e.g. T-04.bed_days_occupied). Multiple columns separated by ';'. No free text."),
    ("User Story", "Step 11", "Enter US-xx. Every Must BR must have at least one story before build starts."),
    ("UI Screen", "Step 12", "Enter UI-xx. A BR with no screen at Step 12 is a scope leak and must be raised with the mentor."),
    ("Test Case", "Step 13", "Enter TC-F-xxx / TC-T-xxx. Every BR needs at least one TC-F. NFRs carry TC-T."),
    ("New BR", "Any later step", "Only with a written change note: 'Added BR-16 because <reason>, source HMW-xx, priority <x>'. Renumbering is forbidden."),
    ("Retired BR", "Any later step", "Never delete the row. Set Status = Deferred and record the reason + stakeholder agreement."),
]
para_note = ("Row vs ID counts: a BR linked to two HMWs appears twice (BR-15). Filter by Req ID to see one row "
             "per requirement, or by HMW to see one row per design problem. Neither view is wrong; state "
             "which one you are quoting in reviews.")
r = 3
for i, row in enumerate(guide):
    for j, v in enumerate(row, start=1):
        cell = hs.cell(row=r + i, column=j, value=v)
        cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if i == 0:
            cell.font = Font(size=9, bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=BLUE)
        else:
            cell.font = Font(size=9)
hs.column_dimensions["A"].width = 22
hs.column_dimensions["B"].width = 32
hs.column_dimensions["C"].width = 90
for rr in range(4, 12):
    hs.row_dimensions[rr].height = 34

hs.cell(row=14, column=1, value="Filtering: row 3 of the RTM sheet carries AutoFilter on all 12 columns. "
                                "Use the Priority filter to isolate the 9 Must rows, or the Status filter "
                                "to see what is still 'Identified'.").font = Font(size=9, italic=True)
hs.merge_cells("A16:C16")
hs.cell(row=16, column=1, value=para_note).font = Font(size=9, italic=True)
hs.cell(row=16, column=1).alignment = Alignment(wrap_text=True, vertical="top")
hs.row_dimensions[16].height = 30

# ================= CHANGE LOG =================
ls = wb.create_sheet("Change Log")
ls.sheet_view.showGridLines = False
ls.merge_cells("A1:D1")
c = ls["A1"]
c.value = "RTM Change Log"
c.font = Font(size=13, bold=True, color=WHITE)
c.fill = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(indent=1)
ls.row_dimensions[1].height = 24
log = [("Version", "Date", "Step", "Change"),
       ("1.0", "19-Nov-2026", "Step 4", "Baseline created: 16 HMW\u2013BR links, 12 HMWs, 15 BRs, 10 OQs. All rows Status = Identified. Feature/DB/Story/UI/TC columns intentionally empty."),
       ("", "", "Step 6", "Pending: add F-xx IDs from the PRD."),
       ("", "", "Step 8", "Pending: add T-xx.column references."),
       ("", "", "Step 11", "Pending: add US-xx IDs."),
       ("", "", "Step 12", "Pending: add UI-xx screen IDs."),
       ("", "", "Step 13", "Pending: add TC-F / TC-T IDs.")]
for i, row in enumerate(log):
    for j, v in enumerate(row, start=1):
        cell = ls.cell(row=3 + i, column=j, value=v)
        cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if i == 0:
            cell.font = Font(size=9, bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=BLUE)
        else:
            cell.font = Font(size=9)
ls.column_dimensions["A"].width = 10
ls.column_dimensions["B"].width = 14
ls.column_dimensions["C"].width = 10
ls.column_dimensions["D"].width = 100
for rr in range(4, 10):
    ls.row_dimensions[rr].height = 30

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print("saved:", OUT)
print("links:", len(LINKS), "| unique BR:", len(BR_META), "| HMW:", len(HMW_TEXT))
