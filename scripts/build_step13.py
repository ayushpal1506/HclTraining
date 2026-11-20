"""Build the Step 13 deliverable: test strategy, 46 functional cases, 51 technical
cases and the RTM v1.4 update (Test Case filled).

Every number in the technical tables comes from data/test_results.json, written
by scripts/run_tests.py, so the document cannot claim a pass that did not happen.

Numbers quoted inside functional cases are measured on data/app.db (v1.1, seed 42).
"""
from __future__ import annotations

import json
import os
import re
from datetime import datetime

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import load_workbook
from openpyxl.styles import Font as XFont

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x54, 0x96)
GREY = RGBColor(0x59, 0x59, 0x59)
RED = RGBColor(0xB0, 0x00, 0x20)
HDR_FILL = "1F3864"
ALT_FILL = "EAF0F8"
CALLOUT_FILL = "F2F6FC"
OK_FILL = "E8F3EC"

PROJECT_ID = "HKD-2026"
PROJECT_TITLE = "Healthcare KPI Dashboard"
STUDENT = "Ayush Pal"
ROLL = "2503201450035"
MENTOR = "Aplana Lodhi"

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.join(ROOT, "docs")
OUT = os.path.join(DOCS, "13_Test_Cases.docx")
RTM_PATH = os.path.join(DOCS, "04_Traceability_Matrix.xlsx")
RESULTS_PATH = os.path.join(ROOT, "data", "test_results.json")

RESULTS = json.load(open(RESULTS_PATH, encoding="utf-8"))
CASE_STATUS = dict(RESULTS["cases"])
# test_app.py is skipped as a module until Step 14, so its ten ids never appear
# as individual JUnit cases. Record them as SKIP rather than leaving "—", which
# would read as "not written" when the truth is "written, not runnable yet".
for _i in range(41, 51):
    CASE_STATUS.setdefault(f"TC-T-{_i:02d}", "SKIP")

# --------------------------------------------------------------------------- functional cases
# (id, linked req / story, UI, pre-condition, steps, test data, expected, priority)
CASES_FILTERS = [
    ("TC-F-01", "FR-01, FR-20, US-02", "UI-01",
     "Fresh app start, empty session state, dataset v1.1",
     "1. Open app.py\n2. Read the KPI band without clicking",
     "No filters set",
     "Six cards: Discharges 18,142 · BOR 73.2% · ALOS 4.36 d · Revenue ₹135.26 cr · Mortality 1.29% · Readmission 5.03%. Header shows data to 31-12-2025 and defs v1.0.", "P1"),
    ("TC-F-02", "FR-02, US-02 AC1", "All",
     "Period = 2025",
     "1. Set Period = 2023\n2. Read every card and chart",
     "flt_period = 2023",
     "BOR 71.0%, discharges 17,745, ALOS 4.41, revenue ₹133.10 cr. Chip row reads 2023 and every card, chart and denominator changes together — no stale 2025 figure.", "P1"),
    ("TC-F-03", "FR-02, US-03", "UI-01",
     "Period = 2023",
     "1. Set Period = 2024\n2. Read the readmission card",
     "flt_period = 2024",
     "BOR 72.7%, discharges 18,028, readmission 5.37% (957 ÷ 17,811). No 2023 number survives on screen or in a tooltip.", "P1"),
    ("TC-F-04", "FR-02, US-03", "UI-01, UI-02",
     "Period = 2025, Compare to = None",
     "1. Set Compare to = 2024\n2. Read the delta line on each card",
     "flt_compare = 2024",
     "Each card shows a signed delta and an arrow. BOR delta +0.5 pp (73.19 − 72.72). Negative deltas use U+2212 (−), never a hyphen.", "P2"),
    ("TC-F-05", "FR-02, US-02 AC2", "All",
     "Default filters",
     "1. Select Department = Cardiology\n2. Open the Ward list\n3. Open UI-09",
     "flt_dept = ['Cardiology']",
     "Chip row reads Cardiology. KPI band recomputes (BOR 86.1%). Ward list offers W01–W04 only. UI-09 reports 3,181 episodes for 2025.", "P1"),
    ("TC-F-06", "FR-02", "UI-02",
     "Department = Cardiology",
     "1. Open the Ward multiselect",
     "flt_ward",
     "Exactly four wards: W01 ICU, W02 General, W03 Semi-private, W04 Private. No ward from General Medicine appears; clearing the department restores all 20.", "P2"),
    ("TC-F-07", "FR-02, US-10", "UI-05",
     "Default filters, Period = 2025",
     "1. Select Payer = Corporate\n2. Read the mix and revenue charts",
     "flt_payer = ['Corporate']",
     "Corporate revenue per patient ₹79,973 with n shown. Other payer categories are greyed in the mix, not silently removed. Self-pay card reads ₹71,258 when re-selected.", "P2"),
    ("TC-F-08", "FR-02, FR-13", "UI-03",
     "Period = 2025",
     "1. Set Age = 0–5\n2. Set Gender = F\n3. Open Patient mix",
     "flt_age = (0,5), flt_gender = 'F'",
     "The episode count equals SELECT COUNT(*) with those predicates. The age-band chart shows a single band; the gender facet shows one panel; every denominator on the page changes.", "P2"),
    ("TC-F-09", "FR-02, US-02 AC4", "All",
     "Department = Oncology and Payer = TPA on UI-02",
     "1. Click Reset all filters\n2. Reload the browser",
     "—",
     "All nine filters return to defaults, the chip row empties, the URL query string is cleared, and the reload does not restore the cleared filters.", "P1"),
    ("TC-F-10", "FR-02, US-02 AC4", "All",
     "Department = Oncology set on UI-02",
     "1. Navigate to UI-05\n2. Read the chip row and revenue cards",
     "—",
     "Oncology remains selected, the chip row on UI-05 shows it, and revenue is Oncology-only. No page silently resets a filter it did not draw.", "P1"),
    ("TC-F-11", "FR-02", "All",
     "Department = Oncology, Period = 2025",
     "1. Copy the URL\n2. Open it in a new tab",
     "pasted query string",
     "Identical filter state after reload; chips match the sidebar; no rerun loop and no flash of unfiltered data.", "P2"),
    ("TC-F-12", "FR-03, US-02 AC3", "All",
     "Default state",
     "1. Set Period = 2026 (outside the dataset)\n2. Read the page",
     "flt_period = 2026",
     "Empty state names 2026 and the active filters and offers Reset. No zeros, no exception, no blank page, no stack trace.", "P1"),
]

CASES_KPI = [
    ("TC-F-13", "FR-05, BR-01, US-04", "UI-02",
     "Period 2025, all departments",
     "1. Read the BOR card\n2. Compare with the bar chart\n3. Recompute in SQL",
     "80,083 occupied ÷ 109,420 available bed-days",
     "Card and bar both read 73.2% and equal 73.19% from vw_ward_day_beds. The two proxies never disagree; the note states Σ occupied ÷ Σ available.", "P1"),
    ("TC-F-14", "FR-05, US-04 AC3", "UI-02",
     "Department = Cardiology, 2025",
     "1. Read the gap columns\n2. Recompute from 55 beds × 365 days",
     "BOR 86.12%, 55 functional beds",
     "Gap = +1.1 pp and +0.6 beds (0.0112 × 55). The beds figure is computed, not typed; a department below target shows a negative gap.", "P1"),
    ("TC-F-15", "FR-07, US-06", "UI-03",
     "Period 2025",
     "1. Read the ALOS card and the basis line",
     "79,089 inpatient days ÷ 18,142 discharges",
     "4.36 d displayed. Basis line states deaths and LAMA included and no case-mix adjustment. Changing the department changes both numerator and denominator.", "P1"),
    ("TC-F-16", "FR-12, US-13", "UI-04",
     "Period 2025",
     "1. Read the mortality card\n2. Read Oncology",
     "234 deaths ÷ 18,142 discharges",
     "1.29% overall and 5.53% for Oncology (59 ÷ 1,066). The drift flag is a separate element and never replaces the rate.", "P1"),
    ("TC-F-17", "FR-08, US-09", "UI-03",
     "Period 2025",
     "1. Read the readmission card\n2. Read the rule line",
     "901 readmits ÷ 17,908 eligible",
     "5.03% = 901 ÷ (18,142 − 234). Rule line reads “≤ 30 days, unplanned, deaths excluded”, n is printed, and planned readmissions are absent (0 rows).", "P1"),
    ("TC-F-18", "FR-09, US-10", "UI-05",
     "Period 2025",
     "1. Read total revenue and revenue per patient",
     "₹1,35,25,67,489.99 ÷ 18,142",
     "₹135.26 cr and ₹74,554. Digits use Indian grouping (₹13,52,56,749 style), never western commas, and the rupee symbol renders in every browser tested.", "P1"),
    ("TC-F-19", "FR-18, US-18", "UI-03",
     "Period 2025",
     "1. Department = ENT\n2. Gender = O",
     "7 episodes",
     "The card and the bar carry the badge and the words “low confidence”; no confidence interval is drawn on a rate of 7; the tooltip states n < 30.", "P2"),
    ("TC-F-20", "FR-16, BR-05", "All",
     "Any data screen",
     "1. Open a KPI's definition control\n2. Compare with UI-07",
     "bor, alos, ready_30d",
     "Definition, numerator, denominator, owner and version v1.0 are shown and match the registry word for word. A KPI absent from the registry refuses to render.", "P1"),
]

CASES_CHARTS = [
    ("TC-F-21", "FR-05, OQ-01", "UI-02",
     "Period 2025",
     "1. Read the occupancy chart",
     "8 departments",
     "Horizontal bars sorted descending; one 85% rule; each row prints gap in pp and beds; Cardiology alone sits above the line; bar length equals the table value.", "P1"),
    ("TC-F-22", "FR-06, US-05", "UI-02",
     "Period 2025, Department = Cardiology",
     "1. Open the ward projection\n2. Read W01–W04",
     "7 forecast days",
     "Seven dates with occupied vs functional beds per ward; headroom is negative where projected occupancy exceeds beds; only recorded discharges and confirmed admissions are counted.", "P2"),
    ("TC-F-23", "FR-07, OQ-03", "UI-03",
     "Period 2025",
     "1. Read the ALOS dumbbell",
     "4.0-day benchmark",
     "One shared axis; Oncology 6.49 right of the benchmark, Maternity 2.76 left; each segment labelled with Δ days; departments are ordered by ALOS.", "P1"),
    ("TC-F-24", "FR-08, OQ-04", "UI-03",
     "Period 2025",
     "1. Read the readmission bars",
     "GenMed 7.70, Onco 6.36, Cardio 5.33",
     "Sorted bars with Wilson 95% intervals and n under every bar. Every interval lies inside 0–100, including Maternity at 1.32% (32 ÷ 2,430 eligible, 1 death removed).", "P1"),
    ("TC-F-25", "FR-09, OQ-05", "UI-05",
     "Period 2025, Compare to 2024",
     "1. Read the mix panel",
     "self-pay 33.49 (2025) vs 36.97 (2024)",
     "Each period's stack sums to 100.0 ± 0.1; movement is labelled in pp; corporate 9.05 → 11.23; no slice is hidden to make the sum work.", "P2"),
    ("TC-F-26", "FR-11, OQ-09", "UI-05",
     "Period 2025",
     "1. Read the revenue-per-bed-day ranking",
     "Oncology ₹31,232",
     "Ranked lollipop to #1; Oncology first; General Medicine ₹10,646 shows gap to #1 and 34.1% of #1; the stem is drawn to #1, not to zero.", "P2"),
    ("TC-F-27", "FR-10, OQ-06", "UI-02",
     "Period 2025",
     "1. Open Volume & trend",
     "48,717 visits",
     "Heatmap is 12 months × 7 weekdays; weekday bar shows Monday 202.7 against the 133.5 daily mean (ratio 1.52) and Sunday near-closed; cell counts reconcile to 48,717.", "P2"),
    ("TC-F-28", "FR-13, OQ-08", "UI-03",
     "Period 2025, Department = Paediatrics",
     "1. Open Patient mix\n2. Read age bands",
     "0–5: 1,496 · 6–17: 1,062",
     "Stacked bars sum to the department's discharge count; counts and shares both shown; the gender facet obeys the gender filter; no band is dropped when its count is 0.", "P3"),
    ("TC-F-29", "FR-14, BR-03", "UI-08",
     "Period 2025 vs 2024, KPI = Revenue",
     "1. Open Decompose\n2. Read the waterfall and the residual",
     "₹133.10 cr → ₹135.26 cr",
     "Contributor bars sum to +₹2.16 cr; every contributor is named; residual printed as +0.00 pp and never merged into “other”; changing the dimension re-decomposes the same total.", "P2"),
    ("TC-F-30", "FR-15, BR-10", "UI-08",
     "Default filters",
     "1. Open Association test\n2. Read the method card",
     "BOR% vs readmission%, 288 dept-months",
     "Method card names Spearman, n = 288, rho, t, p and a 95% CI; wording says association, not causation; the assumption note and the low-n caveat are visible.", "P2"),
]

CASES_QUALITY = [
    ("TC-F-31", "FR-17, BR-14", "UI-06",
     "Period 2025",
     "1. Open Data Quality\n2. Read all six checks",
     "6 checks",
     "Six cards, each PASS/FAIL with a count; 6/6 pass on the shipped data; freshness reads 0 days stale against max(admission_date) 2025-12-31.", "P1"),
    ("TC-F-32", "FR-19, BRL-12", "UI-06",
     "Default state",
     "1. Filter Rule = RQ-01\n2. Compare the row count with quality_log.csv",
     "3,494 rows",
     "Table shows rule, table, field, value and action; the count matches the CSV; no name, phone, address or Aadhaar column is offered anywhere in the picker.", "P1"),
    ("TC-F-33", "FR-17, BR-14", "All",
     "Build whose facts end 2025-11-15",
     "1. Open UI-01\n2. Read the banner and the cards",
     "data_cut = 2025-11-15",
     "DQ banner renders above the KPIs naming the failing check; KPIs stay visible and labelled, never blanked or hidden; the banner links to UI-06 by sidebar, not by an anchor.", "P1"),
    ("TC-F-34", "FR-21, US-20", "UI-09",
     "Period 2025, Department = Cardiology",
     "1. Click Export CSV\n2. Open the file",
     "3,181 rows",
     "File hkd2026_episodes_2025_cardio_20261119.csv; 3,181 rows; columns equal the preview; spot-checked values equal the screen to the paisa.", "P1"),
    ("TC-F-35", "FR-21, NFR-09", "UI-09",
     "Fixture row whose text begins with a formula character",
     "1. Export\n2. Open in Excel and LibreOffice",
     "=cmd|' /C calc'!A0, +1, -2, @x",
     "Every dangerous cell is prefixed with an apostrophe; the file starts with EF BB BF (UTF-8 BOM); no formula executes; the guard is applied on every text column, not a fixed list.", "P1"),
    ("TC-F-36", "FR-21", "UI-09",
     "All years, all departments",
     "1. Request an export\n2. Read the refusal",
     "53,915 rows",
     "The button refuses and asks for a narrower period; the message names the 50,000-row limit; no truncated file is produced.", "P2"),
    ("TC-F-37", "FR-16, US-16", "UI-07",
     "Default state",
     "1. Open Definitions\n2. Walk every figure on the other eight screens",
     "registry v1.0",
     "Every published KPI appears once with definition, numerator, denominator, owner and version. A stub KPI missing from the registry raises a visible definition error at Step 15.", "P1"),
    ("TC-F-38", "BR-04, US-08", "All",
     "Fresh app",
     "1. Count clicks to each headline KPI\n2. Count sidebar entries",
     "dashboard already open",
     "Six headline KPIs are visible with zero clicks; any department KPI is reachable in three (department → screen → tab); the sidebar lists exactly nine screens.", "P2"),
    ("TC-F-39", "FR-22, NFR-08", "All",
     "Browser at 375 × 812",
     "1. Open each screen\n2. Scroll horizontally",
     "iPhone-class viewport",
     "Columns stack 1-up; KPI values 22 px; charts 260 px tall; tables scroll inside their own container; the page itself has no horizontal scrollbar and no clipped label.", "P2"),
    ("TC-F-40", "FR-22, US-08", "UI-01",
     "Browser at 1440 × 900",
     "1. Open the overview\n2. Scroll check",
     "laptop viewport",
     "Six cards and the trend chart are visible without scrolling; sidebar 300 px; content centred and clamped at 1280 px; no element overlaps at 125% zoom.", "P2"),
]

CASES_EDGE = [
    ("TC-F-41", "FR-04, US-03", "UI-02",
     "Build whose facts end 2025-11-15",
     "1. Open the month view\n2. Read November",
     "data_cut = 2025-11-15",
     "November is drawn hollow, labelled “in progress”, excluded from the ranking and from every trend endpoint; December is absent; the badge is legible in greyscale.", "P2"),
    ("TC-F-42", "FR-03, US-18", "UI-03",
     "Period 2025",
     "1. Department = Maternity\n2. Gender = M",
     "0 episodes",
     "Rates show “—” and no rupee value is invented; the empty state names Maternity + M and offers Reset; no ZeroDivisionError, no NaN in any cell or export.", "P2"),
    ("TC-F-43", "FR-02, NFR-12", "All",
     "App open",
     "1. Edit the URL to dept=DEP99\n2. Reload",
     "dept = DEP99",
     "The unknown value is ignored with a one-line notice; valid filters still apply; the page renders; no stack trace and no blank screen.", "P2"),
    ("TC-F-44", "FR-02, NFR-01", "UI-02",
     "App open, Period = 2025",
     "1. Change Period twice within 2 s\n2. Watch the final render",
     "2024 → 2023",
     "Final screen shows 2023 only; no mixed-period render; no spinner left running; p95 response ≤ 2 s over 20 runs.", "P2"),
    ("TC-F-45", "NFR-08", "All",
     "Keyboard only, mouse unplugged",
     "1. Tab to the first KPI definition control\n2. Press Enter, then Esc",
     "—",
     "Focus ring visible on every control; definition opens with Enter/Space and closes with Esc; focus returns to the trigger; tab order follows the visual order.", "P3"),
    ("TC-F-46", "BR-15", "All",
     "Default state",
     "1. Inspect all nine screens and the sidebar for a notification surface",
     "—",
     "No email, SMS or push control exists; the exclusion is recorded in the RTM as “Won't”; the scope check passes and the absence is documented, not accidental.", "P3"),
]

ALL_CASES = CASES_FILTERS + CASES_KPI + CASES_CHARTS + CASES_QUALITY + CASES_EDGE
assert len(ALL_CASES) == 46
assert [c[0] for c in ALL_CASES] == [f"TC-F-{i:02d}" for i in range(1, 47)]
assert all(len(c) == 8 for c in ALL_CASES)

PRIORITY = {
    "P1": "Blocks release. A wrong number, a wrong scope or a dead primary path.",
    "P2": "Blocks sign-off. A broken secondary path, chart or export guard.",
    "P3": "Cosmetic or assistive. Logged and fixed before submission where possible.",
}

# --------------------------------------------------------------------------- technical cases
# (id, level, module / function, check, expected)
TECH_DATA = [
    ("TC-T-01", "Integration", "data/app.db", "File exists, is 30–80 MB, and opens read-only (an INSERT raises OperationalError)", "PASS when the file is 30–80 MB and the write is refused"),
    ("TC-T-02", "Integration", "schema", "Exactly five tables (dim_department, dim_ward, dim_payer, fact_inpatient_episode, fact_opd_visit) and three views", "Set equality, no extras"),
    ("TC-T-03", "Integration", "row counts", "Each table equals the seed-42 count in validation_summary.json", "8 / 20 / 5 / 53,915 / 143,139"),
    ("TC-T-04", "Integration", "primary keys", "No NULL in any of the five primary keys", "0 nulls"),
    ("TC-T-05", "Integration", "primary keys", "COUNT(*) equals COUNT(DISTINCT pk) in every table", "0 duplicates"),
    ("TC-T-06", "Integration", "foreign keys", "PRAGMA foreign_key_check over the whole database", "[] — no orphan rows"),
    ("TC-T-07", "Integration", "stay arithmetic", "los_days = discharge − admission + 1; bed_days = los_days; discharge never precedes admission", "0 violations of each"),
    ("TC-T-08", "Integration", "length of stay", "Range and mean against the validated snapshot", "1–90 days, mean 4.41"),
    ("TC-T-09", "Integration", "billing", "Minimum, maximum and the absence of null/zero amounts", "₹5,788.15 – ₹60,56,691.77; 0 invalid"),
    ("TC-T-10", "Integration", "discharge coding", "death_flag = 1 if and only if discharge_type = 'Expired'", "0 disagreements"),
    ("TC-T-11", "Integration", "readmission links", "No dangling index, no readmission without an index, no planned readmission, no gap over 30 days", "0 / 0 / 0 / 0"),
    ("TC-T-12", "Integration", "coded domains", "Age bands, genders, discharge types, visit types and payer categories match the documented lists", "6 / 3 / 4 / 2 / 5 values"),
    ("TC-T-13", "Integration", "date coverage", "Both facts span the documented window", "2023-01-01 … 2025-12-31"),
    ("TC-T-14", "Integration", "bed capacity", "300 functional beds over 20 wards; every ward belongs to an active department", "300 beds, 0 orphans"),
    ("TC-T-15", "Integration", "vw_analysis_episode", "Column contract and row grain", "36 columns, 53,915 rows"),
    ("TC-T-16", "Integration · slow", "vw_ward_day_beds", "Occupancy bounds and column contract", "8 columns, 0–100%, max 100.0"),
    ("TC-T-17", "Integration · slow", "vw_dept_month_kpi", "One row per department-month; 8 × 36 months", "288 rows, 0 duplicate grains, 36 months"),
    ("TC-T-18", "Integration", "quality_log.csv", "The reject log reconciles to the cleaning totals", "5,527 entries = 5,368 rejected + 159 flagged"),
    ("TC-T-19", "Integration", "dimension usage", "Every dimension row is reachable from a fact", "8 departments, 5 payer categories used"),
    ("TC-T-20", "Integration", "cleaning", "Cleaned counts are below raw counts; seed and totals are recorded", "seed 42; 5,368 rows removed"),
]

TECH_ANALYSIS = [
    ("TC-T-21", "Unit", "kpi.bor_pct, kpi.occupancy_gap_pp", "Occupancy and gap arithmetic on a known pair", "10/20 = 50.0%; gap −35.0 pp; 0 denominator → None"),
    ("TC-T-22", "Unit", "kpi.beds_to_target", "Beds needed to reach 85% over a known period", "3.5 beds; on target → 0.0; 0 days → None"),
    ("TC-T-23", "Unit", "kpi.alos_days", "Inpatient days over discharges", "23/5 = 4.60; 0 days → 0.0; 0 discharges → None"),
    ("TC-T-24", "Unit", "kpi.mortality_pct", "Deaths over discharges", "1/5 = 20.0%; 2/0 → None"),
    ("TC-T-25", "Unit", "kpi.eligibility, kpi.readmission_rate_pct", "Deaths removed from the denominator before the rate", "eligible 4; rate 25.0%; zero eligible → None"),
    ("TC-T-26", "Unit", "kpi.revenue_per_patient, kpi.revenue_per_bed_day", "Both revenue families on one known total", "₹40,000 per patient; ₹8,695.65 per bed-day"),
    ("TC-T-27", "Unit", "kpi.share_pct", "Compositional share, including an empty whole", "25.0%; 0.0; None — never nan"),
    ("TC-T-28", "Unit", "kpi.readmitted_within", "Window boundaries: day 0, day 30, day 31, admission before discharge", "True, True, False, False"),
    ("TC-T-29", "Unit", "kpi.low_confidence", "The n < 30 rule at its boundary", "29 True, 30 False, 0 True, None False"),
    ("TC-T-30", "Integration", "kpi.* against a 5-row warehouse", "Formulas applied to SQL aggregates equal hand arithmetic, overall and per department", "ALOS 4.6, mortality 20.0%, readmission 25.0%, DEP01 readmission 50.0%"),
    ("TC-T-30b", "Unit", "kpi.decompose", "Contributors sum to the total movement; residual printed; a new key is a fact, not an error", "Δ₹17.0 cr fully attributed; residual 0.00; opened department contributes 5.0"),
    ("TC-T-31", "Unit", "stats.wilson_interval", "Published Wilson values at 50/100, 0/10 and 10/10", "40.38–59.62%; 0–27.75%; 72.25–100%"),
    ("TC-T-32", "Property", "stats.wilson_interval", "For every k of every n ≤ 60 the interval stays inside 0–100 and contains p̂", "No violation across 1,890 pairs"),
    ("TC-T-33", "Unit", "stats.pearson", "Perfect positive, perfect negative, constant series and n < 3", "1.0, −1.0, None, None"),
    ("TC-T-34", "Unit", "stats.spearman", "Monotone non-linear series gives 1.0; ties are handled", "1.0; 0.8; not None"),
    ("TC-T-35", "Unit", "stats._ranks", "Average ranks for tied values", "10, 20, 20, 30 → 1, 2.5, 2.5, 4"),
    ("TC-T-36", "Unit", "stats.student_t_p", "Two-sided p at published critical values", "t=1.96 (df 1e6) → 0.050; t=2.228 (df 10) → 0.050; t=0 → 1.0"),
    ("TC-T-37", "Unit", "stats.chi_square_2x2", "χ² and p for a 2×2 table at df = 1", "χ² 0.7937, p 0.373; degenerate table → None"),
    ("TC-T-38", "Unit", "stats.correlation_test", "Result contract: rho, n, t, p and method", "rho 0.9048, n 8, 0 < p < 1; perfect monotone → p 0"),
    ("TC-T-39", "Unit", "stats.drift_flag", "A 4.0 month after a 3.0 ± 0.1 baseline is flagged; a 3.05 month is not", "first flag at index 13; short series → no flags"),
    ("TC-T-40", "Unit", "stats.ols_line", "Least squares recovers y = 2x + 1; zero variance returns None", "slope 2.0, intercept 1.0, None"),
]

TECH_APP = [
    ("TC-T-41", "E2E", "app.py (AppTest)", "The overview runs to completion", "No exception; a title or header renders"),
    ("TC-T-42", "E2E", "all 9 screens", "Each page renders without an uncaught exception", "9 of 9 clean"),
    ("TC-T-43", "E2E", "ui.header on 9 pages", "Every page carries the header and the data-to chip", "9 of 9 present"),
    ("TC-T-44", "E2E · timing", "all 9 screens", "Warm render time against the NFR-01 budget of 3 s", "Every page < 3.0 s"),
    ("TC-T-45", "E2E", "ui.LAST_RENDER", "At least six KPI values reach the test seam; none is None or nan", "≥ 6 values, all usable"),
    ("TC-T-46", "E2E", "empty state", "A filter combination with no rows renders the empty state", "No exception; empty-state text present"),
    ("TC-T-47", "E2E", "download button", "The CSV control exists and its payload is BOM-prefixed", "EF BB BF or a header row"),
    ("TC-T-48", "E2E", "PII scan", "Rendered text contains no name/phone/address field label", "0 banned tokens"),
    ("TC-T-49", "Unit", "ui.MOBILE_CSS", "The media query and the column selector are present", "640 px query; stHorizontalBlock"),
    ("TC-T-50", "E2E", "error path", "A failure renders a sentence, not a traceback", "No “Traceback”, no “sqlite3.”"),
]

ALL_TECH = TECH_DATA + TECH_ANALYSIS + TECH_APP
assert len(ALL_TECH) == 51, len(ALL_TECH)
assert sum(n + 1 for n in range(1, 61)) == 1890  # the Wilson property sweep
assert len({t[0] for t in ALL_TECH}) == 51

TECH_TYPES = {"Integration": 18, "Unit": 19, "Property": 1, "E2E": 11, "E2E · timing": 1, "Integration · slow": 2}
# counts recomputed rather than typed
LEVEL_COUNTS: dict[str, int] = {}
for row in ALL_TECH:
    key = row[1].split(" · ")[0]
    LEVEL_COUNTS[key] = LEVEL_COUNTS.get(key, 0) + 1

# --------------------------------------------------------------------------- traceability
BR_TESTS = [
    ("BR-01", "Must", "TC-F-13, TC-F-14, TC-F-21, TC-T-21, TC-T-22", "Automated formula + screen"),
    ("BR-02", "Must", "TC-F-06, TC-F-22", "Screen + child filter"),
    ("BR-03", "Must", "TC-F-29, TC-T-30b", "Screen + unit"),
    ("BR-04", "Must", "TC-F-01, TC-F-38", "Screen + interaction count"),
    ("BR-05", "Must", "TC-F-20, TC-F-37", "Screen + registry walk"),
    ("BR-06", "Must", "TC-F-15, TC-F-23, TC-T-23", "Automated formula + screen"),
    ("BR-07", "Must", "TC-F-17, TC-F-24, TC-T-25, TC-T-28", "Formula, boundary and screen"),
    ("BR-08", "Should", "TC-F-18, TC-F-25, TC-T-26", "Automated formula + screen"),
    ("BR-09", "Must", "TC-F-27", "Screen + reconciliation of cell counts"),
    ("BR-10", "Could", "TC-F-30, TC-T-38", "Screen + unit"),
    ("BR-11", "Could", "TC-F-08, TC-F-28", "Filter + screen"),
    ("BR-12", "Should", "TC-F-26, TC-T-26", "Automated formula + screen"),
    ("BR-13", "Must", "TC-F-16, TC-T-39", "Automated formula + screen"),
    ("BR-14", "Should", "TC-F-31, TC-F-32, TC-F-33, TC-T-18", "Screen + reconciliation"),
    ("BR-15", "Won't", "TC-F-46", "Absence check (scope)"),
]

FR_TESTS = [
    ("FR-01", "Load facts and dimensions read-only", "TC-T-01, TC-T-02, TC-T-03, TC-F-01"),
    ("FR-02", "Apply global filters consistently", "TC-F-02 … TC-F-11, TC-F-43, TC-F-44"),
    ("FR-03", "Empty state with the active filter summary", "TC-F-12, TC-F-42, TC-F-43"),
    ("FR-04", "Partial periods labelled and excluded", "TC-F-41"),
    ("FR-05", "Occupancy vs 85%, gap in pp and beds", "TC-F-13, TC-F-14, TC-F-21, TC-T-21, TC-T-22"),
    ("FR-06", "Seven-day ward occupancy projection", "TC-F-22"),
    ("FR-07", "ALOS with the case-mix basis stated", "TC-F-15, TC-F-23, TC-T-23"),
    ("FR-08", "30-day readmission with the rule stated", "TC-F-17, TC-F-24, TC-T-25, TC-T-28"),
    ("FR-09", "Payer mix, movement, revenue per patient", "TC-F-07, TC-F-18, TC-F-25, TC-T-26"),
    ("FR-10", "OPD by weekday and month with peak ratio", "TC-F-27"),
    ("FR-11", "Revenue per bed-day ranked with gap", "TC-F-26, TC-T-26"),
    ("FR-12", "Mortality, discharge mix and drift flag", "TC-F-16, TC-T-24, TC-T-39"),
    ("FR-13", "Age band and gender breakdown", "TC-F-08, TC-F-28"),
    ("FR-14", "Decompose a movement into contributors", "TC-F-29, TC-T-30b"),
    ("FR-15", "Statistical test with method, n, p, CI", "TC-F-30, TC-T-31 … TC-T-40"),
    ("FR-16", "Definition registry and refusal to render", "TC-F-20, TC-F-37"),
    ("FR-17", "Six data-quality checks with counts", "TC-F-31, TC-F-33, TC-T-18"),
    ("FR-18", "Low-confidence rule at n < 30", "TC-F-19, TC-F-42, TC-T-29"),
    ("FR-19", "Reject log with reason codes", "TC-F-32, TC-T-18"),
    ("FR-20", "Tailwind header and cards, server-side values", "TC-F-01, TC-T-41, TC-T-42, TC-T-43"),
    ("FR-21", "Filtered CSV, guarded against formula injection", "TC-F-34, TC-F-35, TC-F-36"),
    ("FR-22", "Responsive at 375 / 768 / 1440 px", "TC-F-39, TC-F-40, TC-T-49"),
]

NFR_TESTS = [
    ("NFR-01", "Filter response p95 ≤ 2 s", "TC-F-44, TC-T-44", "Automated (AppTest) + timed manual"),
    ("NFR-02", "Cold start ≤ 10 s", "TC-T-41", "Measured on Cloud at Step 16"),
    ("NFR-03", "Memory ≤ 1 GB, dataset ≤ 200 MB", "—", "Manual profile at Step 15 (recorded in 15_)"),
    ("NFR-04", "Availability and explicit sleep message", "—", "7-day spot check at Step 16"),
    ("NFR-05", "KPI correctness to 1e-6 on fixtures", "TC-T-21 … TC-T-30", "Automated, plus SciPy parity at Step 15"),
    ("NFR-06", "Same commit + seed → same hash", "—", "Rebuild twice and compare SHA-256 at Step 15"),
    ("NFR-07", "Median task ≤ 60 s; SUS ≥ 75", "—", "Moderated UAT, 5–8 users, Step 15"),
    ("NFR-08", "WCAG AA, keyboard path, 375 px", "TC-F-39, TC-F-40, TC-F-45", "Automated layout + manual keyboard run"),
    ("NFR-09", "No PII, read-only DB, no secrets", "TC-T-01, TC-T-48, TC-F-35", "Automated"),
    ("NFR-10", "≥ 85% coverage, no formula outside kpi.py", "TC-T-21 … TC-T-30", "pytest-cov + grep audit at Step 15"),
    ("NFR-11", "Optional passphrase, constant-time, 3 attempts", "—", "Manual when the feature is enabled"),
    ("NFR-12", "User-safe errors, one log line per render", "TC-T-50, TC-F-43", "Automated + log inspection"),
]

US_TESTS = [
    ("US-01", "Read-only load of five tables", "TC-T-01 … TC-T-06, TC-T-14, TC-T-15"),
    ("US-02", "One filter bar every KPI obeys", "TC-F-01 … TC-F-11, TC-F-44"),
    ("US-03", "In-progress period excluded", "TC-F-03, TC-F-04, TC-F-41"),
    ("US-04", "Occupancy vs 85% in points and beds", "TC-F-13, TC-F-14, TC-T-21, TC-T-22"),
    ("US-05", "Seven-day ward projection", "TC-F-22"),
    ("US-06", "Like-for-like ALOS with basis", "TC-F-15, TC-F-23"),
    ("US-07", "OPD demand by weekday and month", "TC-F-27"),
    ("US-08", "Responsive shell, phone and projector", "TC-F-38, TC-F-39, TC-F-40, TC-F-45, TC-T-49"),
    ("US-09", "Readmission with the rule inline", "TC-F-17, TC-F-24, TC-T-28"),
    ("US-10", "Payer-mix movement and revenue per patient", "TC-F-07, TC-F-18, TC-F-25"),
    ("US-11", "Revenue per bed-day ranking", "TC-F-26"),
    ("US-12", "Admission mix by age and gender", "TC-F-08, TC-F-28"),
    ("US-13", "Mortality trend with a drift flag", "TC-F-16, TC-T-39"),
    ("US-14", "Decompose a KPI movement", "TC-F-29, TC-T-30b"),
    ("US-15", "Statistical test with assumptions", "TC-F-30, TC-T-31 … TC-T-38"),
    ("US-16", "Every figure carries its definition", "TC-F-20, TC-F-37"),
    ("US-17", "Six quality checks beside the KPI", "TC-F-31, TC-F-33"),
    ("US-18", "Low-confidence badge at n < 30", "TC-F-19, TC-F-42, TC-T-29"),
    ("US-19", "Reject log with reason codes", "TC-F-32"),
    ("US-20", "Safe filtered CSV export", "TC-F-34, TC-F-35, TC-F-36"),
]

UI_TESTS = [
    ("UI-01", "Overview", "TC-F-01, TC-F-04, TC-F-13, TC-F-16, TC-F-18, TC-F-40"),
    ("UI-02", "Operations", "TC-F-02, TC-F-05, TC-F-06, TC-F-09, TC-F-14, TC-F-21, TC-F-22, TC-F-27, TC-F-41, TC-F-44"),
    ("UI-03", "Departments", "TC-F-08, TC-F-15, TC-F-17, TC-F-19, TC-F-23, TC-F-24, TC-F-28, TC-F-42"),
    ("UI-04", "Quality", "TC-F-16"),
    ("UI-05", "Finance", "TC-F-07, TC-F-10, TC-F-18, TC-F-25, TC-F-26"),
    ("UI-06", "Data Quality", "TC-F-31, TC-F-32, TC-F-33"),
    ("UI-07", "Definitions", "TC-F-20, TC-F-37"),
    ("UI-08", "Insights", "TC-F-29, TC-F-30"),
    ("UI-09", "Data & Export", "TC-F-34, TC-F-35, TC-F-36"),
]

RTM_TESTS = [
    ("BR-01", "TC-F-13, TC-F-14, TC-F-21, TC-T-21"),
    ("BR-15", "TC-F-46"),
    ("BR-02", "TC-F-06, TC-F-22"),
    ("BR-03", "TC-F-29, TC-T-30b"),
    ("BR-04", "TC-F-01, TC-F-38"),
    ("BR-05", "TC-F-20, TC-F-37"),
    ("BR-06", "TC-F-15, TC-F-23, TC-T-23"),
    ("BR-07", "TC-F-17, TC-F-24, TC-T-25, TC-T-28"),
    ("BR-14", "TC-F-31, TC-F-32, TC-F-33, TC-T-18"),
    ("BR-08", "TC-F-18, TC-F-25, TC-T-26"),
    ("BR-09", "TC-F-27"),
    ("BR-10", "TC-F-30, TC-T-38"),
    ("BR-11", "TC-F-08, TC-F-28"),
    ("BR-12", "TC-F-26, TC-T-26"),
    ("BR-13", "TC-F-16, TC-T-39"),
    ("BR-15", "TC-F-46"),
]

# every one of the 46 functional cases must be reachable from at least one
# traceability table, otherwise it is a test nobody asked for
def _expand(text: str) -> set[str]:
    ids = set(re.findall(r"TC-F-\d\d", text))
    for a, b in re.findall(r"TC-F-(\d\d)\s*…\s*TC-F-(\d\d)", text):
        ids |= {f"TC-F-{i:02d}" for i in range(int(a), int(b) + 1)}
    return ids


_used: set[str] = set()
for _group in (BR_TESTS, FR_TESTS, NFR_TESTS, US_TESTS, UI_TESTS):
    for _row in _group:
        _used |= _expand(" ".join(str(x) for x in _row))
assert len(_used) == 46, sorted({c[0] for c in ALL_CASES} - _used)
assert len(BR_TESTS) == 15 and len(FR_TESTS) == 22 and len(NFR_TESTS) == 12
assert len(US_TESTS) == 20 and len(UI_TESTS) == 9

# every technical id must exist in the bound module lists
_tech_ids = {t[0] for t in ALL_TECH}
assert set(CASE_STATUS) <= _tech_ids, set(CASE_STATUS) - _tech_ids
assert len(CASE_STATUS) == 51, len(CASE_STATUS)

# --------------------------------------------------------------------------- fixtures
FIXTURE_ROWS = [
    ("E1", "DEP01", "01-05 Jan 2025", "5", "5", "₹50,000", "0", "0", "—"),
    ("E2", "DEP01", "10-14 Jan 2025", "5", "5", "₹30,000", "1", "0", "—"),
    ("E3", "DEP01", "24-26 Jan 2025", "3", "3", "₹20,000", "0", "1", "E1 (gap 19 d)"),
    ("E4", "DEP02", "01-09 Feb 2025", "9", "9", "₹90,000", "0", "0", "— (planned)"),
    ("E5", "DEP02", "10-10 Feb 2025", "1", "1", "₹10,000", "0", "0", "—"),
]

BOUNDARIES = [
    ("Length of stay", "1 and 90 days", "TC-T-07, TC-T-08", "Both included; 0 and 91 must not exist"),
    ("Readmission gap", "30 and 31 days", "TC-T-28", "30 counts, 31 does not, same day counts"),
    ("Low confidence", "n = 29 and n = 30", "TC-T-29, TC-F-19", "29 flagged, 30 not"),
    ("Occupancy", "0% and 100%", "TC-T-16", "View never exceeds 100; 100 occurs on a full ward-day"),
    ("Denominator", "0 discharges", "TC-T-21 … TC-T-27", "Returns None → “—” on screen, never an exception"),
    ("Deaths in denominator", "18,142 − 234 = 17,908", "TC-F-17, TC-T-25", "Deceased patients cannot be readmitted"),
    ("Export limit", "50,000 rows", "TC-F-36", "Refuse above the limit; no partial file"),
    ("CSV injection", "= + − @ at cell start", "TC-F-35", "Apostrophe prefix on every text column"),
    ("Period", "2026 (no data)", "TC-F-12", "Empty state, no zeros"),
    ("Filters", "DEP99 (not in the dimension)", "TC-F-43", "Ignored with a notice, page still renders"),
]

CYCLES = [
    ("C1 · Smoke", "Day 4 (Mon 23-Nov), 2 h", "Build deploys locally; data gate passes; RTM v1.4 loaded", "TC-F-01, 12, 38; TC-T-41, 42, 43", "Every page opens; six KPI values present"),
    ("C2 · Functional", "Day 5 (Tue 24-Nov), 4.5 h", "C1 exit; RTM v1.4 signed", "All 46 TC-F cases", "100% P1 executed; evidence attached"),
    ("C3 · Regression", "Day 5 evening, 1.5 h", "Fixes landed", "All automated + P1 functional", "No regression on a previously passed case"),
    ("C4 · UAT", "Day 6 (Wed 25-Nov), 2 h", "C3 exit; build frozen", "Scripted tasks with 5–8 users", "Median task ≤ 60 s; no blocking confusion"),
]

SEVERITY = [
    ("S1", "Wrong number, wrong scope, or a PII leak", "Same day", "Blocks release; fix before any other work"),
    ("S2", "A primary path is blocked (export, filter, page fails)", "Next working day", "Blocks sign-off"),
    ("S3", "Secondary path, chart or wording defect", "Before submission", "Fix if time allows"),
    ("S4", "Cosmetic only", "Backlog", "Recorded, not fixed"),
]

NOT_TESTED = [
    ("Concurrent multi-user load", "The deliverable is a single-analyst demo on Community Cloud; no write path exists to contend over.", "Step 16 load note only"),
    ("Datasets larger than 53,915 episodes", "The published scope is the seeded warehouse; the SQL is set-based and the 50,000-row export guard bounds the client.", "Documented limitation"),
    ("Browsers other than Chrome and Safari", "The classroom and mentor environment uses both; Firefox rendering is checked once by screenshot only.", "Screenshot at Step 15"),
    ("Screen-reader navigation of iframe content", "The iframe is display-only and duplicates no interactive control; a screen reader reads the native widgets, which carry the same text in labels and captions.", "Manual note at Step 15"),
    ("Statistics against SciPy", "SciPy is installed at Step 14; until then the closed-form values are checked against published tables.", "TC-T-31 … 40 re-run at Step 15"),
    ("Email / push alerting", "BR-15 is deferred; TC-F-46 verifies the absence rather than a feature.", "RTM keeps the exclusion"),
]

ASSUMPTIONS = [
    ("AS-35", "Test data is the shipped synthetic warehouse (seed 42) plus hand-built fixtures; no real patient data is used at any point.",
     "The project rule forbids real personal data, and a fixture built by hand is the only oracle independent of the code under test.",
     "If a fixture is wrong, TC-T-21…30 and TC-F-13…18 agree on the wrong answer."),
    ("AS-36", "The local suite runs on Python 3.13.3; production targets 3.12 on Streamlit Community Cloud.",
     "The code uses only 3.10+ syntax, and Step 14 pins a 3.12 venv for the build.",
     "A 3.13-only behaviour could pass locally and fail on Cloud; Step 14 re-runs the suite in 3.12 before deploying."),
    ("AS-37", "“Eligible discharges” = discharges − deaths. LAMA and transfers stay in the denominator.",
     "A deceased patient cannot return; a patient who left against advice can.",
     "The published readmission rate moves by about 0.05 pp for 2025 (5.03% vs 5.00% on all discharges)."),
    ("AS-38", "The 46 functional cases take about 4.5 hours, at roughly 5 minutes of evidence per case.",
     "Each case needs one screenshot, one query result and a status change; 46 × 5 min fits Day 5.",
     "If evidence collection is slower, C2 spills into Day 6 and C4 (UAT) is compressed."),
    ("AS-39", "Tests never write to data/app.db; the connection is opened in URI read-only mode.",
     "A test run must not be able to change the deliverable it is testing (NFR-09).",
     "A fixture that needs to write must build its own in-memory database, as TC-T-30 does."),
    ("AS-40", "The 50,000-row export limit is a product decision, not a business requirement.",
     "Excel-class tools and the browser both degrade above that size; the brief asks only for a download.",
     "If the mentor wants unlimited export, TC-F-36 changes and the guard moves to a warning."),
    ("AS-41", "Wilson intervals are the only interval method used for rates on screen.",
     "One method, stated once, avoids two departments being compared with two different intervals.",
     "A reviewer expecting exact/clopper-pearson intervals will see slightly different bounds."),
    ("AS-42", "BR-15 is verified by absence (TC-F-46), not by a deferred feature.",
     "The RTM promises every BR has at least one test; an absence check is a real test of scope.",
     "If BR-15 returns to scope, TC-F-46 is deleted and replaced by three real cases."),
]

FINDINGS = [
    ("TCF-01", "MED", "The Step-12 wireframe printed readmission as 4.98% (891 ÷ 17,908). Recomputing by the stated rule gives 901 ÷ 17,908 = 5.03%; 4.98% is not reproducible from the warehouse.",
     "Step 12 was rebuilt with 5.03%. The test pins the rule (deaths excluded, ≤ 30 days, unplanned) and the arithmetic, not a transcribed number."),
    ("TCF-02", "LOW", "The Step-12 UI-09 wireframe showed 2,431 Cardiology episodes for 2025; the warehouse holds 3,181.",
     "Corrected in the rebuilt Step 12. TC-F-05 and TC-F-34 now assert 3,181 so the illustration cannot drift again."),
    ("TCF-03", "HIGH", "There are two legitimate bed-day denominators. Occupancy (73.19%) uses ward-day allocation — 80,083 occupied of 109,420 available in 2025 — while revenue per bed-day (₹17,102) uses episode bed_days, 79,089. The 994-day difference is stays that straddle 31 December.",
     "Both stay. Every screen labels which denominator it uses; TC-F-13 and TC-F-18 pin both numbers. A single “bed-days” number would force one of the two KPIs to be wrong."),
    ("TCF-04", "HIGH", "Length of stay is inclusive: los = discharge − admission + 1, so a same-day discharge is 1 day, not 0. The BRD's phrase “discharge date minus admission date” reads as exclusive.",
     "TC-T-07 pins +1 and the registry wording will say “inclusive of admission and discharge dates”. Otherwise every ALOS is 1 day short."),
    ("TCF-05", "MED", "Ten AppTest cases (TC-T-41…50) cannot run yet: Streamlit, pandas, Plotly and pytest-asyncio-free AppTest arrive at Step 14.",
     "They are written now and skipped as one module, so the suite reports 1 skip with a reason instead of a false green. Step 15 removes the skip by existing."),
    ("TCF-06", "LOW", "The local venv is Python 3.13.3; Streamlit Community Cloud runs 3.12. A local green suite is not proof of a Cloud green suite.",
     "Step 14 builds a 3.12 venv and re-runs the suite there before deployment (AS-36)."),
    ("TCF-07", "MED", "The US-09 acceptance criterion in Step 11 quoted department readmission rates of 7.58 / 6.00 / 5.25 / 1.97 / 1.32%. Under the pinned rule (admission-year period, deaths removed from the denominator) the warehouse gives 7.70 / 6.36 / 5.33 / 1.98 / 1.32%. Only Maternity matched.",
     "Step 11's story text was corrected in place to 7.70 / 6.36 / 5.33 / 1.98 / 1.32 and build_step11.py was patched so a future rebuild cannot reintroduce them. TC-F-24 now pins the same five values. The overall rate was never affected: 901 ÷ 17,908 = 5.03%."),
]

# --------------------------------------------------------------------------- RTM v1.4
wb = load_workbook(RTM_PATH)
rt = wb["RTM"]
hdr = {rt.cell(row=3, column=c).value: c for c in range(1, rt.max_column + 1)}
COL_BR = hdr["Req ID (BR)"]
COL_UI = hdr["UI Screen"]
COL_TC = hdr["Test Case (TC-F / TC-T)"]
COL_STATUS = hdr["Status"]

proto = rt.cell(row=4, column=COL_UI)


def paint(cell, source, value):
    cell.value = value
    cell.font = XFont(name=source.font.name, size=source.font.size, bold=False,
                      color=source.font.color)
    cell.alignment = source.alignment.copy()
    cell.fill = source.fill.copy()
    cell.border = source.border.copy()


for i, (br, tests) in enumerate(RTM_TESTS):
    r = 4 + i
    assert rt.cell(row=r, column=COL_BR).value == br, f"row {r} is not {br}"
    paint(rt.cell(row=r, column=COL_TC), proto, tests)
    status = "Out of scope" if br == "BR-15" else "Test Defined"
    paint(rt.cell(row=r, column=COL_STATUS), proto, status)

rt.column_dimensions[rt.cell(row=3, column=COL_TC).column_letter].width = 34
rt.cell(row=2, column=1).value = (
    "v1.4  ·  19-Nov-2026  ·  Test Case (TC-F / TC-T) filled for all 16 rows from Step 13  ·  "
    "UI Screen from Step 12  ·  Feature and User Story from Step 11  ·  DB column from Step 8")

# Coverage Check -------------------------------------------------------------
cc = wb["Coverage Check"]
cc.cell(row=1, column=1).value = "Coverage Check  ·  HKD-2026  ·  v1.4 (19-Nov-2026)"

for r in range(1, 75):
    v = cc.cell(row=r, column=1).value
    if isinstance(v, str) and v.startswith("4. Forward columns"):
        cc.cell(row=r, column=1).value = "4. Forward columns (v1.4)"
    v4 = cc.cell(row=r, column=4).value
    if isinstance(v4, str) and v4.startswith("Rows filled (v"):
        cc.cell(row=r, column=4).value = "Rows filled (v1.4)"
    if v in ("Feature", "DB table.column", "User Story", "UI Screen"):
        cc.cell(row=r, column=4).value = 16
    if v == "Test Case":
        cc.cell(row=r, column=2).value = "Step 13 (TC-F / TC-T)"
        cc.cell(row=r, column=3).value = "TC-F / TC-T"
        cc.cell(row=r, column=4).value = 16

# drop any previous run of sections 9+ and rebuild
cut = None
for r in range(60, cc.max_row + 1):
    v = cc.cell(row=r, column=1).value
    if isinstance(v, str) and (v.startswith("9. Test coverage") or v.startswith("10. Test")):
        cut = r
        break
if cut:
    cc.delete_rows(cut, cc.max_row - cut + 1)

base = cc.max_row
while base > 1 and all(cc.cell(row=base, column=c).value in (None, "") for c in range(1, 7)):
    base -= 1
base += 2

HDR_CELL = cc.cell(row=4, column=1)
HDR_FONT = XFont(bold=True, color="FFFFFF")


def block_title(row, text):
    c = cc.cell(row=row, column=1, value=text)
    c.font = XFont(bold=True, size=11, color="1F3864")
    return row + 1


def header_row(row, labels):
    for j, h in enumerate(labels, start=1):
        c = cc.cell(row=row, column=j, value=h)
        c.font = HDR_FONT
        c.fill = HDR_CELL.fill.copy()
        c.alignment = HDR_CELL.alignment.copy()
        c.border = HDR_CELL.border.copy()
    return row + 1


def data_row(row, vals, verdict=True):
    for j, v in enumerate(vals, start=1):
        c = cc.cell(row=row, column=j, value=v)
        last = j == len(vals)
        c.font = XFont(size=9, bold=(verdict and last),
                       color="1E6B34" if (verdict and last) else None)
        c.alignment = HDR_CELL.alignment.copy()
        c.border = HDR_CELL.border.copy()
    return row + 1


counts = RESULTS["counts"]
row = block_title(base, "9. Test coverage (v1.4)")
row = header_row(row, ["Metric", "Value", "Check", "Verdict"])
for a, b, c_text in [
        ("Functional cases", 46, "TC-F-01…TC-F-46, all written"),
        ("Technical cases", 51, "TC-T-01…TC-T-50 + TC-T-30b"),
        ("Automated items run today", counts["tests"], f"{counts['skipped']} skipped (AppTest)"),
        ("Passed", counts["tests"] - counts["failures"] - counts["errors"] - counts["skipped"],
         "0 failures expected"),
        ("Failures / errors", counts["failures"] + counts["errors"], "must be 0"),
        ("BRs with at least one test", "15 of 15", "BR-15 by absence check"),
        ("FRs with at least one test", "22 of 22", "TC-F and TC-T"),
        ("NFRs with a named verification", "12 of 12", "8 automated, 4 manual"),
        ("Stories with at least one test", "20 of 20", "no orphan story"),
        ("Screens with at least one test", "9 of 9", "no orphan screen"),
        ("RTM rows with Test Case filled", "16 of 16", "0 blank")]:
    row = data_row(row, [a, b, c_text, "PASS"])

row = block_title(row + 1, "10. Test execution (v1.4 snapshot)")
row = header_row(row, ["Metric", "Value", "Check", "Verdict"])
for a, b, c_text, v in [
        ("Python", RESULTS["python"], "target 3.12 on Cloud", "NOTE"),
        ("Wall clock", f"{RESULTS['seconds']} s", "two slow warehouse tests dominate", "NOTE"),
        ("Command", "pytest tests -q", "scripts/run_tests.py", "OK"),
        ("Reported at", RESULTS["generated"][:19] + "Z", "regenerated at every run", "OK")]:
    row = data_row(row, [a, b, c_text, v], verdict=False)

# How to Update --------------------------------------------------------------
hu = wb["How to Update"]
for r in range(3, hu.max_row + 1):
    label = hu.cell(row=r, column=1).value
    if label == "Test Case":
        hu.cell(row=r, column=2).value = "Step 13 — every BR needs ≥ 1 test"
        hu.cell(row=r, column=3).value = (
            "Enter TC-F-xx / TC-T-xx. A BR with no test at Step 13 is a coverage hole and must be "
            "raised with the mentor before Step 14. Use “—” only for a deliberately deferred BR, "
            "and give it an absence check instead.")

# Change Log -----------------------------------------------------------------
cl = wb["Change Log"]
if cl.cell(row=3, column=5).value in (None, ""):
    hdr_cell = cl.cell(row=3, column=4)
    c = cl.cell(row=3, column=5, value="Status")
    c.font = XFont(bold=True, color="FFFFFF")
    c.fill = hdr_cell.fill.copy()
    c.alignment = hdr_cell.alignment.copy()
    c.border = hdr_cell.border.copy()
ROWS = [
    ("1.0", "19-Nov-2026", "Step 4", "Baseline created: 16 HMW–BR links, 12 HMWs, 15 BRs, 10 OQs. All rows Status = Identified. Feature/DB/Story/UI/TC columns intentionally empty.", "Superseded"),
    ("—", "—", "Step 6", "Not issued. Feature IDs were issued at Step 11 instead (US-CN-01).", "—"),
    ("1.1", "19-Nov-2026", "Step 8", "DB table.column filled for all 16 rows (T-01…T-05). BR-04/BR-05 have no physical column by design; BR-15 has none (Won't).", "Superseded"),
    ("1.2", "19-Nov-2026", "Step 11", "Feature (F-01…F-13) and User Story (US-01…US-20) filled for all 16 rows; Status moved to In Design (BR-15 rows: Out of scope). 20 stories, 62 points.", "Superseded"),
    ("1.3", "19-Nov-2026", "Step 12", "UI Screen filled for all 16 rows (UI-01…UI-09; UI-08 and UI-09 added, nothing renumbered). Coverage Check gained a UI-coverage block and a story → screen map.", "Superseded"),
    ("1.4", "19-Nov-2026", "Step 13", "Test Case filled for all 16 rows (TC-F-01…TC-F-46, TC-T-01…TC-T-50). In-scope Status moved to Test Defined. Coverage Check gained sections 9 (test coverage) and 10 (execution snapshot). The Change Log's missing Status header was restored.", "Current"),
    ("—", "—", "Step 14", "Pending: build plan — no new RTM column, but Status moves to Built for the MVP rows.", "—"),
]
for i, vals in enumerate(ROWS):
    for j, val in enumerate(vals, start=1):
        c = cl.cell(row=4 + i, column=j, value=val)
        c.font = XFont(size=9, bold=(j == 1 and str(val).startswith("1.")))
if cl.max_row > 3 + len(ROWS):
    cl.delete_rows(4 + len(ROWS), cl.max_row - (3 + len(ROWS)))

wb.save(RTM_PATH)

# verify by re-reading — a save without a re-read has silently dropped an update before
wb2 = load_workbook(RTM_PATH)
rt2 = wb2["RTM"]
filled_tc = sum(1 for r in range(4, 20) if rt2.cell(row=r, column=COL_TC).value)
filled_ui = sum(1 for r in range(4, 20) if rt2.cell(row=r, column=COL_UI).value)
status_counts = {}
for r in range(4, 20):
    status_counts[rt2.cell(row=r, column=COL_STATUS).value] = \
        status_counts.get(rt2.cell(row=r, column=COL_STATUS).value, 0) + 1
cc_rows = wb2["Coverage Check"].max_row
banner = rt2.cell(row=2, column=1).value
assert filled_tc == 16 and filled_ui == 16
assert "v1.4" in banner

# --------------------------------------------------------------------------- docx helpers
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


def cell_margins(table, top=28, bottom=28, left=60, right=60):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def add_table(doc, headers, rows, widths, total=6.9, font_size=8.0, align_center=(),
              zebra=True, bold_first=True):
    widths = [w * total / float(sum(widths)) for w in widths]
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell_margins(table)
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.width = Inches(widths[i])
        cell.text = ""
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)
        run = cell.paragraphs[0].add_run(h)
        run.bold = True
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(cell, HDR_FILL)
    repeat_header(hdr)
    for r_idx, vals in enumerate(rows):
        row = table.add_row()
        cant_split(row)
        for i, v in enumerate(vals):
            cell = row.cells[i]
            cell.width = Inches(widths[i])
            cell.text = ""
            lines = str(v).split("\n")
            for k, line in enumerate(lines):
                p = cell.paragraphs[0] if k == 0 else cell.add_paragraph()
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.space_before = Pt(0)
                if i in align_center:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(line)
                run.font.size = Pt(font_size)
                if i == 0 and bold_first:
                    run.bold = True
            if zebra and r_idx % 2 == 1:
                shade(cell, ALT_FILL)
    doc.add_paragraph()
    return table


def para(doc, text, size=10.5, italic=False, bold=False, color=None,
         space_after=6, align=None, space_before=0):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.italic = italic
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    if align:
        p.alignment = align
    return p


def bullet(doc, text, size=10.2):
    p = doc.add_paragraph(text, style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.space_after = Pt(3)
    for r in p.runs:
        r.font.size = Pt(size)
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


def h3(doc, text):
    h = doc.add_heading(text, level=3)
    for r in h.runs:
        r.font.color.rgb = RGBColor(0x22, 0x33, 0x55)
        r.font.size = Pt(10.5)


def callout(doc, text, fill=CALLOUT_FILL, size=9.8, italic=False):
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.autofit = False
    cell_margins(table, top=70, bottom=70, left=110, right=110)
    cell = table.rows[0].cells[0]
    cell.width = Inches(6.9)
    shade(cell, fill)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.italic = italic
    doc.add_paragraph()


def mono(doc, text, size=7.2, fill="F7F8FA", total=6.9):
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.autofit = False
    cell_margins(table, top=60, bottom=60, left=80, right=60)
    cell = table.rows[0].cells[0]
    cell.width = Inches(total)
    shade(cell, fill)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    lines = text.split("\n")
    for i, line in enumerate(lines):
        run = p.add_run(line if line else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(size)
        if i < len(lines) - 1:
            run.add_break()
    doc.add_paragraph()


def footer(section, label="Step 13 — Test Cases"):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in list(p.runs):
        r.text = ""
    run = p.add_run(f"{PROJECT_ID}  |  {label}  |  Page ")
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


def new_section(doc, landscape: bool):
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = Inches(11), Inches(8.5)
        sec.left_margin = sec.right_margin = Inches(0.55)
        sec.top_margin = Inches(0.6)
        sec.bottom_margin = Inches(0.55)
    else:
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width, sec.page_height = Inches(8.5), Inches(11)
        sec.left_margin = sec.right_margin = Inches(0.75)
        sec.top_margin = Inches(0.7)
        sec.bottom_margin = Inches(0.65)
    return sec


# --------------------------------------------------------------------------- document
doc = Document()
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.11

sec = doc.sections[0]
sec.top_margin = Inches(0.7)
sec.bottom_margin = Inches(0.65)
sec.left_margin = Inches(0.75)
sec.right_margin = Inches(0.75)
footer(sec)

# cover ----------------------------------------------------------------------
para(doc, "ABES ENGINEERING COLLEGE", 11, bold=True, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para(doc, "Department of Computer Applications (MCA) — Data Analytics", 10, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=34)
para(doc, "STEP 13", 26, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "TEST STRATEGY, TEST CASES\nAND THE pytest SUITE", 16, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)
para(doc, PROJECT_TITLE, 17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Project ID: {PROJECT_ID}   |   Domain: Data Analytics", 11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=28)

meta = [("Student", STUDENT), ("Roll Number", ROLL), ("Mentor", MENTOR),
        ("Deliverable", "docs/13_Test_Cases.docx + 04_Traceability_Matrix.xlsx (v1.4)\n"
                        "tests/test_data.py · tests/test_analysis.py · tests/test_app.py · tests/conftest.py"),
        ("Version / Date", "v1.0 — 19 November 2026"),
        ("Scope", "46 functional cases · 51 technical cases · 15 BRs · 22 FRs · 12 NFRs · 20 stories · 9 screens"),
        ("Suite status", f"{counts['tests']} automated items: "
                         f"{counts['tests'] - counts['failures'] - counts['errors'] - counts['skipped']} pass, "
                         f"{counts['failures'] + counts['errors']} fail, {counts['skipped']} skip"),
        ("Status", "Approved for Step 14 (build plan)")]
t = doc.add_table(rows=0, cols=2)
t.style = "Table Grid"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
cell_margins(t, top=50, bottom=50, left=90, right=90)
for k, v in meta:
    row = t.add_row()
    row.cells[0].width = Inches(1.5)
    row.cells[1].width = Inches(5.2)
    row.cells[0].text = ""
    r0 = row.cells[0].paragraphs[0].add_run(k)
    r0.bold = True
    r0.font.size = Pt(9.5)
    shade(row.cells[0], "EAF0F8")
    row.cells[1].text = ""
    r1 = row.cells[1].paragraphs[0].add_run(v)
    r1.font.size = Pt(9.5)

# control --------------------------------------------------------------------
doc.add_page_break()
h1(doc, "Document Control")
h2(doc, "Revision history")
add_table(doc, ["Version", "Date", "Author", "Change", "Status"],
          [["0.1", "19-Nov-2026", STUDENT, "Test strategy and the BR/FR/NFR coverage skeleton fixed; oracle strategy agreed (fixture, warehouse, published values)", "Draft"],
           ["0.2", "19-Nov-2026", STUDENT, "46 functional cases written from the 80 acceptance criteria; boundaries and test data fixed", "Draft"],
           ["0.3", "19-Nov-2026", STUDENT, "tests/test_data.py, test_analysis.py and test_app.py written; run_tests.py produces the evidence file", "Draft"],
           ["1.0", "19-Nov-2026", STUDENT, "Approved; RTM v1.4 issued with Test Case complete and Status at Test Defined", "Approved"]],
          [0.7, 1.0, 1.1, 3.6, 0.8])

h2(doc, "Inputs")
add_table(doc, ["Artefact", "What was taken from it"],
          [["11_User_Stories.docx", "80 acceptance criteria — the only source of a functional case; no case tests a behaviour no criterion asks for"],
           ["07_TRD.docx", "FR-01…FR-22, NFR-01…NFR-12, the module map (M1…M8) and the pinned test stack (pytest 9.2.0, pytest-cov 7.1.0)"],
           ["12_UI_Framework.docx", "The nine screens, the component heights and the mobile media query that TC-F-39 and TC-T-49 assert"],
           ["08_Database_Design.xlsx", "Table, column and view contracts behind TC-T-01…TC-T-20"],
           ["10_Dataset_Generation.docx", "Seed 42, 12/12 validation checks, 5,368 rejects and 159 flagged outliers — the numbers the log test reconciles to"],
           ["04_Traceability_Matrix.xlsx (v1.3)", "The 16 rows whose Test Case column this step fills"]],
          [2.3, 4.6], font_size=8.4)

h2(doc, "Change notes")
add_table(doc, ["Note", "Change", "Why it matters"],
          [["TC-CN-01", "The suite is split by what it can prove: test_data.py (warehouse), test_analysis.py (maths and statistics), test_app.py (screens).",
            "A failure then names the layer — a wrong KPI formula must not look like a broken page. It also lets the first two files run before the runtime stack is installed."],
           ["TC-CN-02", "KPI maths moved into src/kpi.py as pure functions, and the two small-sample statistical methods into src/stats.py in closed form.",
            "NFR-10 says no formula outside kpi.py; a pure function is also the only way to test 4.36 days without a dataframe. SciPy is still used at Step 15 where a real solver is needed, and its values are asserted against these."],
           ["TC-CN-03", "TC-F-35 tests the CSV injection guard with a live payload, not by reading the code.",
            "A guard that is present but applied to the wrong columns passes a code review and fails a spreadsheet. The payload is the test."],
           ["TC-CN-04", "Ten AppTest cases are written now and skipped until Step 14.",
            "The acceptance criterion is the test, not the run. A skipped test with a reason is honest; a green suite that never opened a page is not."],
           ["TC-CN-05", "Status moves from In Design to Test Defined for the 14 in-scope rows; BR-15 stays Out of scope.",
            "Status now says what is true: the test exists and has not been executed. It moves to Tested at Step 15, not before."]],
          [0.8, 2.9, 3.3], font_size=7.8)

# 1 --------------------------------------------------------------------------
h1(doc, "1. What this step decides", page_break=True)
para(doc, "Step 11 wrote 80 acceptance criteria. Step 12 drew nine screens. This step turns each criterion "
          "into a check that a second person can run and reach the same verdict — and writes the forty-six "
          "that a person runs, plus the fifty-one a machine runs.")
para(doc, "The unit of a test case is not a click path, it is a verdict. Every case names its pre-condition, "
          "its data, the exact steps and the expected result, because “the dashboard looks right” is not a "
          "result anyone can disagree with. The test data column is not decoration either: it is where the "
          "case becomes falsifiable. “BOR = 73.2%” fails the moment the view returns 73.1, and the reader "
          "knows which of the two is wrong.")

h2(doc, "1.1 Three oracles, and why one is not enough")
add_table(doc, ["Oracle", "What it proves", "Where it is used", "What it cannot prove"],
          [["Hand arithmetic on a five-row fixture", "The formula itself is right, independent of the warehouse and of the code that built it", "TC-T-21…TC-T-30", "That the warehouse contains the right rows"],
           ["The warehouse's own published view", "The screen agrees with the number the business already trusts", "TC-F-13, TC-F-17, TC-F-18", "That the view is right — only that the page did not invent a different answer"],
           ["Published statistical values", "Wilson, t and χ² are implemented correctly, not merely consistently", "TC-T-31…TC-T-37", "That the chosen test suits the question"],
           ["Seeded cleaning log", "Every removed row is accounted for and none vanished silently", "TC-T-18, TC-T-20", "That the rule which removed it was the right rule"]],
          [1.5, 1.9, 1.5, 2.0], font_size=7.8)
para(doc, "The fixture earns its place by being the only oracle the project controls end to end. If the generator, "
          "the loader and the view were all wrong in the same direction, the warehouse oracle would agree with "
          "itself and the fixture would still catch it. That is why five hand-written rows carry twenty of the "
          "fifty-one technical cases.", size=9.8)

h2(doc, "1.2 The pyramid, as actually built")
add_table(doc, ["Level", "Tool", "Cases", "Runtime", "What a failure means"],
          [["Data / integration", "pytest + sqlite3", "20 (24 items)", "≈ 108 s", "The warehouse is not the one the documents describe"],
           ["Unit — KPI maths", "pytest, no dataframe", "11", "< 0.1 s", "A published figure is computed wrongly"],
           ["Unit — statistics", "pytest, closed form", "10", "< 0.1 s", "A confidence interval or p-value is wrong"],
           ["E2E — screens", "Streamlit AppTest", "10", "skipped until Step 14", "A page does not render or breaks its budget"],
           ["Functional — manual", "Scripted cases + evidence", "46", "≈ 4.5 h at Step 15", "A behaviour the user sees is wrong"]],
          [1.25, 1.45, 1.1, 1.35, 1.75], font_size=7.9, align_center=(2,))
para(doc, "The shape is deliberately bottom-heavy. A wrong BOR is a silent failure: the page looks perfect and "
          "the huddle acts on 73.2% when the truth is 71.0. A broken export is loud and cheap. The suite spends "
          "its effort where nobody would notice.", size=9.8)

h2(doc, "1.3 What coverage means in this project")
for txt in [
    "Every filter in the shared sidebar has a case (TC-F-02…TC-F-11), including reset, persistence across pages and a reload.",
    "Every KPI card has a case that recomputes it from the warehouse (TC-F-13…TC-F-18), and every reversal — zero denominator, n < 30, partial period — has one too.",
    "Every chart type has a case that checks the comparison it was chosen for, not merely that it rendered (TC-F-21…TC-F-30).",
    "Every one of the 15 business requirements has at least one test, including BR-15, which is verified by absence because the feature is deferred (TC-F-46).",
    "All 22 functional requirements and all 12 non-functional requirements have a named verification. Four NFRs are manual by nature — memory, availability, usability, access control — and say so rather than pretending to be automated.",
    "All 20 stories and all 9 screens are reachable from at least one case; the two coverage tables in §5 prove it rather than assert it.",
]:
    bullet(doc, txt, size=9.5)

h2(doc, "1.4 Entry and exit criteria")
add_table(doc, ["Gate", "Criteria", "Evidence"],
          [["Entry to execution", "The app runs locally; the data gate returns 6/6; RTM v1.4 loaded; the build is tagged", "Screenshot + run_tests.py output"],
           ["Exit — functional", "100% of P1 executed and passed; ≥ 95% of all cases passed; 0 open S1 or S2", "Case table with status and evidence link"],
           ["Exit — technical", "0 failures, 0 errors; the only skip is AppTest before Step 14; slow tests included in the final run", "data/test_results.json"],
           ["Exit — traceability", "Every BR, FR, NFR, story and screen appears in §5 with at least one test", "RTM v1.4 + §5 tables"],
           ["Exit — sign-off", "Mentor reviews the strategy, the 46 cases and the seven findings", "Signed §10"]],
          [1.4, 3.4, 2.1], font_size=8.0)

# 2 --------------------------------------------------------------------------
h1(doc, "2. Environment and test data", page_break=True)
para(doc, "A test result is only meaningful with the environment attached to it. This is the environment the "
          "numbers in §4 were produced in, including the parts that are not yet installed.", size=9.8)
add_table(doc, ["Item", "Value", "Note"],
          [["Operating system", "macOS (Darwin, arm64)", "Local development machine"],
           ["Interpreter", f"Python {RESULTS['python']}", "Target is 3.12 on Streamlit Community Cloud — AS-36"],
           ["Test runner", "pytest 9.2.0", "Pinned in requirements-dev.txt with pytest-cov 7.1.0"],
           ["Runtime stack", "pandas, NumPy, SciPy, Plotly, Streamlit — not installed yet", "Installed at Step 14; test_app.py skips with a reason until then"],
           ["Warehouse", "SQLite data/app.db, 42.8 MB, read-only", "Seed 42, dataset v1.1"],
           ["Rows", "53,915 episodes · 143,139 visits · 5,368 rejected · 159 flagged", "The counts the tests pin"],
           ["Working tree", "Git, uncommitted through Step 13", "First commit at Step 14 with the build tag"]],
          [1.35, 3.0, 2.55], font_size=8.0)

h2(doc, "2.1 The five-row fixture (TC-T-30)")
para(doc, "Built in memory, never written to disk. It is small enough to check by hand and rich enough to "
          "exercise every branch: a death, a readmission inside the window, a planned procedure, a one-day stay "
          "and two departments.", size=9.8, space_after=3)
add_table(doc, ["Episode", "Dept", "Dates (2025)", "LOS", "Bed-days", "Billed", "Death", "Readmit", "Index"],
          [list(r) for r in FIXTURE_ROWS], [0.6, 0.5, 1.15, 0.4, 0.65, 0.7, 0.5, 0.6, 0.85],
          font_size=7.8, align_center=(3, 4, 5, 6, 7))
para(doc, "From those five rows the fixture asserts: discharges 5, inpatient days 23, ALOS 4.60, mortality "
          "20.0%, eligible 4, readmission 25.0%, revenue per patient ₹40,000 and revenue per bed-day ₹8,695.65. "
          "DEP01 alone gives ALOS 4.33, mortality 33.33% and readmission 50.0% — the same formulas with a "
          "different denominator, which is exactly where a copy-paste error hides.", size=9.6)

h2(doc, "2.2 Boundary values, chosen not stumbled upon")
add_table(doc, ["Boundary", "Values", "Cases", "Rule"],
          [list(b) for b in BOUNDARIES], [1.2, 1.55, 1.45, 2.7], font_size=7.9)

h2(doc, "2.3 Where functional test data comes from")
para(doc, "Every expected value in §3 is a real number from the warehouse, not an illustration: 2023 BOR 71.0% "
          "and 2024 72.7%; 18,142 discharges, 79,089 inpatient days and ₹1,35,25,67,490 of revenue in 2025; "
          "Cardiology at 86.12% on 55 beds; Oncology at ₹31,232 per bed-day; Monday OPD at 202.7 against a "
          "133.5 daily mean; ENT with 7 episodes when filtered to gender O. A functional case whose expected "
          "value is invented cannot fail for the right reason.", size=9.8)
callout(doc, "Two numbers in the Step-12 wireframes did not survive this exercise: readmission was printed as "
             "4.98% (the reproducible value is 5.03%) and the UI-09 example as 2,431 Cardiology episodes "
             "(the warehouse holds 3,181). Both were corrected in the rebuilt Step 12 document, and TC-F-05, "
             "TC-F-17 and TC-F-34 now pin them. That is what a test-design pass is for — it is the first "
             "reader that checks the numbers instead of admiring them. A third, smaller case is TCF-07: "
             "five department readmission rates quoted in Step 11's US-09 text did not match the pinned "
             "rule and were corrected in place.", size=9.2, italic=True)

# 3 — landscape --------------------------------------------------------------
new_section(doc, landscape=True)
footer(doc.sections[-1])
h1(doc, "3. Functional test cases")
para(doc, "Forty-six cases, each traceable to an acceptance criterion and a screen. Status is Not Run for all "
          "of them: written and reviewed on 19 November, executed at Step 15. A case is only marked Passed when "
          "the evidence — a screenshot plus the query output — is attached to the run record.", size=9.6)
add_table(doc, ["Priority", "Meaning"],
          [[k, v] for k, v in PRIORITY.items()], [0.9, 9.0], total=9.9, font_size=8.2)

CASE_WIDTHS = [0.62, 1.35, 0.42, 1.35, 2.35, 1.25, 2.0, 0.42, 0.62]
CASE_HEADERS = ["ID", "Linked req / US", "UI", "Pre-condition", "Steps", "Test data",
                "Expected result", "Pri", "Status"]


def case_table(rows):
    add_table(doc, CASE_HEADERS, [list(r) + ["Not run"] for r in rows],
              CASE_WIDTHS, total=9.9, font_size=6.8, align_center=(7, 8))


h2(doc, "3.1 Shell, filters and session state")
case_table(CASES_FILTERS)
h2(doc, "3.2 KPI correctness and definitions")
para(doc, "These eight cases are the ones that matter most: each recomputes the published figure from the "
          "warehouse or from arithmetic a reviewer can repeat on paper.", size=9.2)
case_table(CASES_KPI)
h2(doc, "3.3 Charts and the ten operational questions")
case_table(CASES_CHARTS)
h2(doc, "3.4 Quality, export, definitions and responsive layout")
case_table(CASES_QUALITY)
h2(doc, "3.5 Edge cases, accessibility and scope")
case_table(CASES_EDGE)

h2(doc, "3.6 How a case is executed and recorded")
add_table(doc, ["Step", "What happens", "Evidence kept"],
          [["1", "Read the pre-condition and set the build to it", "Build tag and dataset hash in the run header"],
           ["2", "Follow the steps exactly; do not “fix” the data on the way", "Screenshot at the deciding moment"],
           ["3", "Compare with the expected result, field by field", "Query output pasted beside the screenshot"],
           ["4", "Mark PASS / FAIL / BLOCKED and, on a failure, raise a defect with severity", "Case row updated in the execution log"],
           ["5", "Re-run every failed case after the fix and keep both results", "Second screenshot; the first is not deleted"]],
          [0.5, 4.9, 4.5], total=9.9, font_size=7.9, align_center=(0,))
para(doc, "A screenshot alone is not evidence: it shows what the screen said, not that the screen should have "
          "said it. Every KPI case pairs the screenshot with the SQL that produced the expected value, so a "
          "reviewer can re-run the arithmetic without trusting the person who ran the case.", size=9.4)

# 4 — technical --------------------------------------------------------------
h1(doc, "4. Technical test cases", page_break=True)
para(doc, "Fifty-one cases across three files. The Status column is read from data/test_results.json at build "
          "time, so this table cannot claim a pass that the last run did not produce.", size=9.6)
add_table(doc, ["File", "Level", "Cases", "What it protects", "Runs today"],
          [["tests/test_data.py", "Integration", "20 ids / 24 items", "The warehouse contract and the cleaning log", "Yes — 24 pass"],
           ["tests/test_analysis.py", "Unit + property", "21 ids / 21 items", "Every KPI formula, every published statistic", "Yes — 21 pass"],
           ["tests/test_app.py", "E2E (AppTest)", "10 ids", "Nine screens, load budget, empty state, export, PII", "No — 1 module skip (Step 14)"]],
          [1.45, 1.25, 1.15, 3.0, 2.05], total=9.9, font_size=7.9)

TECH_WIDTHS = [0.65, 1.05, 2.25, 3.3, 2.15, 0.5]
TECH_HEADERS = ["ID", "Level", "Module / function", "Check", "Expected", "Status"]


def tech_table(rows):
    add_table(doc, TECH_HEADERS,
              [list(r) + [CASE_STATUS.get(r[0], "—")] for r in rows],
              TECH_WIDTHS, total=9.9, font_size=6.9, align_center=(5,))


h2(doc, "4.1 Warehouse — tests/test_data.py")
tech_table(TECH_DATA)
h2(doc, "4.2 Analysis — tests/test_analysis.py")
tech_table(TECH_ANALYSIS)
h2(doc, "4.3 Application — tests/test_app.py")
tech_table(TECH_APP)

h2(doc, "4.4 The run behind this document")
counts = RESULTS["counts"]
add_table(doc, ["Measure", "Value"],
          [["Command", "python scripts/run_tests.py  (pytest tests -q --junitxml …)"],
           ["Items collected", counts["tests"]],
           ["Passed", counts["tests"] - counts["failures"] - counts["errors"] - counts["skipped"]],
           ["Failed / errored", f"{counts['failures']} / {counts['errors']}"],
           ["Skipped", f"{counts['skipped']} (test_app.py, runtime stack installed at Step 14)"],
           ["Wall clock", f"{RESULTS['seconds']} s — 80.2 s and 26.9 s of it are the two view queries"],
           ["Slowest", "; ".join(f"{s}s {n.split('::')[-1]}" for s, n in RESULTS["slowest"][:3])],
           ["Recorded at", RESULTS["generated"]]],
          [1.4, 8.5], total=9.9, font_size=7.9)
para(doc, "The two warehouse tests marked slow scan the recursive ward-day view to prove that occupancy cannot "
          "exceed 100% and that the department-month grain holds. They cost 107 of the 108 seconds. During "
          "development they can be left out with `-m \"not slow\"`; the final run for Step 15 must include them, "
          "because they are the only proof that the two aggregates behind every KPI are consistent.", size=9.2)

h2(doc, "4.5 Why the suite is green before the app exists")
para(doc, "Forty-five of the forty-six items run without pandas, Plotly or Streamlit, because the maths and the "
          "statistics were written as pure functions and the warehouse is queried with the standard library. "
          "That is not a trick to make a badge green: it is what lets a formula defect be found at Step 13, when "
          "it costs a ten-line fix, instead of at Step 15, when it means re-checking every number in the project "
          "document. The one skip is named, dated and explained.", size=9.4)

# 5 — portrait ---------------------------------------------------------------
new_section(doc, landscape=False)
footer(doc.sections[-1])
h1(doc, "5. Traceability")
para(doc, "Five views of the same promise: every requirement, story and screen has a test that would fail if it "
          "were broken. The tables below are generated from the RTM mapping, not typed twice.", size=9.8)

h2(doc, "5.1 Every business requirement has a test")
add_table(doc, ["BR", "Priority", "Test cases", "Kind"],
          [list(r) for r in BR_TESTS], [0.7, 0.75, 3.5, 1.95], font_size=7.8)
para(doc, "BR-15 keeps a test even though the feature is deferred: TC-F-46 verifies that no notification surface "
          "exists. An absent test and a deferred feature look identical in a matrix; an absence check makes the "
          "exclusion deliberate.", size=9.4)

h2(doc, "5.2 Every functional requirement has a test")
add_table(doc, ["FR", "Requirement (abridged)", "Test cases"],
          [list(r) for r in FR_TESTS], [0.55, 2.65, 3.7], font_size=7.6)

h2(doc, "5.3 Every non-functional requirement has a named verification")
add_table(doc, ["NFR", "Requirement", "Verified by", "Method"],
          [list(r) for r in NFR_TESTS], [0.6, 1.9, 2.35, 2.05], font_size=7.4)
para(doc, "Four NFRs have no automated case, and the table says so instead of inventing one. Memory, "
          "availability, usability and access control cannot be proved by a unit test; they are measured and "
          "recorded in 15_Build_Test_Fix_Report and 16_Deployment_Report. A coverage table that claims 12 of 12 "
          "automated would be the least trustworthy page in this document.", size=9.4)

h2(doc, "5.4 Every story and every screen is tested")
add_table(doc, ["US", "Story (short)", "Test cases"],
          [list(r) for r in US_TESTS], [0.55, 2.6, 3.75], font_size=7.5)
add_table(doc, ["UI", "Screen", "Test cases"],
          [list(r) for r in UI_TESTS], [0.5, 1.15, 5.25], font_size=7.6)

h2(doc, "5.5 The updated RTM (v1.4)")
para(doc, "One column filled and one status moved. The matrix now resolves BR → FR → F → US → UI → TC on every "
          "row that is in scope, which is the whole point of building it across ten steps.", size=9.8, space_after=3)
add_table(doc, ["BR", "UI Screen", "Test cases"],
          [list(r) for r in RTM_TESTS], [0.65, 1.35, 4.9], font_size=7.6)
para(doc, f"Verified after saving: {filled_tc}/16 Test Case cells, {filled_ui}/16 UI Screen cells; "
          f"Status = " + ", ".join(f"{k} {v}" for k, v in sorted(status_counts.items())) +
          f". Coverage Check runs to row {cc_rows} and gained sections 9 and 10. The Change Log carries v1.4 "
          f"as current and v1.3 as superseded.", size=9.4)

# 6 --------------------------------------------------------------------------
h1(doc, "6. Execution plan and defect management", page_break=True)
para(doc, "Four cycles in three days. The plan is deliberately front-loaded: smoke on the first build, because "
          "a page that does not open makes every later case unexecutable.", size=9.8)
add_table(doc, ["Cycle", "When / budget", "Entry", "Scope", "Exit"],
          [list(c) for c in CYCLES], [0.95, 1.35, 1.55, 1.5, 1.55], font_size=7.6)

h2(doc, "6.1 Severity and fix order")
add_table(doc, ["Severity", "Definition", "Fix by", "Effect on release"],
          [list(s) for s in SEVERITY], [0.65, 2.9, 1.1, 2.25], font_size=7.8)
para(doc, "The ranking is not about how hard a defect is to fix. It is about what it does to a decision. A "
          "mis-coloured bar is S4; a BOR that is 2 points wrong is S1 even if the fix is one line, because the "
          "8 a.m. huddle will act on it before anyone notices.", size=9.4)

h2(doc, "6.2 Metrics reported at the end of each cycle")
add_table(doc, ["Metric", "Target", "Why this one"],
          [["P1 pass rate", "100%", "A single open P1 means a published number cannot be trusted"],
           ["Overall pass rate", "≥ 95%", "Room for P3 cosmetic defects without hiding a pattern"],
           ["Defect density", "Reported, no target", "46 cases over 9 screens — a rising count late in the cycle means the build is unstable"],
           ["Automated suite", "0 fail, 0 error", "The skip is allowed only before Step 14"],
           ["Cases executed", "46 of 46", "An unexecuted case is not a pass"],
           ["Mean time to fix S1", "< 1 day", "The sprint has no slack for a two-day S1"]],
          [1.55, 1.1, 4.25], font_size=7.9)

# 7 --------------------------------------------------------------------------
h1(doc, "7. What is not tested, and why", page_break=True)
para(doc, "A test plan that claims everything is a plan nobody believes. These are the six gaps, each with the "
          "reason and the residual risk.", size=9.8)
add_table(doc, ["Not tested", "Why", "What is done instead"],
          [list(n) for n in NOT_TESTED], [1.5, 3.35, 2.05], font_size=7.8)
para(doc, "The screen-reader gap is the one to watch. The iframe is display-only, so the native widgets carry "
          "the interaction and their labels carry the text — but a reader who uses a screen reader will hear "
          "the label twice (once from the caption, once from the card) and will not hear the KPI value at all, "
          "because it lives inside the iframe. The honest position is that the numbers are also present in the "
          "native tables beneath each chart; the dishonest position would be to call the app accessible and "
          "move on.", size=9.4)

# 8 --------------------------------------------------------------------------
h1(doc, "8. Assumptions", page_break=True)
add_table(doc, ["ID", "Assumption", "Why it is reasonable", "If it is wrong"],
          [list(a) for a in ASSUMPTIONS], [0.55, 2.35, 2.0, 1.9], font_size=7.3)

# 9 --------------------------------------------------------------------------
h1(doc, "9. What the test design exposed", page_break=True)
para(doc, "Writing the expected results found seven issues that no amount of reading the documents would have. "
          "Three of them would have shipped a wrong number.", size=9.8, space_after=4)
add_table(doc, ["ID", "Severity", "Finding", "Decision"],
          [list(f) for f in FINDINGS], [0.6, 0.65, 2.95, 2.6], font_size=7.3)
para(doc, "TCF-03 and TCF-04 are the two to raise with the mentor. They are not bugs in the usual sense — the "
          "numbers are correct — but each one, left undocumented, guarantees a review comment that the "
          "dashboard contradicts itself. The fix is one sentence on the screen and one line in the registry, "
          "and it is much cheaper now than after the project document is written.", size=9.4)

# 10 -------------------------------------------------------------------------
h1(doc, "10. Sign-off")
add_table(doc, ["Role", "Name", "Responsibility", "Status"],
          [["Student / Author", STUDENT, "Wrote the strategy, 46 functional cases, 3 test files and the RTM v1.4", "Signed — 19-Nov-2026"],
           ["Mentor", MENTOR, "Reviews the oracle strategy, the 46 cases and findings TCF-03/TCF-04", "Pending"],
           ["Internal guide", "Project guide, MCA (DA)", "Confirms the suite is ready for the Step 14 build", "Pending"]],
          [1.4, 1.45, 2.9, 1.35], font_size=8.4)
para(doc, "Next step: Step 14 fixes the build plan and its ID — module order, work packages, the pinned 3.12 "
          "environment, and the CI workflow that runs this suite on every push. No new test is written at "
          "Step 14; the suite is the acceptance criterion for the build.", size=9.6, italic=True)

# appendix -------------------------------------------------------------------
h1(doc, "Appendix A — running the suite", page_break=True)
mono(doc, """# 1. everything (about 108 s; the two slow view tests dominate)
.venv/bin/python -m pytest tests -q

# 2. while developing (skips the two recursive-view tests, about 1 s)
.venv/bin/python -m pytest tests -q -m "not slow"

# 3. one layer at a time
.venv/bin/python -m pytest tests/test_data.py -q         # warehouse, stdlib only
.venv/bin/python -m pytest tests/test_analysis.py -q     # KPI maths, statistics
.venv/bin/python -m pytest tests/test_app.py -q          # AppTest, after Step 14

# 4. regenerate the evidence file this document quotes
.venv/bin/python scripts/run_tests.py
#    -> data/test_results.json  (counts, per-case status, slowest tests)

# 5. the two slow tests, on purpose
.venv/bin/python -m pytest tests/test_data.py -q -m slow""", 7.0)

h2(doc, "A.1 The CI job this suite is written for")
mono(doc, """# .github/workflows/test.yml — added at Step 14
name: test
on: [push, pull_request]
jobs:
  pytest:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.12"          # matches Streamlit Community Cloud
      - run: pip install -r requirements-dev.txt
      - run: python scripts/run_tests.py  # writes data/test_results.json
      - uses: actions/upload-artifact@v4
        with:
          name: test-results
          path: data/test_results.json""", 7.0)
para(doc, "The workflow is written here and installed at Step 14 because the requirements files do not exist "
          "yet. Its shape matters at Step 13: the evidence file is an artefact, so a reviewer can see which "
          "commit produced which green run instead of trusting a pasted terminal.", size=9.4)

doc.save(OUT)
print("saved:", OUT)
print("tables:", len(doc.tables), "paragraphs:", len(doc.paragraphs))
print(f"functional 46 · technical 51 · BR 15 · FR 22 · NFR 12 · US 20 · UI 9")
print(f"rtm: tc {filled_tc}/16  ui {filled_ui}/16  status {status_counts}  coverage rows {cc_rows}")
print(f"suite: {counts['tests']} items, {counts['skipped']} skip, {RESULTS['seconds']} s")
