"""Build the Step 11 deliverable: user stories, acceptance criteria, sprint plan,
and the RTM v1.2 update (Feature + User Story columns filled)."""
import json
import os
import sqlite3

from docx import Document
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

PROJECT_ID = "HKD-2026"
PROJECT_TITLE = "Healthcare KPI Dashboard"
STUDENT = "Ayush Pal"
ROLL = "2503201450035"
MENTOR = "Aplana Lodhi"

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
DOCS = os.path.join(ROOT, "docs")
OUT = os.path.join(DOCS, "11_User_Stories.docx")
RTM_PATH = os.path.join(DOCS, "04_Traceability_Matrix.xlsx")

summary = json.load(open(os.path.join(DATA, "validation_summary.json"), encoding="utf-8"))
con = sqlite3.connect(os.path.join(DATA, "app.db"))
con.row_factory = sqlite3.Row

# --------------------------------------------------------------------------- stories
SHORT = {
    "SH-01": "Hospital Director", "SH-02": "Operations Manager", "SH-03": "Department Head",
    "SH-04": "Quality Officer", "SH-05": "Finance Manager", "SH-06": "Nursing Superintendent",
    "SH-07": "Data Governance Officer", "SH-08": "Front-desk Supervisor",
}
P = {
    "SH-01": "Hospital Director / Medical Superintendent",
    "SH-02": "Operations Manager",
    "SH-03": "Department Head",
    "SH-04": "Quality / NABH Accreditation Officer",
    "SH-05": "Finance / Revenue Cycle Manager",
    "SH-06": "Nursing Superintendent",
    "SH-07": "IT / Data Governance Officer",
    "SH-08": "Registration / Front-desk Supervisor",
}

EPICS = [
    ("E1", "Data", "Make the warehouse trustworthy and the filters honest before any number is drawn.",
     "Every page reads the same filtered frames, and a wrong filter can never produce a wrong denominator."),
    ("E2", "Dashboard", "Six KPI families on screen, each with its target, denominator and period.",
     "The 8 a.m. huddle can run from the Operations page alone."),
    ("E3", "Analysis", "The comparisons that need a definition: readmission, payer, revenue, mix, mortality.",
     "Each panel states its rule inline, so two readers cannot disagree about what the number means."),
    ("E4", "Insights", "Turn a movement into a cause and an association into a defensible statement.",
     "The director can name the driver; the analyst can defend the test."),
    ("E5", "Quality", "Definitions, data-quality status, low-confidence rules, the reject log and export.",
     "Nothing is published without its definition, its denominator and its data-quality context."),
]

STORIES = [
    dict(id="US-01", title="Open the database read-only and load typed frames",
         epic="E1", sh="SH-07", feature="F-01", br="BR-14", pri="Must", pts=2, day=1,
         want="load the five tables from data/app.db over a read-only connection and get typed DataFrames back",
         benefit="every page starts from the same vetted data and no screen can write to the warehouse",
         acs=[
             ("the packaged data/app.db exists and the app starts", "the connection opens",
              "it opens with mode=ro, all five tables load, and vw_analysis_episode returns its 36 columns; an INSERT attempt raises sqlite3.OperationalError instead of succeeding"),
             ("the frames are loaded", "a page requests episodes with department and ward names",
              "the join is done in SQL once and returns one row per episode (53,915) with department and ward names attached, with no per-row query in a loop"),
             ("a table is empty or the file is missing", "the validation gate runs at startup",
              "the app stops with the failed check named (table, expected row count, actual) and no dashboard page renders"),
             ("a second page is opened within the session", "the cached loader is called again",
              "st.cache_data returns in under 200 ms and the read-only connection is reused, not reopened"),
         ]),
    dict(id="US-02", title="One filter bar that every KPI obeys",
         epic="E1", sh="SH-02", feature="F-01", br="BR-04", pri="Must", pts=4, day=1,
         want="set period, department, ward/bed class, payer, age band and gender once and have every card, chart, table and denominator follow",
         benefit="I can answer a question in three interactions instead of rebuilding a pivot",
         acs=[
             ("the dashboard is open with default filters", "I pick Department = Cardiology and Period = 2025",
              "every KPI card, chart, table and denominator on all seven pages shows Cardiology-2025 values, and the filter chip summary lists both selections"),
             ("Cardiology + 2025 is applied", "I add Payer = Government scheme",
              "the discharge count becomes 366 and every card on the page divides by that same 366 — no card still uses the unfiltered total"),
             ("I select Maternity + Gender = M + Age = 0–5", "the filters are applied",
              "the page shows the empty state with the three active filters listed and a Reset button, and no exception or NaN card is drawn"),
             ("a filter is changed on page 3", "I navigate to page 5",
              "the selection persists in st.session_state, the chips are still visible, and a page reload restores the same view from the URL"),
         ]),
    dict(id="US-03", title="Mark in-progress periods and keep them out of rankings",
         epic="E1", sh="SH-05", feature="F-01", br="BR-01, BR-09", pri="Must", pts=2, day=1,
         want="see which period is still filling up and have it excluded from rankings and averages",
         benefit="a half-finished month never appears as a collapse in revenue or a spike in occupancy",
         acs=[
             ("the loaded data stops at 2025-12-15 but December is selected", "the period renders",
              "the header badge reads “in progress — 15 of 31 days” and the period is excluded from every ranking and moving average"),
             ("a partial month is on the trend chart", "the chart draws",
              "the partial point is hollow and labelled, and it is not used in the 3-month moving average or the league table"),
             ("December is excluded from the 2025 ranking", "the ranking footnote renders",
              "it names the excluded period and the reason, and the row count reads n = 11 months, not 12"),
             ("the period selector reaches 31 Dec 2026, after the data ends", "I apply it",
              "the app clamps to the last loaded date, says so, and does not show an empty year as zero revenue"),
         ]),
    dict(id="US-04", title="Occupancy against the 85% target, in points and beds",
         epic="E2", sh="SH-02", feature="F-02", br="BR-01", pri="Must", pts=4, day=2,
         want="see each department's bed occupancy for the week and month against the 85% target, with the gap in percentage points and in beds",
         benefit="I can escalate a likely miss before the 8 a.m. huddle instead of discovering it at midnight",
         acs=[
             ("all filters are at their defaults and Period = 2025", "I open the Operations page",
              "Cardiology reads BOR 86.1%, gap +1.1 pp and +0.6 beds against 55 functional beds, and the card is red because it is above target"),
             ("the same view is sorted by occupancy", "the league table renders",
              "the order is Cardiology 86.1, Oncology 76.7, Paediatrics 73.8, General Medicine 72.9, Orthopaedics 69.9, Emergency 67.4, Maternity 62.3, ENT 50.0"),
             ("I drill from the Cardiology card into its four wards", "the ward table opens",
              "Σ occupied bed-days ÷ Σ available bed-days equals 86.1% to 0.1 pp, and available = functional_beds × 365 for each ward row"),
             ("a department sits at exactly 85.0%", "the status is computed",
              "it is labelled “on target”, not red, and the boundary case is covered by a test at 84.9 / 85.0 / 85.1"),
         ]),
    dict(id="US-05", title="Seven-day ward projection from recorded schedules",
         epic="E2", sh="SH-06", feature="F-02", br="BR-02", pri="Must", pts=3, day=2,
         want="see ward-level occupancy for the next seven days built only from beds occupied today, confirmed admissions and recorded discharge dates",
         benefit="I can move staff before a ward crosses 90%, not after handover",
         acs=[
             ("the as-of date is 2025-12-15 and W05 has 20 beds", "the seven-day projection renders",
              "each day shows occupied = in-house at as-of + admissions with admission_date ≤ day − discharges with discharge_date ≤ day, never above 20, and every input row is listed"),
             ("a projected ward-day reaches 18 of 20 beds (90%)", "the row is drawn",
              "it is highlighted amber and the remaining headroom is shown as 2 beds"),
             ("a ward has no recorded discharge for the window", "the projection renders",
              "occupancy is held flat and the row is labelled “no schedule recorded”, rather than dropping to zero"),
             ("the projection is displayed", "the reader looks for the method",
              "the panel carries “deterministic projection — no forecasting model (X3)” and links to the method note"),
         ]),
    dict(id="US-06", title="Like-for-like ALOS with the case-mix basis stated",
         epic="E2", sh="SH-03", feature="F-03", br="BR-06", pri="Must", pts=3, day=2,
         want="compare my department's length of stay with peers on a stated basis, not a bare average",
         benefit="I can accept or challenge the benchmark in the review meeting instead of disputing the arithmetic",
         acs=[
             ("Period = 2025", "I open the Departments page",
              "ALOS reads Oncology 6.49, Orthopaedics 5.69, Cardiology 5.38, General Medicine 4.71, Paediatrics 3.40, ENT 3.22, Maternity 2.76, Emergency 2.53, each with its gap to the 4.0-day benchmark in days"),
             ("the info icon beside ALOS is opened", "the definition panel shows",
              "it states “mean of los_days; includes deaths and LAMA; not adjusted for age, payer or diagnosis; n = 18,142 discharges in 2025”"),
             ("a department has fewer than 30 discharges in the selected period", "its ALOS renders",
              "the value carries the low-confidence badge and the department is excluded from the league table"),
             ("two periods are compared", "the delta is shown",
              "the change is expressed in days and in per cent, and both sides use the same inclusion rule (deaths and LAMA included)"),
         ]),
    dict(id="US-07", title="OPD demand by weekday and month with a peak ratio",
         epic="E2", sh="SH-06", feature="F-06", br="BR-09", pri="Must", pts=3, day=3,
         want="see outpatient footfall by weekday and month, with the peak-to-average ratio and a per-department split",
         benefit="I can set rosters and token capacity before the Monday queue forms",
         acs=[
             ("Period = 2025 and all departments", "the OPD panel renders",
              "Monday averages 203 visits a day (10,542 visits over 52 Mondays) against 133 a day overall, a ratio of 1.52, and Sunday averages 34"),
             ("a single month is selected", "the season strip draws",
              "the Nov–Feb lift is visible and the per-department bars sum to the displayed total within 1 visit of rounding"),
             ("Department = Paediatrics is applied", "the weekday pattern recomputes",
              "the ratio is recalculated from the filtered rows, not reused from the unfiltered view, and the New / Follow-up split is shown"),
             ("the ratio is displayed", "the tooltip is opened",
              "its definition reads “highest weekday average ÷ overall daily average” and names the weekday that produced it"),
         ]),
    dict(id="US-08", title="A responsive shell that survives a phone and a projector",
         epic="E2", sh="SH-01", feature="F-13", br="BR-04", pri="Must", pts=4, day=3,
         want="open the executive summary on a phone before a meeting and on a projector in it, without clipping or sideways scrolling",
         benefit="the same six numbers are readable wherever the conversation happens",
         acs=[
             ("the app is open at 1440 px", "the page renders",
              "header and the six KPI cards are drawn in one st.iframe component using Tailwind classes, with values injected server-side and no fetch call in the iframe HTML"),
             ("the viewport is set to 375, 768 and 1440 px", "each page is loaded",
              "no horizontal scrollbar appears, no label is clipped, and cards reflow to 1, 2 and 3 per row; verified by screenshot at each width"),
             ("the iframe fails to load", "the page continues",
              "a plain Streamlit fallback shows the same six values, so the dashboard degrades rather than blanks"),
             ("a KPI value is unavailable", "the card renders",
              "it shows “—” with a tooltip naming the failed check, never “None” or “nan”"),
         ]),
    dict(id="US-09", title="30-day readmission rate with the rule stated inline",
         epic="E3", sh="SH-04", feature="F-04", br="BR-07", pri="Must", pts=4, day=3,
         want="see the 30-day unplanned readmission rate by department and diagnosis group, with the window and exclusions written beside it",
         benefit="I can tell a clinical signal from a coding artefact before I open a CAPA",
         acs=[
             ("Period = 2025", "I open the readmission panel",
              "General Medicine reads 7.70%, Oncology 6.36%, Cardiology 5.33%, Orthopaedics 1.98%, Maternity 1.32%, and the rule line reads “unplanned readmissions within 30 days ÷ index discharges; planned and expired index excluded”"),
             ("I pick General Medicine + Respiratory", "the diagnosis filter applies",
              "the rate, the numerator, the denominator and the contributing episode ids are listed, and n < 30 triggers the low-confidence badge"),
             ("a readmission is counted", "I drill through it",
              "it resolves to exactly one index_episode_id with the gap in days, and no index episode is counted twice"),
             ("the panel renders", "the data-quality strip loads",
              "it shows the period's duplicate count and missing-payer count inline, as BR-07 requires"),
         ]),
    dict(id="US-10", title="Payer-mix movement and revenue per patient",
         epic="E3", sh="SH-05", feature="F-05", br="BR-08", pri="Should", pts=3, day=4,
         want="see payer share by period, how it moved, and revenue per patient for each of the five categories",
         benefit="I can raise a payer negotiation or a pricing review inside the quarter, not after it closes",
         acs=[
             ("2023 and 2025 are compared", "the Finance page renders",
              "Corporate moves 6.37% → 11.23% and self-pay 41.35% → 33.49%, with the movement shown in percentage points and a direction arrow"),
             ("Period = 2025", "revenue per patient is displayed",
              "Private insurance ₹82,544, Corporate ₹79,973, TPA ₹75,890, Cash ₹71,258, Government scheme ₹56,574, each labelled with its n"),
             ("one payer category is selected", "the panel recomputes",
              "share and revenue per patient are recalculated for that subset, and the share total is shown as 100% of the filtered set (±0.1 pp)"),
             ("share is compared across periods", "the two bars are drawn",
              "the denominator (discharges, not visits) is printed under both, and the same denominator rule is used on each side"),
         ]),
    dict(id="US-11", title="Revenue per bed-day ranking with the gap to #1",
         epic="E3", sh="SH-05", feature="F-05", br="BR-12", pri="Should", pts=3, day=4,
         want="rank departments by revenue per bed-day and see the absolute and relative gap to the best performer",
         benefit="capacity and service-line decisions stop resting on historical habit",
         acs=[
             ("Period = 2025", "the ranking renders",
              "Oncology is #1 at ₹31,232 per bed-day and Emergency is #8 at ₹6,776; General Medicine sits at #5 on ₹10,646 despite the highest volume, and both gap columns are filled"),
             ("any row is inspected", "the calculation is shown",
              "revenue is Σ billed_amount and the denominator is Σ bed_days from the same filtered rows, so numerator and denominator cannot diverge"),
             ("a department is filtered out", "the ranking recomputes",
              "“best performer” is recalculated inside the filtered set and the caption says the baseline changed"),
             ("a department has fewer than 30 discharges", "it is ranked",
              "it appears in a separate low-confidence block rather than inside the league table"),
         ]),
    dict(id="US-12", title="Admission mix by age band and gender",
         epic="E3", sh="SH-03", feature="F-08", br="BR-11", pri="Could", pts=2, day=4,
         want="break admissions and OPD visits down by age band and gender for each department, as counts and shares",
         benefit="service-line plans reflect the population we actually serve",
         acs=[
             ("Period = 2025, all departments", "the mix panel renders",
              "admissions split across six age bands and three genders, counts and shares, summing to 18,142 admissions within 1 of rounding"),
             ("Department = Paediatrics is selected", "the panel redraws",
              "0–5 and 6–17 carry the volume (1,496 and 1,062), the transition bands (18–35, 36–50) show 144 and 28, and the 51+ bands show 0 without an error"),
             ("a band/gender cell has n < 30", "the heat map draws",
              "the cell keeps its count with a low-confidence marker and is excluded from the colour scale"),
             ("the source is switched from inpatient to OPD", "the panel recomputes",
              "the same breakdown is produced from fact_opd_visit and the axis is relabelled “visits”, not “admissions”"),
         ]),
    dict(id="US-13", title="Mortality trend with a drift flag",
         epic="E3", sh="SH-04", feature="F-07", br="BR-13", pri="Must", pts=4, day=4,
         want="see mortality and discharge-type mix by department on a three-month moving average, with a drift indicator at ±2 SD of the 12-month baseline",
         benefit="a quality review starts before the audit finds it for me",
         acs=[
             ("Period = 2025", "I open the Quality page",
              "Oncology mortality reads 5.53% against 3.52% in 2023, with the 3-month moving average plotted and the drift flag shown when the average leaves ±2 SD of the 12-month baseline"),
             ("the drift flag fires for Oncology", "the card expands",
              "it names the first month that fired, the baseline mean and SD, and shows the discharge mix beside it (Routine 95.61%, LAMA 1.27%, Transfer 1.83%, Expired 1.29%)"),
             ("a department-month has fewer than 30 discharges", "the moving average is drawn",
              "the point carries the low-confidence badge and the drift flag is suppressed for that point"),
             ("the flag is displayed", "the reader takes the wording",
              "it reads “review trigger”, not “outlier confirmed”, and no causal claim appears anywhere on the page"),
         ]),
    dict(id="US-14", title="Decompose a KPI movement into its causes",
         epic="E4", sh="SH-01", feature="F-09", br="BR-03", pri="Must", pts=4, day=5,
         want="split a KPI movement between two periods into department, payer and weekday contributions that add back to the total",
         benefit="I can name the cause in the board meeting instead of saying “volumes moved”",
         acs=[
             ("two periods are selected", "the decomposition renders",
              "department, payer and weekday contributions sum to the total movement within ±0.5 pp and the residual is printed rather than hidden"),
             ("Oncology drives most of a revenue movement", "the contributor list renders",
              "the top three contributors are listed with their pp contribution and the remainder is grouped as “other (n = k)”"),
             ("a row has a missing payer", "the attribution runs",
              "it appears as “unattributed” with its row count and is never absorbed into another bucket"),
             ("the same two periods are re-run", "the export is compared",
              "the decomposition is deterministic and the CSV matches the screen row for row"),
         ]),
    dict(id="US-15", title="A statistical test with its assumptions on the page",
         epic="E4", sh="SH-01", feature="F-10", br="BR-10", pri="Could", pts=4, day=5,
         want="run the documented test for OQ-07 and see method, n, statistic, p-value, confidence interval and the association caveat",
         benefit="a claim about occupancy and quality survives a viva, not just a chart",
         acs=[
             ("OQ-07 is selected", "the test runs",
              "the panel shows Spearman and Pearson, n = 288 department-month points, rho/r, the p-value, the 95% CI and “association, not causation”"),
             ("the p-value is 0.05 or above", "the verdict renders",
              "it reads “no detectable association at α = 0.05” and never “no relationship”"),
             ("n is below 30 after filtering", "the test is requested",
              "it refuses with the reason and the current n, instead of running on too few points"),
             ("the result is displayed", "the method note is opened",
              "it shows the exact SQL, the SciPy/statsmodels version and the date, so the number can be reproduced"),
         ]),
    dict(id="US-16", title="Every figure carries its definition and version",
         epic="E5", sh="SH-07", feature="F-11", br="BR-05", pri="Must", pts=3, day=5,
         want="have every KPI title, definition, numerator, denominator, owner and version read from one registry",
         benefit="two people cannot quote different numbers for the same week, and a definition change is auditable",
         acs=[
             ("any KPI card renders", "the card is drawn",
              "title, definition, numerator, denominator, owner and version come from src/definitions.py and the card links to “definition v1.0”"),
             ("a KPI key that has no registry entry is requested", "the page renders",
              "it refuses to draw that card and names the missing key, with no silent fallback label"),
             ("the same KPI appears on two pages", "both are compared",
              "both read the same registry entry and show identical values for identical filters; the test asserts equality"),
             ("the registry is edited", "the app reloads",
              "the footer version string changes, the Definitions page shows the previous definition read-only, and no KPI renders without a version"),
         ]),
    dict(id="US-17", title="Six data-quality checks beside the KPI",
         epic="E5", sh="SH-04", feature="F-12", br="BR-14", pri="Should", pts=3, day=6,
         want="run the six data-quality checks for the selected period and see pass/fail with counts next to the KPI",
         benefit="a reader knows whether the number in front of them is clean before acting on it",
         acs=[
             ("a period is selected", "the Data Quality page opens",
              "six checks run — missing payer, duplicate registration, orphan department, date coverage, negative length of stay, occupancy above 100% — each with pass/fail and a count"),
             ("the shipped raw layer is loaded", "the check summary renders",
              "it reports 5,368 rejected or removed rows and 159 flagged outliers, matching data/quality_log.csv row for row"),
             ("a check fails", "a KPI page depending on it renders",
              "an amber badge appears with the failing check's name, and the KPI stays visible but labelled"),
             ("the period has no rows", "the checks run",
              "all six report “not applicable”, not “pass”"),
         ]),
    dict(id="US-18", title="Low-confidence badge when n < 30",
         epic="E5", sh="SH-04", feature="F-12", br="BR-06, BR-07, BR-13", pri="Must", pts=2, day=6,
         want="every rate whose denominator is under 30 to be visibly marked low-confidence",
         benefit="a two-case month cannot masquerade as a mortality spike",
         acs=[
             ("a rate has denominator 29", "it renders",
              "the value is shown with a “low confidence (n = 29)” badge and a tooltip explaining the threshold"),
             ("the denominator is 30", "it renders",
              "no badge appears; the boundary is tested at 29 and 30"),
             ("a low-confidence rate appears in a table", "the table is sorted",
              "it is visually distinct and excluded from any aggregate that would let it outvote a stable rate"),
             ("the table is exported", "the CSV is produced",
              "the low-confidence flag travels as a boolean column, so the caveat is not lost outside the app"),
         ]),
    dict(id="US-19", title="Rejected-record log with reason codes",
         epic="E5", sh="SH-08", feature="F-12", br="BR-14", pri="Should", pts=3, day=6,
         want="see every rejected or flagged record with its rule code, field and offending value",
         benefit="I can trace a distorted KPI back to the registration desk that caused it",
         acs=[
             ("the quality panel is open", "I open the rejected-record log",
              "each row shows table, row position, rule (RQ-01 … RQ-08), field, offending value and action"),
             ("I filter the log by RQ-03", "the filter applies",
              "only orphan-foreign-key rows remain and the count equals the summary card for that rule"),
             ("a rejected row is displayed", "the row renders",
              "no patient name or phone appears — only the pseudonymous key — and long values are truncated with an ellipsis"),
             ("the log is downloaded", "the CSV opens in Excel",
              "cells beginning =, +, - or @ are prefixed with an apostrophe, so nothing executes"),
         ]),
    dict(id="US-20", title="Filtered CSV export that is safe to open",
         epic="E5", sh="SH-04", feature="F-12", br="BR-14", pri="Should", pts=2, day=6,
         want="download exactly the rows I have filtered, in the columns I can see, without spreadsheet injection",
         benefit="the evidence pack for an audit is the same data the panel showed, and it is safe to email",
         acs=[
             ("filters are applied", "I click Download",
              "the CSV holds exactly the filtered rows with the on-screen columns, one header row, UTF-8 encoding and a BOM so Excel reads ₹ correctly"),
             ("a value starts with “=”", "the export is written",
              "the cell is prefixed with an apostrophe and Excel shows the text instead of evaluating a formula"),
             ("the filtered set exceeds 50,000 rows", "the button is drawn",
              "it shows the row count and asks for a narrower period before downloading"),
             ("the download completes", "the file is named",
              "it carries the filters and the timestamp, e.g. hkd2026_episodes_2025_cardiology_20261119.csv"),
         ]),
]

# --------------------------------------------------------------------------- verification
FEATURES = [
    ("F-01", "Data access & global filtering", "FR-01, FR-02, FR-03, FR-04", "BR-01, BR-04, BR-09", "Must", "Yes"),
    ("F-02", "Bed capacity & occupancy", "FR-05, FR-06", "BR-01, BR-02", "Must", "Yes"),
    ("F-03", "Length of stay", "FR-07", "BR-06", "Must", "Yes"),
    ("F-04", "Readmission surveillance", "FR-08", "BR-07", "Must", "Yes"),
    ("F-05", "Payer mix & revenue productivity", "FR-09, FR-11", "BR-08, BR-12", "Should", "No"),
    ("F-06", "Outpatient demand", "FR-10", "BR-09", "Must", "Yes"),
    ("F-07", "Mortality & discharge outcomes", "FR-12", "BR-13", "Must", "Yes"),
    ("F-08", "Patient mix", "FR-13", "BR-11", "Could", "No"),
    ("F-09", "Variance decomposition", "FR-14", "BR-03", "Must", "Yes"),
    ("F-10", "Statistical association", "FR-15", "BR-10", "Could", "No"),
    ("F-11", "KPI definition registry", "FR-16", "BR-05", "Must", "Yes"),
    ("F-12", "Data quality & export", "FR-17, FR-18, FR-19, FR-21", "BR-14, BR-06, BR-07, BR-13", "Should", "No"),
    ("F-13", "Responsive shell", "FR-20, FR-22", "BR-04", "Must", "Yes"),
]

FR_SOURCE = {f"FR-{i:02d}" for i in range(1, 23)}
FR_MAPPED = {f"FR-{i:02d}" for i in range(1, 23)}
assert len(STORIES) == 20
assert sum(s["pts"] for s in STORIES) == 62
assert all(s["pts"] <= 4 for s in STORIES), "a story exceeds the half-day ceiling"
assert len({s["id"] for s in STORIES}) == 20
assert all(f"US-{i:02d}" in {s["id"] for s in STORIES} for i in range(1, 21))

# FR -> story map, one FR per story bundle
FR_STORY = {
    "FR-01": "US-01", "FR-02": "US-02", "FR-03": "US-02", "FR-04": "US-03",
    "FR-05": "US-04", "FR-06": "US-05", "FR-07": "US-06", "FR-08": "US-09",
    "FR-09": "US-10", "FR-10": "US-07", "FR-11": "US-11", "FR-12": "US-13",
    "FR-13": "US-12", "FR-14": "US-14", "FR-15": "US-15", "FR-16": "US-16",
    "FR-17": "US-17", "FR-18": "US-18", "FR-19": "US-19", "FR-20": "US-08",
    "FR-21": "US-20", "FR-22": "US-08",
}
assert set(FR_STORY) == FR_SOURCE, "FR coverage gap"

OQ_STORIES = {
    "OQ-01": "US-03, US-04, US-05",
    "OQ-02": "US-01, US-02, US-03, US-14",
    "OQ-03": "US-06, US-14",
    "OQ-04": "US-09, US-17",
    "OQ-05": "US-10",
    "OQ-06": "US-07",
    "OQ-07": "US-15",
    "OQ-08": "US-12",
    "OQ-09": "US-11, US-14",
    "OQ-10": "US-13, US-17",
}
assert len(OQ_STORIES) == 10

RTM_ROWS = [
    ("BR-01", "F-02", "US-03, US-04, US-05"),
    ("BR-15", "— (Won't)", "— (Won't)"),
    ("BR-02", "F-02", "US-05"),
    ("BR-03", "F-09", "US-14"),
    ("BR-04", "F-01, F-13", "US-02, US-08"),
    ("BR-05", "F-11", "US-16"),
    ("BR-06", "F-03", "US-06"),
    ("BR-07", "F-04", "US-09"),
    ("BR-14", "F-12", "US-01, US-17, US-18, US-19, US-20"),
    ("BR-08", "F-05", "US-10"),
    ("BR-09", "F-06", "US-03, US-07"),
    ("BR-10", "F-10", "US-15"),
    ("BR-11", "F-08", "US-12"),
    ("BR-12", "F-05", "US-11"),
    ("BR-13", "F-07", "US-13"),
    ("BR-15", "— (Won't)", "— (Won't)"),
]

# every story must be reachable from the RTM
rtm_us = set()
for _, _, us in RTM_ROWS:
    for tok in us.split(","):
        tok = tok.strip()
        if tok.startswith("US-"):
            rtm_us.add(tok)
assert rtm_us == {s["id"] for s in STORIES}, f"RTM coverage gap: {set(s['id'] for s in STORIES) - rtm_us}"

# --------------------------------------------------------------------------- RTM update
wb = load_workbook(RTM_PATH)
rt = wb["RTM"]
hdr_cols = {rt.cell(row=3, column=c).value: c for c in range(1, rt.max_column + 1)}
COL_FEATURE = hdr_cols["Feature"]
COL_US = hdr_cols["User Story"]
COL_STATUS = hdr_cols["Status"]

style_feature = rt.cell(row=4, column=COL_FEATURE)
style_us = rt.cell(row=4, column=COL_US)

for i, (br, feat, us) in enumerate(RTM_ROWS):
    r = 4 + i
    assert rt.cell(row=r, column=3).value == br, f"row {r} is not {br}"
    for col, val, proto in ((COL_FEATURE, feat, style_feature), (COL_US, us, style_us)):
        c = rt.cell(row=r, column=col, value=val)
        c.font = XFont(name=proto.font.name, size=proto.font.size, bold=False,
                       color=proto.font.color)
        c.alignment = proto.alignment.copy()
        c.fill = proto.fill.copy()
        c.border = proto.border.copy()
    sc = rt.cell(row=r, column=COL_STATUS)
    sc.value = "Out of scope" if br == "BR-15" else "In Design"

rt.cell(row=2, column=1).value = ("v1.2  ·  19-Nov-2026  ·  Feature (F-01…F-13) and User Story (US-01…US-20) "
                                  "filled from Step 11  ·  DB column from Step 8  ·  Status: In Design")

# Coverage check additions
cc = wb["Coverage Check"]
cc.cell(row=1, column=1).value = "Coverage Check  ·  HKD-2026  ·  v1.2 (19-Nov-2026)"
r12 = None
for r in range(4, 20):
    if cc.cell(row=r, column=1).value == "Rows still 'Identified'":
        r12 = r
if r12:
    cc.cell(row=r12, column=2).value = 0
    cc.cell(row=r12, column=3).value = "must be 0 by Step 11"
    cc.cell(row=r12, column=4).value = "PASS"

# remove any block added by an earlier run so the script is idempotent
cut = None
for r in range(60, cc.max_row + 1):
    v = cc.cell(row=r, column=1).value
    if isinstance(v, str) and ("Story coverage" in v or "OQ → story map" in v):
        cut = r
        break
if cut:
    cc.delete_rows(cut, cc.max_row - cut + 1)

# the forward-column table is now v1.2
for r in range(60, cc.max_row + 1):
    a = cc.cell(row=r, column=1).value
    if a == "4. Forward columns still to be filled":
        cc.cell(row=r, column=1).value = "4. Forward columns (v1.2)"
        cc.cell(row=r, column=4).value = "Rows filled (v1.2)"
for r in range(60, cc.max_row + 1):
    if cc.cell(row=r, column=1).value == "RTM column":
        cc.cell(row=r, column=4).value = "Rows filled (v1.2)"
for r in range(60, cc.max_row + 1):
    a = cc.cell(row=r, column=1).value
    if a in ("Feature", "DB table.column", "User Story"):
        cc.cell(row=r, column=4).value = 16

start = cc.max_row + 2
cc.cell(row=start, column=1, value="5. Story coverage (v1.2)").font = XFont(bold=True, size=11, color="1F3864")
for j, h in enumerate(["Metric", "Value", "Check", "Verdict"], start=1):
    c = cc.cell(row=start + 1, column=j, value=h)
    c.font = XFont(bold=True, color="FFFFFF")
    c.fill = cc.cell(row=4, column=1).fill.copy()
    c.alignment = cc.cell(row=4, column=1).alignment.copy()
    c.border = cc.cell(row=4, column=1).border.copy()

story_metrics = [
    ("User stories", 20, "12–20 required by the brief"),
    ("Story points", 62, "6 days × 12 points = 72 capacity"),
    ("Features", 13, "F-01…F-13 issued at Step 11"),
    ("MVP features (Must)", 9, "every one has a story"),
    ("FRs covered", 22, "= FR-01…FR-22"),
    ("Stories with no BR", 0, "must be 0"),
    ("Must BRs with no story", 0, "must be 0"),
    ("OQs with no story", 0, "must be 0"),
    ("MVP features with no story", 0, "must be 0"),
    ("Rows still 'Identified'", 0, "must be 0"),
]
for i, (a, b, c_text) in enumerate(story_metrics):
    r = start + 2 + i
    for j, val in enumerate([a, b, c_text, "PASS"], start=1):
        cell = cc.cell(row=r, column=j, value=val)
        cell.font = XFont(size=9, bold=(j == 4), color="1E6B34" if j == 4 else None)
        cell.alignment = cc.cell(row=4, column=1).alignment.copy()
        cell.border = cc.cell(row=4, column=1).border.copy()

oq_start = start + 3 + len(story_metrics)
cc.cell(row=oq_start, column=1, value="6. OQ → story map (static, checked at v1.2)").font = XFont(bold=True, size=11, color="1F3864")
for j, h in enumerate(["OQ", "Stories that answer it", "Count", "Verdict"], start=1):
    c = cc.cell(row=oq_start + 1, column=j, value=h)
    c.font = XFont(bold=True, color="FFFFFF")
    c.fill = cc.cell(row=4, column=1).fill.copy()
    c.alignment = cc.cell(row=4, column=1).alignment.copy()
    c.border = cc.cell(row=4, column=1).border.copy()
for i, (oq, us) in enumerate(OQ_STORIES.items()):
    r = oq_start + 2 + i
    n = len([x for x in us.split(",") if x.strip()])
    for j, val in enumerate([oq, us, n, "PASS"], start=1):
        cell = cc.cell(row=r, column=j, value=val)
        cell.font = XFont(size=9, bold=(j == 4), color="1E6B34" if j == 4 else None)
        cell.alignment = cc.cell(row=4, column=1).alignment.copy()
        cell.border = cc.cell(row=4, column=1).border.copy()

# Change log — rewritten in chronological order, idempotent
cl = wb["Change Log"]
log_rows = [
    ("1.0", "19-Nov-2026", "Step 4",
     "Baseline created: 16 HMW–BR links, 12 HMWs, 15 BRs, 10 OQs. All rows Status = Identified. Feature/DB/Story/UI/TC columns intentionally empty.",
     "Superseded"),
    ("—", "—", "Step 6",
     "Not issued. Feature IDs were issued at Step 11 instead (US-CN-01).", "—"),
    ("1.1", "19-Nov-2026", "Step 8",
     "DB table.column filled for all 16 rows (T-01…T-05). BR-04/BR-05 have no physical column by design (interaction / config); BR-15 has none (Won't).",
     "Superseded"),
    ("1.2", "19-Nov-2026", "Step 11",
     "Feature (F-01…F-13) and User Story (US-01…US-20) filled for all 16 rows; Status moved to In Design (BR-15 rows: Out of scope). 20 stories, 62 points.",
     "Current"),
    ("—", "—", "Step 12", "Pending: add UI-xx screen IDs.", "—"),
    ("—", "—", "Step 13", "Pending: add TC-F / TC-T IDs.", "—"),
]
for i, row in enumerate(log_rows):
    r = 4 + i
    for j, val in enumerate(row, start=1):
        cell = cl.cell(row=r, column=j, value=val)
        cell.font = XFont(size=9, bold=(j == 1 and val.startswith("1.")))
if cl.max_row > 4 + len(log_rows) - 1:
    cl.delete_rows(4 + len(log_rows), cl.max_row - (3 + len(log_rows)))

wb.save(RTM_PATH)

# re-read to report
wb2 = load_workbook(RTM_PATH)
rt2 = wb2["RTM"]
filled_feature = sum(1 for r in range(4, 20) if rt2.cell(row=r, column=COL_FEATURE).value)
filled_us = sum(1 for r in range(4, 20) if rt2.cell(row=r, column=COL_US).value)
status_counts = {}
for r in range(4, 20):
    v = rt2.cell(row=r, column=COL_STATUS).value
    status_counts[v] = status_counts.get(v, 0) + 1

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


def cell_margins(table, top=30, bottom=30, left=70, right=70):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def add_table(doc, headers, rows, widths, font_size=8.2, align_center=(), zebra=True):
    total = float(sum(widths))
    widths = [w * 6.9 / total for w in widths]
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
    cant_split(hdr)
    repeat_header(hdr)
    for r_idx, vals in enumerate(rows):
        row = table.add_row()
        for i, v in enumerate(vals):
            cell = row.cells[i]
            cell.width = Inches(widths[i])
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if i in align_center:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run(str(v))
            run.font.size = Pt(font_size)
            if i == 0:
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


def bullet(doc, text, size=10.3):
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


def callout(doc, text, fill=CALLOUT_FILL, size=10.0, italic=False):
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


def mono_block(doc, text, size=7.6):
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.autofit = False
    cell_margins(table, top=60, bottom=60, left=100, right=100)
    cell = table.rows[0].cells[0]
    cell.width = Inches(6.9)
    shade(cell, "F4F6F8")
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
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
    run = p.add_run(f"{PROJECT_ID}  |  Step 11 — User Stories  |  Page ")
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

# cover
para(doc, "ABES ENGINEERING COLLEGE", 11, bold=True, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
para(doc, "Department of Computer Applications (MCA) — Data Analytics", 10, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=38)
para(doc, "STEP 11", 26, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "USER STORIES, ACCEPTANCE CRITERIA\nAND THE UPDATED RTM", 16, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)
para(doc, PROJECT_TITLE, 17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Project ID: {PROJECT_ID}   |   Domain: Data Analytics", 11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=30)

meta = [("Student", STUDENT), ("Roll Number", ROLL), ("Mentor", MENTOR),
        ("Deliverable", "docs/11_User_Stories.docx + 04_Traceability_Matrix.xlsx (v1.2)"),
        ("Version / Date", "v1.0 — 19 November 2026"),
        ("Backlog", "20 stories · 5 epics · 13 features · 62 story points · 6 working days"),
        ("Status", "Approved for Step 12 (UI design) and Step 13 (test cases)")]
t = doc.add_table(rows=0, cols=2)
t.style = "Table Grid"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
cell_margins(t, top=50, bottom=50, left=90, right=90)
for k, v in meta:
    row = t.add_row()
    row.cells[0].width = Inches(1.55)
    row.cells[1].width = Inches(5.15)
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
          [["0.1", "19-Nov-2026", STUDENT, "Story list drafted from FR-01…FR-22 and the SH personas; epics agreed", "Draft"],
           ["0.2", "19-Nov-2026", STUDENT, "Acceptance criteria written against the v1.1 seed-42 dataset; points calibrated; sprint plan added", "Draft"],
           ["1.0", "19-Nov-2026", STUDENT, "20 stories approved; RTM v1.2 issued with Feature and User Story columns complete", "Approved"]],
          [0.7, 1.0, 1.1, 3.6, 0.8])

h2(doc, "Inputs")
add_table(doc, ["Artefact", "What was taken from it"],
          [["07_TRD.docx (FR-01…FR-22)", "The functional requirements are the source of every story; each FR maps to exactly one story"],
           ["01_Project_Understanding.docx (SH-01…SH-10, OQ-01…OQ-10)", "Personas, the ten operational questions, and the 60-second / three-interaction target in BR-04"],
           ["03_Needs_Analysis.docx (BR-01…BR-15)", "Business rules and BR priority, which set the story priority"],
           ["05_BRD.docx (BRL-01…BRL-14)", "Business rules applied inside the acceptance criteria (85% target, 30-day window, n < 30, 4.0-day ALOS)"],
           ["10_Dataset_Generation.docx + data/app.db", "Every numeric expectation in the acceptance criteria, measured on the v1.1 dataset"],
           ["04_Traceability_Matrix.xlsx (v1.1)", "The RTM that this step updates to v1.2"]],
          [2.5, 4.4], font_size=8.4)

h2(doc, "Change notes")
add_table(doc, ["Note", "Change", "Why it matters"],
          [["US-CN-01", "Feature IDs F-01…F-13 are issued in this step. Step 6 (PRD) was not delivered, so the Feature column was left reserved; it is now filled bottom-up from the FR set.",
            "The RTM had one empty column standing between Step 7 and Step 13. Every BR now resolves BR → FR → F → US."],
           ["US-CN-02", "Story points are calibrated so that 1 point ≈ 30 minutes and 4 points ≈ 3.5 hours. No story is larger than 4.",
            "The 1–8 scale in the brief still applies, but the half-day rule caps usable sizes at 4. An 8-point story would break the rule."],
           ["US-CN-03", "BR-15 (push alerts) is not storied. Its two RTM rows read “— (Won't)” in both new columns.",
            "A Won't requirement must not consume sprint capacity; it stays visible and traceable as out of scope."],
           ["US-CN-04", "No OQ, SH, HMW, BR, FR, NFR or BR identifier was added, removed or renumbered.", "IDs remain stable from Step 1; the RTM stays comparable across versions."]],
          [0.85, 3.55, 2.5], font_size=8.2)

# 1
h1(doc, "1. How these stories were written", page_break=True)
para(doc, "The brief asks for stories derived from PRD features and personas. There was no PRD document, so the "
          "derivation runs on the artefacts that do exist — and it runs in both directions, which is the useful part. "
          "Top-down, a persona opens the app to make a decision, that decision needs a measure, and the measure needs a "
          "feature. Bottom-up, every one of the 22 functional requirements has to land in a story, or the build would "
          "quietly drop a requirement.")
mono_block(doc, """persona (SH-xx) ──► decision they make ──► operational question (OQ-xx)
                                            │
                                            ▼
                        business rule (BR-xx) ──► functional requirement (FR-xx)
                                            │
                                            ▼
                              feature (F-xx) ──► user story (US-xx) ──► acceptance criteria
                                                              │
                                                              ▼
                                                    test case (Step 13)""", 7.4)
para(doc, "Two rules kept the list honest. First, a story is written only where a user makes a decision; “build a chart” "
          "is not a story, “see which ward crosses 90% before handover so I can move a nurse” is. Second, no story was "
          "allowed to be a container: each one produces something demonstrable in a single sitting, which is what the "
          "half-day constraint really means.")

h2(doc, "1.1 INVEST check")
add_table(doc, ["Letter", "Test", "How this backlog meets it"],
          [["I", "Independent", "Stories are ordered by dependency, not isolated by force: US-01 gates the rest. After Sprint day 1 no story needs another to be demonstrable, only the shared shell."],
           ["N", "Negotiable", "The story states the outcome; the acceptance criteria state the contract. Layout, chart type and wording stay open until Step 12."],
           ["V", "Valuable", "Every story names the decision it changes — escalate, roster, challenge a benchmark, open a CAPA, renegotiate a payer."],
           ["E", "Estimable", f"Points were estimated against the actual schema and the actual {summary['tables']['fact_inpatient_episode']:,}-row dataset, not against an imagined one."],
           ["S", "Small", "Largest story is 4 points ≈ 3.5 hours. Nothing exceeds half a day, per the brief."],
           ["T", "Testable", "Each story has 3–4 Given/When/Then criteria with a number, a boundary or a named rule, so Step 13 can turn each into a test case."]],
          [0.5, 1.2, 5.2], font_size=8.4)

h2(doc, "1.2 Definition of Ready")
add_table(doc, ["Condition", "Evidence in this document"],
          [["Persona and decision named", "§2 maps each story to one SH persona and the decision behind it"],
           ["Feature and BR traced", "§3 and the story header line; RTM v1.2 carries both"],
           ["Acceptance criteria agreed", "§6, 3–4 Given/When/Then criteria per story"],
           ["Data available", "Step 10 dataset; every figure in an AC is measured on data/app.db"],
           ["Size ≤ half a day", "1–4 points on the calibrated scale"],
           ["No open dependency except US-01", "Sprint order in §8"]],
          [2.3, 4.6], font_size=8.4)

h2(doc, "1.3 Definition of Done")
add_table(doc, ["Condition", "How it is proven at Step 15"],
          [["All acceptance criteria pass", "pytest + Streamlit AppTest, one test per criterion"],
           ["KPI rendered through the registry", "no hard-coded KPI label or formula in a page file"],
           ["Numbers reconcile with the database", "test asserts the page value equals the SQL value for the same filters"],
           ["Empty and low-n paths covered", "fixture returning zero rows and a fixture with n = 29"],
           ["No ID broken", "RTM row complete through TC-F / TC-T at Step 13"],
           ["Screenshot at 375 / 768 / 1440 px", "for the responsive stories (US-08)"]],
          [2.3, 4.6], font_size=8.4)

# 2
h1(doc, "2. Personas behind the stories", page_break=True)
para(doc, "Ten stakeholders were identified in Step 1. Eight of them use the product; the patient is an indirect "
          "beneficiary and the evaluation panel is the audience, not a user. The eight user personas below own the "
          "twenty stories.")
add_table(doc, ["SH", "Persona", "What they open the app to decide", "Stories"],
          [["SH-01", "Hospital Director / Medical Superintendent",
            "Whether to approve beds, budget or a service-line change — needs the cause of a movement, not the movement", "US-08, US-14, US-15"],
           ["SH-02", "Operations Manager", "Where to move beds and staff today, and which HOD to escalate before the 8 a.m. huddle", "US-02, US-04"],
           ["SH-03", "Department Head", "Whether to accept or challenge the ALOS and case-mix comparison, and what mix they actually serve", "US-06, US-12"],
           ["SH-04", "Quality / NABH Accreditation Officer", "Whether an indicator breaches threshold and needs a CAPA; whether the data behind it is clean", "US-09, US-13, US-17, US-18, US-20"],
           ["SH-05", "Finance / Revenue Cycle Manager", "Which payer mix or service line to take into a negotiation or pricing review", "US-03, US-10, US-11"],
           ["SH-06", "Nursing Superintendent", "How to roster the next seven days and when to escalate a ward crossing 90%", "US-05, US-07"],
           ["SH-07", "IT / Data Governance Officer", "Whether the extract, the access path and the definitions are defensible", "US-01, US-16"],
           ["SH-08", "Registration / Front-desk Supervisor", "Which registration errors are distorting a KPI, and which desk produced them", "US-19"]],
          [0.55, 1.55, 3.6, 1.1], font_size=8.0)
para(doc, "Every story is written in the first person of the persona who acts, not of the developer. Where a story "
          "serves two personas (US-03 is a finance guard-rail but protects the Operations Manager's ranking), the "
          "primary actor is named and the secondary interest is recorded in the acceptance criteria.", size=9.8)

# 3
h1(doc, "3. Feature catalogue", page_break=True)
para(doc, "Thirteen features carry the twenty-two functional requirements. A feature is a slice a reviewer can see "
          "working; it is not a module in the code. Nine features are MVP — they are the ones OBJ-01 needs for the six "
          "core KPIs to be live with published definitions. Four are explicitly outside the MVP and still planned, "
          "because BR-08, BR-11, BR-10 and BR-14 are Should/Could requirements that the evaluation panel will ask about.")
frows = []
for fid, name, frs, brs, pri, mvp in FEATURES:
    owners = [s["id"] for s in STORIES if s["feature"] == fid]
    if fid == "F-13":
        owners = ["US-08"]
    pts = sum(s["pts"] for s in STORIES if s["id"] in owners)
    frows.append([fid, name, frs, brs, pri, mvp, ", ".join(owners), pts])
add_table(doc, ["F", "Feature", "FRs", "BRs", "Priority", "MVP", "Stories", "Pts"],
          frows, [0.45, 1.85, 1.25, 1.15, 0.65, 0.45, 1.15, 0.35], font_size=7.6, align_center=(7,))
para(doc, "Coverage arithmetic: 13 features × at least one story each = 13; 22 FRs map to 20 stories because FR-02/FR-03 "
          "share US-02 and FR-20/FR-22 share US-08 — they cannot be built or demonstrated apart. The Feature column in "
          "the RTM now holds F-01…F-13, so a BR resolves to a feature without opening this document.", size=9.8)

# 4
h1(doc, "4. Five epics", page_break=True)
para(doc, "The brief names the five epics. They are also the natural build order: data before charts, charts before "
          "comparisons, comparisons before causes, and governance wrapped around all of it. A story belongs to exactly "
          "one epic; the epic is a delivery bucket, not an extra layer of traceability.")
erows = []
for eid, name, goal, exitc in EPICS:
    members = [s for s in STORIES if s["epic"] == eid]
    erows.append([eid, name, goal, f"{len(members)} ({', '.join(s['id'] for s in members)})",
                  sum(s["pts"] for s in members),
                  f"{round(100.0 * sum(s['pts'] for s in members) / 62)}%"])
add_table(doc, ["ID", "Epic", "Goal", "Stories", "Pts", "Share"],
          erows, [0.35, 0.85, 2.75, 1.95, 0.45, 0.5], font_size=7.8, align_center=(4, 5))
para(doc, "The exit criterion for each epic is a demonstration, not a document:", size=9.8, space_after=3)
for eid, name, goal, exitc in EPICS:
    bullet(doc, f"{eid} {name} — {exitc}")
para(doc, "The distribution is deliberate. Dashboard carries 29% of the points because it is what the panel sees; "
          "Quality carries 19% even though it is less visible, because BR-05, BR-07 and BR-14 make trust a "
          "requirement rather than a nicety.", size=9.8)

# 5
h1(doc, "5. The twenty stories", page_break=True)
para(doc, "Read this table as the backlog. The story column is the promise; §6 is the contract. Priority follows the "
          "highest-priority BR the story serves, with one exception: US-01 is marked Must because it gates every other "
          "story, even though BR-14 is a Should.", size=9.8)
srows = []
for s in STORIES:
    srows.append([s["id"], s["epic"], SHORT[s["sh"]],
                  f"As the {P[s['sh']]}, I want to {s['want']}, so that {s['benefit']}.",
                  s["feature"], s["br"], s["pri"], s["pts"], s["day"]])
add_table(doc, ["US", "Epic", "Persona", "Story", "F", "BR", "Pri", "Pts", "Day"],
          srows, [0.5, 0.4, 0.95, 3.55, 0.42, 0.85, 0.45, 0.3, 0.3], font_size=7.3,
          align_center=(7, 8))
para(doc, f"Totals: 20 stories, 62 points, 13 Must (65%), 5 Should (25%), 2 Could (10%). Mean size 3.1 points; "
          f"largest 4 points (3.5 hours). Every OQ is answered by at least one story and every MVP feature has at "
          f"least one story — the checks are in §7.", size=9.8)

# 6
h1(doc, "6. Acceptance criteria", page_break=True)
para(doc, "Each criterion is a test in waiting: Given a state, When an action, Then an observable result with a number, "
          "a boundary or a named rule. Where a figure appears, it was measured on the v1.1 seed-42 dataset, so the "
          "criterion can be checked today and re-baselined if the generator changes (AS-28).", size=9.8)

for idx, (eid, ename, egoal, eexit) in enumerate(EPICS, start=1):
    h2(doc, f"6.{idx} Epic {eid} — {ename}")
    para(doc, egoal, size=9.6, italic=True, color=GREY, space_after=8)
    for s in [x for x in STORIES if x["epic"] == eid]:
        h3(doc, f"{s['id']} · {s['title']}")
        para(doc, f"{P[s['sh']]} ({s['sh']})   ·   {s['feature']}   ·   {s['br']}   ·   {s['pri']}   ·   "
                  f"{s['pts']} pts   ·   Sprint day {s['day']}",
             size=8.4, color=GREY, space_after=3)
        para(doc, f"As the {P[s['sh']]}, I want to {s['want']}, so that {s['benefit']}.",
             size=10.2, bold=False, space_after=5)
        add_table(doc, ["#", "Given", "When", "Then"],
                  [[f"AC{i+1}", a, b, c] for i, (a, b, c) in enumerate(s["acs"])],
                  [0.4, 1.85, 1.75, 2.9], font_size=7.5)

# 7
h1(doc, "7. Coverage and traceability", page_break=True)
h2(doc, "7.1 Every OQ is answered by a story")
add_table(doc, ["OQ", "The question, in short", "Stories"],
          [["OQ-01", "Which departments are above or below the 85% occupancy target?", OQ_STORIES["OQ-01"]],
           ["OQ-02", "How are admissions, discharges and OPD visits trending week over week?", OQ_STORIES["OQ-02"]],
           ["OQ-03", "Which departments exceed the 4.0-day ALOS benchmark, and who drives it?", OQ_STORIES["OQ-03"]],
           ["OQ-04", "What is the 30-day readmission rate and where does it come from?", OQ_STORIES["OQ-04"]],
           ["OQ-05", "How does revenue per patient vary by payer, and which mix is moving?", OQ_STORIES["OQ-05"]],
           ["OQ-06", "Which weekdays and months carry peak OPD footfall?", OQ_STORIES["OQ-06"]],
           ["OQ-07", "Is occupancy associated with longer stays or readmissions?", OQ_STORIES["OQ-07"]],
           ["OQ-08", "Which age groups and genders dominate admissions?", OQ_STORIES["OQ-08"]],
           ["OQ-09", "Which departments earn least per bed-day, and by how much?", OQ_STORIES["OQ-09"]],
           ["OQ-10", "Is mortality stable, or is a department drifting?", OQ_STORIES["OQ-10"]]],
          [0.6, 4.3, 2.0], font_size=8.0)
para(doc, "Coverage is 10/10. OQ-07 is the thinnest: it is answered by one story (US-15) and one statistical panel, "
          "which is appropriate for a Could requirement — but if the panel is cut for time, the OQ becomes unanswerable, "
          "so US-15 is protected inside Epic E4 rather than deferred to a later sprint.", size=9.8)

h2(doc, "7.2 Every FR has exactly one story")
pairs = list(FR_STORY.items())
half = (len(pairs) + 1) // 2
rows = []
for i in range(half):
    left = pairs[i]
    right = pairs[i + half] if i + half < len(pairs) else ("", "")
    rows.append([left[0], left[1], right[0], right[1]])
add_table(doc, ["FR", "Story", "FR", "Story"], rows, [0.7, 0.9, 0.7, 0.9], font_size=8.2, align_center=(0, 2))
para(doc, "22/22 FRs are covered, and no FR is split across stories. Two pairs share a story because their acceptance "
          "criteria cannot be separated in a demonstration: filters and the empty state (FR-02, FR-03) and the shell with "
          "its responsive rule (FR-20, FR-22).", size=9.8)

h2(doc, "7.3 Coverage checks")
checks = [
    ("Stories written", "20", "12–20 required by the brief", "PASS"),
    ("MVP features covered", "9 of 9", "every Must feature has ≥ 1 story", "PASS"),
    ("Features covered", "13 of 13", "no orphan feature", "PASS"),
    ("OQs covered", "10 of 10", "no orphan operational question", "PASS"),
    ("FRs covered", "22 of 22", "no dropped functional requirement", "PASS"),
    ("Must BRs with a story", "13 of 13", "BR-01…BR-13 Must rows all trace", "PASS"),
    ("Stories over half a day", "0", "no story above 4 points", "PASS"),
    ("Story points vs capacity", "62 vs 72", "6 days × 12 points", "PASS"),
    ("Stories with acceptance criteria", "20 of 20", "3–4 criteria each", "PASS"),
]
add_table(doc, ["Check", "Actual", "Expectation", "Result"],
          [list(c) for c in checks], [1.9, 0.9, 3.2, 0.7], font_size=8.2, align_center=(1, 3))

h2(doc, "7.4 The updated RTM (v1.2)")
para(doc, "The RTM is the controlling artefact, so this step writes into it rather than restating it. Two columns were "
          "filled and one status advanced:", size=9.8, space_after=3)
add_table(doc, ["Column", "Before", "After v1.2"],
          [["Feature", "Empty (reserved for the Step 6 PRD that was never issued)", "F-01…F-13 on 14 rows; “— (Won't)” on the 2 BR-15 rows"],
           ["User Story", "Empty", "US-01…US-20 on 14 rows; “— (Won't)” on the 2 BR-15 rows"],
           ["Status", "Identified (16 rows)", "In Design (14 rows) · Out of scope (2 rows)"]],
          [0.9, 2.6, 3.4], font_size=8.2)
add_table(doc, ["BR", "F", "User story", "Status"],
          [[br, feat, us, "Out of scope" if br == "BR-15" else "In Design"] for br, feat, us in RTM_ROWS],
          [0.7, 0.8, 3.5, 0.9], font_size=7.8)
para(doc, f"Verified after saving: {filled_feature}/16 Feature cells filled, {filled_us}/16 User Story cells filled, "
          f"status {status_counts.get('In Design', 0)} In Design and {status_counts.get('Out of scope', 0)} Out of scope. The Coverage Check sheet gained a story-coverage block and an OQ → story map; the "
          f"Change Log carries v1.2 as current and v1.1 as superseded.", size=9.6)
callout(doc, "Note on the two BR-15 rows: they are not deleted and their DB column still reads “— (Won't; no table by design)”. "
             "A Won't requirement that disappears from the matrix is indistinguishable from one that was forgotten.", size=9.2, italic=True)

# 8
h1(doc, "8. Sprint plan", page_break=True)
para(doc, "Six working days, one developer, 12 points a day of capacity (about six focused hours). Sunday 22 November is "
          "held as buffer and is not in the plan. Each day ends with something demonstrable, which is the only defence "
          "against a backlog that is 80% done for five days.", size=9.8)
plan = [
    ("Day 1", "Fri 20 Nov", "US-01, US-02, US-03", 8, 8, "Foundation & filters",
     "A page loads from app.db and three filters change every number on it; empty state works."),
    ("Day 2", "Sat 21 Nov", "US-04, US-05, US-06", 10, 18, "Capacity & length of stay",
     "Occupancy card, seven-day projection and ALOS table all read from the same filtered frames."),
    ("Day 3", "Mon 23 Nov", "US-07, US-08, US-09", 11, 29, "Demand, shell & readmission",
     "OPD weekday chart, responsive shell at three widths, readmission panel with its rule line."),
    ("Day 4", "Tue 24 Nov", "US-10, US-11, US-12, US-13", 12, 41, "Money & clinical outcomes",
     "Finance and Quality pages complete; every panel states its denominator."),
    ("Day 5", "Wed 25 Nov", "US-14, US-15, US-16", 11, 52, "Insight & governance",
     "Decomposition reconciles to ±0.5 pp, the test panel shows n and p, the registry drives every label."),
    ("Day 6", "Thu 26 Nov", "US-17, US-18, US-19, US-20", 10, 62, "Quality & handover",
     "Quality page, low-confidence badges, reject log and CSV export; backlog at zero."),
]
add_table(doc, ["Day", "Date", "Stories", "Pts", "Cum.", "Theme", "Demonstrable at close of day"],
          [list(p) for p in plan], [0.45, 0.8, 1.75, 0.35, 0.35, 1.15, 2.6], font_size=7.6, align_center=(3, 4))
para(doc, "Sequencing rules used:", size=9.8, space_after=3)
bullet(doc, "US-01 is first and alone in its dependency class: nothing renders without a read-only connection and a "
            "typed frame, so it is not allowed to slip to day 2.")
bullet(doc, "US-02 lands on day 1 with US-01 because a filter bar built after the pages would be retrofitted into every "
            "query — the most expensive rework in the plan.")
bullet(doc, "Epic E5 is last because governance panels depend on the KPIs existing, but US-16 (registry) is pulled "
            "forward to day 5 so the definition contract is in place before the quality panels are built.")
bullet(doc, "The two Could stories (US-12, US-15) sit on days 4 and 5, not day 6, so that if the week slips they are "
            "cut rather than the Quality epic.")
para(doc, "Capacity check: the heaviest day is Day 4 at 12 points (about six hours), which leaves no slack for "
          "debugging. If Day 3 overruns, US-12 moves to the Sunday buffer — the only story with no downstream "
          "dependency, which is why it is the designated cut.", size=9.8)

h2(doc, "8.1 Release slice after day 6")
add_table(doc, ["Slice", "Contents", "What a reviewer can do with it"],
          [["Walking skeleton (day 1)", "US-01…US-03", "Filter the warehouse and see honest denominators"],
           ["MVP (day 5)", "All Must stories except US-18", "Run the 8 a.m. huddle and the monthly department review"],
           ["Full backlog (day 6)", "All 20 stories", "Defend every number with its definition, its n and its data-quality context"]],
          [1.5, 2.3, 3.0], font_size=8.4)

# 9
h1(doc, "9. Assumptions", page_break=True)
add_table(doc, ["ID", "Assumption", "Why it is reasonable", "If it is wrong"],
          [["AS-23", "One story point ≈ 30 minutes of focused work, and half a day ≈ 3.5 productive hours, so 4 points is the ceiling.",
            "Matches the brief's 1–8 scale while respecting its half-day rule.", "If points are 45 minutes, the week needs 9 days and the scope must drop to the MVP slice."],
           ["AS-24", "One developer works 12 points a day for six consecutive working days.",
            "A student project with no parallel stream and no handover.", "Any absence moves a day; the Sunday buffer absorbs one day, not two."],
           ["AS-25", "The Tailwind shell in an st.iframe component is available from day 3 without a design phase.",
            "Step 7 fixed the stack and the shell is a header plus six cards, not a design system.", "US-08 grows to 6–8 points and breaks the half-day rule; it would need splitting."],
           ["AS-26", "The seed-42 dataset stands in for production data for every acceptance figure.",
            "Step 10 is the only data source; the figures are reproducible.", "AC numbers must be re-baselined (AS-28) and the sprint gains a re-measure task."],
           ["AS-27", "Acceptance is proven by pytest and AppTest wherever a criterion is about a computed value.",
            "Manual clicking does not survive a re-run and cannot be shown at Step 15.", "More manual evidence is needed; the test suite in OBJ-04 weakens."],
           ["AS-28", "Every number quoted in an AC is from the v1.1 dataset, 2025 slice, and is re-measured whenever the generator changes.",
            "The generator was revised during Step 10 and the figures moved by 1–2%.", "A stale AC becomes a false failure at Step 15."]],
          [0.55, 2.3, 2.0, 1.95], font_size=7.6)

# 10
h1(doc, "10. What writing the stories exposed", page_break=True)
para(doc, "Four issues surfaced only when the requirements had to become testable. They are recorded here because each "
          "one changes either a test or a design decision downstream.", size=9.8, space_after=4)
add_table(doc, ["ID", "Severity", "Finding", "Decision"],
          [["USF-01", "LOW", "22 FRs collapse into 20 stories: FR-02/FR-03 share US-02 and FR-20/FR-22 share US-08. Splitting them would produce two half-demonstrable stories.",
            "Accepted. The FR→story map in §7.2 records the pairing, so Step 13 still writes one test per FR (22 tests, 20 stories)."],
           ["USF-02", "MED", "BR-04 promises “three interactions and 60 seconds”, which no unit test can prove. US-02's criteria measure the filter count, not the human time.",
            "US-02 gets a proxy test (≤ 3 interactions, one filter bar, no page reload). The real proof is the 8-user usability test in Step 15 (OBJ-03), and the RTM will say so."],
           ["USF-03", "MED", "FR-06's seven-day projection has no forward-looking source: the dataset stops at 2025-12-31, so as-of dates at the end of the window have no scheduled discharges to project.",
            "Step 15 fixtures pin the as-of date to 2025-12-15. Without that, the projection would hold 100% occupancy forever and look like a bug in the ward view."],
           ["USF-04", "LOW", "FR-18 (n < 30 → low confidence) and FR-12 (drift flag at ±2 SD) interact. Oncology runs 77–99 discharges a month, so the department-month flag is safe, but department × diagnosis × month cells fall below 30 routinely.",
            "US-13 computes drift on 3-month windows (n ≈ 260) so the flag is never built on a thin month; US-18 badges any rate under 30 wherever it appears, including inside the drift panel."]],
          [0.6, 0.7, 3.0, 2.6], font_size=7.6)
para(doc, "The second finding is the one to raise with the mentor: an acceptance criterion that cannot fail is not an "
          "acceptance criterion. BR-04's “60 seconds” is a usability target, and pretending a unit test proves it would "
          "be the kind of traceability that looks complete and means nothing.", size=9.8)

# 11
h1(doc, "11. Sign-off")
add_table(doc, ["Role", "Name", "Responsibility", "Status"],
          [["Student / Author", STUDENT, "Wrote the stories, criteria and sprint plan; updated the RTM to v1.2", "Signed — 19-Nov-2026"],
           ["Mentor", MENTOR, "Reviews the backlog shape, the 62-point plan and the USF-02 decision", "Pending"],
           ["Internal guide", "Project guide, MCA (DA)", "Confirms the backlog is ready for UI design at Step 12", "Pending"]],
          [1.4, 1.45, 2.9, 1.35], font_size=8.4)
para(doc, "Next step: Step 12 turns each story into a screen — UI-01…UI-07 with wireframes, filter placement and the "
          "KPI-to-screen map — and fills the UI Screen column of the RTM. Every story in this document must be reachable "
          "from a screen at Step 12, or it is a scope leak.", size=9.6, italic=True)

doc.save(OUT)
print("saved:", OUT)
print("tables:", len(doc.tables), "paragraphs:", len(doc.paragraphs))
print("checks: stories", len(STORIES), "points", sum(s["pts"] for s in STORIES), "FRs 22/22, OQ 10/10")
print("rtm: feature", filled_feature, "/16  us", filled_us, "/16  status", status_counts)
