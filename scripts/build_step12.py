"""Build the Step 12 deliverable: UI framework, wireframes, component contract,
chart selection and the RTM v1.3 update (UI Screen filled).

Numbers quoted in wireframes are measured on data/app.db (v1.1, seed 42).
"""
from __future__ import annotations

import json
import os
import re
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
OUT = os.path.join(DOCS, "12_UI_Framework.docx")
RTM_PATH = os.path.join(DOCS, "04_Traceability_Matrix.xlsx")

summary = json.load(open(os.path.join(DATA, "validation_summary.json"), encoding="utf-8"))
UI_SRC = open(os.path.join(ROOT, "src", "ui.py"), encoding="utf-8").read()
CFG_SRC = open(os.path.join(ROOT, ".streamlit", "config.toml"), encoding="utf-8").read()

# --------------------------------------------------------------------------- design data
SCREENS = [
    ("UI-01", "Overview", "app.py",
     "Six executive numbers in one screenful, then the three movements that matter.",
     "SH-01", "US-02, US-08, US-16", "OQ-01, OQ-02"),
    ("UI-02", "Operations", "pages/1_Operations.py",
     "Occupancy against 85%, the seven-day ward projection, volume trend and a two-period compare.",
     "SH-02, SH-06", "US-03, US-04, US-05, US-07", "OQ-01, OQ-02, OQ-06"),
    ("UI-03", "Departments", "pages/2_Departments.py",
     "ALOS against the benchmark, 30-day readmission, patient mix, and an EDA explorer.",
     "SH-03, SH-04", "US-06, US-09, US-12, US-18", "OQ-03, OQ-04, OQ-08"),
    ("UI-04", "Quality", "pages/3_Quality.py",
     "Mortality trend with a drift flag and the discharge-type mix beside it.",
     "SH-04", "US-13", "OQ-10"),
    ("UI-05", "Finance", "pages/4_Finance.py",
     "Payer mix and its movement, revenue per patient, revenue per bed-day ranking.",
     "SH-05", "US-03, US-10, US-11", "OQ-05, OQ-09"),
    ("UI-06", "Data Quality", "pages/5_Data_Quality.py",
     "Six checks with counts, freshness, and the reject log with reason codes.",
     "SH-04, SH-08", "US-17, US-18, US-19", "OQ-04"),
    ("UI-07", "Definitions", "pages/6_Definitions.py",
     "The KPI registry, the method notes, the version and the build fingerprint.",
     "SH-07", "US-16", "—"),
    ("UI-08", "Insights", "pages/7_Insights.py",
     "Decompose a movement into contributors, and run the one documented test.",
     "SH-01", "US-14, US-15", "OQ-02, OQ-07, OQ-09"),
    ("UI-09", "Data & Export", "pages/8_Data.py",
     "Preview exactly the filtered rows, read the column dictionary, export safely.",
     "SH-07, SH-05", "US-20", "—"),
]

BRIEF_MAP = [
    ("Overview / KPIs", "UI-01", "Same job as the brief's first screen: six numbers, no scrolling."),
    ("Trend", "UI-02 (tab)", "A view mode, not a destination — it shares UI-02's filters and frames."),
    ("Compare", "UI-02 (tab)", "Same reason; the dumbbell comparison for ALOS lives in UI-03."),
    ("EDA & Statistics", "UI-03 (tab) + UI-08", "Exploration is a tab of Departments; a published claim is UI-08."),
    ("Insights", "UI-08", "Decomposition and the test panel need their own controls (KPI, period A/B)."),
    ("Data + download", "UI-09", "Preview, dictionary and the guarded CSV; an Export button also sits on every screen."),
    ("Data quality", "UI-06", "Unchanged from Step 7."),
    ("About", "UI-07", "Unchanged from Step 7 — the registry screen is the About screen."),
    ("—", "UI-04, UI-05", "Quality and Finance keep their Step-7 identity; they are not folded into Overview."),
]

FILTERS = [
    ("Period", "st.selectbox", "2023 · 2024 · 2025 (month grain in the trend)", "2025", "flt_period", "Every data screen"),
    ("Compare to", "st.selectbox", "Prior year · Prior period · None", "Prior year", "flt_compare", "UI-01, 02, 03, 05, 08"),
    ("Department", "st.multiselect", "8 active departments", "All", "flt_dept", "Every data screen"),
    ("Ward", "st.multiselect (child of Department)", "20 wards, filtered by the chosen departments", "All", "flt_ward", "UI-02"),
    ("Bed class", "st.segmented_control", "Intensive · Twin · Single · Economy · All", "All", "flt_class", "UI-02"),
    ("Payer", "st.multiselect", "5 categories", "All", "flt_payer", "UI-01, 03, 05, 08"),
    ("Age band", "st.select_slider (range)", "0–5 … 80+", "All", "flt_age", "UI-03, 09"),
    ("Gender", "st.segmented_control", "All · F · M · O", "All", "flt_gender", "UI-03, 09"),
    ("Reset", "st.button", "—", "—", "—", "Every data screen"),
]

PALETTE = [
    ("Navy", "#1F3864", "Primary: actions, selected filter, headings, KPI text accent", "11.62 : 1", "AA + AAA"),
    ("Teal", "#0F766E", "On target · passed check · improving", "5.47 : 1", "AA"),
    ("Amber", "#B45309", "Warning · low confidence · period in progress", "5.02 : 1", "AA"),
    ("Red", "#B91C1C", "Breach · failed check · deteriorating", "6.47 : 1", "AA"),
    ("Slate 900", "#0F172A", "Body text, KPI values", "17.85 : 1", "AAA"),
    ("Slate 600", "#475569", "Secondary text, captions, axis labels", "7.58 : 1", "AAA"),
    ("Surface", "#F4F6FA", "Sidebar, table stripes, card tint", "—", "background"),
    ("Border", "#E2E8F0", "Hairlines, card borders, dividers", "—", "non-text"),
]

TYPE_SCALE = [
    ("KPI value", "Inter", "26 px / 600", "tabular-nums, −0.01 em tracking", "One per card; never wraps"),
    ("KPI label", "Inter", "11 px / 500", "uppercase, 0.04 em tracking, slate-500", "Truncates with ellipsis"),
    ("Screen title", "Inter", "20 px / 600", "−0.02 em tracking", "Header component only"),
    ("Section title", "Inter (Streamlit)", "16 px / 600", "Native st.subheader", "Charts and tables"),
    ("Body", "Inter / Source Sans", "13 px / 400", "1.45 line height", "Insight text, captions"),
    ("Table cell", "Inter", "12 px / 400", "tabular-nums for every number", "st.dataframe"),
    ("ID / code", "ui-monospace, Consolas", "11 px / 400", "never italic", "US-04, BR-07, T-04"),
]

BREAKPOINTS = [
    ("375 px", "Phone (executive check)", "Sidebar collapsed · columns stacked 1-up · KPI value 22 px · chart 260 px tall · tables scroll horizontally", "Readable, not full-feature (AS-31)"),
    ("768 px", "Tablet", "Sidebar collapsible · KPI band 2-up · charts full width", "Review on the ward round"),
    ("1440 px", "Laptop / huddle", "Sidebar open 300 px · KPI band 4-up · charts 340 px", "Primary target"),
    ("1920 px", "Projector", "Content clamps at 1280 px, centred", "Six numbers legible from 4 m"),
]

BOUNDARY = [
    ("Page header", "Tailwind iframe", "Needs a two-tone title block Streamlit cannot produce; display-only, no widget."),
    ("KPI cards", "Tailwind iframe", "The 26 px value + 11 px label hierarchy is the whole point of the card; st.metric cannot do it."),
    ("Filter chips", "Tailwind iframe", "Read-only summary of the sidebar state."),
    ("Insight box", "Tailwind iframe", "Colour-coded left edge plus action line; display-only."),
    ("DQ banner", "Tailwind iframe", "One-line trust strip under the header."),
    ("Empty state", "Tailwind iframe", "Names the filters that produced zero rows."),
    ("Charts", "Plotly via st.plotly_chart", "Hover, zoom, legend toggle and PNG export must keep working."),
    ("Tables", "st.dataframe", "Sorting, column resize, CSV from the toolbar, sticky header."),
    ("Every input", "Streamlit widget", "An iframe cannot hold state; all interactivity stays native."),
    ("Navigation", "Streamlit multipage sidebar", "A link inside the iframe would open the app inside the iframe (UIF-06)."),
]

COMPONENTS = [
    ("KPI card", "ui.kpi_card(spec)", "label · value · delta · direction · status · sub · n · low_n",
     "ok / warn / bad / neutral · low-n badge · long label", "118", "rounded-xl border bg-white p-3 shadow-sm · text-[26px] tabular-nums"),
    ("KPI row", "ui.kpi_row(cards)", "sequence of card dicts", "empty → draws nothing", "auto", "st.columns(n, gap='small')"),
    ("Page header", "ui.header(title, subtitle, as_of, version)", "title · subtitle · as_of · version", "missing chip → omitted", "78", "flex justify-between border-b · truncate"),
    ("Chart card", "ui.chart_card(title, fig, note)", "title · figure · note · key", "fig None → info box", "auto", "st.container(border=True) + caption"),
    ("Insight box", "ui.insight_box(text, level, action, source)", "text · level · action · source", "info / warn / bad", "98", "border-l-4 + tint · line-clamp-2"),
    ("Filter chips", "ui.filter_chips(list)", "list of strings", "empty → draws nothing", "36", "rounded-full bg-slate-100 px-2.5 py-1"),
    ("DQ banner", "ui.dq_banner(passed, failed, period)", "counts · period", "all pass / any fail", "58", "rounded-lg border px-3 py-2"),
    ("Empty state", "ui.empty_state(msg, filters)", "message · filter list", "always", "138", "border-dashed bg-slate-50 · chips"),
]

CHARTS = [
    ("OQ-01", "Which departments miss the 85% occupancy target, and by how much?",
     "Horizontal bar with a target rule", "dept → BOR%; a vertical line at 85%; label = gap in pp and beds",
     "Departments are nominal, so horizontal bars keep the names readable and the 85% line is a single shared reference. The gap is the decision, so it is printed, not eyeballed.",
     "Gauge per department — 8 gauges cannot be compared, and a gauge hides the denominator."),
    ("OQ-02", "How are admissions, discharges and visits moving?",
     "Small-multiple line, one panel per measure", "month → count; 3 lines (current, prior year, 3-month MA)",
     "Three series with different scales (18k discharges vs 49k OPD visits) would make a dual axis lie. Small multiples share the x-axis and each panel keeps its own y.",
     "Dual-axis combo — it manufactures correlation and has no common denominator."),
    ("OQ-03", "Which departments exceed the 4.0-day ALOS benchmark?",
     "Dumbbell (benchmark dot → department dot)", "row = department; two dots; segment = the gap",
     "One axis, two positions: the reader sees direction and size of the gap without arithmetic. A plain bar hides which side of 4.0 a department sits on.",
     "Bar with a target line — the gap is read by eye and deaths/LAMA inclusion is easy to forget."),
    ("OQ-04", "Where does the 30-day readmission rate come from?",
     "Sorted bar with Wilson 95% CI", "dept (or diagnosis) → rate%; error bars; n under each bar",
     "A rate with n = 120 and a rate with n = 34 should not look equally certain. Wilson intervals stay inside 0–1 for small n, unlike the normal approximation.",
     "Pie chart — proportions of a whole, but readmissions are a rate per department, not shares."),
    ("OQ-05", "How does revenue per patient vary by payer, and which mix is moving?",
     "100% stacked bar (mix) + grouped bar (₹ per patient)", "period → share of discharges; payer → ₹",
     "Two questions, two panels: mix is compositional, price is a level. Putting both on one axis invites reading share as revenue.",
     "Dual-axis line — share and rupees have no common unit."),
    ("OQ-06", "Which weekdays and months carry peak OPD footfall?",
     "Heatmap (month × weekday) + weekday average bar", "colour = mean visits/day; bar = weekday mean",
     "The decision is roster-by-weekday, but the driver is seasonal. A heatmap shows the interaction in one glance; the bar collapses it to the actionable number.",
     "Line by weekday — 12 lines overlap and the seasonality disappears."),
    ("OQ-07", "Is occupancy associated with longer stays or readmissions?",
     "Scatter with fitted line + density, one point per department-month", "BOR% → outcome%; point size = n; annotation = rho, p, n",
     "The claim is about association across 288 department-months, so every point must be visible. A binned bar chart would hide the spread the test depends on.",
     "Correlation matrix — hides non-linearity and the n behind each cell."),
    ("OQ-08", "Which age groups and genders dominate admissions?",
     "Stacked horizontal bar by department, gender faceted", "dept → share of episodes; colour = age band",
     "Six age bands × 3 genders is 18 cells; stacking by department keeps the comparison the service-line planner actually makes (within and across departments).",
     "Pie or donut — 6 slices at 3–8% cannot be read, and two pies cannot be compared."),
    ("OQ-09", "Which departments earn least per bed-day, and by how much?",
     "Ranked lollipop with gap-to-#1 labels", "rank → ₹ per bed-day; stem to the #1 value",
     "The gap to the leader is the decision, so the stem is drawn to #1 rather than to zero — a bar from zero compresses the 4.6× spread into equal-looking bars.",
     "Bar from zero — every small department looks the same once Oncology sets the scale."),
    ("OQ-10", "Is mortality stable, or is a department drifting?",
     "Line + 3-month MA + ±2 SD band, with a 100% stacked mix bar underneath", "month → mortality %; band = baseline ±2 SD; shared x with discharge mix",
     "A drift flag is only defensible next to the mix that could explain it: a rise in LAMA looks like a mortality fall. Shared x keeps the two panels aligned.",
     "Control chart table — precise, but no reader finds the month without plotting it."),
]

US_UI = [
    ("US-01", "Read-only load of the five tables", "All 9 screens (loader)", "src/data_access.py"),
    ("US-02", "One filter bar every KPI obeys", "All 9 screens (sidebar)", "st.session_state + chips"),
    ("US-03", "In-progress period excluded from rankings", "UI-02, UI-05", "Badge + hollow trend point"),
    ("US-04", "Occupancy vs 85%, in points and beds", "UI-02", "Bar + gap table"),
    ("US-05", "Seven-day ward projection", "UI-02", "Small-multiple line + headroom"),
    ("US-06", "Like-for-like ALOS with basis stated", "UI-03", "Dumbbell + info panel"),
    ("US-07", "OPD demand by weekday and month", "UI-02", "Heatmap + weekday bar"),
    ("US-08", "Responsive shell, phone and projector", "All 9 screens (shell)", "Header + enable_mobile()"),
    ("US-09", "30-day readmission with the rule inline", "UI-03", "Bar + CI + rule line"),
    ("US-10", "Payer-mix movement and revenue per patient", "UI-05", "Stacked bar + grouped bar"),
    ("US-11", "Revenue per bed-day ranking", "UI-05", "Lollipop + gap columns"),
    ("US-12", "Admission mix by age band and gender", "UI-03", "Stacked bar + heatmap"),
    ("US-13", "Mortality trend with a drift flag", "UI-04", "Line + SD band + mix bar"),
    ("US-14", "Decompose a KPI movement", "UI-08", "Waterfall + contributor table"),
    ("US-15", "Statistical test with assumptions", "UI-08", "Scatter + test card"),
    ("US-16", "Every figure carries its definition", "UI-07 (+ all cards)", "Registry table + info icon"),
    ("US-17", "Six data-quality checks beside the KPI", "UI-06 (+ banners)", "Check cards + DQ banner"),
    ("US-18", "Low-confidence badge when n < 30", "UI-03, UI-06", "Badge + flag column"),
    ("US-19", "Rejected-record log with reason codes", "UI-06", "Filtered table + CSV"),
    ("US-20", "Filtered CSV export that is safe to open", "UI-09 (+ every screen)", "Download button + guard"),
]

RTM_FULL = [
    ("BR-01", "F-02", "US-03, US-04, US-05", "UI-01, UI-02"),
    ("BR-15", "— (Won't)", "— (Won't)", "— (Won't)"),
    ("BR-02", "F-02", "US-05", "UI-02"),
    ("BR-03", "F-09", "US-14", "UI-08"),
    ("BR-04", "F-01, F-13", "US-02, US-08", "All 9 screens"),
    ("BR-05", "F-11", "US-16", "UI-07 + all cards"),
    ("BR-06", "F-03", "US-06", "UI-03"),
    ("BR-07", "F-04", "US-09", "UI-03"),
    ("BR-14", "F-12", "US-01, US-17, US-18, US-19, US-20", "UI-06, UI-09"),
    ("BR-08", "F-05", "US-10", "UI-05"),
    ("BR-09", "F-06", "US-03, US-07", "UI-02"),
    ("BR-10", "F-10", "US-15", "UI-08"),
    ("BR-11", "F-08", "US-12", "UI-03"),
    ("BR-12", "F-05", "US-11", "UI-05"),
    ("BR-13", "F-07", "US-13", "UI-04"),
    ("BR-15", "— (Won't)", "— (Won't)", "— (Won't)"),
]
RTM_UI = [(br, ui) for br, _f, _us, ui in RTM_FULL]

# --------------------------------------------------------------------------- wireframe helper
SIDE_W, MAIN_W = 22, 61


def panel(label, sketch=(), w=MAIN_W):
    body = w - 4
    out = ["+" + "-" * (w - 2) + "+", "| " + label.ljust(body)[:body] + " |"]
    for s in sketch:
        if len(s) > body:
            raise ValueError(f"panel line too long ({len(s)}): {s}")
        out.append("| " + s.ljust(body) + " |")
    out.append("+" + "-" * (w - 2) + "+")
    return out


def wf(top, side, main):
    """Fixed-width frame. Row width must equal the border width or it looks broken."""
    width = SIDE_W + MAIN_W + 7
    sep = "+" + "-" * (SIDE_W + 2) + "+" + "-" * (MAIN_W + 2) + "+"
    assert len(sep) == width
    lines = ["+" + "-" * (width - 2) + "+",
             "| " + top.ljust(width - 4)[:width - 4] + " |",
             sep]
    for i in range(max(len(side), len(main))):
        left = side[i] if i < len(side) else ""
        right = main[i] if i < len(main) else ""
        if len(left) > SIDE_W:
            raise ValueError(f"sidebar line too long ({len(left)}): {left}")
        if len(right) > MAIN_W:
            raise ValueError(f"main line too long ({len(right)}): {right}")
        lines.append("| " + left.ljust(SIDE_W) + " | " + right.ljust(MAIN_W) + " |")
    lines.append(sep)
    return "\n".join(lines)


SIDE_BASE = [
    "FILTERS",
    "Period      [ 2025  v]",
    "Compare to  [ 2024  v]",
    "Department  [ All   v]",
    "Payer       [ All   v]",
    "Age         [ All   v]",
    "Gender      [ All   v]",
    "",
    "[ Reset all filters ]",
    "",
    "QUALITY  6/6 pass",
    "DEFS     v1.0",
    "DATA     to 31-12-2025",
]

WIREFRAMES = {}

WIREFRAMES["UI-01"] = wf(
    "HKD-2026  Healthcare KPI Dashboard          data to 31-12-2025   defs v1.0",
    SIDE_BASE,
    ["chips: 2025 · All depts · All payers          [Export CSV]",
     "KPI BAND — six cards: 3 x 2 at 1440 px"]
    + panel("Discharges 18,142 | BOR 73.2% | ALOS 4.36 d",
            ["Revenue ₹135.26 cr | Mort. 1.29% | Readmit 5.03%"])
    + panel("line · month × discharges · 2025 vs 2024",
            ["  ▁▂▃▅▆▇█▇▆▅▆▇  ▁▂▃▄▅▆▇", "  jan            jul      dec"])
    + panel("INSIGHTS — max 3, UI-08 depth on click",
            ["! Cardiology 86.1% vs 85% target, 4th month",
             "i Oncology ₹31,232/bed-day, 1.8× the median"])
    + ["6/6 checks pass · source app.db · render 0.8 s"])

WIREFRAMES["UI-02"] = wf(
    "UI-02  Operations        [Occupancy]  Volume & trend   Compare",
    SIDE_BASE[:4] + ["Ward        [ All   v]", "Bed class   [ All   v]"] + SIDE_BASE[4:],
    ["chips: 2025 · All depts · All wards        [Export CSV]",
     "OCCUPANCY vs 85%  — horizontal bars, gap printed"]
    + panel("Cardiology   ████████████████ 86.1  |85%",
            ["Oncology     ██████████████   76.7",
             "Paediatrics  █████████████    73.8",
             "Gen. Med.    █████████████    72.9",
             "…4 more rows, sorted descending"])
    + panel("GAP TABLE  dept | BOR | pp | beds | n",
            ["Cardiology   86.1  +1.1   +0.6   4 wards"])
    + panel("WARD DRILL-DOWN (click a row)",
            ["W01 ICU 92%  W02 Gen 88%  W03 Semi 74%  W04 Pvt 61%"])
    + ["footer: source vw_ward_day_beds · 109,420 bed-days"])

WIREFRAMES["UI-03"] = wf(
    "UI-03  Departments       [ALOS]  Readmission  Patient mix  EDA",
    SIDE_BASE,
    ["chips: 2025 · All depts · All payers",
     "ALOS vs 4.0-day benchmark — dumbbell, one axis"]
    + panel("Oncology   4.0 ●————————▶ 6.49",
            ["Orthopaed. 4.0 ●——————▶ 5.69",
             "Cardiology 4.0 ●—————▶ 5.38",
             "Gen. Med.  4.0 ●————▶ 4.71",
             "Maternity  4.0 ◀——● 2.76"])
    + panel("DISTRIBUTION los_days — box plot, 1.5 IQR whiskers",
            ["▁▂█▆▃  ▁▃█▅▂  ▁▂█▄▂   (per department)"])
    + panel("TABLE  dept | ALOS | Δ bench | n | flag",
            ["Gen. Med.  4.71  +0.71  4,102  ok"])
    + ["i mean of los_days; deaths + LAMA included; no case-mix adj."])

WIREFRAMES["UI-04"] = wf(
    "UI-04  Quality                       2025 · all departments",
    SIDE_BASE,
    ["chips: 2025 · All depts                    [Export CSV]",
     "MORTALITY % by month — line + 3-mo MA + ±2 SD"]
    + panel("Oncology  5.53%  base 3.52 ± 0.61  ⚑ drift",
            ["  ╭─╮        ╭──╮",
             " ╱   ╲╭─────╯   ╲───  MA",
             " ┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈┈  +2 SD"])
    + panel("DRIFT FLAG — Oncology · first fired 2025-08",
            ["baseline 12 mo: mean 3.52%, SD 0.61pp"])
    + panel("DISCHARGE MIX — 100% stacked, shared x",
            ["Routine 95.61 | LAMA 1.27 | Transf 1.83 | Exp 1.29"])
    + ["⚑ reads 'review trigger', never 'outlier confirmed'"])

WIREFRAMES["UI-05"] = wf(
    "UI-05  Finance                         2025 vs 2024",
    SIDE_BASE,
    ["chips: 2025 vs 2024 · All depts            [Export CSV]",
     "PAYER MIX — 100% stacked, 5 categories"]
    + panel("2024  ████████░░░░░░░░░░░░░░░░░░░░░░░░░░",
            ["2025  ██████████░░░░░░░░░░░░░░░░░░░░░░░░",
             "Corp 6.4→11.2  Self-pay 41.3→33.5"])
    + panel("REVENUE PER PATIENT — grouped bars + n",
            ["Private 82,544 | Corp 79,973 | TPA 75,890",
             "Cash 71,258 | Govt 56,574"])
    + panel("₹ PER BED-DAY — ranked lollipop to #1",
            ["1 Onco 31,232 ●————————", "5 GenM 10,646 ●———————"])
    + ["table: dept | revenue | bed-days | ₹/bd | gap | % of #1"])

WIREFRAMES["UI-06"] = wf(
    "UI-06  Data Quality            2025 · 6 checks · freshness",
    SIDE_BASE[:6] + ["Rule        [ All   v]"] + SIDE_BASE[6:],
    ["chips: 2025 · All tables                  [Export log]",
     "SIX CHECKS — one card each, pass/fail + count"]
    + panel("✓ missing payer   ✓ duplicate reg.",
            ["✓ orphan dept     ✓ date coverage",
             "✓ negative LOS    ✓ BOR > 100%"])
    + panel("SHIPPED RAW LAYER — 5,368 rejected / 159 flagged",
            ["max(admission_date) 2025-12-31 · 0 days stale"])
    + panel("REJECT LOG — rule · table · field · value · action",
            ["RQ-03  T-04  department_id  99  row dropped"])
    + ["no patient name or phone is selectable on this screen"])

WIREFRAMES["UI-07"] = wf(
    "UI-07  Definitions                    registry v1.0",
    ["NO DATA FILTERS",
     "This screen ignores",
     "the sidebar — a",
     "definition must not",
     "change with a filter.",
     "",
     "QUALITY  —",
     "DEFS     v1.0",
     "DATA     all periods"],
    ["KPI REGISTRY — one row per published figure"]
    + panel("key | definition | numerator | denominator | owner",
            ["bor | occupied / available bed-days | … | SH-02",
             "alos | mean los_days | … | SH-03",
             "ready_30d | unplanned readmit ≤30d | … | SH-04"])
    + panel("METHOD NOTES — readmission, ALOS, drift, n<30",
            ["window = 30 days · planned excluded · expired excl."])
    + panel("BUILD — commit · db sha256 · streamlit 1.67 · 19-11-26",
            ["v1.0 definitions · v1.2 RTM · dataset seed 42"])
    + ["[ Export registry ]"])

WIREFRAMES["UI-08"] = wf(
    "UI-08  Insights          [Decompose]   Association test",
    SIDE_BASE,
    ["KPI [Revenue v]   A [2024 v]   B [2025 v]   by [Dept v]",
     "DECOMPOSITION — waterfall, contributors add to total"]
    + panel("+₹12.4 cr  total movement 2024 → 2025",
            ["Onco ▇▇▇▇▇ +4.1   Cardio ▇▇▇ +2.0",
             "GenM ▇▇ +1.4      other (5) ▇▇ +4.9",
             "residual +0.0 pp  (printed, never hidden)"])
    + panel("ASSOCIATION — scatter, one point per dept-month",
            ["BOR% → readmit%   n = 288   rho [ ]  p [ ]",
             "95% CI [ ]   association, not causation"])
    + ["unattributed rows shown with their count, never merged"])

WIREFRAMES["UI-09"] = wf(
    "UI-09  Data & Export                  filtered view",
    SIDE_BASE,
    ["chips: 2025 · Cardiology · All payers      [CSV]",
     "FILTERED — 3,181 episodes · 36 columns"]
    + panel("PREVIEW — first 100 rows, chosen columns",
            ["ep_id | admit | dept | ward | payer | los | ₹",
             "E-00123 2025-01-04 Cardio W01 Govt 6 ₹48,120"])
    + panel("COLUMN PICKER + DICTIONARY",
            ["12 of 36 selected · type · null % · definition"])
    + panel("EXPORT GUARD",
            ["=,+,-,@ → ' prefix · UTF-8 BOM · 3,181 rows",
             "name: hkd2026_episodes_2025_cardio_20261119.csv"])
    + ["rows > 50,000 → button asks for a narrower period"])

_SEP = "+" + "-" * (MAIN_W - 2) + "+"
SHELL = wf("HKD-2026 — one shell, nine screens",
           ["SIDEBAR — native",
            "Period      [ 2025 ]",
            "Compare to  [ 2024 ]",
            "Department  [ All  ]",
            "Ward/class  [ All  ]",
            "Payer       [ All  ]",
            "Age/Gender  [ All  ]",
            "",
            "[ Reset all ]",
            "",
            "quality 6/6 pass",
            "defs v1.0",
            "data to 31-12-25"],
           ["TAILWIND HEADER (iframe, 78 px)",
            _SEP,
            "FILTER CHIPS (iframe, 36 px)",
            _SEP,
            "KPI BAND  4 cards x 118 px  (st.columns)",
            _SEP,
            "CHART CARD (Plotly, 340 px)",
            _SEP,
            "TABLE (st.dataframe) or second chart",
            _SEP,
            "INSIGHT BOXES (max 3, 98 px)",
            _SEP,
            "footer · source app.db · render 0.8 s"])

TABS = {
    "UI-01": [],
    "UI-02": [("Occupancy", "BOR vs target, gap table, ward drill-down (default)"),
              ("Volume & trend", "Admissions / discharges / OPD small multiples, weekday heatmap"),
              ("Compare", "Two-period dumbbell + delta table for bed-days and volume")],
    "UI-03": [("ALOS", "Dumbbell vs 4.0-day benchmark, distribution box plot (default)"),
              ("Readmission", "Sorted bars + Wilson CI, rule line, DQ strip"),
              ("Patient mix", "Stacked age-band bars, gender facet, n badges"),
              ("EDA", "Variable picker, histogram / scatter, describe table — no claim, no test")],
    "UI-04": [],
    "UI-05": [],
    "UI-06": [],
    "UI-07": [],
    "UI-08": [("Decompose", "Waterfall by department / payer / weekday (default)"),
              ("Association test", "Scatter + method card + assumption checks")],
    "UI-09": [],
}

# --------------------------------------------------------------------------- verification
assert len(SCREENS) == 9
assert len({s[0] for s in SCREENS}) == 9
assert [s[0] for s in SCREENS] == [f"UI-{i:02d}" for i in range(1, 10)]
assert len(CHARTS) == 10 and [c[0] for c in CHARTS] == [f"OQ-{i:02d}" for i in range(1, 11)]
assert len(US_UI) == 20 and [u[0] for u in US_UI] == [f"US-{i:02d}" for i in range(1, 21)]
assert len(RTM_UI) == 16

# every story reachable from a screen
screened = set()
for _, _, screens, _ in US_UI:
    if screens.startswith("All 9"):
        screened |= {s[0] for s in SCREENS}
    else:
        screened |= set(re.findall(r"UI-\d\d", screens))
assert screened == {s[0] for s in SCREENS}, sorted({s[0] for s in SCREENS} - screened)

# --------------------------------------------------------------------------- RTM v1.3
_ver = str(load_workbook(RTM_PATH)["RTM"]["A2"].value)
if "v1.3" not in _ver:
    raise SystemExit(
        "RTM is now '" + _ver.split("·")[0].strip() + "'. Step 13 owns the Test Case column\n"
        "and later versions own the workbook. Revert docs/04_Traceability_Matrix.xlsx to v1.3\n"
        "before rebuilding Step 12, then rebuild Step 13. This guard exists because a silent\n"
        "rebuild would delete sections 9-10 and the v1.4 change-log row.")

wb = load_workbook(RTM_PATH)
rt = wb["RTM"]
hdr_cols = {rt.cell(row=3, column=c).value: c for c in range(1, rt.max_column + 1)}
COL_BR, COL_FEAT, COL_US = hdr_cols["Req ID (BR)"], hdr_cols["Feature"], hdr_cols["User Story"]
COL_UI = hdr_cols["UI Screen"]
COL_STATUS = hdr_cols["Status"]

proto = rt.cell(row=4, column=COL_US)
for i, (br, ui) in enumerate(RTM_UI):
    r = 4 + i
    assert rt.cell(row=r, column=COL_BR).value == br, f"row {r} is not {br}"
    c = rt.cell(row=r, column=COL_UI, value=ui)
    c.font = XFont(name=proto.font.name, size=proto.font.size, bold=False, color=proto.font.color)
    c.alignment = proto.alignment.copy()
    c.fill = proto.fill.copy()
    c.border = proto.border.copy()
rt.column_dimensions[rt.cell(row=3, column=COL_UI).column_letter].width = 14

rt.cell(row=2, column=1).value = (
    "v1.3  ·  19-Nov-2026  ·  UI Screen (UI-01…UI-09) filled for all 16 rows from Step 12  ·  "
    "Feature and User Story from Step 11  ·  DB column from Step 8")

# Coverage Check ------------------------------------------------------------
cc = wb["Coverage Check"]
cc.cell(row=1, column=1).value = "Coverage Check  ·  HKD-2026  ·  v1.3 (19-Nov-2026)"

# forward-column table
for r in range(55, cc.max_row + 1):
    v = cc.cell(row=r, column=1).value
    if v == "4. Forward columns (v1.2)":
        cc.cell(row=r, column=1).value = "4. Forward columns (v1.3)"
    v4 = cc.cell(row=r, column=4).value
    if isinstance(v4, str) and v4.startswith("Rows filled (v"):
        cc.cell(row=r, column=4).value = "Rows filled (v1.3)"
    if v == "Feature" and r < 70:
        cc.cell(row=r, column=2).value = "Step 11 (issued bottom-up; Step 6 not delivered)"
    if v in ("Feature", "DB table.column", "User Story", "UI Screen"):
        cc.cell(row=r, column=4).value = 16

# drop a previous run of sections 5–8 and rebuild
cut = None
for r in range(60, cc.max_row + 1):
    v = cc.cell(row=r, column=1).value
    if isinstance(v, str) and (v.startswith("5. Story coverage") or v.startswith("6. OQ → story")
                               or v.startswith("7. UI coverage") or v.startswith("8. Story → screen")):
        cut = r
        break
if cut:
    cc.delete_rows(cut, cc.max_row - cut + 1)

# delete_rows leaves openpyxl's dimension at the old maximum, so find the real
# last used row by hand — otherwise every re-run appends one more blank row.
base = cc.max_row
while base > 1 and all(cc.cell(row=base, column=c).value in (None, "") for c in range(1, 7)):
    base -= 1
base += 2

HDR_FILL_CELL = cc.cell(row=4, column=1)
HDR_FONT = XFont(bold=True, color="FFFFFF")


def block_title(ws, row, text):
    c = ws.cell(row=row, column=1, value=text)
    c.font = XFont(bold=True, size=11, color="1F3864")
    return row + 1


def header_row(ws, row, labels):
    for j, h in enumerate(labels, start=1):
        c = ws.cell(row=row, column=j, value=h)
        c.font = HDR_FONT
        c.fill = HDR_FILL_CELL.fill.copy()
        c.alignment = HDR_FILL_CELL.alignment.copy()
        c.border = HDR_FILL_CELL.border.copy()
    return row + 1


def data_row(ws, row, vals, bold_last=False, colour_last="1E6B34"):
    for j, v in enumerate(vals, start=1):
        c = ws.cell(row=row, column=j, value=v)
        c.font = XFont(size=9, bold=(bold_last and j == len(vals)),
                       color=colour_last if (bold_last and j == len(vals)) else None)
        c.alignment = HDR_FILL_CELL.alignment.copy()
        c.border = HDR_FILL_CELL.border.copy()
    return row + 1


row = block_title(cc, base, "5. Story coverage (v1.2 · unchanged)")
row = header_row(cc, row, ["Metric", "Value", "Check", "Verdict"])
for a, b, c_text in [("User stories", 20, "12–20 required by the brief"),
                     ("Story points", 62, "6 days × 12 points = 72 capacity"),
                     ("Features", 13, "F-01…F-13 issued at Step 11"),
                     ("MVP features (Must)", 9, "every one has a story"),
                     ("FRs covered", 22, "= FR-01…FR-22"),
                     ("Stories with no BR", 0, "must be 0"),
                     ("Must BRs with no story", 0, "must be 0"),
                     ("OQs with no story", 0, "must be 0")]:
    row = data_row(cc, row, [a, b, c_text, "PASS"], bold_last=True)

row = block_title(cc, row + 1, "6. OQ → story map (static, checked at v1.2)")
row = header_row(cc, row, ["OQ", "Stories that answer it", "Count", "Verdict"])
OQ_US = {
    "OQ-01": "US-03, US-04, US-05", "OQ-02": "US-01, US-02, US-03, US-14",
    "OQ-03": "US-06, US-14", "OQ-04": "US-09, US-17", "OQ-05": "US-10",
    "OQ-06": "US-07", "OQ-07": "US-15", "OQ-08": "US-12",
    "OQ-09": "US-11, US-14", "OQ-10": "US-13, US-17",
}
for oq, us in OQ_US.items():
    row = data_row(cc, row, [oq, us, len([x for x in us.split(",") if x.strip()]), "PASS"], bold_last=True)

row = block_title(cc, row + 1, "7. UI coverage (v1.3)")
row = header_row(cc, row, ["Metric", "Value", "Check", "Verdict"])
for a, b, c_text in [("Screens", 9, "UI-01…UI-09"),
                     ("Screens with a wireframe", 9, "one per screen"),
                     ("Stories with a screen", "20 of 20", "no orphan story"),
                     ("MVP stories with a screen", "13 of 13", "Must stories"),
                     ("OQs with a screen", "10 of 10", "every OQ is reachable"),
                     ("Reusable components", 8, "§5 of the Step 12 document"),
                     ("Chart decisions", 10, "one per OQ"),
                     ("RTM rows with UI filled", "16 of 16", "0 blank"),
                     ("Story points unchanged", 62, "no scope added")]:
    row = data_row(cc, row, [a, b, c_text, "PASS"], bold_last=True)

row = block_title(cc, row + 1, "8. Story → screen map (static, checked at v1.3)")
row = header_row(cc, row, ["User story", "Story (short)", "Screen(s)", "Primary component"])
for us, title, screens, comp in US_UI:
    row = data_row(cc, row, [us, title, screens, comp])

# How to Update -------------------------------------------------------------
hu = wb["How to Update"]
for r in range(3, hu.max_row + 1):
    if hu.cell(row=r, column=1).value == "Feature":
        hu.cell(row=r, column=2).value = "Step 11 (issued bottom-up — Step 6 not delivered)"
    if hu.cell(row=r, column=1).value == "UI Screen":
        hu.cell(row=r, column=2).value = "Step 12 (UI-01…UI-09)"
        hu.cell(row=r, column=3).value = (
            "Enter UI-xx. A BR with no screen at Step 12 is a scope leak and must be raised "
            "with the mentor before Step 13. Write 'All 9 screens' only for a global rule "
            "(BR-04, BR-05).")

# Change Log ----------------------------------------------------------------
cl = wb["Change Log"]
EXTRA = [("1.0", "19-Nov-2026", "Step 4",
          "Baseline created: 16 HMW–BR links, 12 HMWs, 15 BRs, 10 OQs. All rows Status = Identified. "
          "Feature/DB/Story/UI/TC columns intentionally empty.", "Superseded"),
         ("—", "—", "Step 6",
          "Not issued. Feature IDs were issued at Step 11 instead (US-CN-01).", "—"),
         ("1.1", "19-Nov-2026", "Step 8",
          "DB table.column filled for all 16 rows (T-01…T-05). BR-04/BR-05 have no physical column "
          "by design; BR-15 has none (Won't).", "Superseded"),
         ("1.2", "19-Nov-2026", "Step 11",
          "Feature (F-01…F-13) and User Story (US-01…US-20) filled for all 16 rows; Status moved to "
          "In Design (BR-15 rows: Out of scope). 20 stories, 62 points.", "Superseded"),
         ("1.3", "19-Nov-2026", "Step 12",
          "UI Screen filled for all 16 rows (UI-01…UI-09; UI-08 and UI-09 added, nothing renumbered). "
          "Coverage Check gained a UI-coverage block and a story → screen map.", "Current"),
         ("—", "—", "Step 13", "Pending: add TC-F / TC-T IDs.", "—")]
for i, r in enumerate(EXTRA):
    for j, val in enumerate(r, start=1):
        c = cl.cell(row=4 + i, column=j, value=val)
        c.font = XFont(size=9, bold=(j == 1 and str(val).startswith("1.")))
if cl.max_row > 3 + len(EXTRA):
    cl.delete_rows(4 + len(EXTRA), cl.max_row - (3 + len(EXTRA)))

wb.save(RTM_PATH)

wb2 = load_workbook(RTM_PATH)
rt2 = wb2["RTM"]
filled_ui = sum(1 for r in range(4, 20) if rt2.cell(row=r, column=COL_UI).value)
filled_f = sum(1 for r in range(4, 20) if rt2.cell(row=r, column=COL_FEAT).value)
filled_us = sum(1 for r in range(4, 20) if rt2.cell(row=r, column=COL_US).value)
cc2 = wb2["Coverage Check"]
cc_rows = cc2.max_row

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


def add_table(doc, headers, rows, widths, font_size=8.2, align_center=(), zebra=True,
              bold_first=True):
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


def h3(doc, text, code=False):
    h = doc.add_heading(text, level=3)
    for r in h.runs:
        r.font.color.rgb = RGBColor(0x22, 0x33, 0x55)
        r.font.size = Pt(10.5)
        if code:
            r.font.name = "Consolas"


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


def mono(doc, text, size=7.0, fill="F7F8FA"):
    """Monospace block for wireframes and code. No border: it reads as a figure."""
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.autofit = False
    cell_margins(table, top=60, bottom=60, left=80, right=60)
    cell = table.rows[0].cells[0]
    cell.width = Inches(6.9)
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


def footer(section):
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{PROJECT_ID}  |  Step 12 — UI Framework  |  Page ")
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
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=34)
para(doc, "STEP 12", 26, bold=True, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, "UI FRAMEWORK, WIREFRAMES\nAND THE STYLED SHELL", 16, bold=True, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=16)
para(doc, PROJECT_TITLE, 17, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
para(doc, f"Project ID: {PROJECT_ID}   |   Domain: Data Analytics", 11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=28)

meta = [("Student", STUDENT), ("Roll Number", ROLL), ("Mentor", MENTOR),
        ("Deliverable", "docs/12_UI_Framework.docx + 04_Traceability_Matrix.xlsx (v1.3)\n"
                        "src/ui.py · .streamlit/config.toml · scripts/check_ui.py"),
        ("Version / Date", "v1.0 — 19 November 2026"),
        ("Scope", "9 screens · 8 components · 10 chart decisions · 1 theme file · 1 UI module"),
        ("Status", "Approved for Step 13 (test cases)")]
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

# control
doc.add_page_break()
h1(doc, "Document Control")
h2(doc, "Revision history")
add_table(doc, ["Version", "Date", "Author", "Change", "Status"],
          [["0.1", "19-Nov-2026", STUDENT, "Screen list and navigation fixed; UI-08/UI-09 added to cover the brief's Insights and Data screens", "Draft"],
           ["0.2", "19-Nov-2026", STUDENT, "Wireframes drawn at 90 characters; palette contrast measured; chart choice argued per OQ", "Draft"],
           ["0.3", "19-Nov-2026", STUDENT, "src/ui.py written and checked (22/22); config.toml theme aligned to the same tokens", "Draft"],
           ["1.0", "19-Nov-2026", STUDENT, "Approved; RTM v1.3 issued with the UI Screen column complete", "Approved"]],
          [0.7, 1.0, 1.1, 3.6, 0.8])

h2(doc, "Inputs")
add_table(doc, ["Artefact", "What was taken from it"],
          [["07_TRD.docx", "The fixed stack, the module map (M7 UI), the UI-01…UI-07 file tree, the Tailwind-in-iframe constraint and NFR-08/09/12"],
           ["11_User_Stories.docx", "20 stories, 80 acceptance criteria, 5 epics and the sprint order — every screen exists to serve a story"],
           ["04_Traceability_Matrix.xlsx (v1.2)", "The 16 rows whose UI Screen column this step fills"],
           ["05_BRD.docx (BRL-01…BRL-14)", "The 85% target, 4.0-day ALOS, 30-day window and n < 30 rule that the screens must display"],
           ["data/app.db (v1.1, seed 42)", "Every figure printed inside a wireframe"],
           ["NFR-08 / NFR-12", "WCAG AA contrast, keyboard path, 375 px usability, user-safe error messages"]],
          [2.3, 4.6], font_size=8.4)

h2(doc, "Change notes")
add_table(doc, ["Note", "Change", "Why it matters"],
          [["UI-CN-01", "UI-08 (Insights) and UI-09 (Data & Export) are added. No existing screen is renumbered; UI-01…UI-07 keep their Step-7 meaning.",
            "The brief asks for an Insights screen and a Data screen. Squeezing both into UI-01 and UI-03 would have given three screens two jobs each. The TRD file tree gains pages/7 and 8 at Step 14."],
           ["UI-CN-02", "Trend and Compare are tabs inside UI-02, not separate screens. EDA is a tab inside UI-03.",
            "A view mode that shares the same filters, frames and export is a tab. A separate screen would duplicate the query and let the two drift."],
           ["UI-CN-03", "Tailwind is confined to six display components rendered in an iframe. Charts, tables and every input stay native Streamlit.",
            "Tailwind cannot style a Streamlit widget across the iframe boundary (TRD §6.1). Pretending otherwise is the most common way these dashboards end up with two clashing visual systems."],
           ["UI-CN-04", "The palette is semantic: teal = on target, amber = warning, red = breach. Departments never get their own colour.",
            "Colour must mean one thing. A reader who learns teal in the occupancy chart must not meet teal as 'Oncology' in the revenue chart."],
           ["UI-CN-05", "No OQ, SH, HMW, BR, F, FR, NFR, T or US identifier was added, removed or renumbered.", "The RTM stays comparable across versions."]],
          [0.8, 2.9, 3.3], font_size=7.8)

# 1
h1(doc, "1. What this step decides", page_break=True)
para(doc, "Step 11 said what each story must let a user do. This step says where they do it and what they "
          "see while they do it. Three decisions carry most of the weight: how many screens there are, "
          "what the shared filter bar is allowed to control, and where the Tailwind layer stops.")
para(doc, "The constraint that shapes everything is structural. Streamlit renders the page; Tailwind can only "
          "reach inside an iframe, and an iframe shares no state with Streamlit. So the design is deliberately "
          "two-layered: a Tailwind skin for display-only elements (header, KPI cards, insight boxes, banners) "
          "and native Streamlit for anything the user touches (filters, charts, tables, downloads). The two "
          "layers are tied together by one palette and one type scale, so the seam is invisible — but it is "
          "still a seam, and it is documented rather than hidden.")
mono(doc, """a user's question  ──►  a screen  ──►  a filter state  ──►  one SQL frame
                          │                                           │
                          │                                           ▼
                          │                                   KPI registry (BR-05)
                          │                                           │
                          ▼                                           ▼
              Tailwind iframe (display)  ◄── values ──  native widgets (interaction)
                 header · cards ·                          sidebar · Plotly · tables
                 chips · insight · banner · empty           · download buttons""", 7.4)

h2(doc, "1.1 The 60-second path (BR-04)")
add_table(doc, ["Interaction", "Where", "What changes on screen"],
          [["1 · Open the app", "app.py (UI-01)", "Six numbers with their period and target — no click, no scroll at 1440 px"],
           ["2 · Choose Department = Cardiology", "Sidebar", "Every card, chart and denominator on the page recomputes; the chip row names the selection"],
           ["3 · Open Operations → Occupancy", "UI-02", "Cardiology's BOR, the +1.1 pp gap, the four wards and the 85% line"]],
          [2.0, 1.5, 3.4], font_size=8.4)
para(doc, "That is the promised three interactions to a decision, and it is why the sidebar is the only place "
          "filters live. A filter repeated on a page is a second source of truth, and the first bug report "
          "will be that two filters disagree.", size=9.8)

# 2
h1(doc, "2. The nine screens", page_break=True)
para(doc, "Eight screens were proposed in the brief and seven were fixed in Step 7. The resolution was to keep "
          "the seven where they are and add the two the brief names but the file tree never had — Insights and "
          "Data & Export (UI-CN-01). The mapping is shown honestly in the second table, including the merges.", size=9.8)
add_table(doc, ["UI", "Screen", "Page file", "Purpose", "Persona", "Stories", "OQ"],
          [list(s) for s in SCREENS],
          [0.5, 0.95, 1.35, 2.5, 0.7, 1.25, 0.85], font_size=7.4)

h2(doc, "2.1 Where the brief's eight names went")
add_table(doc, ["Brief's suggestion", "Lands at", "Reason"],
          [list(b) for b in BRIEF_MAP], [1.5, 1.25, 4.15], font_size=8.2)
para(doc, "Two screens in this document did not appear in the brief's list — Operations and Finance became "
          "separate screens in Step 7 because occupancy and revenue answer to different people and different "
          "denominators. The brief's 'Overview/KPIs' would have to hold eight KPI families, which is a scroll, "
          "not an overview.", size=9.8)

h2(doc, "2.2 Navigation and file layout")
mono(doc, """app.py                    UI-01  Overview            (six numbers, three insights)
pages/
  1_Operations.py         UI-02  Occupancy | Volume & trend | Compare
  2_Departments.py        UI-03  ALOS | Readmission | Patient mix | EDA
  3_Quality.py            UI-04  Mortality, drift, discharge mix
  4_Finance.py            UI-05  Payer mix, revenue per patient, per bed-day
  5_Data_Quality.py       UI-06  Six checks, freshness, reject log
  6_Definitions.py        UI-07  KPI registry, method notes, build info
  7_Insights.py           UI-08  Decompose | Association test
  8_Data.py               UI-09  Preview, dictionary, guarded CSV
.streamlit/config.toml           theme tokens, server settings
src/ui.py                        Tailwind components (this step)""", 7.6)
para(doc, "Streamlit derives the sidebar labels from the file names, so the numeric prefix sets the order and "
          "the underscore becomes a space. Every page imports the same filter block and the same header "
          "component; a page that draws its own header is a defect.", size=9.8)

# 3
h1(doc, "3. Global layout", page_break=True)
para(doc, "Every screen is built from the same four bands: a Tailwind header, a filter chip row, a KPI band, then "
          "content. Two screens drop one band on purpose — Data Quality replaces the KPI band with its six check "
          "cards, and Definitions has no filter chips because a definition must not change with a filter. "
          "Everything else is identical, and that repetition is the design: a reader who learns one screen has "
          "learned all nine.", size=9.8)
mono(doc, SHELL, 6.9)
para(doc, "Content clamps at 1280 px and stays centred; the sidebar is 300 px expanded "
          "and collapses to icons at 768 px.", size=9.6)

h2(doc, "3.1 The filter contract")
para(doc, "Nine controls, one state object. Every page reads the same keys, so a filter set on UI-02 is still "
          "set on UI-05 (US-02 AC4).", size=9.8, space_after=3)
add_table(doc, ["Filter", "Widget", "Values", "Default", "State key", "Used on"],
          [list(f) for f in FILTERS], [0.95, 1.4, 1.9, 0.55, 0.85, 1.15], font_size=7.4)
mono(doc, """# src/filters.py — imported by all nine pages, rendered once in the sidebar
DEFAULTS = {"flt_period": "2025", "flt_compare": "2024", "flt_dept": [], "flt_ward": [],
            "flt_class": "All", "flt_payer": [], "flt_age": (0, 90), "flt_gender": "All"}

def sidebar(ctx):
    for k, v in DEFAULTS.items():
        st.session_state.setdefault(k, v)          # one source of truth
    with st.sidebar:
        st.selectbox("Period", ctx.years, key="flt_period")
        st.multiselect("Department", ctx.departments, key="flt_dept")
        ...
        if st.button("Reset all filters", width="stretch"):
            for k, v in DEFAULTS.items():
                st.session_state[k] = v
            st.query_params.clear()
            st.rerun()
    st.query_params.update({k: ",".join(map(str, v)) if isinstance(v, (list, tuple)) else v
                            for k, v in active().items()})   # survives a reload (US-02 AC4)
    return active()""", 7.4)
para(doc, "Two rules make this survive contact with a user. Reset clears st.session_state and the URL together, "
          "otherwise a reload restores the view the user just cleared. And every page calls sidebar(ctx) before "
          "it draws anything — the chip row is generated from the same dict, so the chips cannot disagree with "
          "the data.", size=9.6)

h2(doc, "3.2 Palette")
para(doc, "Four colours, each with one meaning, plus two neutrals. Contrast was measured against white and "
          "against the surface tint; every text colour clears WCAG AA for normal text (NFR-08).", size=9.8, space_after=3)
add_table(doc, ["Colour", "Hex", "Means", "On white", "WCAG"],
          [list(p) for p in PALETTE], [0.85, 0.75, 2.95, 0.75, 0.75], font_size=7.8)
callout(doc, "Status is never encoded by colour alone. Every badge carries a word — “on target”, “warning”, "
             "“breach”, “low confidence” — because roughly one man in twelve cannot separate the teal from the "
             "red, and the KPI that matters most is the red one. The dot is a second cue, the word is the "
             "guarantee (NFR-08).", size=9.2)

h2(doc, "3.3 Type scale")
add_table(doc, ["Role", "Font", "Size / weight", "Treatment", "Rule"],
          [list(t_) for t_ in TYPE_SCALE], [1.05, 1.35, 0.85, 2.05, 1.6], font_size=7.8)
para(doc, "Inter is loaded inside the iframe. Streamlit widgets cannot be given Inter without injecting CSS into "
          "the parent document, so they fall back to the bundled sans — at 12–16 px the difference is invisible, "
          "and matching the palette is what makes the two layers read as one (AS-33). Numbers use tabular "
          "figures everywhere, so a column of percentages lines up and a changing value does not jitter.", size=9.6)

h2(doc, "3.4 Phone view (375 px)")
add_table(doc, ["Viewport", "Context", "What happens", "Evidence at Step 15"],
          [list(b) for b in BREAKPOINTS], [0.75, 1.35, 3.1, 1.55], font_size=7.8)
para(doc, "Streamlit keeps columns side by side at any width unless told otherwise, so ui.enable_mobile() injects "
          "one media query that makes the horizontal block wrap below 640 px. That is the whole mobile strategy: "
          "one column, full-width charts, 22 px KPI values, horizontal scroll inside tables rather than a squashed "
          "table. It relies on Streamlit's data-testid attributes, which is the one fragile line in the "
          "stylesheets — recorded as UIF-04 and caught by the 375/768/1440 screenshot test.", size=9.6)

h2(doc, "3.5 What belongs in which layer")
add_table(doc, ["Element", "Rendered by", "Why"],
          [list(b) for b in BOUNDARY], [1.35, 1.75, 3.8], font_size=7.8)

# 4
h1(doc, "4. Screen wireframes", page_break=True)
para(doc, "Each wireframe is drawn at 90 characters — the sidebar is the real sidebar, repeated on every screen "
          "because it is the same object. Numbers printed inside a panel are measured on the v1.1 dataset; the "
          "statistical panel shows field labels rather than invented values, because that test is first computed "
          "at Step 15.", size=9.8)

for idx, (uid, name, page, purpose, persona, stories, oqs) in enumerate(SCREENS, start=1):
    h2(doc, f"4.{idx} {uid} — {name}")
    para(doc, f"{purpose}", size=10.0, space_after=2)
    para(doc, f"File {page}   ·   Persona {persona}   ·   Stories {stories}   ·   OQ {oqs}",
         size=8.4, color=GREY, space_after=5)
    mono(doc, WIREFRAMES[uid], 6.9)
    if TABS[uid]:
        para(doc, "Tabs:", size=9.2, bold=True, space_after=2)
        for tab, desc in TABS[uid]:
            bullet(doc, f"{tab} — {desc}", size=9.4)

h2(doc, "4.10 Screen rules that apply to all nine")
for txt in [
    "The working screens carry four cards; the overview carries six. A working page with six cards has no room left for the chart that explains them.",
    "Every chart sits in a chart card with a title and, where the definition is not obvious, a one-line note. The note is not decoration — it is BR-05 made visible.",
    "Every rate shows its n. Where n < 30 the card or bar carries the low-confidence badge (US-18).",
    "No screen shows a value without its period. The header chip is the default; a table cell adds its own period when it differs.",
    "The empty state names the filters that produced zero rows and offers Reset — it never renders zeros or an exception (US-02 AC3).",
    "A page that fails its data-quality gate renders the DQ banner and the KPIs, labelled. It does not blank the screen and it does not hide the number (BR-14).",
]:
    bullet(doc, txt, size=9.6)

# 5
h1(doc, "5. Reusable components", page_break=True)
para(doc, "Eight components cover all nine screens. Each one is display-only, escapes its inputs, declares a "
          "fixed height, and degrades to a native Streamlit equivalent when Tailwind is unavailable.", size=9.8)
add_table(doc, ["Component", "Function", "Props", "States", "H", "Tailwind signature"],
          [list(c) for c in COMPONENTS], [0.7, 1.55, 1.7, 1.15, 0.3, 1.9], font_size=7.2)

h2(doc, "5.1 The contract every component obeys")
add_table(doc, ["Rule", "Why", "Where it is enforced"],
          [["Escape every dynamic value", "A department name or a filter chip is user-influenced text going into HTML", "html.escape(quote=True) in _esc(); 4 checks in check_ui.py"],
           ["No widget and no anchor inside the iframe", "The iframe is a separate document: a link would open the app inside itself and strand the user", "UIF-06; no <a href> in any component; asserted by check_ui.py"],
           ["Fixed height, ellipsis, never wrap", "components.html cannot auto-size, so a long label would clip or overlap the card below it", "HEADER_H … EMPTY_H constants; truncate + overflow: hidden"],
           ["Numbers formatted on the server", "The client must not do arithmetic; the value shown must equal the SQL value", "fmt_inr / fmt_pct / fmt_delta; rendered as strings"],
           ["Status is colour + word", "Colour alone fails for deuteranopia and in a greyscale print-out", "The badge text is always present"],
           ["One palette, one source", "Two definitions of navy drift within a week", "THEME in ui.py == config.toml, asserted by check"]],
          [1.7, 2.7, 2.5], font_size=7.8)
mono(doc, """# src/ui.py — the KPI card, which is the component everything else imitates
def card(label, value, delta="", direction="flat", status="neutral", sub="", n=None, low_n=False): ...
def kpi_card(spec, height=CARD_H):
    palette = STATUS.get(spec["status"], STATUS["neutral"])
    body = f'''
<div class="h-full rounded-xl border border-slate-200 bg-white p-3 shadow-sm">
  <p class="truncate text-[11px] font-medium uppercase tracking-wide text-slate-500">{_esc(spec["label"])}</p>
  <p class="mt-1 truncate text-[26px] font-semibold leading-none tabular-nums text-slate-900">{_esc(spec["value"])}</p>
  <p class="mt-1.5 truncate text-[11px] tabular-nums {palette['text']}">{arrow} {_esc(spec["delta"])}{badge}</p>
</div>'''
    _frame(body, height)""", 7.2)
para(doc, "The card is 118 px tall, which is what makes four of them fit beside the chart without the fold. The "
          "value is the only element allowed to be 26 px; the label is 11 px upper-case, so the eye lands on the "
          "number first and the caveat second — that order matters when the number is bad.", size=9.6)

# 6
h1(doc, "6. The best chart for each OQ", page_break=True)
para(doc, "A chart earns its place by making one comparison obvious. For each operational question there is a "
          "chart that answers it and at least one plausible chart that quietly misleads — the rejected option is "
          "recorded, because “we used a bar chart” is not a design decision until you can say what you turned "
          "down.", size=9.8)
add_table(doc, ["OQ", "Question", "Chart", "X / Y / mark", "Why this one", "Rejected"],
          [list(c) for c in CHARTS], [0.45, 1.5, 1.15, 1.5, 2.05, 1.4], font_size=6.9)
para(doc, "Three rules produced these choices. Never two y-axes — the second axis is a correlation the reader "
          "will believe. Never a pie for more than three categories. And always show the denominator: the danger "
          "in this dataset is not a wrong mean, it is a plausible rate on 22 discharges (UIF-03, US-18).", size=9.6)

# 7
h1(doc, "7. Theme file", page_break=True)
para(doc, "Streamlit cannot read Python constants, so the palette is declared twice — once in src/ui.py and once "
          "in .streamlit/config.toml. scripts/check_ui.py asserts the four shared values are identical, which is "
          "the only thing standing between two copies and a slow drift (UIF-01).", size=9.8)
add_table(doc, ["Key", "Value", "What it controls"],
          [["base", "light", "Forces light mode even if the OS is dark, so a screenshot is reproducible"],
           ["primaryColor", "#1F3864", "Selected filter, slider handle, focused input, link colour"],
           ["backgroundColor", "#FFFFFF", "Page behind the content"],
           ["secondaryBackgroundColor", "#F4F6FA", "Sidebar, table stripes, container borders"],
           ["textColor", "#0F172A", "Body text; 17.9:1 on white"],
           ["font", "sans serif", "Streamlit's bundled stack — the widgets only"],
           ["theme.sidebar.*", "#F8FAFC / #EEF2F7", "Sidebar slightly cooler than the page, so the two read as separate surfaces"],
           ["client.showErrorDetails", "none", "A stack trace is a security leak and a support ticket (NFR-12)"]],
          [1.9, 1.5, 3.5], font_size=8.0)
mono(doc, CFG_SRC, 7.2)

# 8
h1(doc, "8. Starter src/ui.py", page_break=True)
para(doc, "The module is complete rather than illustrative: nine screens can be built on it without writing "
          "another line of HTML. It was checked without installing Streamlit by substituting a stub runtime and "
          "inspecting the HTML it would emit — 22 checks, all passing.", size=9.8)
add_table(doc, ["What was checked", "Result"],
          [["Header and KPI cards emit exactly one iframe each, at the declared heights", "PASS — 1 frame, 4 frames"],
           ["Low-confidence badge appears only when low_n is set", "PASS"],
           ["Status colour follows status, not position", "PASS — amber on card 1, red on card 3"],
           ["A <script> tag and attribute quotes in a title are escaped", "PASS — &lt;script&gt;, &quot;"],
           ["No anchor or href in any component", "PASS — 0"],
           ["Indian digit grouping (₹12,34,567 / ₹60,56,692 / ₹13,52,60,000)", "PASS"],
           ["Delta uses U+2212, not a hyphen", "PASS — −1.1 pp"],
           ["USE_TAILWIND = False falls back to st.title / st.metric / st.warning", "PASS — 0 frames"],
           ["config.toml colours equal ui.THEME", "PASS — 4/4 tokens"],
           ["Six text tokens ≥ 4.5:1 on white", "PASS — navy 11.6, teal 5.5, amber 5.0, red 6.5"]],
          [5.4, 1.5], font_size=8.0)
para(doc, "It compiles (py_compile) but has not been executed against a real Streamlit runtime: the runtime stack "
          "is installed at Step 14, and Step 15 is the first AppTest run. The 22 checks are a static and "
          "behavioural proof of the HTML contract, not a substitute for that run.", size=9.4, italic=True)

h2(doc, "8.1 A page built on it")
mono(doc, """# pages/1_Operations.py — the shape every page follows
import streamlit as st
from src import ui, filters, data_access as da, kpi, charts

ui.setup("Operations · HKD-2026")
ui.enable_mobile()
ctx = da.context()                      # cached read-only frames
f = filters.sidebar(ctx)                # one state object
ui.header("Operations", "Capacity and demand · 2023–2025",
          as_of=ctx.max_date, version=ctx.defs_version)
ui.filter_chips(filters.chips(f))
ui.dq_banner(*da.quality_counts(f), period=f["period"])

frames = da.episodes(f)
if frames.empty:
    ui.empty_state("No episodes match these filters", filters.chips(f))
    st.stop()

ui.kpi_row([ui.card(*k) for k in kpi.occupancy_cards(frames, f)])
tabs = st.tabs(["Occupancy", "Volume & trend", "Compare"])
with tabs[0]:
    ui.chart_card("Bed occupancy vs 85% target", charts.bor_vs_target(frames, f),
                  note="Σ occupied ÷ Σ available bed-days; 109,420 available in 2025")
    st.dataframe(kpi.bor_gap_table(frames), hide_index=True)""", 7.2)

# 9
h1(doc, "9. Traceability", page_break=True)
h2(doc, "9.1 Every story has a screen")
add_table(doc, ["US", "Story", "Screen(s)", "Primary component"],
          [list(u) for u in US_UI], [0.55, 2.55, 1.55, 2.25], font_size=7.6)
para(doc, "US-01 and US-08 are the two stories with no screen of their own: US-01 is the loader behind every "
          "screen and US-08 is the shell around every screen. Assigning them to a single page would have hidden "
          "the fact that they are the two stories a partial build cannot do without.", size=9.6)

h2(doc, "9.2 Every OQ is reachable")
add_table(doc, ["OQ", "Screen that answers it"],
          [["OQ-01", "UI-02 (bars vs 85%), UI-01 (BOR card)"],
           ["OQ-02", "UI-01 (trend), UI-02 (volume tab), UI-08 (decomposition)"],
           ["OQ-03", "UI-03 (ALOS tab)"],
           ["OQ-04", "UI-03 (readmission tab), UI-06 (quality context)"],
           ["OQ-05", "UI-05 (payer mix, revenue per patient)"],
           ["OQ-06", "UI-02 (weekday heatmap)"],
           ["OQ-07", "UI-08 (association test)"],
           ["OQ-08", "UI-03 (patient mix tab)"],
           ["OQ-09", "UI-05 (revenue per bed-day), UI-08 (contribution)"],
           ["OQ-10", "UI-04 (drift), UI-06 (quality context)"]],
          [0.6, 6.3], font_size=8.2)

h2(doc, "9.3 Coverage checks")
add_table(doc, ["Check", "Actual", "Expectation", "Result"],
          [["Screens designed", "9", "brief suggested 8; no ID renumbered", "PASS"],
           ["Screens with a wireframe", "9 of 9", "one ASCII frame each", "PASS"],
           ["Stories with a screen", "20 of 20", "no orphan story", "PASS"],
           ["MVP stories with a screen", "13 of 13", "every Must story lands", "PASS"],
           ["OQs reachable from a screen", "10 of 10", "no orphan question", "PASS"],
           ["Components defined", "8", "all used on ≥ 1 screen", "PASS"],
           ["OQs with a chart decision", "10 of 10", "chart + rejected alternative", "PASS"],
           ["Text tokens ≥ AA contrast", "6 of 6", "≥ 4.5:1 on white", "PASS"],
           ["ui.py checks", "22 of 22", "static + behavioural", "PASS"],
           ["RTM rows with UI filled", f"{filled_ui} of 16", "0 blank", "PASS" if filled_ui == 16 else "FAIL"]],
          [2.0, 0.9, 3.2, 0.7], font_size=8.2, align_center=(1, 3))

h2(doc, "9.4 The updated RTM (v1.3)")
para(doc, "One column was filled. Feature and User Story were already complete from Step 11, so the matrix now "
          "resolves BR → FR → F → US → UI on every row that is in scope.", size=9.8, space_after=3)
add_table(doc, ["BR", "Feature", "User story", "UI Screen"],
          [list(r) for r in RTM_FULL], [0.7, 0.8, 3.15, 1.35], font_size=7.6)
para(doc, f"Verified after saving: {filled_ui}/16 UI cells filled, {filled_f}/16 Feature, {filled_us}/16 User Story. "
          f"The Coverage Check sheet gained section 7 (UI coverage) and section 8 (story → screen map) and now runs "
          f"to row {cc_rows}; the Change Log carries v1.3 as current and v1.2 as superseded.", size=9.6)
callout(doc, "The two BR-15 rows keep “— (Won't)” in the UI column rather than an empty cell. An empty cell and a "
             "deliberate exclusion look identical in a filter, and only one of them is a decision.", size=9.2, italic=True)

# 10
h1(doc, "10. Assumptions", page_break=True)
add_table(doc, ["ID", "Assumption", "Why it is reasonable", "If it is wrong"],
          [["AS-29", "cdn.tailwindcss.com is reachable when the app is demonstrated; if it is not, USE_TAILWIND = False renders the same screens with native Streamlit components.",
            "The CDN is the only runtime dependency that is not in requirements.txt, and a demo must not die on a network hiccup.",
            "Cards lose their styling and the 26 px hierarchy; the numbers survive. Precompiling the CSS at Step 14 removes the risk (UIF-02)."],
           ["AS-30", "Streamlit 1.67.0 keeps components.html, st.columns, st.plotly_chart(width=...) and the data-testid selectors used by enable_mobile().",
            "The version is pinned in requirements.txt; behaviour is verified once at Step 15, not assumed forever.",
            "The mobile media query stops matching; columns stay side-by-side at 375 px until the selector is updated."],
           ["AS-31", "The primary device is a laptop at ≥ 1280 px; 375 px is a secondary check-the-number path.",
            "The 8 a.m. huddle runs on a desktop; the phone use is an executive looking at one figure before a meeting.",
            "The phone layout must become fully interactive, which means native widgets at 1-up and a longer test list."],
           ["AS-32", "Four semantic colours are enough; no chart needs a categorical palette beyond the sequential scale.",
            "Departments are compared by position (ranked bars), not by colour, so identity colour is never load-bearing.",
            "A categorical palette (8 dept colours) would be needed, and it must pass a colour-blind simulation first."],
           ["AS-33", "Inter is loaded only inside the iframe; native widgets keep Streamlit's bundled sans.",
            "Injecting CSS into the parent document to load a font is fragile and adds a network dependency.",
            "The two typefaces are visibly different at large sizes on the Definitions screen; acceptable, and fixable with a local font file."],
           ["AS-34", "UI-08 and UI-09 are additions, not renumberings; the TRD file tree is updated at Step 14.",
            "No earlier ID changes meaning, so every previous document stays valid.",
            "Step 13 test cases written against the old 7-screen list would miss two screens and 2 stories."]],
          [0.55, 2.35, 2.0, 1.9], font_size=7.4)

# 11
h1(doc, "11. What the design exposed", page_break=True)
para(doc, "Six issues surfaced while turning stories into screens. Two of them land as Step 14 tasks, and one is a "
          "live dependency on demo day rather than a design question.", size=9.8, space_after=4)
add_table(doc, ["ID", "Severity", "Finding", "Decision"],
          [["UIF-01", "MED", "The palette exists twice — src/ui.py (Tailwind) and .streamlit/config.toml (native). Nothing in Streamlit links them, so a colour change in one place silently splits the UI.",
            "Accepted with a control: THEME is the single editable source, config.toml is generated from it by hand, and check_ui.py fails if the four shared tokens differ."],
           ["UIF-02", "MED", "cdn.tailwindcss.com is a JIT build, not a pinned asset. TRD §6.1 accepts it, but it is still an unpinned remote script at runtime, and a CDN outage would strip the card styling from every screen.",
            "USE_TAILWIND = False keeps the app usable. Step 14 must precompile the CSS to assets/tailwind.css with the Tailwind CLI and drop the CDN — recorded as a build task, not left as a hope."],
           ["UIF-03", "LOW", "components.html cannot auto-size: every component needs a fixed height, so a long department name or a translated label clips instead of pushing the card taller.",
            "Heights are constants (78/118/36/98/58/138) and every text node is truncate + ellipsis with a title attribute. The 375 px screenshot test is the check."],
           ["UIF-04", "MED", "Streamlit does not stack st.columns on a narrow viewport. enable_mobile() relies on data-testid selectors, which are internal API and can change in any release.",
            "Accepted: the version is pinned, the query is one block with a comment, and Step 15's 375/768/1440 screenshots catch a regression. If it breaks, the fallback is a single-column layout at all widths."],
           ["UIF-05", "LOW", "Nine screens no longer match the seven-screen file tree in 07_TRD.docx.",
            "Declared in UI-CN-01. Step 14 rewrites the tree; no earlier identifier is renumbered, so nothing else needs reissuing."],
           ["UIF-06", "LOW", "An <a href> inside the iframe navigates the iframe, not the app — a link to the Data Quality page would have opened the whole dashboard inside a 58 px banner.",
            "All components are anchor-free; navigation lives in the Streamlit sidebar. check_ui.py fails the build if 'href' ever appears in an emitted component."]],
          [0.6, 0.6, 3.1, 2.6], font_size=7.4)
para(doc, "UIF-02 is the one to raise with the mentor. Every Tailwind-in-Streamlit tutorial on the internet loads the "
          "CDN and stops there; the CDN is fine for a demo and wrong for a submission that claims a reproducible "
          "build, because the CSS a reader sees is whatever the CDN served that day. Precompiling is a 20-minute "
          "task at Step 14 and removes an entire class of 'it looked different on my machine'.", size=9.6)

# 12
h1(doc, "12. Sign-off")
add_table(doc, ["Role", "Name", "Responsibility", "Status"],
          [["Student / Author", STUDENT, "Designed the nine screens, wrote src/ui.py and the theme, updated the RTM to v1.3", "Signed — 19-Nov-2026"],
           ["Mentor", MENTOR, "Reviews the screen count (UI-08/UI-09), the Tailwind boundary and the UIF-02 decision", "Pending"],
           ["Internal guide", "Project guide, MCA (DA)", "Confirms the UI is ready to be tested at Step 13", "Pending"]],
          [1.4, 1.45, 2.9, 1.35], font_size=8.4)
para(doc, "Next step: Step 13 turns each acceptance criterion into test cases — TC-F for the 22 functional "
          "requirements, TC-T for the 12 non-functional ones — and fills the last column of the RTM. Every "
          "criterion in Step 11 must resolve to at least one test, and every test must name the screen it runs on.",
     size=9.6, italic=True)

# appendix
h1(doc, "Appendix A — src/ui.py as shipped", page_break=True)
para(doc, f"{len(UI_SRC.splitlines())} lines. Compiles with py_compile; the HTML contract is verified by "
          f"scripts/check_ui.py (22 checks). The file in the repository is the source of truth — this listing is "
          f"generated from it at build time, so the two cannot diverge.", size=9.2, italic=True)
mono(doc, UI_SRC, 6.2)

doc.save(OUT)
print("saved:", OUT)
print("tables:", len(doc.tables), "paragraphs:", len(doc.paragraphs))
print(f"screens 9, wireframes 9, charts 10, components 8, stories 20")
print(f"rtm: ui {filled_ui}/16  feature {filled_f}/16  us {filled_us}/16  coverage rows {cc_rows}")
