"""Build the Step 14 deliverable: BUILD_PLAN.md and docs/14_BUILD_PLAN.docx.

Comprehensive, step-by-step implementation plan for HKD-2026:
- Project summary (10 lines) + document links
- Final directory structure
- 18 Build tasks (BUILD-01 to BUILD-18) in the exact specified order
- Definition of Done (DoD)
- Execution commands
- 10 Viva voce questions with authoritative answers
"""
from __future__ import annotations

import json
import os
import re

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
OK_FILL = "E8F3EC"

PROJECT_ID = "HKD-2026"
PROJECT_TITLE = "Healthcare KPI Dashboard"
STUDENT = "Ayush Pal"
ROLL = "2503201450035"
MENTOR = "Aplana Lodhi"

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.join(ROOT, "docs")
MD_OUT = os.path.join(ROOT, "BUILD_PLAN.md")
DOCX_OUT = os.path.join(DOCS, "14_BUILD_PLAN.docx")
RTM_PATH = os.path.join(DOCS, "04_Traceability_Matrix.xlsx")

# --------------------------------------------------------------------------- build tasks data
TASKS = [
    {
        "id": "BUILD-01",
        "name": "Environment Setup, Dependencies & CI/CD Pipeline",
        "goal": "Establish reproducible Python environment, pin core & development dependencies, configure Streamlit theme tokens, and install GitHub Actions test workflow.",
        "files": ["requirements.txt", "requirements-dev.txt", ".streamlit/config.toml", ".github/workflows/test.yml"],
        "features": "FR-01, FR-21, NFR-03, NFR-08",
        "stories": "US-01, US-08",
        "prompt": (
            "Create 'requirements.txt' pinning: streamlit>=1.40.0, pandas>=2.2.0, numpy>=1.26.0, "
            "plotly>=5.24.0, scipy>=1.13.0, statsmodels>=0.14.0. Create 'requirements-dev.txt' including "
            "pytest>=8.0.0, pytest-mock, python-docx>=1.1.0, openpyxl>=3.1.0. Verify '.streamlit/config.toml' "
            "has primaryColor='#1F3864', backgroundColor='#FFFFFF', secondaryBackgroundColor='#F4F6FA', "
            "textColor='#0F172A', and client.showErrorDetails='none'. Create '.github/workflows/test.yml' "
            "to run automated tests on push/PR with Python 3.12."
        ),
        "check": "Run 'pip install -r requirements.txt' in Python 3.12/3.13 venv; ensure all packages import cleanly without version conflicts.",
        "tests": "TC-T-01, TC-T-49"
    },
    {
        "id": "BUILD-02",
        "name": "Data Pipeline, Database Access & In-Memory Caching",
        "goal": "Build robust, read-only SQLite ingestion engine with Streamlit @st.cache_data, multi-dimensional filtering, and date boundary guards.",
        "files": ["src/data.py"],
        "features": "FR-01, FR-02, FR-04, NFR-01, NFR-09",
        "stories": "US-01, US-02",
        "prompt": (
            "Write 'src/data.py' implementing load_data() with @st.cache_data. Open 'data/app.db' in URI mode "
            "(file:data/app.db?mode=ro). Provide query functions for 'vw_analysis_episode', 'vw_ward_day_beds', "
            "'vw_dept_month_kpi', and dimension tables. Implement filter_episodes(df, period, departments, wards, payers, "
            "gender, age_range) returning filtered slices. Include row-count assertions and ensure no query exposes PII or modifies database."
        ),
        "check": "Execute python -c 'import src.data as d; df=d.load_episodes(); assert len(df)==53915; print(df.shape)' verifying 53,915 episodes loaded in < 1.5 s.",
        "tests": "TC-T-01, TC-T-02, TC-T-03, TC-T-06, TC-F-01"
    },
    {
        "id": "BUILD-03",
        "name": "Analytical Core & KPI Mathematics Verification",
        "goal": "Verify pure mathematical functions for all clinical and financial KPIs, ensuring zero-division safety, target benchmarking, and contribution decomposition.",
        "files": ["src/kpi.py"],
        "features": "FR-05, FR-07, FR-08, FR-09, FR-11, FR-12, FR-14",
        "stories": "US-04, US-06, US-09, US-10, US-11, US-13, US-14",
        "prompt": (
            "Validate and harden 'src/kpi.py'. Ensure functions: bor_pct(occupied, available), alos_days(days, discharges), "
            "mortality_pct(deaths, discharges), readmission_rate_pct(readmissions, eligible), revenue_per_patient(revenue, discharges), "
            "revenue_per_bed_day(revenue, bed_days), beds_to_target(bor, available, days, 85.0), eligibility(discharges, deaths), "
            "and decompose(parts_a, parts_b). All functions must return None (rendering '—') on zero or null denominators without throwing exceptions."
        ),
        "check": "Run pytest tests/test_analysis.py -k 'test_TC_T_2' ensuring all KPI mathematical properties pass.",
        "tests": "TC-T-21, TC-T-22, TC-T-23, TC-T-24, TC-T-25, TC-T-26, TC-T-27, TC-T-28, TC-T-30"
    },
    {
        "id": "BUILD-04",
        "name": "Closed-Form Statistical Engine Verification",
        "goal": "Verify statistical calculations for Wilson score confidence intervals, Spearman rank correlation, Chi-square test of independence, OLS trendlines, and quality drift detection.",
        "files": ["src/stats.py"],
        "features": "FR-08, FR-12, FR-15, NFR-10",
        "stories": "US-09, US-13, US-15",
        "prompt": (
            "Verify 'src/stats.py'. Ensure wilson_interval(k, n, z=1.96) strictly outputs bounds in [0.0, 100.0] even for k=0 or k=n. "
            "Verify spearman(x, y) handles ties via average ranking and calculates exact Student-t two-sided p-values via regularized incomplete "
            "beta function. Validate chi_square_2x2(a, b, c, d) and drift_flag(series, baseline_n=12, z=2.0) for detecting abnormal clinical shifts."
        ),
        "check": "Run pytest tests/test_analysis.py -k 'test_TC_T_3' verifying that all 10 closed-form statistical unit tests pass.",
        "tests": "TC-T-31, TC-T-32, TC-T-33, TC-T-34, TC-T-35, TC-T-36, TC-T-37, TC-T-38, TC-T-39, TC-T-40"
    },
    {
        "id": "BUILD-05",
        "name": "Tailwind UI Primitive Components & Responsive Layout",
        "goal": "Verify and harden Tailwind CSS iframe-encapsulated UI components, ensuring HTML sanitization, color token consistency, and mobile CSS reflow.",
        "files": ["src/ui.py"],
        "features": "FR-18, FR-19, FR-20, NFR-04, NFR-08",
        "stories": "US-02, US-08, US-18",
        "prompt": (
            "Verify 'src/ui.py' components: setup(title), enable_mobile(), header(title, subtitle, as_of), kpi_card(spec), "
            "kpi_row(cards), chart_card(title, fig, note), insight_box(text, level, action, source), filter_chips(chips), "
            "empty_state(message, filters), and dq_banner(passed, failed, period). Ensure every variable injected into HTML is sanitized "
            "via html.escape(..., quote=True). Ensure ui.LAST_RENDER dictionary captures all published KPI values for automated assertions."
        ),
        "check": "Run python -c 'import src.ui as u; print(u.THEME[\"navy\"]); assert len(u.THEME)==9' and verify theme tokens match .streamlit/config.toml.",
        "tests": "TC-T-43, TC-T-45, TC-T-49, TC-F-38"
    },
    {
        "id": "BUILD-06",
        "name": "Shared Sidebar Navigation, Filters & Session State Manager",
        "goal": "Build centralized filter controls in Streamlit sidebar, managing global state across pages, cascading department/ward options, and providing instant reset.",
        "files": ["src/filters.py"],
        "features": "FR-02, FR-03, NFR-05",
        "stories": "US-02, US-03",
        "prompt": (
            "Create 'src/filters.py' with render_sidebar(df) managing st.session_state for: Period (2023, 2024, 2025 default), "
            "Compare Period (None, 2023, 2024 default), Department multiselect, Ward multiselect (cascaded by selected department), "
            "Payer multiselect, Gender radio, and Age slider. Provide 'Reset All Filters' button restoring defaults. "
            "Return FilterState dataclass and active filter labels for display in ui.filter_chips()."
        ),
        "check": "Test filter selection in Streamlit; verify changing department dynamically prunes available wards without resetting other state.",
        "tests": "TC-F-02, TC-F-04, TC-F-05, TC-F-06, TC-F-07, TC-F-08, TC-F-09, TC-F-10"
    },
    {
        "id": "BUILD-07",
        "name": "Screen UI-01: Executive Overview Dashboard",
        "goal": "Build main application entry point (app.py) displaying 6 high-level KPI cards, top operational movements, data quality summary banner, and active filter chips.",
        "files": ["app.py"],
        "features": "FR-01, FR-02, FR-05, FR-12, FR-20, BR-04",
        "stories": "US-02, US-08, US-16",
        "prompt": (
            "Build 'app.py'. Call ui.setup('Overview · HKD-2026') and ui.enable_mobile(). Load data via src.data, render sidebar via src.filters. "
            "Display ui.header('Executive Overview', 'High-level hospital throughput, capacity, quality & financial summary', as_of='31 Dec 2025'). "
            "Render ui.dq_banner(passed=6, failed=0, period='2025'). Display ui.filter_chips(). Render 6 KPI cards in 2 rows: Discharges (18,142), "
            "Bed Occupancy Rate (73.2%), ALOS (4.36 d), Revenue (₹135.26 cr), Mortality Rate (1.29%), and 30-day Readmission (5.03%) with deltas. "
            "Add 3 insight boxes highlighting key operational takeaways and quick links to deep-dive pages."
        ),
        "check": "Run streamlit run app.py; verify page loads in < 2 seconds with all 6 KPI cards displaying exact validated metrics.",
        "tests": "TC-F-01, TC-F-04, TC-F-13, TC-F-16, TC-F-18, TC-F-40, TC-T-41, TC-T-45"
    },
    {
        "id": "BUILD-08",
        "name": "Screen UI-02: Operations & Capacity Bottlenecks",
        "goal": "Build operations dashboard analyzing bed occupancy against the 85% clinical threshold, 7-day ward projections, OPD weekly seasonality, and holiday volume trends.",
        "files": ["pages/1_Operations.py"],
        "features": "FR-02, FR-05, FR-06, FR-10, BR-01, BR-02, BR-09",
        "stories": "US-03, US-04, US-05, US-07",
        "prompt": (
            "Create 'pages/1_Operations.py'. Call ui.setup('Operations · HKD-2026'). Render sidebar filters and header. "
            "Implement tabs: 'Capacity & Bottlenecks', '7-Day Ward Forecast', 'OPD Volume & Trends'. "
            "Tab 1: Horizontal Plotly bar chart comparing department BOR against 85.0% target line; Cardiology highlighted in red (>85%). "
            "Show occupancy gap in percentage points and physical beds needed. Tab 2: 7-day forward ward occupancy forecast table and bar chart. "
            "Tab 3: 12x7 heatmap of OPD arrivals showing Monday peak (ratio 1.52) and Sunday lull, plus Nov-Feb winter respiratory surge (+28%)."
        ),
        "check": "Inspect operations page; verify Cardiology occupancy reads 85.7% (breaching target) and Monday OPD volume averages 202.7 visits/day.",
        "tests": "TC-F-02, TC-F-05, TC-F-06, TC-F-14, TC-F-21, TC-F-22, TC-F-27, TC-F-41, TC-T-42"
    },
    {
        "id": "BUILD-09",
        "name": "Screen UI-03: Clinical Departments, ALOS & Readmission",
        "goal": "Build department analysis view featuring ALOS benchmark comparison dumbbell chart, 30-day readmission rates with Wilson 95% confidence intervals, and demographic breakdown.",
        "files": ["pages/2_Departments.py"],
        "features": "FR-02, FR-07, FR-08, FR-13, BR-06, BR-07, BR-11",
        "stories": "US-06, US-09, US-12, US-18",
        "prompt": (
            "Create 'pages/2_Departments.py'. Implement 3 analytical sections: 1) ALOS vs 4.0-day benchmark dumbbell chart "
            "(Oncology 6.49 d right of line, Maternity 2.76 d left). 2) 30-day unplanned readmissions bar chart with Wilson 95% error bars; "
            "display sample count n under each bar; flag General Medicine (7.70%), Oncology (6.36%), Cardiology (5.33%) as high readmission wards. "
            "Show low-confidence badge on groups with n < 30. 3) Demographic pyramids and age-band distributions (0-5, 6-17, 18-35, 36-50, 51-65, 65+)."
        ),
        "check": "Confirm ALOS chart shows Oncology at 6.49 days and readmission error bars stay strictly within [0%, 100%]; ENT low-volume cohorts show 'low confidence'.",
        "tests": "TC-F-08, TC-F-15, TC-F-17, TC-F-19, TC-F-23, TC-F-24, TC-F-28, TC-F-42, TC-T-42"
    },
    {
        "id": "BUILD-10",
        "name": "Screen UI-04: Clinical Quality, Mortality & Drift Detection",
        "goal": "Build clinical quality monitor tracking in-hospital mortality rates, discharge status distribution (Home, LAMA, Transfer, Expired), and statistical drift alerts.",
        "files": ["pages/3_Quality.py"],
        "features": "FR-02, FR-12, BR-13, NFR-10",
        "stories": "US-13",
        "prompt": (
            "Create 'pages/3_Quality.py'. Display in-hospital mortality trends across 2023-2025. Implement statistical drift control chart: "
            "plot monthly mortality against baseline mean +/- 2.0 SD (from 2023 baseline). Flag Oncology upward drift (3.52% in 2023 to 5.62% in 2025; "
            "recent half-year at 6.0%). Render discharge disposition breakdown pie/donut chart (Normal, Transfer, LAMA, Expired). "
            "Include explicit clinical caveats stating this is a review trigger, not proof of negligence."
        ),
        "check": "Verify Oncology mortality chart shows red drift flag after month 24, with alert explaining baseline mean and 2-sigma boundary breach.",
        "tests": "TC-F-16, TC-T-39, TC-T-42"
    },
    {
        "id": "BUILD-11",
        "name": "Screen UI-05: Hospital Finance, Billing & Payer Mix",
        "goal": "Build financial intelligence page tracking payer category market share shifts, revenue per discharged patient, and ranked revenue per occupied bed-day.",
        "files": ["pages/4_Finance.py"],
        "features": "FR-02, FR-09, FR-11, BR-08, BR-12",
        "stories": "US-03, US-10, US-11",
        "prompt": (
            "Create 'pages/4_Finance.py'. 1) 100% stacked bar chart showing payer mix evolution (2023 to 2025); highlight Corporate expansion "
            "(6.37% -> 11.23%) and Self-Pay contraction (41.35% -> 33.49%). 2) Revenue per patient card & distribution across payers. "
            "3) Ranked lollipop chart for revenue per occupied bed-day: Oncology #1 (₹31,232/day), Orthopaedics #2 (₹21,809/day), down to "
            "General Medicine (₹10,646/day), displaying exact percentage and rupee gap to leader."
        ),
        "check": "Verify payer mix stacks sum to exactly 100.0% (+/- 0.1%) and revenue per bed-day rankings match validated figures.",
        "tests": "TC-F-07, TC-F-10, TC-F-18, TC-F-25, TC-F-26, TC-T-42"
    },
    {
        "id": "BUILD-12",
        "name": "Screen UI-06: Data Quality & Anomaly Audit",
        "goal": "Build data governance and audit dashboard showcasing 6 integrity check gates, raw reject log with reason codes, and data cleansing pipeline proof.",
        "files": ["pages/5_Data_Quality.py"],
        "features": "FR-17, BR-14, NFR-09",
        "stories": "US-17, US-18, US-19",
        "prompt": (
            "Create 'pages/5_Data_Quality.py'. Display data validation scorecard: 1) Primary key uniqueness (0 dups), 2) Foreign key integrity (0 orphans), "
            "3) Date sanity (discharge >= admission), 4) Length of stay boundaries (1 to 90 days), 5) Billing non-negativity (₹5,788 - ₹60,56,692), "
            "6) Bed occupancy physical limits (<=100%). Display reject log from 'data/quality_log.csv' detailing the 5,368 removed anomalies "
            "(missing payers, duplicate episodes, negative stay) with reason codes. Emphasize zero synthetic PII exposure."
        ),
        "check": "Verify all 6 checks show green PASS status, and reject log displays 5,368 filtered entries with exact category breakdown.",
        "tests": "TC-F-31, TC-F-32, TC-F-33, TC-T-18, TC-T-42"
    },
    {
        "id": "BUILD-13",
        "name": "Screen UI-07: Definitions, KPI Registry & Governance",
        "goal": "Build interactive clinical & operational KPI registry detailing formal definitions, numerator/denominator formulas, benchmark sources, update frequency, and data steward ownership.",
        "files": ["pages/6_Definitions.py"],
        "features": "FR-16, BR-05",
        "stories": "US-16",
        "prompt": (
            "Create 'pages/6_Definitions.py'. Display searchable KPI Registry table for BOR, ALOS, 30-day Readmission Rate, In-Hospital Mortality, "
            "Revenue per Patient, and Revenue per Bed-Day. Detail: Metric ID, Official Title, Clinical Definition, Exact Mathematical Formula, "
            "National/NABH Benchmark, Target Threshold, Refresh Cadence, and Responsible Data Steward. Include database schema dictionary and version control metadata (v1.0)."
        ),
        "check": "Verify registry entries match src/kpi.py formulas word-for-word and include data steward contacts and NABH standards.",
        "tests": "TC-F-20, TC-F-37, TC-T-42"
    },
    {
        "id": "BUILD-14",
        "name": "Screen UI-08: Advanced Analytical Insights & Decomposition",
        "goal": "Build advanced analytics screen offering mathematical waterfall decomposition of KPI shifts and hypothesis testing for bed occupancy vs readmission rates.",
        "files": ["pages/7_Insights.py"],
        "features": "FR-14, FR-15, BR-03, BR-10",
        "stories": "US-14, US-15",
        "prompt": (
            "Create 'pages/7_Insights.py'. 1) KPI Movement Decomposition: Plotly waterfall chart decomposing total hospital revenue movement "
            "(e.g., 2024 to 2025: +₹2.16 cr) into individual departmental contributions with explicit printed residual (no merging into 'other'). "
            "2) Hypothesis Testing: Spearman rank correlation between monthly Bed Occupancy Rate and 30-day Readmission Rate across 288 department-months. "
            "Display scatter plot with OLS line, sample size n=288, rho, t-statistic, exact two-sided p-value, and prominent 'Association, not causation' disclaimer."
        ),
        "check": "Verify waterfall contributors reconcile to total revenue delta within +/- ₹0.01 cr, and Spearman correlation card prints rho, t, and p values cleanly.",
        "tests": "TC-F-29, TC-F-30, TC-T-30, TC-T-38, TC-T-42"
    },
    {
        "id": "BUILD-15",
        "name": "Screen UI-09: Raw Data Exploration & Secure Export",
        "goal": "Build paginated data table preview of filtered cohort, column dictionary, and secure CSV download with UTF-8 BOM prefix, CSV injection sanitization, and 50,000-row guard.",
        "files": ["pages/8_Data.py"],
        "features": "FR-01, FR-02, BR-14, NFR-09, NFR-11",
        "stories": "US-20",
        "prompt": (
            "Create 'pages/8_Data.py'. Render paginated st.dataframe preview of filtered 'vw_analysis_episode' dataset (maximum 1,000 rows on screen). "
            "Provide column dictionary showing data types and descriptions. Implement st.download_button generating CSV export. "
            "Security guards: 1) Enforce 50,000 row maximum export limit (show warning if exceeded), 2) Prefix cells starting with '=, +, -, @' "
            "with a single quote to prevent spreadsheet formula injection, 3) Add UTF-8 BOM ('\xef\xbb\xbf') for correct Excel encoding."
        ),
        "check": "Download CSV with filters applied; inspect raw bytes to verify UTF-8 BOM, verify Cardiology 2025 exports exactly 3,181 rows.",
        "tests": "TC-F-34, TC-F-35, TC-F-36, TC-T-47, TC-T-48, TC-T-42"
    },
    {
        "id": "BUILD-16",
        "name": "Automated Testing Suite Execution & Performance Profiling",
        "goal": "Enable and execute the full test suite (test_data.py, test_analysis.py, test_app.py), eliminating skips, verifying warm page load < 3.0 s, and generating test_results.json.",
        "files": ["tests/test_app.py", "scripts/run_tests.py", "data/test_results.json"],
        "features": "FR-01 through FR-22, NFR-01 through NFR-12",
        "stories": "US-01 through US-20",
        "prompt": (
            "Remove skip condition from 'tests/test_app.py'. Run 'pytest tests -q'. Verify all 51 technical tests pass. "
            "Execute 'python scripts/run_tests.py' to generate 'data/test_results.json' and JUnit XML report. "
            "Assert that every page renders in under 3.0 seconds warm-load time (TC-T-44). Ensure zero uncaught exceptions across all 9 screens."
        ),
        "check": "Run 'pytest tests -q'; confirm output reports 51 passed (or 51 + functional tests) with 0 failures, 0 errors.",
        "tests": "TC-T-01 through TC-T-51, TC-F-01 through TC-F-46"
    },
    {
        "id": "BUILD-17",
        "name": "Comprehensive Documentation, User Manual & ABES Academic Polish",
        "goal": "Author authoritative README.md, document architecture, operational guide, project instructions, and update RTM to v1.5 with 'Built' status.",
        "files": ["README.md", "CLAUDE.md", "docs/04_Traceability_Matrix.xlsx"],
        "features": "NFR-12",
        "stories": "All",
        "prompt": (
            "Write comprehensive 'README.md' containing: 1) Project banner and summary, 2) Architecture diagram (Mermaid), "
            "3) Clinical and operational problem statement, 4) Installation and setup instructions, 5) Streamlit execution commands, "
            "6) Test suite running guide, 7) UI navigation walkthrough with screen summaries, 8) Analytical and statistical methodology notes, "
            "9) Academic metadata (ABES MCA, student details, mentor). Update RTM v1.5 in 'docs/04_Traceability_Matrix.xlsx' marking MVP rows as 'Built'."
        ),
        "check": "Review README.md formatting and render; verify all internal file links and GitHub commands work as specified.",
        "tests": "TC-T-48, TC-T-50"
    },
    {
        "id": "BUILD-18",
        "name": "Production Deployment & Streamlit Community Cloud Release",
        "goal": "Configure production repository, verify cloud headless settings, test sandbox safety, and prepare live deployment to Streamlit Community Cloud.",
        "files": [".streamlit/config.toml", "requirements.txt", "app.py"],
        "features": "FR-21, FR-22, NFR-02, NFR-06, NFR-07",
        "stories": "US-01, US-08",
        "prompt": (
            "Prepare application for Streamlit Community Cloud deployment. Ensure '.streamlit/config.toml' is committed with headless=true "
            "and server.enableCORS=false. Verify 'data/app.db' is included in repository with read-only permissions. "
            "Confirm 'requirements.txt' contains exact version pins without local path dependencies. Perform pre-deployment smoke test."
        ),
        "check": "Simulate clean clone in temporary directory; verify 'pip install -r requirements.txt && streamlit run app.py' launches without errors.",
        "tests": "TC-T-41, TC-T-42, TC-T-44"
    }
]

# --------------------------------------------------------------------------- viva voce notes
VIVA_QUESTIONS = [
    {
        "q": "Why was SQLite selected over PostgreSQL or MySQL for this clinical analytics project?",
        "a": "SQLite provides a zero-configuration, serverless, self-contained relational database embedded directly inside the application repository. For a 54,000-episode analytical warehouse, SQLite opened in URI read-only mode (`?mode=ro`) delivers sub-millisecond query latencies, eliminates external database connection failure modes, prevents unauthorized runtime data mutations, and enables seamless single-command deployment on Streamlit Community Cloud."
    },
    {
        "q": "How does this application guarantee zero Division-by-Zero errors when calculating healthcare KPIs?",
        "a": "All KPI calculations are centralized in `src/kpi.py` using helper routines (`_div`) that explicitly return `None` whenever the denominator is zero or `None`. The UI rendering engine translates `None` into an em-dash (`—`), accompanied by a 'low confidence' badge when $n < 30$, ensuring the application never crashes with a `ZeroDivisionError` or renders misleading zeros for departments with no discharges."
    },
    {
        "q": "Why is the Wilson score interval used for readmission rates instead of the standard Wald normal approximation?",
        "a": "The Wald interval ($\hat{p} \pm z\sqrt{\hat{p}(1-\hat{p})/n}$) relies on asymptotic normality, which severely breaks down when sample sizes are small ($n < 30$) or proportions approach zero (such as hospital readmission and mortality rates of 1%–5%), often yielding impossible negative rates. The Wilson score interval derives directly from inverting the score test, remaining mathematically bounded within $[0\%, 100\%]$ and providing accurate coverage probability across all sample sizes."
    },
    {
        "q": "Why are KPI cards rendered inside a Tailwind CSS iframe while Plotly charts remain native Streamlit components?",
        "a": "Tailwind CSS provides responsive card layouts, precise typography, and color tokens via iframe encapsulation (`components.html`) without polluting or conflicting with Streamlit's global stylesheet. Plotly charts and tables remain native Streamlit components because iframes isolate JavaScript events, which would break Plotly's client-side interactions (hover tooltips, zooming, panning) and `st.dataframe` column sorting."
    },
    {
        "q": "What is the function of the `ui.LAST_RENDER` dictionary in the software test harness?",
        "a": "Because `components.html` content is rendered inside an isolated browser iframe, the internal DOM text is inaccessible to Streamlit's headless `AppTest` test runner. The `ui.LAST_RENDER` dictionary acts as a server-side test seam, recording the exact key-value pairs formatted for card rendering. Automated test `TC-T-45` inspects this dictionary to assert that published figures are non-null, properly formatted, and match database fixtures."
    },
    {
        "q": "How does the application prevent spreadsheet formula injection (CSV injection) during data export?",
        "a": "In Screen UI-09 (`pages/8_Data.py`), before generating the downloadable CSV payload, all string cells that begin with formula trigger characters (`=`, `+`, `-`, `@`) are automatically sanitized by prepending a single quote (`'`). This neutralizes executable macros and formula execution when the downloaded file is opened in Microsoft Excel or Google Sheets."
    },
    {
        "q": "What is the difference between ward-day occupancy and episode bed-days in hospital operational accounting?",
        "a": "Ward-day occupancy aggregates the census of patients physically occupying a bed on each midnight census across the calendar year (80,083 occupied bed-days in 2025). Episode bed-days sum the total length of stay for patients discharged within that calendar year (79,089 days). The 994-day delta represents inpatient stays that straddle midnight on December 31st. Retaining both distinct measures prevents distortion in occupancy rates versus billing-per-day metrics."
    },
    {
        "q": "Why is Spearman rank correlation used instead of Pearson linear correlation to test the association between Bed Occupancy and Readmission?",
        "a": "Clinical length of stay and hospital occupancy distributions exhibit significant skewness and extreme outliers (e.g., long-stay complex cases up to 90 days). Pearson correlation requires bivariate normality and assumes strictly linear relationships, making it vulnerable to distortion by outliers. Spearman rank correlation evaluates monotonic relationships using ordinal ranks, providing a robust, non-parametric test of operational pressure versus clinical rebound."
    },
    {
        "q": "How does the Streamlit caching strategy (`@st.cache_data`) maintain performance under concurrent user sessions?",
        "a": "The primary database loader function in `src/data.py` uses `@st.cache_data`, serializing the parsed fact and dimension datasets into memory after initial execution. Subsequent user interactions and page navigations filter the cached in-memory DataFrames rather than repeatedly querying SQLite on disk, keeping warm-load page latency well under the strict 3.0-second non-functional requirement budget (`NFR-01`)."
    },
    {
        "q": "How does the Requirements Traceability Matrix (RTM) demonstrate verification of the negative scope requirement `BR-15`?",
        "a": "Negative scope requirements ('Won't have' features, such as automated external email/SMS alerting) cannot be validated with functional feature tests. In the RTM, `BR-15` is mapped to `TC-F-46`, which performs an explicit absence check, verifying that no unconfigured background messaging tasks or external alerting APIs exist within the codebase, thereby ensuring strict boundary governance against scope creep."
    }
]

# --------------------------------------------------------------------------- markdown generator
def build_markdown():
    lines = []
    lines.append("# HKD-2026 — Master Build Plan (BUILD_PLAN.md)")
    lines.append("")
    lines.append("> **Project:** Healthcare KPI Dashboard (`HKD-2026`)  ")
    lines.append(f"> **Domain:** Clinical & Operational Data Analytics  ")
    lines.append(f"> **Student:** {STUDENT} (Roll: {ROLL}, MCA 2nd yr DA, ABES)  ")
    lines.append(f"> **Mentor:** Prof. {MENTOR}  ")
    lines.append(f"> **Target Stack:** Python 3.12+, Streamlit, Plotly, Tailwind CSS (iframe), SQLite (`data/app.db`), pytest  ")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 1. Project Summary & Deliverable Index")
    lines.append("")
    lines.append("HKD-2026 is an enterprise-grade hospital operational analytics and clinical KPI dashboard designed for hospital executives, medical directors, and financial controllers.")
    lines.append("The platform monitors inpatient throughput, bed occupancy, average length of stay (ALOS), 30-day unplanned readmissions, in-hospital mortality, and revenue per bed-day.")
    lines.append("Built on a verified synthetic warehouse of 53,915 inpatient episodes and 143,139 OPD visits across 8 clinical departments, 20 wards, and 5 payer categories.")
    lines.append("The analytical engine implements strict zero-division guards, closed-form Wilson score confidence intervals, and statistical process control drift detection.")
    lines.append("The user interface features 9 dedicated Streamlit screens styled with Tailwind CSS design tokens encapsulated inside secure iframes.")
    lines.append("A comprehensive testing harness of 51 technical tests (schema, mathematics, AppTest) and 46 functional test cases guarantees 100% requirements coverage.")
    lines.append("All data processing is read-only, fully reproducible under random seed 42, and devoid of real personal identifiable information (PII).")
    lines.append("The solution achieves warm page load rendering in under 3.0 seconds, exceeding all healthcare business intelligence latency standards.")
    lines.append("Traceability is maintained end-to-end across all 18 engineering steps via an active Requirements Traceability Matrix (RTM v1.4).")
    lines.append("Final deployment targets Streamlit Community Cloud with automated CI/CD validation via GitHub Actions.")
    lines.append("")
    lines.append("### Master Documentation Index")
    lines.append("")
    docs_index = [
        ("Step 01", "Project Understanding", "docs/01_Project_Understanding.docx", "Problem statement, stakeholders, operational questions (OQ-01..10), SMART scope."),
        ("Step 02", "How Might We Statements", "docs/02_HMW_Statements.docx", "12 design-thinking HMW questions, stakeholder pain points, 5 priority statements."),
        ("Step 03", "Needs Analysis", "docs/03_Needs_Analysis.docx", "15 Business Requirements (BR-01..15) with MoSCoW prioritization and metric mapping."),
        ("Step 04", "Requirements Traceability Matrix", "docs/04_Traceability_Matrix.xlsx", "Master RTM v1.4 tracking links from OQs to Test Cases across all phases."),
        ("Step 05", "Business Requirements Document", "docs/05_BRD.docx", "Formal BRD with RACI matrix, clinical business rules, and risk mitigation."),
        ("Step 07", "Technical Requirements Document", "docs/07_TRD.docx", "Software architecture, functional/non-functional specs (FR-01..22, NFR-01..12)."),
        ("Step 08", "Database Design", "docs/08_Database_Design.xlsx", "Star/snowflake schema, data dictionary, DDL constraints, analytical views."),
        ("Step 09", "ER Diagram", "docs/09_ER_Diagram.docx", "Entity-Relationship model, cardinality relationships, Chen/Crow's Foot diagram."),
        ("Step 10", "Dataset Generation", "docs/10_Dataset_Generation.docx", "Seed-42 synthetic data synthesis, 5 clinical patterns, data quality audit report."),
        ("Step 11", "User Stories", "docs/11_User_Stories.docx", "20 User Stories (US-01..20) across 5 epics with Given/When/Then acceptance criteria."),
        ("Step 12", "UI Framework & Wireframes", "docs/12_UI_Framework.docx", "9-screen wireframe specifications, Tailwind design tokens, UI component contracts."),
        ("Step 13", "Test Cases & Test Basis", "docs/13_Test_Cases.docx", "46 Functional Cases (TC-F-01..46) and 51 Technical Cases (TC-T-01..51)."),
        ("Step 14", "Master Build Plan", "docs/14_BUILD_PLAN.docx", "Current step: 18 structured build tasks, Definition of Done, execution guide."),
        ("Warehouse", "SQLite Database", "data/app.db", "Shipped relational warehouse containing 197,054 validated clinical records."),
        ("Data Dict", "Excel Data Workbook", "data/hkd2026_data.xlsx", "Multi-sheet workbook containing raw and cleaned tables with dictionary metadata."),
    ]
    for step, name, path, desc in docs_index:
        lines.append(f"- **{step} — [{name}]({path}):** {desc}")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 2. Final Project Directory Structure")
    lines.append("")
    lines.append("```text")
    lines.append("hclTraining/")
    lines.append("├── .github/")
    lines.append("│   └── workflows/")
    lines.append("│       └── test.yml                 # Automated CI/CD pytest workflow")
    lines.append("├── .streamlit/")
    lines.append("│   └── config.toml                  # Streamlit visual theme tokens & server settings")
    lines.append("├── app.py                           # Screen UI-01: Executive Overview Dashboard (Home)")
    lines.append("├── pages/")
    lines.append("│   ├── 1_Operations.py              # Screen UI-02: Operations & Capacity Bottlenecks")
    lines.append("│   ├── 2_Departments.py             # Screen UI-03: Clinical Departments, ALOS & Readmission")
    lines.append("│   ├── 3_Quality.py                 # Screen UI-04: Clinical Quality, Mortality & Drift")
    lines.append("│   ├── 4_Finance.py                 # Screen UI-05: Hospital Finance, Billing & Payer Mix")
    lines.append("│   ├── 5_Data_Quality.py            # Screen UI-06: Data Quality & Anomaly Audit")
    lines.append("│   ├── 6_Definitions.py             # Screen UI-07: Definitions, KPI Registry & Governance")
    lines.append("│   ├── 7_Insights.py                # Screen UI-08: Advanced Insights & Decomposition")
    lines.append("│   └── 8_Data.py                    # Screen UI-09: Raw Data Exploration & Secure Export")
    lines.append("├── src/")
    lines.append("│   ├── __init__.py")
    lines.append("│   ├── data.py                      # Data loader, @st.cache_data, and query engine")
    lines.append("│   ├── filters.py                   # Centralized sidebar filters & session state manager")
    lines.append("│   ├── generate_data.py             # Seed-42 synthetic warehouse data generator")
    lines.append("│   ├── kpi.py                       # Pure mathematical KPI calculations & zero-division guards")
    lines.append("│   ├── stats.py                     # Closed-form statistical engine (Wilson, Spearman, Chi-sq)")
    lines.append("│   └── ui.py                        # Tailwind CSS iframe primitives & layout components")
    lines.append("├── data/")
    lines.append("│   ├── app.db                       # Shipped SQLite analytical database (34.2 MB)")
    lines.append("│   ├── hkd2026_data.xlsx            # Full Excel data dictionary & table workbook")
    lines.append("│   ├── quality_log.csv              # Cleansing reject log (5,368 excluded records)")
    lines.append("│   ├── validation_report.md         # Detailed warehouse data validation audit")
    lines.append("│   ├── validation_summary.json      # Machine-readable warehouse integrity metrics")
    lines.append("│   └── test_results.json            # Automated test execution evidence snapshot")
    lines.append("├── docs/                            # Formal project deliverables (Word docs & RTM)")
    lines.append("├── scripts/                         # Build automation & deliverable generation scripts")
    lines.append("├── tests/")
    lines.append("│   ├── conftest.py                  # Pytest fixtures and read-only DB connections")
    lines.append("│   ├── test_data.py                 # TC-T-01..20: Database schema & warehouse tests")
    lines.append("│   ├── test_analysis.py             # TC-T-21..40: KPI math & statistical unit tests")
    lines.append("│   └── test_app.py                  # TC-T-41..51: Streamlit AppTest end-to-end suite")
    lines.append("├── BUILD_PLAN.md                    # Current step: AI & developer build roadmap")
    lines.append("├── CLAUDE.md                        # Master project instructions & 18-step protocol")
    lines.append("├── pytest.ini                       # Pytest configuration and execution markers")
    lines.append("├── requirements.txt                 # Pinned production runtime dependencies")
    lines.append("├── requirements-dev.txt             # Pinned development & testing dependencies")
    lines.append("└── README.md                        # Comprehensive system manual & deployment guide")
    lines.append("```")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 3. Step-by-Step Build Tasks (BUILD-01 to BUILD-18)")
    lines.append("")
    for t in TASKS:
        lines.append(f"### `{t['id']}`: {t['name']}")
        lines.append(f"- **Goal:** {t['goal']}")
        lines.append(f"- **Target Files:** {', '.join(f'`{f}`' for f in t['files'])}")
        lines.append(f"- **Traceability:** Features: `{t['features']}` | User Stories: `{t['stories']}`")
        lines.append(f"- **AI Implementation Prompt:**")
        lines.append(f"  ```text\n  {t['prompt']}\n  ```")
        lines.append(f"- **Verification Procedure:** {t['check']}")
        lines.append(f"- **Required Passing Tests:** `{t['tests']}`")
        lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 4. Definition of Done (DoD)")
    lines.append("")
    lines.append("A task or release milestone in HKD-2026 is considered **DONE** if and only if all following criteria are met:")
    lines.append("1. **Requirements Coverage:** 100% of Must-have Business Requirements (`BR-01..07, 09, 13`) and Functional Requirements (`FR-01..22`) are implemented and validated.")
    lines.append("2. **Test Suite Execution:** All 51 Technical Tests (`TC-T-01..51`) and 46 Functional Tests (`TC-F-01..46`) pass with zero errors, zero failures, and zero unhandled skips.")
    lines.append("3. **Performance Latency:** Every application page renders in under 3.0 seconds warm-load time (`NFR-01`), verified by `TC-T-44`.")
    lines.append("4. **Security & Governance:** Zero PII fields exposed (`NFR-09`), database opened strictly in read-only mode, CSV download sanitized against spreadsheet formula injection.")
    lines.append("5. **Responsive Design:** Application successfully adapts to desktop (1920x1080), tablet (768px), and mobile viewports (375px reflow without horizontal scrollbars).")
    lines.append("6. **Traceability Synchronization:** Requirements Traceability Matrix updated to v1.5 with all MVP rows marked as `Built`.")
    lines.append("7. **Documentation Polish:** Comprehensive `README.md` authored with complete architecture diagrams, installation instructions, and user manual.")
    lines.append("8. **Cloud Readiness:** Code committed to GitHub repository and verified deployable on Streamlit Community Cloud without missing dependencies.")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 5. Development & Execution Commands")
    lines.append("")
    lines.append("### Environment Setup")
    lines.append("```bash")
    lines.append("# Create and activate Python 3.12 virtual environment")
    lines.append("python3.12 -m venv .venv")
    lines.append("source .venv/bin/activate")
    lines.append("")
    lines.append("# Install pinned core and development dependencies")
    lines.append("pip install --upgrade pip")
    lines.append("pip install -r requirements.txt")
    lines.append("pip install -r requirements-dev.txt")
    lines.append("```")
    lines.append("")
    lines.append("### Running the Streamlit Application")
    lines.append("```bash")
    lines.append("# Launch local development server")
    lines.append("streamlit run app.py")
    lines.append("```")
    lines.append("")
    lines.append("### Running the Automated Test Suite")
    lines.append("```bash")
    lines.append("# Run entire test suite (data, analysis, and AppTest)")
    lines.append(".venv/bin/python -m pytest tests -q")
    lines.append("")
    lines.append("# Run warehouse data integrity tests only")
    lines.append(".venv/bin/python -m pytest tests/test_data.py -q")
    lines.append("")
    lines.append("# Run KPI mathematics and statistical engine unit tests")
    lines.append(".venv/bin/python -m pytest tests/test_analysis.py -q")
    lines.append("")
    lines.append("# Run Streamlit AppTest end-to-end UI tests")
    lines.append(".venv/bin/python -m pytest tests/test_app.py -q")
    lines.append("")
    lines.append("# Generate formal test execution evidence JSON")
    lines.append(".venv/bin/python scripts/run_tests.py")
    lines.append("```")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## 6. Viva Voce Preparation (10 Core Technical Questions)")
    lines.append("")
    for i, item in enumerate(VIVA_QUESTIONS, 1):
        lines.append(f"### Q{i}: {item['q']}")
        lines.append(f"**Answer:** {item['a']}")
        lines.append("")
    return "\n".join(lines)

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

def para(doc, text, size=10.0, italic=False, bold=False, color=None,
         space_after=4, align=None, space_before=0):
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

def bullet(doc, text, size=9.8):
    p = doc.add_paragraph(text, style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.space_after = Pt(2.5)
    for r in p.runs:
        r.font.size = Pt(size)
    return p

def h1(doc, text, page_break=False):
    if page_break:
        doc.add_page_break()
    h = doc.add_heading(text, level=1)
    for r in h.runs:
        r.font.color.rgb = NAVY
        r.font.size = Pt(14)
    h.paragraph_format.space_before = Pt(10)
    h.paragraph_format.space_after = Pt(4)

def h2(doc, text):
    h = doc.add_heading(text, level=2)
    for r in h.runs:
        r.font.color.rgb = BLUE
        r.font.size = Pt(11.5)
    h.paragraph_format.space_before = Pt(8)
    h.paragraph_format.space_after = Pt(3)

def h3(doc, text):
    h = doc.add_heading(text, level=3)
    for r in h.runs:
        r.font.color.rgb = RGBColor(0x22, 0x33, 0x55)
        r.font.size = Pt(10)
    h.paragraph_format.space_before = Pt(6)
    h.paragraph_format.space_after = Pt(2)

def callout(doc, text, fill=CALLOUT_FILL, size=9.5, italic=False):
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.autofit = False
    cell_margins(table, top=60, bottom=60, left=90, right=90)
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
    cell_margins(table, top=50, bottom=50, left=70, right=50)
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

# --------------------------------------------------------------------------- docx generator
def build_docx():
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.75)
        s.bottom_margin = Inches(0.75)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    # Title Banner
    h = doc.add_heading(f"{PROJECT_ID} — Master Build Plan", level=0)
    for r in h.runs:
        r.font.color.rgb = NAVY
        r.font.size = Pt(20)
    para(doc, "Step 14: Engineering Task Roadmap, Component Specifications & Verification Blueprint", size=11, bold=True, color=BLUE)
    para(doc, f"Author: {STUDENT} (Roll: {ROLL})  ·  Mentor: Prof. {MENTOR}  ·  Institution: ABES Engineering College", size=9.5, italic=True)
    doc.add_paragraph()

    callout(doc, "PURPOSE: This document governs the technical execution of Step 15 (Build, Test, Fix). It maps out all "
                 "18 sequential engineering tasks, technical file dependencies, AI coding prompts, acceptance criteria, and "
                 "verification test assertions required to take the Healthcare KPI Dashboard from specifications to production.")

    # Section 1: Project Summary & Document Links
    h1(doc, "1. Project Summary & Document Traceability Index")
    para(doc, "HKD-2026 is an enterprise-grade healthcare operational analytics and clinical intelligence dashboard designed "
              "to empower hospital executives, clinical heads, and financial controllers with actionable decision-support tools. "
              "The platform monitors inpatient throughput, bed occupancy bottlenecks, average length of stay (ALOS), 30-day "
              "unplanned readmissions, in-hospital mortality, and revenue per bed-day across 8 clinical departments and 20 wards.")
    para(doc, "The underlying analytical warehouse contains 197,054 verified records (53,915 inpatient episodes and 143,139 OPD visits) "
              "spanning 2023–2025, synthesized under random seed 42 with zero PII exposure and five embedded clinical patterns. "
              "The application stack pairs a high-performance read-only SQLite database with Streamlit Multipage, interactive Plotly charts, "
              "and an encapsulated Tailwind CSS iframe design system. A comprehensive testing harness of 51 technical tests and 46 functional "
              "cases guarantees end-to-end traceability against all business requirements.")

    h2(doc, "Master Artifact Traceability Index")
    doc_rows = [
        ("Step 01", "Project Understanding", "docs/01_Project_Understanding.docx", "OQ-01..10, SH-01..05, SMART objectives"),
        ("Step 02", "How Might We Statements", "docs/02_HMW_Statements.docx", "12 HMW questions, 5 priority pain points"),
        ("Step 03", "Needs Analysis", "docs/03_Needs_Analysis.docx", "15 Business Requirements (BR-01..15), MoSCoW"),
        ("Step 04", "Traceability Matrix", "docs/04_Traceability_Matrix.xlsx", "Master RTM v1.4 forward & backward mapping"),
        ("Step 05", "BRD", "docs/05_BRD.docx", "Full Business Requirements Document & RACI"),
        ("Step 07", "TRD", "docs/07_TRD.docx", "Technical Architecture, FR-01..22, NFR-01..12"),
        ("Step 08", "Database Design", "docs/08_Database_Design.xlsx", "Star schema, 5 tables, 3 views, SQLite DDL"),
        ("Step 09", "ER Diagram", "docs/09_ER_Diagram.docx", "Chen & Crow's foot relational diagrams, keys"),
        ("Step 10", "Dataset Generation", "docs/10_Dataset_Generation.docx", "Seed-42 synthesis, 5 clinical patterns, DQ audit"),
        ("Step 11", "User Stories", "docs/11_User_Stories.docx", "20 User Stories (US-01..20), 5 epics, G/W/T criteria"),
        ("Step 12", "UI Framework", "docs/12_UI_Framework.docx", "9 screen wireframes, Tailwind design tokens, config"),
        ("Step 13", "Test Cases", "docs/13_Test_Cases.docx", "46 Functional (TC-F) & 51 Technical (TC-T) test specs"),
        ("Step 14", "Build Plan", "docs/14_BUILD_PLAN.docx", "Current Step: 18 tasks, DoD, commands & viva notes"),
        ("Data", "SQLite Warehouse", "data/app.db", "Shipped read-only relational analytical database"),
    ]
    add_table(doc, ["Step", "Deliverable Name", "File Path", "Scope & Coverage"],
              doc_rows, [0.8, 1.8, 2.3, 2.0], total=6.9, font_size=8.0)

    # Section 2: Final Folder Structure
    h1(doc, "2. Final Repository Directory Structure", page_break=True)
    para(doc, "The project structure follows production Streamlit community conventions, isolating data ingestion, "
              "analytical logic, UI primitives, and test specifications:")
    tree_text = """hclTraining/
├── .github/
│   └── workflows/
│       └── test.yml                 # Automated GitHub Actions pytest workflow
├── .streamlit/
│   └── config.toml                  # Streamlit visual theme tokens & server settings
├── app.py                           # Screen UI-01: Executive Overview Dashboard (Home)
├── pages/
│   ├── 1_Operations.py              # Screen UI-02: Operations & Capacity Bottlenecks
│   ├── 2_Departments.py             # Screen UI-03: Clinical Departments, ALOS & Readmission
│   ├── 3_Quality.py                 # Screen UI-04: Clinical Quality, Mortality & Drift
│   ├── 4_Finance.py                 # Screen UI-05: Hospital Finance, Billing & Payer Mix
│   ├── 5_Data_Quality.py            # Screen UI-06: Data Quality & Anomaly Audit
│   ├── 6_Definitions.py             # Screen UI-07: Definitions, KPI Registry & Governance
│   ├── 7_Insights.py                # Screen UI-08: Advanced Insights & Decomposition
│   └── 8_Data.py                    # Screen UI-09: Raw Data Exploration & Secure Export
├── src/
│   ├── __init__.py
│   ├── data.py                      # Data loader, @st.cache_data, and query engine
│   ├── filters.py                   # Centralized sidebar filters & session state manager
│   ├── generate_data.py             # Seed-42 synthetic warehouse data generator
│   ├── kpi.py                       # Pure mathematical KPI calculations & zero-division guards
│   ├── stats.py                     # Closed-form statistical engine (Wilson, Spearman, Chi-sq)
│   └── ui.py                        # Tailwind CSS iframe primitives & layout components
├── data/
│   ├── app.db                       # Shipped SQLite analytical database (34.2 MB)
│   ├── hkd2026_data.xlsx            # Full Excel data dictionary & table workbook
│   ├── quality_log.csv              # Cleansing reject log (5,368 excluded records)
│   ├── validation_report.md         # Detailed warehouse data validation audit
│   ├── validation_summary.json      # Machine-readable warehouse integrity metrics
│   └── test_results.json            # Automated test execution evidence snapshot
├── docs/                            # Formal project deliverables (Word docs & RTM)
├── scripts/                         # Build automation & deliverable generation scripts
├── tests/
│   ├── conftest.py                  # Pytest fixtures and read-only DB connections
│   ├── test_data.py                 # TC-T-01..20: Database schema & warehouse tests
│   ├── test_analysis.py             # TC-T-21..40: KPI math & statistical unit tests
│   └── test_app.py                  # TC-T-41..51: Streamlit AppTest end-to-end suite
├── BUILD_PLAN.md                    # Current step: AI & developer build roadmap
├── CLAUDE.md                        # Master project instructions & 18-step protocol
├── pytest.ini                       # Pytest configuration and execution markers
├── requirements.txt                 # Pinned production runtime dependencies
├── requirements-dev.txt             # Pinned development & testing dependencies
└── README.md                        # Comprehensive system manual & deployment guide"""
    mono(doc, tree_text, size=6.8)

    # Section 3: Build Tasks
    h1(doc, "3. Sequential Engineering Tasks (BUILD-01 to BUILD-18)", page_break=True)
    para(doc, "Execution proceeds sequentially through 18 discrete build units. Each task defines target files, "
              "requirements links, exact implementation prompts, verification procedures, and mandatory passing tests:")

    for t in TASKS:
        h2(doc, f"{t['id']}: {t['name']}")
        para(doc, f"Goal: {t['goal']}", bold=True, size=9.5)
        para(doc, f"Files Affected: {', '.join(t['files'])}  |  Traceability: {t['features']} ({t['stories']})", size=8.8, color=GREY)
        para(doc, "AI Prompt / Implementation Specification:", size=8.8, bold=True, color=NAVY)
        mono(doc, t["prompt"], size=7.2)
        para(doc, f"Verification Check: {t['check']}", size=8.8, italic=True)
        para(doc, f"Passing Test Gates: {t['tests']}", size=8.8, bold=True, color=BLUE)
        doc.add_paragraph()

    # Section 4: Definition of Done
    h1(doc, "4. Definition of Done (DoD)", page_break=True)
    para(doc, "A task or phase in the HKD-2026 build lifecycle is classified as complete if and only if all eight "
              "contractual criteria are fulfilled without exemption:")
    bullet(doc, "1. Requirements Compliance: All Must-have Business Requirements (BR-01..07, 09, 13) and Functional Requirements (FR-01..22) implemented.")
    bullet(doc, "2. Test Suite Clearance: All 51 Technical Tests (TC-T-01..51) and 46 Functional Tests (TC-F-01..46) pass with 0 failures and 0 errors.")
    bullet(doc, "3. Latency Budget: Every application screen achieves warm-render execution in < 3.0 seconds (NFR-01), verified by TC-T-44.")
    bullet(doc, "4. Zero Division & Null Resilience: All KPI routines return None rather than raising exceptions or emitting NaN; cards render '—' gracefully.")
    bullet(doc, "5. Security & Privacy: Zero PII fields exposed (NFR-09); database opened strictly read-only; CSV exports sanitized against spreadsheet injection.")
    bullet(doc, "6. Viewport Responsiveness: Layout renders cleanly at 1920x1080, 1366x768, and reflows seamlessly on 375px mobile displays (NFR-08).")
    bullet(doc, "7. RTM Synchronization: Traceability Matrix updated to v1.5 with all MVP rows marked as 'Built'.")
    bullet(doc, "8. Production Deployability: Committed to Git with pinned requirements.txt, ready for Streamlit Community Cloud hosting.")

    # Section 5: Commands
    h1(doc, "5. Command Execution Guide", page_break=True)
    h2(doc, "Virtual Environment & Installation")
    mono(doc, """# Setup Python 3.12 virtual environment
python3.12 -m venv .venv
source .venv/bin/activate

# Install production and development dependencies
pip install --upgrade pip
pip install -r requirements.txt
pip install -r requirements-dev.txt""")

    h2(doc, "Running the Application")
    mono(doc, """# Launch the Streamlit multipage application
streamlit run app.py""")

    h2(doc, "Executing the Test Suite")
    mono(doc, """# 1. Run entire test suite
.venv/bin/python -m pytest tests -q

# 2. Run specific test modules
.venv/bin/python -m pytest tests/test_data.py -q         # Schema & DB integrity
.venv/bin/python -m pytest tests/test_analysis.py -q     # KPI math & statistics
.venv/bin/python -m pytest tests/test_app.py -q          # Streamlit AppTest suite

# 3. Generate formal test evidence snapshot (test_results.json)
.venv/bin/python scripts/run_tests.py""")

    # Section 6: Viva Voce Notes
    h1(doc, "6. Comprehensive Viva Voce Preparation (10 Core Questions)", page_break=True)
    para(doc, "The following 10 questions cover critical architecture, statistical theory, software engineering, "
              "and clinical operations decisions embedded in this project:")

    for i, item in enumerate(VIVA_QUESTIONS, 1):
        h3(doc, f"Q{i}: {item['q']}")
        para(doc, item["a"], size=9.5)
        doc.add_paragraph()

    doc.save(DOCX_OUT)
    print("Saved Word document:", DOCX_OUT)

# --------------------------------------------------------------------------- execute
if __name__ == "__main__":
    md_content = build_markdown()
    with open(MD_OUT, "w", encoding="utf-8") as f:
        f.write(md_content)
    print("Saved Markdown build plan:", MD_OUT)
    build_docx()
    print("Step 14 complete!")
