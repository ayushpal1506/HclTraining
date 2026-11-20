# HKD-2026 — Master Build Plan (BUILD_PLAN.md)

> **Project:** Healthcare KPI Dashboard (`HKD-2026`)  
> **Domain:** Clinical & Operational Data Analytics  
> **Student:** Ayush Pal (Roll: 2503201450035, MCA 2nd yr DA, ABES)  
> **Mentor:** Prof. Aplana Lodhi  
> **Target Stack:** Python 3.12+, Streamlit, Plotly, Tailwind CSS (iframe), SQLite (`data/app.db`), pytest  

---

## 1. Project Summary & Deliverable Index

HKD-2026 is an enterprise-grade hospital operational analytics and clinical KPI dashboard designed for hospital executives, medical directors, and financial controllers.
The platform monitors inpatient throughput, bed occupancy, average length of stay (ALOS), 30-day unplanned readmissions, in-hospital mortality, and revenue per bed-day.
Built on a verified synthetic warehouse of 53,915 inpatient episodes and 143,139 OPD visits across 8 clinical departments, 20 wards, and 5 payer categories.
The analytical engine implements strict zero-division guards, closed-form Wilson score confidence intervals, and statistical process control drift detection.
The user interface features 9 dedicated Streamlit screens styled with Tailwind CSS design tokens encapsulated inside secure iframes.
A comprehensive testing harness of 51 technical tests (schema, mathematics, AppTest) and 46 functional test cases guarantees 100% requirements coverage.
All data processing is read-only, fully reproducible under random seed 42, and devoid of real personal identifiable information (PII).
The solution achieves warm page load rendering in under 3.0 seconds, exceeding all healthcare business intelligence latency standards.
Traceability is maintained end-to-end across all 18 engineering steps via an active Requirements Traceability Matrix (RTM v1.4).
Final deployment targets Streamlit Community Cloud with automated CI/CD validation via GitHub Actions.

### Master Documentation Index

- **Step 01 — [Project Understanding](docs/01_Project_Understanding.docx):** Problem statement, stakeholders, operational questions (OQ-01..10), SMART scope.
- **Step 02 — [How Might We Statements](docs/02_HMW_Statements.docx):** 12 design-thinking HMW questions, stakeholder pain points, 5 priority statements.
- **Step 03 — [Needs Analysis](docs/03_Needs_Analysis.docx):** 15 Business Requirements (BR-01..15) with MoSCoW prioritization and metric mapping.
- **Step 04 — [Requirements Traceability Matrix](docs/04_Traceability_Matrix.xlsx):** Master RTM v1.4 tracking links from OQs to Test Cases across all phases.
- **Step 05 — [Business Requirements Document](docs/05_BRD.docx):** Formal BRD with RACI matrix, clinical business rules, and risk mitigation.
- **Step 07 — [Technical Requirements Document](docs/07_TRD.docx):** Software architecture, functional/non-functional specs (FR-01..22, NFR-01..12).
- **Step 08 — [Database Design](docs/08_Database_Design.xlsx):** Star/snowflake schema, data dictionary, DDL constraints, analytical views.
- **Step 09 — [ER Diagram](docs/09_ER_Diagram.docx):** Entity-Relationship model, cardinality relationships, Chen/Crow's Foot diagram.
- **Step 10 — [Dataset Generation](docs/10_Dataset_Generation.docx):** Seed-42 synthetic data synthesis, 5 clinical patterns, data quality audit report.
- **Step 11 — [User Stories](docs/11_User_Stories.docx):** 20 User Stories (US-01..20) across 5 epics with Given/When/Then acceptance criteria.
- **Step 12 — [UI Framework & Wireframes](docs/12_UI_Framework.docx):** 9-screen wireframe specifications, Tailwind design tokens, UI component contracts.
- **Step 13 — [Test Cases & Test Basis](docs/13_Test_Cases.docx):** 46 Functional Cases (TC-F-01..46) and 51 Technical Cases (TC-T-01..51).
- **Step 14 — [Master Build Plan](docs/14_BUILD_PLAN.docx):** Current step: 18 structured build tasks, Definition of Done, execution guide.
- **Warehouse — [SQLite Database](data/app.db):** Shipped relational warehouse containing 197,054 validated clinical records.
- **Data Dict — [Excel Data Workbook](data/hkd2026_data.xlsx):** Multi-sheet workbook containing raw and cleaned tables with dictionary metadata.

---

## 2. Final Project Directory Structure

```text
hclTraining/
├── .github/
│   └── workflows/
│       └── test.yml                 # Automated CI/CD pytest workflow
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
└── README.md                        # Comprehensive system manual & deployment guide
```

---

## 3. Step-by-Step Build Tasks (BUILD-01 to BUILD-18)

### `BUILD-01`: Environment Setup, Dependencies & CI/CD Pipeline
- **Goal:** Establish reproducible Python environment, pin core & development dependencies, configure Streamlit theme tokens, and install GitHub Actions test workflow.
- **Target Files:** `requirements.txt`, `requirements-dev.txt`, `.streamlit/config.toml`, `.github/workflows/test.yml`
- **Traceability:** Features: `FR-01, FR-21, NFR-03, NFR-08` | User Stories: `US-01, US-08`
- **AI Implementation Prompt:**
  ```text
  Create 'requirements.txt' pinning: streamlit>=1.40.0, pandas>=2.2.0, numpy>=1.26.0, plotly>=5.24.0, scipy>=1.13.0, statsmodels>=0.14.0. Create 'requirements-dev.txt' including pytest>=8.0.0, pytest-mock, python-docx>=1.1.0, openpyxl>=3.1.0. Verify '.streamlit/config.toml' has primaryColor='#1F3864', backgroundColor='#FFFFFF', secondaryBackgroundColor='#F4F6FA', textColor='#0F172A', and client.showErrorDetails='none'. Create '.github/workflows/test.yml' to run automated tests on push/PR with Python 3.12.
  ```
- **Verification Procedure:** Run 'pip install -r requirements.txt' in Python 3.12/3.13 venv; ensure all packages import cleanly without version conflicts.
- **Required Passing Tests:** `TC-T-01, TC-T-49`

### `BUILD-02`: Data Pipeline, Database Access & In-Memory Caching
- **Goal:** Build robust, read-only SQLite ingestion engine with Streamlit @st.cache_data, multi-dimensional filtering, and date boundary guards.
- **Target Files:** `src/data.py`
- **Traceability:** Features: `FR-01, FR-02, FR-04, NFR-01, NFR-09` | User Stories: `US-01, US-02`
- **AI Implementation Prompt:**
  ```text
  Write 'src/data.py' implementing load_data() with @st.cache_data. Open 'data/app.db' in URI mode (file:data/app.db?mode=ro). Provide query functions for 'vw_analysis_episode', 'vw_ward_day_beds', 'vw_dept_month_kpi', and dimension tables. Implement filter_episodes(df, period, departments, wards, payers, gender, age_range) returning filtered slices. Include row-count assertions and ensure no query exposes PII or modifies database.
  ```
- **Verification Procedure:** Execute python -c 'import src.data as d; df=d.load_episodes(); assert len(df)==53915; print(df.shape)' verifying 53,915 episodes loaded in < 1.5 s.
- **Required Passing Tests:** `TC-T-01, TC-T-02, TC-T-03, TC-T-06, TC-F-01`

### `BUILD-03`: Analytical Core & KPI Mathematics Verification
- **Goal:** Verify pure mathematical functions for all clinical and financial KPIs, ensuring zero-division safety, target benchmarking, and contribution decomposition.
- **Target Files:** `src/kpi.py`
- **Traceability:** Features: `FR-05, FR-07, FR-08, FR-09, FR-11, FR-12, FR-14` | User Stories: `US-04, US-06, US-09, US-10, US-11, US-13, US-14`
- **AI Implementation Prompt:**
  ```text
  Validate and harden 'src/kpi.py'. Ensure functions: bor_pct(occupied, available), alos_days(days, discharges), mortality_pct(deaths, discharges), readmission_rate_pct(readmissions, eligible), revenue_per_patient(revenue, discharges), revenue_per_bed_day(revenue, bed_days), beds_to_target(bor, available, days, 85.0), eligibility(discharges, deaths), and decompose(parts_a, parts_b). All functions must return None (rendering '—') on zero or null denominators without throwing exceptions.
  ```
- **Verification Procedure:** Run pytest tests/test_analysis.py -k 'test_TC_T_2' ensuring all KPI mathematical properties pass.
- **Required Passing Tests:** `TC-T-21, TC-T-22, TC-T-23, TC-T-24, TC-T-25, TC-T-26, TC-T-27, TC-T-28, TC-T-30`

### `BUILD-04`: Closed-Form Statistical Engine Verification
- **Goal:** Verify statistical calculations for Wilson score confidence intervals, Spearman rank correlation, Chi-square test of independence, OLS trendlines, and quality drift detection.
- **Target Files:** `src/stats.py`
- **Traceability:** Features: `FR-08, FR-12, FR-15, NFR-10` | User Stories: `US-09, US-13, US-15`
- **AI Implementation Prompt:**
  ```text
  Verify 'src/stats.py'. Ensure wilson_interval(k, n, z=1.96) strictly outputs bounds in [0.0, 100.0] even for k=0 or k=n. Verify spearman(x, y) handles ties via average ranking and calculates exact Student-t two-sided p-values via regularized incomplete beta function. Validate chi_square_2x2(a, b, c, d) and drift_flag(series, baseline_n=12, z=2.0) for detecting abnormal clinical shifts.
  ```
- **Verification Procedure:** Run pytest tests/test_analysis.py -k 'test_TC_T_3' verifying that all 10 closed-form statistical unit tests pass.
- **Required Passing Tests:** `TC-T-31, TC-T-32, TC-T-33, TC-T-34, TC-T-35, TC-T-36, TC-T-37, TC-T-38, TC-T-39, TC-T-40`

### `BUILD-05`: Tailwind UI Primitive Components & Responsive Layout
- **Goal:** Verify and harden Tailwind CSS iframe-encapsulated UI components, ensuring HTML sanitization, color token consistency, and mobile CSS reflow.
- **Target Files:** `src/ui.py`
- **Traceability:** Features: `FR-18, FR-19, FR-20, NFR-04, NFR-08` | User Stories: `US-02, US-08, US-18`
- **AI Implementation Prompt:**
  ```text
  Verify 'src/ui.py' components: setup(title), enable_mobile(), header(title, subtitle, as_of), kpi_card(spec), kpi_row(cards), chart_card(title, fig, note), insight_box(text, level, action, source), filter_chips(chips), empty_state(message, filters), and dq_banner(passed, failed, period). Ensure every variable injected into HTML is sanitized via html.escape(..., quote=True). Ensure ui.LAST_RENDER dictionary captures all published KPI values for automated assertions.
  ```
- **Verification Procedure:** Run python -c 'import src.ui as u; print(u.THEME["navy"]); assert len(u.THEME)==9' and verify theme tokens match .streamlit/config.toml.
- **Required Passing Tests:** `TC-T-43, TC-T-45, TC-T-49, TC-F-38`

### `BUILD-06`: Shared Sidebar Navigation, Filters & Session State Manager
- **Goal:** Build centralized filter controls in Streamlit sidebar, managing global state across pages, cascading department/ward options, and providing instant reset.
- **Target Files:** `src/filters.py`
- **Traceability:** Features: `FR-02, FR-03, NFR-05` | User Stories: `US-02, US-03`
- **AI Implementation Prompt:**
  ```text
  Create 'src/filters.py' with render_sidebar(df) managing st.session_state for: Period (2023, 2024, 2025 default), Compare Period (None, 2023, 2024 default), Department multiselect, Ward multiselect (cascaded by selected department), Payer multiselect, Gender radio, and Age slider. Provide 'Reset All Filters' button restoring defaults. Return FilterState dataclass and active filter labels for display in ui.filter_chips().
  ```
- **Verification Procedure:** Test filter selection in Streamlit; verify changing department dynamically prunes available wards without resetting other state.
- **Required Passing Tests:** `TC-F-02, TC-F-04, TC-F-05, TC-F-06, TC-F-07, TC-F-08, TC-F-09, TC-F-10`

### `BUILD-07`: Screen UI-01: Executive Overview Dashboard
- **Goal:** Build main application entry point (app.py) displaying 6 high-level KPI cards, top operational movements, data quality summary banner, and active filter chips.
- **Target Files:** `app.py`
- **Traceability:** Features: `FR-01, FR-02, FR-05, FR-12, FR-20, BR-04` | User Stories: `US-02, US-08, US-16`
- **AI Implementation Prompt:**
  ```text
  Build 'app.py'. Call ui.setup('Overview · HKD-2026') and ui.enable_mobile(). Load data via src.data, render sidebar via src.filters. Display ui.header('Executive Overview', 'High-level hospital throughput, capacity, quality & financial summary', as_of='31 Dec 2025'). Render ui.dq_banner(passed=6, failed=0, period='2025'). Display ui.filter_chips(). Render 6 KPI cards in 2 rows: Discharges (18,142), Bed Occupancy Rate (73.2%), ALOS (4.36 d), Revenue (₹135.26 cr), Mortality Rate (1.29%), and 30-day Readmission (5.03%) with deltas. Add 3 insight boxes highlighting key operational takeaways and quick links to deep-dive pages.
  ```
- **Verification Procedure:** Run streamlit run app.py; verify page loads in < 2 seconds with all 6 KPI cards displaying exact validated metrics.
- **Required Passing Tests:** `TC-F-01, TC-F-04, TC-F-13, TC-F-16, TC-F-18, TC-F-40, TC-T-41, TC-T-45`

### `BUILD-08`: Screen UI-02: Operations & Capacity Bottlenecks
- **Goal:** Build operations dashboard analyzing bed occupancy against the 85% clinical threshold, 7-day ward projections, OPD weekly seasonality, and holiday volume trends.
- **Target Files:** `pages/1_Operations.py`
- **Traceability:** Features: `FR-02, FR-05, FR-06, FR-10, BR-01, BR-02, BR-09` | User Stories: `US-03, US-04, US-05, US-07`
- **AI Implementation Prompt:**
  ```text
  Create 'pages/1_Operations.py'. Call ui.setup('Operations · HKD-2026'). Render sidebar filters and header. Implement tabs: 'Capacity & Bottlenecks', '7-Day Ward Forecast', 'OPD Volume & Trends'. Tab 1: Horizontal Plotly bar chart comparing department BOR against 85.0% target line; Cardiology highlighted in red (>85%). Show occupancy gap in percentage points and physical beds needed. Tab 2: 7-day forward ward occupancy forecast table and bar chart. Tab 3: 12x7 heatmap of OPD arrivals showing Monday peak (ratio 1.52) and Sunday lull, plus Nov-Feb winter respiratory surge (+28%).
  ```
- **Verification Procedure:** Inspect operations page; verify Cardiology occupancy reads 85.7% (breaching target) and Monday OPD volume averages 202.7 visits/day.
- **Required Passing Tests:** `TC-F-02, TC-F-05, TC-F-06, TC-F-14, TC-F-21, TC-F-22, TC-F-27, TC-F-41, TC-T-42`

### `BUILD-09`: Screen UI-03: Clinical Departments, ALOS & Readmission
- **Goal:** Build department analysis view featuring ALOS benchmark comparison dumbbell chart, 30-day readmission rates with Wilson 95% confidence intervals, and demographic breakdown.
- **Target Files:** `pages/2_Departments.py`
- **Traceability:** Features: `FR-02, FR-07, FR-08, FR-13, BR-06, BR-07, BR-11` | User Stories: `US-06, US-09, US-12, US-18`
- **AI Implementation Prompt:**
  ```text
  Create 'pages/2_Departments.py'. Implement 3 analytical sections: 1) ALOS vs 4.0-day benchmark dumbbell chart (Oncology 6.49 d right of line, Maternity 2.76 d left). 2) 30-day unplanned readmissions bar chart with Wilson 95% error bars; display sample count n under each bar; flag General Medicine (7.70%), Oncology (6.36%), Cardiology (5.33%) as high readmission wards. Show low-confidence badge on groups with n < 30. 3) Demographic pyramids and age-band distributions (0-5, 6-17, 18-35, 36-50, 51-65, 65+).
  ```
- **Verification Procedure:** Confirm ALOS chart shows Oncology at 6.49 days and readmission error bars stay strictly within [0%, 100%]; ENT low-volume cohorts show 'low confidence'.
- **Required Passing Tests:** `TC-F-08, TC-F-15, TC-F-17, TC-F-19, TC-F-23, TC-F-24, TC-F-28, TC-F-42, TC-T-42`

### `BUILD-10`: Screen UI-04: Clinical Quality, Mortality & Drift Detection
- **Goal:** Build clinical quality monitor tracking in-hospital mortality rates, discharge status distribution (Home, LAMA, Transfer, Expired), and statistical drift alerts.
- **Target Files:** `pages/3_Quality.py`
- **Traceability:** Features: `FR-02, FR-12, BR-13, NFR-10` | User Stories: `US-13`
- **AI Implementation Prompt:**
  ```text
  Create 'pages/3_Quality.py'. Display in-hospital mortality trends across 2023-2025. Implement statistical drift control chart: plot monthly mortality against baseline mean +/- 2.0 SD (from 2023 baseline). Flag Oncology upward drift (3.52% in 2023 to 5.62% in 2025; recent half-year at 6.0%). Render discharge disposition breakdown pie/donut chart (Normal, Transfer, LAMA, Expired). Include explicit clinical caveats stating this is a review trigger, not proof of negligence.
  ```
- **Verification Procedure:** Verify Oncology mortality chart shows red drift flag after month 24, with alert explaining baseline mean and 2-sigma boundary breach.
- **Required Passing Tests:** `TC-F-16, TC-T-39, TC-T-42`

### `BUILD-11`: Screen UI-05: Hospital Finance, Billing & Payer Mix
- **Goal:** Build financial intelligence page tracking payer category market share shifts, revenue per discharged patient, and ranked revenue per occupied bed-day.
- **Target Files:** `pages/4_Finance.py`
- **Traceability:** Features: `FR-02, FR-09, FR-11, BR-08, BR-12` | User Stories: `US-03, US-10, US-11`
- **AI Implementation Prompt:**
  ```text
  Create 'pages/4_Finance.py'. 1) 100% stacked bar chart showing payer mix evolution (2023 to 2025); highlight Corporate expansion (6.37% -> 11.23%) and Self-Pay contraction (41.35% -> 33.49%). 2) Revenue per patient card & distribution across payers. 3) Ranked lollipop chart for revenue per occupied bed-day: Oncology #1 (₹31,232/day), Orthopaedics #2 (₹21,809/day), down to General Medicine (₹10,646/day), displaying exact percentage and rupee gap to leader.
  ```
- **Verification Procedure:** Verify payer mix stacks sum to exactly 100.0% (+/- 0.1%) and revenue per bed-day rankings match validated figures.
- **Required Passing Tests:** `TC-F-07, TC-F-10, TC-F-18, TC-F-25, TC-F-26, TC-T-42`

### `BUILD-12`: Screen UI-06: Data Quality & Anomaly Audit
- **Goal:** Build data governance and audit dashboard showcasing 6 integrity check gates, raw reject log with reason codes, and data cleansing pipeline proof.
- **Target Files:** `pages/5_Data_Quality.py`
- **Traceability:** Features: `FR-17, BR-14, NFR-09` | User Stories: `US-17, US-18, US-19`
- **AI Implementation Prompt:**
  ```text
  Create 'pages/5_Data_Quality.py'. Display data validation scorecard: 1) Primary key uniqueness (0 dups), 2) Foreign key integrity (0 orphans), 3) Date sanity (discharge >= admission), 4) Length of stay boundaries (1 to 90 days), 5) Billing non-negativity (₹5,788 - ₹60,56,692), 6) Bed occupancy physical limits (<=100%). Display reject log from 'data/quality_log.csv' detailing the 5,368 removed anomalies (missing payers, duplicate episodes, negative stay) with reason codes. Emphasize zero synthetic PII exposure.
  ```
- **Verification Procedure:** Verify all 6 checks show green PASS status, and reject log displays 5,368 filtered entries with exact category breakdown.
- **Required Passing Tests:** `TC-F-31, TC-F-32, TC-F-33, TC-T-18, TC-T-42`

### `BUILD-13`: Screen UI-07: Definitions, KPI Registry & Governance
- **Goal:** Build interactive clinical & operational KPI registry detailing formal definitions, numerator/denominator formulas, benchmark sources, update frequency, and data steward ownership.
- **Target Files:** `pages/6_Definitions.py`
- **Traceability:** Features: `FR-16, BR-05` | User Stories: `US-16`
- **AI Implementation Prompt:**
  ```text
  Create 'pages/6_Definitions.py'. Display searchable KPI Registry table for BOR, ALOS, 30-day Readmission Rate, In-Hospital Mortality, Revenue per Patient, and Revenue per Bed-Day. Detail: Metric ID, Official Title, Clinical Definition, Exact Mathematical Formula, National/NABH Benchmark, Target Threshold, Refresh Cadence, and Responsible Data Steward. Include database schema dictionary and version control metadata (v1.0).
  ```
- **Verification Procedure:** Verify registry entries match src/kpi.py formulas word-for-word and include data steward contacts and NABH standards.
- **Required Passing Tests:** `TC-F-20, TC-F-37, TC-T-42`

### `BUILD-14`: Screen UI-08: Advanced Analytical Insights & Decomposition
- **Goal:** Build advanced analytics screen offering mathematical waterfall decomposition of KPI shifts and hypothesis testing for bed occupancy vs readmission rates.
- **Target Files:** `pages/7_Insights.py`
- **Traceability:** Features: `FR-14, FR-15, BR-03, BR-10` | User Stories: `US-14, US-15`
- **AI Implementation Prompt:**
  ```text
  Create 'pages/7_Insights.py'. 1) KPI Movement Decomposition: Plotly waterfall chart decomposing total hospital revenue movement (e.g., 2024 to 2025: +₹2.16 cr) into individual departmental contributions with explicit printed residual (no merging into 'other'). 2) Hypothesis Testing: Spearman rank correlation between monthly Bed Occupancy Rate and 30-day Readmission Rate across 288 department-months. Display scatter plot with OLS line, sample size n=288, rho, t-statistic, exact two-sided p-value, and prominent 'Association, not causation' disclaimer.
  ```
- **Verification Procedure:** Verify waterfall contributors reconcile to total revenue delta within +/- ₹0.01 cr, and Spearman correlation card prints rho, t, and p values cleanly.
- **Required Passing Tests:** `TC-F-29, TC-F-30, TC-T-30, TC-T-38, TC-T-42`

### `BUILD-15`: Screen UI-09: Raw Data Exploration & Secure Export
- **Goal:** Build paginated data table preview of filtered cohort, column dictionary, and secure CSV download with UTF-8 BOM prefix, CSV injection sanitization, and 50,000-row guard.
- **Target Files:** `pages/8_Data.py`
- **Traceability:** Features: `FR-01, FR-02, BR-14, NFR-09, NFR-11` | User Stories: `US-20`
- **AI Implementation Prompt:**
  ```text
  Create 'pages/8_Data.py'. Render paginated st.dataframe preview of filtered 'vw_analysis_episode' dataset (maximum 1,000 rows on screen). Provide column dictionary showing data types and descriptions. Implement st.download_button generating CSV export. Security guards: 1) Enforce 50,000 row maximum export limit (show warning if exceeded), 2) Prefix cells starting with '=, +, -, @' with a single quote to prevent spreadsheet formula injection, 3) Add UTF-8 BOM ('ï»¿') for correct Excel encoding.
  ```
- **Verification Procedure:** Download CSV with filters applied; inspect raw bytes to verify UTF-8 BOM, verify Cardiology 2025 exports exactly 3,181 rows.
- **Required Passing Tests:** `TC-F-34, TC-F-35, TC-F-36, TC-T-47, TC-T-48, TC-T-42`

### `BUILD-16`: Automated Testing Suite Execution & Performance Profiling
- **Goal:** Enable and execute the full test suite (test_data.py, test_analysis.py, test_app.py), eliminating skips, verifying warm page load < 3.0 s, and generating test_results.json.
- **Target Files:** `tests/test_app.py`, `scripts/run_tests.py`, `data/test_results.json`
- **Traceability:** Features: `FR-01 through FR-22, NFR-01 through NFR-12` | User Stories: `US-01 through US-20`
- **AI Implementation Prompt:**
  ```text
  Remove skip condition from 'tests/test_app.py'. Run 'pytest tests -q'. Verify all 51 technical tests pass. Execute 'python scripts/run_tests.py' to generate 'data/test_results.json' and JUnit XML report. Assert that every page renders in under 3.0 seconds warm-load time (TC-T-44). Ensure zero uncaught exceptions across all 9 screens.
  ```
- **Verification Procedure:** Run 'pytest tests -q'; confirm output reports 51 passed (or 51 + functional tests) with 0 failures, 0 errors.
- **Required Passing Tests:** `TC-T-01 through TC-T-51, TC-F-01 through TC-F-46`

### `BUILD-17`: Comprehensive Documentation, User Manual & ABES Academic Polish
- **Goal:** Author authoritative README.md, document architecture, operational guide, project instructions, and update RTM to v1.5 with 'Built' status.
- **Target Files:** `README.md`, `CLAUDE.md`, `docs/04_Traceability_Matrix.xlsx`
- **Traceability:** Features: `NFR-12` | User Stories: `All`
- **AI Implementation Prompt:**
  ```text
  Write comprehensive 'README.md' containing: 1) Project banner and summary, 2) Architecture diagram (Mermaid), 3) Clinical and operational problem statement, 4) Installation and setup instructions, 5) Streamlit execution commands, 6) Test suite running guide, 7) UI navigation walkthrough with screen summaries, 8) Analytical and statistical methodology notes, 9) Academic metadata (ABES MCA, student details, mentor). Update RTM v1.5 in 'docs/04_Traceability_Matrix.xlsx' marking MVP rows as 'Built'.
  ```
- **Verification Procedure:** Review README.md formatting and render; verify all internal file links and GitHub commands work as specified.
- **Required Passing Tests:** `TC-T-48, TC-T-50`

### `BUILD-18`: Production Deployment & Streamlit Community Cloud Release
- **Goal:** Configure production repository, verify cloud headless settings, test sandbox safety, and prepare live deployment to Streamlit Community Cloud.
- **Target Files:** `.streamlit/config.toml`, `requirements.txt`, `app.py`
- **Traceability:** Features: `FR-21, FR-22, NFR-02, NFR-06, NFR-07` | User Stories: `US-01, US-08`
- **AI Implementation Prompt:**
  ```text
  Prepare application for Streamlit Community Cloud deployment. Ensure '.streamlit/config.toml' is committed with headless=true and server.enableCORS=false. Verify 'data/app.db' is included in repository with read-only permissions. Confirm 'requirements.txt' contains exact version pins without local path dependencies. Perform pre-deployment smoke test.
  ```
- **Verification Procedure:** Simulate clean clone in temporary directory; verify 'pip install -r requirements.txt && streamlit run app.py' launches without errors.
- **Required Passing Tests:** `TC-T-41, TC-T-42, TC-T-44`

---

## 4. Definition of Done (DoD)

A task or release milestone in HKD-2026 is considered **DONE** if and only if all following criteria are met:
1. **Requirements Coverage:** 100% of Must-have Business Requirements (`BR-01..07, 09, 13`) and Functional Requirements (`FR-01..22`) are implemented and validated.
2. **Test Suite Execution:** All 51 Technical Tests (`TC-T-01..51`) and 46 Functional Tests (`TC-F-01..46`) pass with zero errors, zero failures, and zero unhandled skips.
3. **Performance Latency:** Every application page renders in under 3.0 seconds warm-load time (`NFR-01`), verified by `TC-T-44`.
4. **Security & Governance:** Zero PII fields exposed (`NFR-09`), database opened strictly in read-only mode, CSV download sanitized against spreadsheet formula injection.
5. **Responsive Design:** Application successfully adapts to desktop (1920x1080), tablet (768px), and mobile viewports (375px reflow without horizontal scrollbars).
6. **Traceability Synchronization:** Requirements Traceability Matrix updated to v1.5 with all MVP rows marked as `Built`.
7. **Documentation Polish:** Comprehensive `README.md` authored with complete architecture diagrams, installation instructions, and user manual.
8. **Cloud Readiness:** Code committed to GitHub repository and verified deployable on Streamlit Community Cloud without missing dependencies.

---

## 5. Development & Execution Commands

### Environment Setup
```bash
# Create and activate Python 3.12 virtual environment
python3.12 -m venv .venv
source .venv/bin/activate

# Install pinned core and development dependencies
pip install --upgrade pip
pip install -r requirements.txt
pip install -r requirements-dev.txt
```

### Running the Streamlit Application
```bash
# Launch local development server
streamlit run app.py
```

### Running the Automated Test Suite
```bash
# Run entire test suite (data, analysis, and AppTest)
.venv/bin/python -m pytest tests -q

# Run warehouse data integrity tests only
.venv/bin/python -m pytest tests/test_data.py -q

# Run KPI mathematics and statistical engine unit tests
.venv/bin/python -m pytest tests/test_analysis.py -q

# Run Streamlit AppTest end-to-end UI tests
.venv/bin/python -m pytest tests/test_app.py -q

# Generate formal test execution evidence JSON
.venv/bin/python scripts/run_tests.py
```

---

## 6. Viva Voce Preparation (10 Core Technical Questions)

### Q1: Why was SQLite selected over PostgreSQL or MySQL for this clinical analytics project?
**Answer:** SQLite provides a zero-configuration, serverless, self-contained relational database embedded directly inside the application repository. For a 54,000-episode analytical warehouse, SQLite opened in URI read-only mode (`?mode=ro`) delivers sub-millisecond query latencies, eliminates external database connection failure modes, prevents unauthorized runtime data mutations, and enables seamless single-command deployment on Streamlit Community Cloud.

### Q2: How does this application guarantee zero Division-by-Zero errors when calculating healthcare KPIs?
**Answer:** All KPI calculations are centralized in `src/kpi.py` using helper routines (`_div`) that explicitly return `None` whenever the denominator is zero or `None`. The UI rendering engine translates `None` into an em-dash (`—`), accompanied by a 'low confidence' badge when $n < 30$, ensuring the application never crashes with a `ZeroDivisionError` or renders misleading zeros for departments with no discharges.

### Q3: Why is the Wilson score interval used for readmission rates instead of the standard Wald normal approximation?
**Answer:** The Wald interval ($\hat{p} \pm z\sqrt{\hat{p}(1-\hat{p})/n}$) relies on asymptotic normality, which severely breaks down when sample sizes are small ($n < 30$) or proportions approach zero (such as hospital readmission and mortality rates of 1%–5%), often yielding impossible negative rates. The Wilson score interval derives directly from inverting the score test, remaining mathematically bounded within $[0\%, 100\%]$ and providing accurate coverage probability across all sample sizes.

### Q4: Why are KPI cards rendered inside a Tailwind CSS iframe while Plotly charts remain native Streamlit components?
**Answer:** Tailwind CSS provides responsive card layouts, precise typography, and color tokens via iframe encapsulation (`components.html`) without polluting or conflicting with Streamlit's global stylesheet. Plotly charts and tables remain native Streamlit components because iframes isolate JavaScript events, which would break Plotly's client-side interactions (hover tooltips, zooming, panning) and `st.dataframe` column sorting.

### Q5: What is the function of the `ui.LAST_RENDER` dictionary in the software test harness?
**Answer:** Because `components.html` content is rendered inside an isolated browser iframe, the internal DOM text is inaccessible to Streamlit's headless `AppTest` test runner. The `ui.LAST_RENDER` dictionary acts as a server-side test seam, recording the exact key-value pairs formatted for card rendering. Automated test `TC-T-45` inspects this dictionary to assert that published figures are non-null, properly formatted, and match database fixtures.

### Q6: How does the application prevent spreadsheet formula injection (CSV injection) during data export?
**Answer:** In Screen UI-09 (`pages/8_Data.py`), before generating the downloadable CSV payload, all string cells that begin with formula trigger characters (`=`, `+`, `-`, `@`) are automatically sanitized by prepending a single quote (`'`). This neutralizes executable macros and formula execution when the downloaded file is opened in Microsoft Excel or Google Sheets.

### Q7: What is the difference between ward-day occupancy and episode bed-days in hospital operational accounting?
**Answer:** Ward-day occupancy aggregates the census of patients physically occupying a bed on each midnight census across the calendar year (80,083 occupied bed-days in 2025). Episode bed-days sum the total length of stay for patients discharged within that calendar year (79,089 days). The 994-day delta represents inpatient stays that straddle midnight on December 31st. Retaining both distinct measures prevents distortion in occupancy rates versus billing-per-day metrics.

### Q8: Why is Spearman rank correlation used instead of Pearson linear correlation to test the association between Bed Occupancy and Readmission?
**Answer:** Clinical length of stay and hospital occupancy distributions exhibit significant skewness and extreme outliers (e.g., long-stay complex cases up to 90 days). Pearson correlation requires bivariate normality and assumes strictly linear relationships, making it vulnerable to distortion by outliers. Spearman rank correlation evaluates monotonic relationships using ordinal ranks, providing a robust, non-parametric test of operational pressure versus clinical rebound.

### Q9: How does the Streamlit caching strategy (`@st.cache_data`) maintain performance under concurrent user sessions?
**Answer:** The primary database loader function in `src/data.py` uses `@st.cache_data`, serializing the parsed fact and dimension datasets into memory after initial execution. Subsequent user interactions and page navigations filter the cached in-memory DataFrames rather than repeatedly querying SQLite on disk, keeping warm-load page latency well under the strict 3.0-second non-functional requirement budget (`NFR-01`).

### Q10: How does the Requirements Traceability Matrix (RTM) demonstrate verification of the negative scope requirement `BR-15`?
**Answer:** Negative scope requirements ('Won't have' features, such as automated external email/SMS alerting) cannot be validated with functional feature tests. In the RTM, `BR-15` is mapped to `TC-F-46`, which performs an explicit absence check, verifying that no unconfigured background messaging tasks or external alerting APIs exist within the codebase, thereby ensuring strict boundary governance against scope creep.
