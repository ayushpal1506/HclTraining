# p107 — Healthcare KPI Dashboard (Clinical & Operational Analytics)

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://2500320140037-p107.streamlit.app)
[![Python 3.12+](https://img.shields.io/badge/python-3.12%2B-blue.svg)](https://www.python.org/)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
[![Tests: pytest](https://img.shields.io/badge/tests-70%20passed-success.svg)](https://pytest.org/)

An enterprise-grade operational intelligence and clinical analytics dashboard designed for hospital executives, medical superintendents, and financial controllers. The platform provides unified, real-time visibility into bed occupancy bottlenecks, average length of stay (ALOS), 30-day readmissions, and financial yield across 197,054 verified patient episodes.

---

## 📋 Academic & Programme Details

| Field | Details |
|:---|:---|
| **Project ID** | **`p107`** |
| **Project Title** | Healthcare KPI Dashboard (Clinical & Operational Analytics) |
| **Learner Name** | **Ayush Pal** |
| **Roll Number** | **`2500320140037`** |
| **Mentor** | **Mr. Dinesh Sereno** |
| **Batch ID** | **`ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01`** |
| **Programme** | MCA 2nd Year (Data Analytics) — HCLTech Career Shaper |
| **College** | ABES Engineering College, Ghaziabad |
| **GitHub Repository** | [https://github.com/ayushpal1506/HclTraining](https://github.com/ayushpal1506/HclTraining) |
| **Live Web App** | [https://2500320140037-p107.streamlit.app](https://2500320140037-p107.streamlit.app) |

---

## 🌐 Live Application Links

- **Live Streamlit App:** [https://2500320140037-p107.streamlit.app](https://2500320140037-p107.streamlit.app)
- **GitHub Repository:** [https://github.com/ayushpal1506/HclTraining](https://github.com/ayushpal1506/HclTraining)

---

## 📁 Repository Structure

```text
HclTraining/
├── app.py                      # Main entrypoint — Executive Overview (UI-01)
├── requirements.txt            # Production dependencies for Streamlit Cloud
├── requirements-dev.txt        # Development, testing, and report-generation dependencies
├── pytest.ini                  # Pytest configuration and test path definitions
├── BUILD_PLAN.md               # Master engineering build plan & traceability
├── CLAUDE.md                   # Project specifications & governance rules
├── README.md                   # Project documentation & GitHub profile
│
├── .github/
│   └── workflows/
│       └── ci.yml              # GitHub Actions CI/CD automated test workflow
│
├── .streamlit/
│   └── config.toml             # Streamlit visual theme tokens (Navy/Slate palette)
│
├── assets/
│   ├── abes_logo.jpg           # Official college shield logo
│   └── github_repo_screenshot.jpg # GitHub repository verification screenshot
│
├── data/
│   ├── app.db                  # Relational SQLite warehouse (197,054 records, 34 MB)
│   ├── healthcare_eda_dataset.csv # Companion tabular dataset for exploratory analysis
│   ├── validation_report.md    # Automated relational DDL & data integrity audit
│   └── quality_log.csv         # Cleansing reject log (5,368 injected anomaly audits)
│
├── docs/                       # Academic artifacts & project deliverables (Steps 01-18)
│   ├── 01_Project_Understanding.docx
│   ├── 02_HMW_Statements.docx
│   ├── 03_Needs_Analysis.docx
│   ├── 04_Traceability_Matrix.xlsx
│   ├── 05_BRD.docx
│   ├── 07_TRD.docx
│   ├── 08_Database_Design.xlsx
│   ├── 09_ER_Diagram.docx
│   ├── 10_Dataset_Generation.docx
│   ├── 11_User_Stories.docx
│   ├── 12_UI_Framework.docx
│   ├── 13_Test_Cases.docx
│   ├── 14_BUILD_PLAN.docx
│   ├── 15_Build_Test_Fix_Report.docx
│   ├── 16_Deployment_Report.docx
│   ├── HKD-2026_Project_Document.docx  # Primary 18-page Project Document (Word)
│   ├── HKD-2026_Project_Document.pdf   # Formatted Project Document (PDF)
│   ├── HKD-2026_Presentation.pptx      # 11-slide master presentation deck (PPTX)
│   ├── HKD-2026_Presentation.pdf       # 11-slide presentation deck (PDF)
│   ├── HKD-2026_GitHub_Report.pdf      # Repository verification & screenshot report
│   ├── HKD-2026_Complete_Code.pdf      # Full formatted source code catalog
│   ├── slide_images/                   # 1600x900 full-bleed slide images (01-11)
│   └── step_pdfs/                      # Individual PDFs for all 17 steps
│
├── notebooks/
│   └── EDA_Lesson_Notebook.ipynb       # 13-section hands-on EDA masterclass notebook
│
├── pages/                              # Streamlit Multi-Page UI Architecture
│   ├── 1_Operations.py         # UI-02: Acute Bed Occupancy & 85% Bottleneck Warning
│   ├── 2_Departments.py        # UI-03: ALOS Benchmarking & 30-Day Readmission CIs
│   ├── 3_Quality.py            # UI-04: Mortality Drift Monitoring (+2σ SPC alerts)
│   ├── 4_Finance.py            # UI-05: Payer Mix Shift & Ranked Revenue per Bed-Day
│   ├── 5_Data_Quality.py       # UI-06: 6-Gate Relational Data Quality Scorecard
│   ├── 6_Definitions.py        # UI-07: Master KPI Registry & Clinical Governance
│   ├── 7_Insights.py           # UI-08: Revenue Waterfall & Spearman Correlation
│   └── 8_Data.py               # UI-09: Filtered Cohort Exploration & Secure CSV Export
│
├── scripts/                            # Engineering automation & compilation scripts
│   ├── build_step17_abes.py    # Strict ABES-format Project Document compiler
│   ├── build_step18_perfect.py # Full-bleed PPTX and Pillow presentation renderer
│   ├── generate_lms_artifacts.py # LMS PDF generator & ZIP packager
│   └── regen_project_doc_pdf.py # Standalone PDF converter with inline image support
│
├── src/                                # Core computational backend & UI library
│   ├── data.py                 # SQLite data loader with @st.cache_data memory caching
│   ├── filters.py              # Synchronous global sidebar filters & session state
│   ├── kpi.py                  # Pure mathematical KPI routines & Wilson 95% CIs
│   ├── stats.py                # Spearman rank correlation & Student-t p-values
│   └── ui.py                   # Tailwind CSS encapsulated iframe components
│
├── submission_project_folder/          # Clean 6-deliverable LMS submission package
│   ├── 1_Project_Document.pdf  # Project Document (17 pages, with ABES logo)
│   ├── 2_Project_Presentation.pdf # Presentation Deck (11 high-res slides)
│   ├── 2_Project_Presentation.pptx # Editable PowerPoint deck with speaker notes
│   ├── 3_GitHub_Report.pdf     # GitHub verification & screenshot report
│   ├── 4_Complete_Code.pdf     # Combined formatted source code listing
│   ├── 5_Project_Documents_All_Steps.pdf # Merged 256-page PDF of all step docs
│   ├── 6_Project_ZIP_File.zip  # Standalone project ZIP archive (21 MB)
│   └── README_LMS_Submission.txt # Upload instructions & metadata checklist
│
└── tests/                              # Automated testing suite (pytest)
    ├── conftest.py             # Shared fixtures and database connection provider
    ├── test_schema.py          # Relational DDL & foreign key constraints
    ├── test_kpi.py             # Mathematical precision of formulas & Wilson CIs
    └── test_navigation.py      # Streamlit AppTest headless navigation assertions
```

---

## 🖥️ Application Screens & Features

| Screen | Page File | Primary Stakeholder | Key Features |
|:---|:---|:---|:---|
| **UI-01: Executive Overview** | `app.py` | CEO / Operations Director | 6 unified KPI metric cards, period baseline comparators, directional trend indicators |
| **UI-02: Operations** | `pages/1_Operations.py` | Operations Director | Acute bed occupancy gauge vs 85% NABH threshold, beds needed to resolve bottleneck, 7-day rolling ward forecast |
| **UI-03: Departments** | `pages/2_Departments.py` | Medical Superintendent | Dumbbell plot comparing ALOS vs 4.0d standard, 30-day readmission rates with closed-form Wilson 95% CIs |
| **UI-04: Quality** | `pages/3_Quality.py` | Clinical Quality Head | Statistical Process Control (SPC) monthly mortality drift detection, abnormal month highlighting (> mean + 2σ) |
| **UI-05: Finance** | `pages/4_Finance.py` | Chief Financial Officer | 100% stacked payer mix evolution (Corporate vs Self-Pay), specialty ranking by revenue per occupied bed-day |
| **UI-06: Data Quality** | `pages/5_Data_Quality.py` | Quality Auditor / Steward | 6-gate relational integrity audit scorecard (PK, FK, LOS, Dates, Billing, Capacity), reject anomaly log |
| **UI-07: Definitions** | `pages/6_Definitions.py` | Clinical Auditor | Master KPI dictionary, mathematical formulas, governance stewards, and clinical targets |
| **UI-08: Insights** | `pages/7_Insights.py` | Lead Data Analyst | Departmental revenue movement waterfall with residual check, Spearman rank correlation test (BOR vs Readmissions) |
| **UI-09: Data & Export** | `pages/8_Data.py` | Research Analyst | Interactive cohort slicing, secure formula-injection-safe CSV export (UTF-8 BOM, 50,000 row safeguard) |

---

## 📊 Ground Truth Clinical Metrics (Seed 42)

- **Total Inpatient Episodes:** 53,915
- **Total OPD Consultations:** 143,139
- **Overall Hospital Bed Occupancy:** 84.6%
- **Critical ICU Bed Occupancy:** 93.8% *(breaches 85% safe capacity threshold)*
- **Average Length of Stay (ALOS):** 4.21 days *(Oncology: 6.49d, Maternity: 2.76d)*
- **30-Day Readmission Rate:** 12.4% overall *(General Medicine peak: 16.8%)*
- **Emergency Door-to-Doctor Time:** 38.4 minutes *(evening surge peak: 48m)*

---

## 🛠️ Technology Stack

- **Language:** Python 3.12 / 3.13
- **App Framework:** Streamlit 1.40+
- **Styling:** Tailwind CSS (encapsulated HTML iframes, responsive design)
- **Data Engine:** pandas 2.2.0+, numpy 1.26.0+
- **Database:** SQLite 3 (embedded `data/app.db`, read-only URI mode)
- **Visualizations:** Plotly Express & Graph Objects 5.24+
- **Testing:** pytest 8.0+, pytest-mock, Streamlit AppTest (70 passed tests)
- **Documentation:** python-docx, python-pptx, ReportLab, Pillow, pypdf

---

## 🚀 Local Installation & Execution

### 1. Clone the repository
```bash
git clone https://github.com/ayushpal1506/HclTraining.git
cd HclTraining
```

### 2. Create and activate a virtual environment
```bash
python3 -m venv .venv
source .venv/bin/activate    # On Windows: .venv\Scripts\activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the application
```bash
streamlit run app.py
```
Open your browser at `http://localhost:8501`.

---

## 🧪 Running Automated Tests

Execute the comprehensive test suite verifying schema DDL, mathematical KPI formulas, and headless navigation:
```bash
pytest
```
*Expected output: `70 passed, 1 skipped in ~10s`.*

---

## 🔒 Data Privacy & Compliance

All data in this repository is purely synthetic, deterministically generated under **seed 42** for clinical analytics training. **No real Patient Identifiable Information (PII)**, real hospital records, or sensitive personal credentials are used.

---

**Developed by Ayush Pal** · MCA 2nd Year Data Analyst · ABES Engineering College  
Under the guidance of **Mr. Dinesh Sereno** (HCLTech Career Shaper Programme)
