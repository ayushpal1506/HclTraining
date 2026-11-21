"""Build Step 8 deliverables: Database Design workbook + updated RTM (DB table.column filled)."""
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "docs", "08_Database_Design.xlsx")
RTM = os.path.join(ROOT, "docs", "04_Traceability_Matrix.xlsx")

NAVY = "1F3864"
BLUE = "2E5496"
LIGHT = "EAF0F8"
CODE = "F4F6F8"
WHITE = "FFFFFF"

thin = Side(style="thin", color="B4C6E7")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
hdr_font = Font(name="Consolas", size=9, bold=True, color=WHITE)
body_font = Font(name="Calibri", size=9)
mono_font = Font(name="Consolas", size=8.5)

# ------------------------------------------------------------------ data
TABLES = [
    ("T-01", "dim_department", "Master", "One row per clinical department",
     "department_id", "none", 6, 8, 8,
     "Closed list of the 8 clinical departments; drives department filters and joins."),
    ("T-02", "dim_ward", "Master", "One row per ward (functional beds per ward)",
     "ward_id", "department_id \u2192 T-01", 9, 20, 20,
     "Ward base capacity; supplies functional_beds, the denominator of occupancy (BRL-01)."),
    ("T-03", "dim_payer", "Master", "One row per payer category",
     "payer_id", "none", 5, 5, 5,
     "Closed five-value payer list (BRL-06); drives payer mix and revenue analysis."),
    ("T-04", "fact_inpatient_episode", "Fact", "One row per inpatient episode (admission \u2192 discharge)",
     "episode_id", "department_id \u2192 T-01; ward_id \u2192 T-02; payer_id \u2192 T-03; index_episode_id \u2192 T-04 (self)", 19, 260, 57000,
     "Main fact table: volume, ALOS, readmissions, revenue, mortality, demographics."),
    ("T-05", "fact_opd_visit", "Fact", "One row per outpatient visit",
     "visit_id", "department_id \u2192 T-01; payer_id \u2192 T-03", 10, 300, 150000,
     "Outpatient footfall, weekday/seasonal pattern, OPD demographics."),
]

DD = {}

DD["T-01"] = [
    ("department_id", "TEXT", "PK", "No", "\u2014", "DEP01\u2026DEP08", "Department surrogate key", "OQ-01\u2013OQ-10, BR-01\u2026BR-13"),
    ("department_name", "TEXT", "UNIQUE", "No", "\u2014", "Cardiology, General Medicine, Orthopaedics, Paediatrics, Maternity, Oncology, Emergency, ENT", "Display name used in filters and charts", "OQ-01\u2013OQ-10, BR-01"),
    ("specialty", "TEXT", "FK (to specialty dictionary in code)", "No", "\u2014", "Clinical specialty label", "Case-mix context for ALOS comparison (BR-06)", "BR-06"),
    ("floor_block", "TEXT", "none", "No", "\u2014", "Block A\u2026D", "Physical location; used in ward tooltips only", "BR-02"),
    ("is_active", "INTEGER", "none", "No", "1", "0, 1", "Soft-retire flag so historical data keeps a valid department", "BR-01, BR-14"),
    ("effective_from", "TEXT (ISO date)", "none", "No", "\u2014", "YYYY-MM-DD", "When the department was opened / entered the model", "BR-05"),
]

DD["T-02"] = [
    ("ward_id", "TEXT", "PK", "No", "\u2014", "W01\u2026W20", "Ward surrogate key", "BR-01, BR-02"),
    ("ward_name", "TEXT", "UNIQUE", "No", "\u2014", "e.g. Cardio-ICU, GenMed-A", "Ward label shown on the nursing view", "BR-02"),
    ("ward_type", "TEXT", "none", "No", "\u2014", "General, Semi-private, Private, ICU, HDU, Maternity, Paediatric, Day-care", "Care intensity; drives 90% saturation threshold logic", "BR-02"),
    ("bed_class", "TEXT", "none", "No", "\u2014", "Economy, Twin, Single, Intensive", "Tariff class used in revenue context", "BR-08, BR-12"),
    ("department_id", "TEXT", "FK \u2192 T-01", "No", "\u2014", "DEP01\u2026DEP08", "Owning department; enforces ward\u2192department roll-up", "BR-01, BR-02"),
    ("functional_beds", "INTEGER", "none", "No", "\u2014", "1\u201360", "Staffed beds available; denominator of BOR (BRL-01)", "BR-01, BR-02, BR-12"),
    ("is_active", "INTEGER", "none", "No", "1", "0, 1", "Ward closed or decommissioned", "BR-01"),
    ("effective_from", "TEXT (ISO date)", "none", "No", "\u2014", "YYYY-MM-DD", "SCD-2 start of the bed count", "BR-01, BR-05"),
    ("effective_to", "TEXT (ISO date)", "none", "Yes", "NULL", "YYYY-MM-DD or NULL", "SCD-2 end; NULL means currently effective", "BR-01"),
]

DD["T-03"] = [
    ("payer_id", "TEXT", "PK", "No", "\u2014", "P01\u2026P05", "Payer surrogate key", "BR-08, BR-12"),
    ("payer_name", "TEXT", "UNIQUE", "No", "\u2014", "Self-pay, Star Health, MedAssist TPA, PM-JAY, Infosys Corporate", "Specific payer within a category", "BR-08"),
    ("payer_category", "TEXT", "none", "No", "\u2014", "Cash, Private insurance, TPA, Government scheme, Corporate", "Closed list enforced by CHECK (BRL-06); sums to 100%", "BR-08, BR-14"),
    ("settlement_days", "INTEGER", "none", "No", "0", "0\u2013180", "Typical reimbursement delay; context for revenue discussion", "BR-08"),
    ("is_active", "INTEGER", "none", "No", "1", "0, 1", "Payer still contracted", "BR-08"),
]

DD["T-04"] = [
    ("episode_id", "INTEGER", "PK", "No", "autoincrement", "Positive integer", "Episode identifier; also the self-reference target for readmissions", "OQ-02\u2013OQ-04, OQ-08, OQ-10"),
    ("patient_key", "TEXT", "INDEX", "No", "\u2014", "12-char pseudonym (e.g. PT0000123456)", "Non-reversible synthetic patient key; enables readmission self-join and repeat-visit logic", "OQ-04, BR-07"),
    ("department_id", "TEXT", "FK \u2192 T-01", "No", "\u2014", "DEP01\u2026DEP08", "Treating department at discharge", "OQ-01\u2013OQ-10"),
    ("ward_id", "TEXT", "FK \u2192 T-02", "No", "\u2014", "W01\u2026W20", "Ward of stay (last ward for transfers)", "BR-02"),
    ("payer_id", "TEXT", "FK \u2192 T-03", "No", "\u2014", "P01\u2026P05", "Payer for this episode; NOT NULL is a data-quality rule", "OQ-05, BR-08"),
    ("admission_date", "TEXT (ISO date)", "INDEX", "No", "\u2014", "YYYY-MM-DD within 2023-01-01\u20262025-12-31", "Trend origin, weekday/season attributes, occupancy start", "OQ-02, OQ-06, OQ-08"),
    ("discharge_date", "TEXT (ISO date)", "INDEX", "No", "\u2014", "YYYY-MM-DD \u2265 admission_date", "Trend completion, occupancy end, readmission window origin", "OQ-01, OQ-03, OQ-04, OQ-10"),
    ("los_days", "INTEGER", "none", "No", "\u2014", "1\u2013120", "Length of stay in days; same-day discharge = 1 (BRL-03)", "OQ-03, BR-06"),
    ("discharge_type", "TEXT", "none", "No", "\u2014", "Routine, Transfer, LAMA, Expired", "Outcome category; drives discharge mix", "OQ-10, BR-13"),
    ("death_flag", "INTEGER", "none", "No", "0", "0, 1", "1 = in-hospital death; CHECK forces consistency with Expired", "OQ-10, BR-13"),
    ("readmit_30d_flag", "INTEGER", "INDEX", "No", "0", "0, 1", "1 = unplanned readmission within 30 days of the index discharge (BRL-04)", "OQ-04, BR-07"),
    ("index_episode_id", "INTEGER", "FK \u2192 T-04 (self)", "Yes", "NULL", "episode_id or NULL", "The earlier discharge that this readmission follows; NULL when not a readmission", "OQ-04, BR-07"),
    ("diagnosis_group", "TEXT", "INDEX", "No", "\u2014", "20-value grouped list (e.g. Cardiac, Orthopaedic, Respiratory, Obstetric, Oncology)", "Attribution dimension for readmission and ALOS review", "OQ-04, BR-07"),
    ("is_planned_procedure", "INTEGER", "none", "No", "0", "0, 1", "1 = planned procedure; excluded from readmission counting (BRL-04)", "OQ-04, BR-07"),
    ("age_band", "TEXT", "none", "No", "\u2014", "0-5, 6-17, 18-35, 36-50, 51-65, 66+", "Patient segment for service-line analysis", "OQ-08, BR-11"),
    ("gender", "TEXT", "none", "No", "\u2014", "M, F, O", "Patient segment; O = other/undisclosed", "OQ-08, BR-11"),
    ("billed_amount", "REAL", "none", "No", "\u2014", "\u2265 0, INR", "Gross billed amount before deductions (BRL-05)", "OQ-05, OQ-09, BR-08, BR-12"),
    ("bed_days", "INTEGER", "none", "No", "\u2014", "1\u2013121", "Midnights occupied, derived from los_days (BRL-02)", "OQ-01, OQ-09"),
    ("created_at", "TEXT (ISO timestamp)", "none", "No", "datetime('now')", "YYYY-MM-DD HH:MM:SS", "Audit of when the synthetic row was generated", "BR-14"),
]

DD["T-05"] = [
    ("visit_id", "INTEGER", "PK", "No", "autoincrement", "Positive integer", "Outpatient visit identifier", "OQ-06, OQ-08"),
    ("patient_key", "TEXT", "INDEX", "No", "\u2014", "12-char pseudonym", "Synthetic patient key; links an OPD patient to inpatient episodes", "OQ-08, BR-11"),
    ("visit_date", "TEXT (ISO date)", "INDEX", "No", "\u2014", "YYYY-MM-DD within 2023-01-01\u20262025-12-31", "Footfall date; source of weekday and month attributes", "OQ-06, BR-09"),
    ("department_id", "TEXT", "FK \u2192 T-01", "No", "\u2014", "DEP01\u2026DEP08", "OPD department of the visit", "OQ-06, BR-09"),
    ("payer_id", "TEXT", "FK \u2192 T-03", "No", "\u2014", "P01\u2026P05", "Payer for the consultation", "OQ-05, BR-08"),
    ("age_band", "TEXT", "none", "No", "\u2014", "0-5, 6-17, 18-35, 36-50, 51-65, 66+", "OPD patient segment", "OQ-08, BR-11"),
    ("gender", "TEXT", "none", "No", "\u2014", "M, F, O", "OPD patient segment", "OQ-08, BR-11"),
    ("visit_type", "TEXT", "none", "No", "\u2014", "New, Follow-up", "Visit mix; follow-ups are a capacity driver", "OQ-06, BR-09"),
    ("consultation_fee", "REAL", "none", "No", "\u2014", "\u2265 0, INR", "Fee charged; kept separate from inpatient revenue", "BR-08"),
    ("created_at", "TEXT (ISO timestamp)", "none", "No", "datetime('now')", "YYYY-MM-DD HH:MM:SS", "Generation audit", "BR-14"),
]

RELATIONSHIPS = [
    ("T-01 dim_department", "T-02 dim_ward", "department_id", "1 : N", "RESTRICT", "A department owns many wards; a ward belongs to exactly one department"),
    ("T-01 dim_department", "T-04 fact_inpatient_episode", "department_id", "1 : N", "RESTRICT", "Every episode is attributed to one department at discharge"),
    ("T-01 dim_department", "T-05 fact_opd_visit", "department_id", "1 : N", "RESTRICT", "Every OPD visit belongs to one department"),
    ("T-02 dim_ward", "T-04 fact_inpatient_episode", "ward_id", "1 : N", "RESTRICT", "Every episode occupies one ward (last ward after transfers)"),
    ("T-03 dim_payer", "T-04 fact_inpatient_episode", "payer_id", "1 : N", "RESTRICT", "One payer per episode; NULL not allowed (data-quality rule)"),
    ("T-03 dim_payer", "T-05 fact_opd_visit", "payer_id", "1 : N", "RESTRICT", "One payer per visit"),
    ("T-04 fact_inpatient_episode", "T-04 fact_inpatient_episode", "index_episode_id", "0..1 : N", "SET NULL", "Self-reference: a readmission points to the earlier index discharge"),
    ("T-02 dim_ward", "T-01 dim_department", "department_id", "\u2014 (inverse)", "\u2014", "Ward-to-department consistency is enforced by trigger trg_episode_ward_dept on T-04"),
]

KPIS = [
    ("Bed occupancy rate (BOR)", "100 * SUM(bed_days_occupied) / SUM(functional_beds * days_in_period)", "OQ-01", "Department \u00d7 week/month", "T-02, T-04", "%"),
    ("Occupancy gap to target", "BOR \u2212 85", "OQ-01", "Department \u00d7 week/month", "T-02, T-04", "pp"),
    ("Shortfall in beds", "ROUND((85/100 * bed_days_available \u2212 bed_days_occupied) / days_in_period, 0)", "OQ-01", "Department \u00d7 week", "T-02, T-04", "beds"),
    ("Average length of stay (ALOS)", "1.0 * SUM(los_days) / COUNT(*) [discharges]", "OQ-03", "Department \u00d7 month", "T-04", "days"),
    ("Case-mix context", "Share of each diagnosis_group within the department's discharges", "OQ-03", "Department \u00d7 month", "T-04", "%"),
    ("30-day readmission rate", "100 * SUM(readmit_30d_flag) / COUNT(*) [index discharges]", "OQ-04", "Department / diagnosis \u00d7 month", "T-04", "%"),
    ("Readmission contribution", "Department readmissions / total readmissions", "OQ-04", "Department \u00d7 period", "T-04", "%"),
    ("Revenue per discharge", "SUM(billed_amount) / COUNT(*)", "OQ-05", "Payer \u00d7 month", "T-03, T-04", "INR"),
    ("Payer share", "SUM(billed_amount) by payer / SUM(billed_amount) total", "OQ-05", "Payer \u00d7 month", "T-03, T-04", "%"),
    ("OPD peak-to-average ratio", "MAX(avg visits by weekday) / overall average daily visits", "OQ-06", "Department \u00d7 month", "T-05", "ratio"),
    ("Revenue per bed-day", "SUM(billed_amount) / SUM(bed_days)", "OQ-09", "Department \u00d7 month", "T-02, T-04", "INR"),
    ("Gap to best revenue per bed-day", "MAX(dept revenue per bed-day) \u2212 dept value", "OQ-09", "Department \u00d7 month", "T-02, T-04", "INR"),
    ("In-hospital mortality rate", "100 * SUM(death_flag) / COUNT(*)", "OQ-10", "Department \u00d7 month", "T-04", "%"),
    ("Discharge-type mix", "Share of each discharge_type within discharges", "OQ-10", "Department \u00d7 month", "T-04", "%"),
    ("Drift flag", "1 if value outside mean \u00b1 2*SD of the previous 12 months", "OQ-10 (and BR-13)", "Department \u00d7 month", "T-04", "flag"),
    ("Week-on-week change", "(current week \u2212 prior week) / prior week", "OQ-02", "Week", "T-04, T-05", "%"),
    ("Movement contribution", "(department delta) / (total delta) for the same KPI and periods", "OQ-02 (BR-03)", "Department / payer / weekday", "T-04, T-05", "%"),
    ("Low-confidence flag", "1 if denominator < 30 (BRL-10)", "All rates", "Any", "T-04, T-05", "flag"),
    ("Occupancy validity", "BOR > 100 \u2192 reject as data defect (BRL-01)", "OQ-01", "Department \u00d7 month", "T-02, T-04", "flag"),
]

DDL = [
    "PRAGMA foreign_keys = ON;",
    "",
    "CREATE TABLE dim_department (",
    "  department_id   TEXT    PRIMARY KEY CHECK (length(department_id) = 5),",
    "  department_name TEXT    NOT NULL UNIQUE,",
    "  specialty       TEXT    NOT NULL,",
    "  floor_block     TEXT    NOT NULL,",
    "  is_active       INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0,1)),",
    "  effective_from  TEXT    NOT NULL CHECK (date(effective_from) IS NOT NULL)",
    ") STRICT;",
    "",
    "CREATE TABLE dim_ward (",
    "  ward_id         TEXT    PRIMARY KEY,",
    "  ward_name       TEXT    NOT NULL UNIQUE,",
    "  ward_type       TEXT    NOT NULL CHECK (ward_type IN",
    "                    ('General','Semi-private','Private','ICU','HDU','Maternity','Paediatric','Day-care')),",
    "  bed_class       TEXT    NOT NULL CHECK (bed_class IN ('Economy','Twin','Single','Intensive')),",
    "  department_id   TEXT    NOT NULL REFERENCES dim_department(department_id)",
    "                    ON UPDATE CASCADE ON DELETE RESTRICT,",
    "  functional_beds INTEGER NOT NULL CHECK (functional_beds BETWEEN 1 AND 60),",
    "  is_active       INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0,1)),",
    "  effective_from  TEXT    NOT NULL,",
    "  effective_to    TEXT    NULL,",
    "  CHECK (effective_to IS NULL OR effective_to >= effective_from)",
    ") STRICT;",
    "",
    "CREATE TABLE dim_payer (",
    "  payer_id        TEXT    PRIMARY KEY,",
    "  payer_name      TEXT    NOT NULL UNIQUE,",
    "  payer_category  TEXT    NOT NULL CHECK (payer_category IN",
    "                    ('Cash','Private insurance','TPA','Government scheme','Corporate')),",
    "  settlement_days INTEGER NOT NULL DEFAULT 0 CHECK (settlement_days BETWEEN 0 AND 180),",
    "  is_active       INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0,1))",
    ") STRICT;",
    "",
    "CREATE TABLE fact_inpatient_episode (",
    "  episode_id          INTEGER PRIMARY KEY,",
    "  patient_key         TEXT    NOT NULL CHECK (length(patient_key) = 12),",
    "  department_id       TEXT    NOT NULL REFERENCES dim_department(department_id),",
    "  ward_id             TEXT    NOT NULL REFERENCES dim_ward(ward_id),",
    "  payer_id            TEXT    NOT NULL REFERENCES dim_payer(payer_id),",
    "  admission_date      TEXT    NOT NULL CHECK (date(admission_date) IS NOT NULL),",
    "  discharge_date      TEXT    NOT NULL CHECK (date(discharge_date) IS NOT NULL),",
    "  los_days            INTEGER NOT NULL CHECK (los_days BETWEEN 1 AND 120),",
    "  discharge_type      TEXT    NOT NULL CHECK (discharge_type IN ('Routine','Transfer','LAMA','Expired')),",
    "  death_flag          INTEGER NOT NULL CHECK (death_flag IN (0,1)),",
    "  readmit_30d_flag    INTEGER NOT NULL DEFAULT 0 CHECK (readmit_30d_flag IN (0,1)),",
    "  index_episode_id    INTEGER NULL REFERENCES fact_inpatient_episode(episode_id) ON DELETE SET NULL,",
    "  diagnosis_group     TEXT    NOT NULL,",
    "  is_planned_procedure INTEGER NOT NULL DEFAULT 0 CHECK (is_planned_procedure IN (0,1)),",
    "  age_band            TEXT    NOT NULL CHECK (age_band IN ('0-5','6-17','18-35','36-50','51-65','66+')),",
    "  gender              TEXT    NOT NULL CHECK (gender IN ('M','F','O')),",
    "  billed_amount       REAL    NOT NULL CHECK (billed_amount >= 0),",
    "  bed_days            INTEGER NOT NULL CHECK (bed_days BETWEEN 1 AND 121),",
    "  created_at          TEXT    NOT NULL DEFAULT (datetime('now')),",
    "  CHECK (discharge_date >= admission_date),",
    "  CHECK (death_flag = 0 OR discharge_type = 'Expired'),",
    "  CHECK (discharge_type <> 'Expired' OR death_flag = 1),",
    "  CHECK (readmit_30d_flag = 0 OR index_episode_id IS NOT NULL),",
    "  CHECK (index_episode_id IS NULL OR index_episode_id <> episode_id),",
    "  CHECK (los_days = CAST(julianday(discharge_date) - julianday(admission_date) AS INTEGER) + 1),",
    "  CHECK (bed_days = los_days),",
    "  UNIQUE (patient_key, admission_date)",
    ") STRICT;",
    "",
    "CREATE TABLE fact_opd_visit (",
    "  visit_id         INTEGER PRIMARY KEY,",
    "  patient_key      TEXT    NOT NULL CHECK (length(patient_key) = 12),",
    "  visit_date       TEXT    NOT NULL CHECK (date(visit_date) IS NOT NULL),",
    "  department_id    TEXT    NOT NULL REFERENCES dim_department(department_id),",
    "  payer_id         TEXT    NOT NULL REFERENCES dim_payer(payer_id),",
    "  age_band         TEXT    NOT NULL CHECK (age_band IN ('0-5','6-17','18-35','36-50','51-65','66+')),",
    "  gender           TEXT    NOT NULL CHECK (gender IN ('M','F','O')),",
    "  visit_type       TEXT    NOT NULL CHECK (visit_type IN ('New','Follow-up')),",
    "  consultation_fee REAL    NOT NULL CHECK (consultation_fee >= 0),",
    "  created_at       TEXT    NOT NULL DEFAULT (datetime('now')),",
    "  UNIQUE (patient_key, visit_date, department_id)",
    ") STRICT;",
    "",
    "-- Indexes tuned to the filter and join paths used by the app",
    "CREATE INDEX idx_ep_dept_adm    ON fact_inpatient_episode(department_id, admission_date);",
    "CREATE INDEX idx_ep_dept_dis    ON fact_inpatient_episode(department_id, discharge_date);",
    "CREATE INDEX idx_ep_discharge   ON fact_inpatient_episode(discharge_date);",
    "CREATE INDEX idx_ep_payer       ON fact_inpatient_episode(payer_id);",
    "CREATE INDEX idx_ep_patient     ON fact_inpatient_episode(patient_key);",
    "CREATE INDEX idx_ep_diag        ON fact_inpatient_episode(diagnosis_group);",
    "CREATE INDEX idx_ep_readmit     ON fact_inpatient_episode(readmit_30d_flag) WHERE readmit_30d_flag = 1;",
    "CREATE INDEX idx_opd_date       ON fact_opd_visit(visit_date);",
    "CREATE INDEX idx_opd_dept_date  ON fact_opd_visit(department_id, visit_date);",
    "CREATE INDEX idx_opd_payer      ON fact_opd_visit(payer_id);",
    "",
    "-- Ward must belong to the same department as the episode",
    "CREATE TRIGGER trg_episode_ward_dept",
    "BEFORE INSERT ON fact_inpatient_episode",
    "FOR EACH ROW",
    "WHEN (SELECT department_id FROM dim_ward WHERE ward_id = NEW.ward_id) <> NEW.department_id",
    "BEGIN",
    "  SELECT RAISE(ABORT, 'ward_id does not belong to department_id');",
    "END;",
]

VIEW_ANALYSIS = [
    "-- vw_analysis_episode : 36 columns, episode grain, one row per episode.",
    "-- week_of_year is strftime('%W') (Monday-based); ISO-8601 week numbers are computed in pandas,",
    "-- because SQLite cannot produce ISO week 53 correctly (BRL-08).",
    "CREATE VIEW vw_analysis_episode AS",
    "SELECT",
    "  e.episode_id, e.patient_key, e.admission_date, e.discharge_date,",
    "  e.los_days, e.discharge_type, e.death_flag, e.readmit_30d_flag,",
    "  e.index_episode_id, e.diagnosis_group, e.is_planned_procedure,",
    "  e.age_band, e.gender, e.billed_amount, e.bed_days,",
    "  e.payer_id, p.payer_name, p.payer_category,",
    "  e.department_id, d.department_name, d.specialty,",
    "  e.ward_id, w.ward_name, w.ward_type, w.bed_class, w.functional_beds,",
    "  CAST(strftime('%Y', e.admission_date) AS INTEGER) AS admission_year,",
    "  strftime('%Y-%m', e.admission_date)                 AS admission_month,",
    "  CAST(strftime('%W', e.admission_date) AS INTEGER)   AS week_of_year,",
    "  CASE strftime('%w', e.admission_date)",
    "    WHEN '0' THEN 'Sunday'   WHEN '1' THEN 'Monday'",
    "    WHEN '2' THEN 'Tuesday'  WHEN '3' THEN 'Wednesday'",
    "    WHEN '4' THEN 'Thursday' WHEN '5' THEN 'Friday'",
    "    ELSE 'Saturday' END                                AS weekday_name,",
    "  'Q' || ((CAST(strftime('%m', e.admission_date) AS INTEGER) - 1) / 3 + 1) AS admission_quarter,",
    "  CASE WHEN strftime('%w', e.admission_date) IN ('0','6') THEN 1 ELSE 0 END AS is_weekend_admission,",
    "  CASE",
    "    WHEN CAST(strftime('%m', e.admission_date) AS INTEGER) IN (12,1,2) THEN 'Winter'",
    "    WHEN CAST(strftime('%m', e.admission_date) AS INTEGER) IN (3,4,5)  THEN 'Summer'",
    "    WHEN CAST(strftime('%m', e.admission_date) AS INTEGER) IN (6,7,8,9) THEN 'Monsoon'",
    "    ELSE 'Post-monsoon' END                            AS season,",
    "  CASE WHEN e.los_days <= 2 THEN '0-2' WHEN e.los_days <= 4 THEN '3-4'",
    "       WHEN e.los_days <= 7 THEN '5-7' ELSE '8+' END   AS los_band,",
    "  CASE WHEN e.billed_amount < 25000  THEN 'Low'",
    "       WHEN e.billed_amount < 75000  THEN 'Medium'",
    "       WHEN e.billed_amount < 200000 THEN 'High'",
    "       ELSE 'Very high' END                            AS revenue_band,",
    "  CASE WHEN COUNT(*) OVER (",
    "         PARTITION BY e.patient_key",
    "         ORDER BY e.admission_date, e.episode_id",
    "         ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING) > 0",
    "       THEN 1 ELSE 0 END                               AS is_repeat_patient",
    "FROM fact_inpatient_episode e",
    "JOIN dim_department d ON d.department_id = e.department_id",
    "JOIN dim_ward       w ON w.ward_id       = e.ward_id",
    "JOIN dim_payer      p ON p.payer_id      = e.payer_id;",
]

VIEW_DEPT_MONTH = [
    "-- vw_dept_month_kpi : one row per department-month, all six KPI families pre-aggregated.",
    "-- Requires vw_ward_day_beds to exist first (bed-days are allocated to the month they occur in,",
    "-- not to the discharge month, so occupancy cannot exceed 100% at month boundaries).",
    "CREATE VIEW vw_dept_month_kpi AS",
    "WITH beds AS (",
    "  SELECT department_id,",
    "         strftime('%Y-%m', day) AS month,",
    "         SUM(functional_beds)   AS bed_days_available,",
    "         SUM(beds_occupied)     AS bed_days_occupied",
    "  FROM vw_ward_day_beds",
    "  GROUP BY 1,2",
    "),",
    "ep AS (",
    "  SELECT department_id,",
    "         strftime('%Y-%m', discharge_date) AS month,",
    "         COUNT(*)                           AS discharges,",
    "         SUM(los_days)                      AS inpatient_days,",
    "         SUM(billed_amount)                 AS revenue,",
    "         SUM(death_flag)                    AS deaths,",
    "         SUM(readmit_30d_flag)              AS readmissions",
    "  FROM fact_inpatient_episode",
    "  GROUP BY 1,2",
    "),",
    "opd AS (",
    "  SELECT department_id, strftime('%Y-%m', visit_date) AS month, COUNT(*) AS opd_visits",
    "  FROM fact_opd_visit GROUP BY 1,2",
    ")",
    "SELECT b.department_id, b.month,",
    "       COALESCE(ep.discharges, 0)      AS discharges,",
    "       COALESCE(ep.inpatient_days, 0)  AS inpatient_days,",
    "       b.bed_days_occupied,",
    "       b.bed_days_available,",
    "       ROUND(100.0 * b.bed_days_occupied / NULLIF(b.bed_days_available,0), 2) AS bor_pct,",
    "       ROUND(1.0 * ep.inpatient_days / NULLIF(ep.discharges,0), 2)            AS alos,",
    "       COALESCE(ep.readmissions, 0)    AS readmissions,",
    "       ROUND(100.0 * ep.readmissions / NULLIF(ep.discharges,0), 2)            AS readmission_rate,",
    "       COALESCE(ep.revenue, 0)         AS revenue,",
    "       ROUND(1.0 * ep.revenue / NULLIF(ep.discharges,0), 2)                   AS revenue_per_discharge,",
    "       ROUND(1.0 * ep.revenue / NULLIF(b.bed_days_occupied,0), 2)             AS revenue_per_bed_day,",
    "       COALESCE(ep.deaths, 0)          AS deaths,",
    "       ROUND(100.0 * ep.deaths / NULLIF(ep.discharges,0), 2)                  AS mortality_rate,",
    "       COALESCE(opd.opd_visits, 0)     AS opd_visits,",
    "       CASE WHEN COALESCE(ep.discharges,0) < 30 THEN 1 ELSE 0 END             AS low_confidence",
    "FROM beds b",
    "LEFT JOIN ep  ON ep.department_id  = b.department_id  AND ep.month  = b.month",
    "LEFT JOIN opd ON opd.department_id = b.department_id  AND opd.month = b.month;",
]

VIEW_WARD_DAY = [
    "-- vw_ward_day_beds : ward-day occupancy, used by BR-02 (next-7-day ward view)",
    "CREATE VIEW vw_ward_day_beds AS",
    "WITH RECURSIVE cal(day) AS (",
    "  SELECT MIN(admission_date) FROM fact_inpatient_episode",
    "  UNION ALL",
    "  SELECT date(day, '+1 day') FROM cal",
    "  WHERE day < (SELECT MAX(discharge_date) FROM fact_inpatient_episode)",
    "),",
    "stays AS (",
    "  SELECT e.ward_id, c.day,",
    "         COUNT(DISTINCT e.episode_id) AS beds_occupied,",
    "         MAX(w.functional_beds)       AS functional_beds,",
    "         MAX(w.ward_name)             AS ward_name,",
    "         MAX(w.ward_type)             AS ward_type,",
    "         MAX(w.department_id)         AS department_id",
    "  FROM fact_inpatient_episode e",
    "  JOIN cal c ON c.day >= e.admission_date",
    "             AND c.day <  date(e.discharge_date, '+1 day')",
    "  JOIN dim_ward w ON w.ward_id = e.ward_id",
    "                 AND c.day BETWEEN w.effective_from",
    "                               AND COALESCE(w.effective_to, '9999-12-31')",
    "  GROUP BY 1,2",
    ")",
    "SELECT day, ward_id, ward_name, ward_type, department_id,",
    "       functional_beds, beds_occupied,",
    "       ROUND(100.0 * beds_occupied / NULLIF(functional_beds,0), 2) AS bor_pct",
    "FROM stays;",
]

OQ_MAP = [
    ("OQ-01", "Department occupancy vs 85% target", "T-02 + T-04", "vw_dept_month_kpi, vw_ward_day_beds"),
    ("OQ-02", "Admissions/discharges/OPD trend week over week", "T-04 + T-05", "vw_analysis_episode"),
    ("OQ-03", "ALOS by department and case mix", "T-04", "vw_dept_month_kpi"),
    ("OQ-04", "30-day readmissions by department/diagnosis", "T-04 (self-join)", "vw_analysis_episode"),
    ("OQ-05", "Revenue per patient by payer", "T-03 + T-04", "vw_dept_month_kpi"),
    ("OQ-06", "OPD weekday/season peaks", "T-05", "vw_analysis_episode (weekday_name) / pandas"),
    ("OQ-07", "Occupancy vs ALOS/readmission association", "Derived", "vw_dept_month_kpi"),
    ("OQ-08", "Age/gender mix by department", "T-04 + T-05", "vw_analysis_episode"),
    ("OQ-09", "Revenue per bed-day ranking", "T-02 + T-04", "vw_dept_month_kpi"),
    ("OQ-10", "Mortality and discharge mix drift", "T-04", "vw_dept_month_kpi"),
]

# ------------------------------------------------------------------ workbook
wb = Workbook()


def sheet_title(ws, title, subtitle, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(name="Calibri", size=13, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 24
    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
        c = ws.cell(row=2, column=1, value=subtitle)
        c.font = Font(name="Calibri", size=9, italic=True, color="404040")
        c.fill = PatternFill("solid", fgColor=LIGHT)
        c.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[2].height = 22


def write_table(ws, start_row, headers, rows, widths, mono_cols=(), wrap_cols=(),
                fits=True, freeze=None):
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=j, value=h)
        cell.font = hdr_font
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[start_row].height = 28
    for i, r in enumerate(rows):
        rr = start_row + 1 + i
        for j, v in enumerate(r, start=1):
            cell = ws.cell(row=rr, column=j, value=v)
            cell.font = mono_font if j in mono_cols else body_font
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=(j in wrap_cols or not fits),
                                       horizontal="left")
    if freeze:
        ws.freeze_panes = freeze
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.auto_filter.ref = (f"A{start_row}:"
                          f"{get_column_letter(len(headers))}{start_row + len(rows)}")


# ---------------- 1 Overview ----------------
ws = wb.active
ws.title = "1 Overview"
ws.sheet_view.showGridLines = False
sheet_title(ws, "p107  \u00b7  Database Design  \u00b7  v1.0  (19-Nov-2026)",
            "Five tables (3 masters + 2 facts), two analysis views and one derived KPI view. SQLite, read-only at runtime. "
            "IDs T-01\u2026T-05 are frozen; columns may only be added by change note.", 6)

r = 4
ws.cell(row=r, column=1, value="1. Design at a glance").font = Font(size=11, bold=True, color=NAVY)
r += 1
facts = [
    ("Tables", "5 (T-01 dim_department, T-02 dim_ward, T-03 dim_payer, T-04 fact_inpatient_episode, T-05 fact_opd_visit)"),
    ("Grain", "T-04 = one inpatient episode; T-05 = one outpatient visit; masters = one row per entity"),
    ("Keys", "Surrogate PKs on every table; FK enforcement ON; self-FK for readmission; CHECK, UNIQUE and a ward\u2192department trigger"),
    ("Analysis view", "vw_analysis_episode \u2014 36 columns, episode grain, used by OQ-02/04/06/08"),
    ("KPI view", "vw_dept_month_kpi \u2014 17 columns, one row per department-month, used by OQ-01/03/05/07/09/10"),
    ("Ward view", "vw_ward_day_beds \u2014 ward-day occupancy for BR-02"),
    ("Definition registry", "Not a table: a versioned config block (src/definitions.py) so the schema stays at five tables; BRL-14 gates access"),
    ("Runtime mode", "Opened read-only (mode=ro&immutable=1); writes happen only in scripts/build_db.py"),
]
rows = [[k, v] for k, v in facts]
write_table(ws, r, ["Item", "Decision"], rows, [22, 105], wrap_cols=(2,))
r = r + len(rows) + 3

ws.cell(row=r, column=1, value="2. OQ \u2192 table map (every question has a physical home)").font = Font(size=11, bold=True, color=NAVY)
r += 1
write_table(ws, r, ["OQ", "Operational question", "Tables used", "View / access path"],
            [[a, b, c, d] for a, b, c, d in OQ_MAP], [9, 48, 22, 46], wrap_cols=(2, 4), mono_cols=())
r = r + len(OQ_MAP) + 3

ws.cell(row=r, column=1, value="3. Row plan and why 150\u2013300 is not the production size").font = Font(size=11, bold=True, color=NAVY)
r += 1
notes = [
    ("Template expectation", "Step 8 asks for a main fact table of 150\u2013300 planned rows."),
    ("Project reality", "OBJ-02 commits to \u2265 18,000 admissions and \u2265 60,000 OPD rows over three years; a 300-row fact would break SC-02, the seasonality in BR-09 and the n = 288 correlation in OQ-07."),
    ("Resolution (change note DB-CN-01)", "150\u2013300 is kept as the design-review sample (T-04 = 260 rows, T-05 = 300 rows) used for schema review, ERD and pytest fixtures. Production volume is planned separately at ~57,000 episodes and ~150,000 visits."),
    ("Why that volume", "300 functional beds \u00d7 1,095 days = 328,500 bed-days available; at ~72% average occupancy = 236,520 bed-days occupied; \u00f7 4.1 ALOS \u2248 57,700 episodes. OPD at ~140 visits/day \u2248 153,000 visits."),
    ("Storage estimate", "Episode row \u2248 180 bytes \u2192 ~10 MB; OPD row \u2248 90 bytes \u2192 ~14 MB; indexes ~6 MB; total database \u2248 30\u201335 MB (under the 100 MB GitHub limit, TR-03)."),
]
write_table(ws, r, ["Point", "Explanation"], [[a, b] for a, b in notes], [32, 95], wrap_cols=(2,))
r = r + len(notes) + 3

ws.cell(row=r, column=1, value="4. Conventions").font = Font(size=11, bold=True, color=NAVY)
r += 1
conv = [
    ("Naming", "dim_* for masters, fact_* for facts, vw_* for views, idx_* for indexes, trg_* for triggers."),
    ("IDs", "TEXT surrogate keys with a fixed prefix (DEP/W/P) so they stay readable in screenshots and viva."),
    ("Dates", "Stored as ISO-8601 TEXT (YYYY-MM-DD); SQLite has no date type, and ISO text sorts correctly."),
    ("Money", "REAL in INR (2 dp), gross billed, never negative (BRL-05). STRICT tables do not accept NUMERIC."),
    ("Flags", "INTEGER 0/1 with CHECK constraints; NULL is never used for a boolean."),
    ("Time grain", "Day for facts; week/month derived in views (ISO week, calendar month)."),
    ("Deletes", "RESTRICT on all dimension deletes; history is never orphaned. No cascade deletes in this schema."),
    ("Change control", "A new column requires: change note, DDL update, dictionary update, RTM update, rebuild and full test run."),
]
write_table(ws, r, ["Convention", "Rule"], [[a, b] for a, b in conv], [20, 107], wrap_cols=(2,))
r = r + len(conv) + 3

ws.cell(row=r, column=1, value="5. Change notes issued with this design").font = Font(size=11, bold=True, color=NAVY)
r += 1
changes = [
    ("DB-CN-01", "Row plan split into design sample (150\u2013300 rows) and production (~57k episodes / ~150k visits). No change to OBJ-02."),
    ("DB-CN-02", "KPI definition registry kept as configuration (src/definitions.py), not a sixth table, to respect the 3\u20135 table brief."),
    ("DB-CN-03", "Bed capacity is derived from dim_ward.functional_beds (SCD-2), not stored as a monthly fact, so occupancy cannot contradict the bed base (fixes R-01)."),
    ("DB-CN-04", "OPD has no ward attribution; ward-level questions use only inpatient data (BR-02)."),
    ("DB-CN-06 (v1.1)", "View column iso_week renamed week_of_year: SQLite %W is Monday-based, not ISO-8601; ISO weeks are derived in pandas for BRL-08."),
    ("DB-CN-05 (v1.1)", "ERD review fixes applied: CHECK (bed_days = los_days); idx_ep_discharge for trend queries; SCD-2 aware bed allocation in vw_ward_day_beds; window-function repeat-patient flag (replaces a correlated subquery)."),
    ("DB-CN-07 (v1.1)", "Documentation correction: Tables_Summary showed 7 columns for T-01 while the dictionary and DDL both had 6. Total column count corrected from 50 to 49 (ERD verification counts every column)."),
]
write_table(ws, r, ["Note", "Change"], [[a, b] for a, b in changes], [12, 115], wrap_cols=(2,))
ws.column_dimensions["A"].width = 34

# ---------------- 2 Tables_Summary ----------------
ws = wb.create_sheet("2 Tables_Summary")
ws.sheet_view.showGridLines = False
sheet_title(ws, "Tables Summary", "Five tables: three masters and two facts. Row counts shown separately for the design-review "
                                 "sample and the planned production build.",
            9)
rows = []
for t in TABLES:
    rows.append([t[0], t[1], t[2], t[3], t[4], t[5], t[6], t[7], t[8], t[9]])
write_table(ws, 4, ["ID", "Table", "Type", "Grain (one row =)", "PK", "Foreign keys", "# columns",
                    "Rows (design sample)", "Rows (planned prod)", "Purpose"],
            rows, [7, 26, 8, 34, 14, 40, 10, 14, 14, 60], wrap_cols=(4, 6, 10), mono_cols=(2,))
ws.freeze_panes = "C5"
tot = ws.max_row + 2
ws.cell(row=tot, column=1, value="Totals").font = Font(bold=True)
ws.cell(row=tot, column=7, value=sum(t[6] for t in TABLES)).font = Font(bold=True)
ws.cell(row=tot, column=8, value=sum(t[7] for t in TABLES)).font = Font(bold=True)
ws.cell(row=tot, column=9, value=f"{sum(t[8] for t in TABLES):,}").font = Font(bold=True)

# ---------------- 3-7 data dictionaries ----------------
for idx, t in enumerate(TABLES, start=3):
    name = f"{idx} {t[0]} {t[1]}"
    ws = wb.create_sheet(name[:31])
    ws.sheet_view.showGridLines = False
    sheet_title(ws, f"{t[0]}  {t[1]}  \u00b7  data dictionary",
                f"{t[2]} table \u00b7 grain: {t[3]} \u00b7 PK: {t[4]} \u00b7 rows: {t[7]} (design sample) / {t[8]:,} (production) \u00b7 {t[9]}",
                8)
    write_table(ws, 4, ["Column", "Type", "Key", "Nullable", "Default", "Allowed values", "Description", "OQ / BR supported"],
                DD[t[0]], [24, 18, 20, 8, 12, 42, 55, 42], wrap_cols=(6, 7, 8), mono_cols=(1,))
    ws.freeze_panes = "B5"

# ---------------- 8 Relationships ----------------
ws = wb.create_sheet("8 Relationships")
ws.sheet_view.showGridLines = False
sheet_title(ws, "Relationships", "Cardinality and referential actions. All dimension deletes are RESTRICT so history is never orphaned.",
            6)
write_table(ws, 4, ["Parent", "Child", "Key", "Cardinality", "On delete", "Rule / meaning"],
            [[a, b, c, d, e, f] for a, b, c, d, e, f in RELATIONSHIPS],
            [24, 32, 18, 16, 12, 70], wrap_cols=(6,), mono_cols=(3,))
ws.freeze_panes = "A5"

# ---------------- 9 KPIs ----------------
ws = wb.create_sheet("9 KPIs_Measures")
ws.sheet_view.showGridLines = False
sheet_title(ws, "KPI measures and formulas", "Every measure maps to an OQ and to physical columns. Formulas are implemented once in "
                                            "src/kpi.py and tested at 1e-6 tolerance (NFR-05).", 6)
write_table(ws, 4, ["Measure", "Formula (as implemented)", "OQ", "Grain", "Tables", "Unit"],
            [[a, b, c, d, e, f] for a, b, c, d, e, f in KPIS],
            [28, 62, 16, 26, 18, 8], wrap_cols=(2, 4), mono_cols=(2,))
ws.freeze_panes = "A5"

# ---------------- 10 DDL ----------------
ws = wb.create_sheet("10 SQLite_DDL")
ws.sheet_view.showGridLines = False
sheet_title(ws, "SQLite DDL \u2014 tables, constraints, indexes and trigger",
            "Execute in order. PRAGMA foreign_keys is ON, so the app must open read-only and the build script must load in FK order: "
            "T-01 \u2192 T-02/T-03 \u2192 T-04 \u2192 T-05.", 2)
for i, line in enumerate(DDL):
    cell = ws.cell(row=4 + i, column=1, value=line)
    cell.font = mono_font
    cell.alignment = Alignment(vertical="top")
    if line.strip().startswith(("CREATE TABLE", "CREATE INDEX", "--")):
        cell.font = Font(name="Consolas", size=8.5, bold=True, color=NAVY)
ws.column_dimensions["A"].width = 120

# ---------------- 11 Views ----------------
ws = wb.create_sheet("11 Views")
ws.sheet_view.showGridLines = False
sheet_title(ws, "Analysis and aggregation views", "Three views: the 36-column episode view, the department-month KPI view and the "
                                                "ward-day bed view.", 2)
row = 4
for block, caption in ((VIEW_ANALYSIS, "1. vw_analysis_episode \u2014 36 columns"),
                       (VIEW_WARD_DAY, "2. vw_ward_day_beds \u2014 ward-day occupancy"),
                       (VIEW_DEPT_MONTH, "3. vw_dept_month_kpi \u2014 17 columns (depends on view 2)")):
    c = ws.cell(row=row, column=1, value=caption)
    c.font = Font(size=10, bold=True, color=NAVY)
    row += 1
    for line in block:
        cell = ws.cell(row=row, column=1, value=line)
        cell.font = mono_font
        row += 1
    row += 1
ws.column_dimensions["A"].width = 130

# ---------------- 12 Row_Plan ----------------
ws = wb.create_sheet("12 Row_Plan")
ws.sheet_view.showGridLines = False
sheet_title(ws, "Row plan \u2014 design sample vs production", "Column counts and volumes drive index choice, cache strategy and the "
                                                             "deployment size check (TR-03).", 5)
plan = [
    ("T-01", "dim_department", 8, 8, "Static"),
    ("T-02", "dim_ward", 20, 20, "Static, SCD-2 tracked"),
    ("T-03", "dim_payer", 5, 5, "Static"),
    ("T-04", "fact_inpatient_episode", 260, 57000, "3 years \u00d7 8 departments"),
    ("T-05", "fact_opd_visit", 300, 150000, "3 years \u00d7 8 departments"),
]
write_table(ws, 4, ["ID", "Table", "Rows (design sample)", "Rows (planned production)", "Basis"],
            [[a, b, c, d, e] for a, b, c, d, e in plan], [8, 28, 20, 24, 60])
r = 4 + len(plan) + 2
ws.cell(row=r, column=1, value="Volume arithmetic").font = Font(size=10, bold=True, color=NAVY)
r += 1
calc = [
    "Bed-days available  = 300 functional beds \u00d7 1,095 days = 328,500",
    "Bed-days occupied   = 328,500 \u00d7 0.72 average occupancy = 236,520",
    "Discharges          = 236,520 \u00f7 4.1 average ALOS \u2248 57,688  \u2192 plan 57,000",
    "OPD visits          = 140 visits/day \u00d7 1,095 days \u2248 153,300 \u2192 plan 150,000",
    "Database size       \u2248 57,000 \u00d7 180 B + 150,000 \u00d7 90 B + indexes \u2248 30\u201335 MB",
]
for line in calc:
    c = ws.cell(row=r, column=1, value=line)
    c.font = mono_font
    r += 1

# ---------------- 13 Change_Log ----------------
ws = wb.create_sheet("13 Change_Log")
ws.sheet_view.showGridLines = False
sheet_title(ws, "Change log", "The schema is baselined at v1.0. Additions require a note here and a matching RTM update.", 5)
log = [
    ("1.0", "19-Nov-2026", "Step 8", "Initial design: 5 tables, 3 views, 36-column analysis view, DDL with PK/FK/CHECK/UNIQUE/trigger.", "Baseline"),
    ("1.1", "19-Nov-2026", "Step 9", "Design review: bed_days = los_days CHECK, discharge_date index, SCD-2 bed allocation, window-function repeat flag.", "Current"),
    ("", "", "Later", "Any new column or table: append a note, update the dictionary sheet, update the RTM, rebuild and retest.", "Rule"),
]
write_table(ws, 4, ["Version", "Date", "Step", "Change", "Status"], [[a, b, c, d, e] for a, b, c, d, e in log],
            [10, 14, 10, 95, 12], wrap_cols=(4,))

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print("saved:", OUT)

# ------------------------------------------------------------------ update RTM
DB_MAP = {
    "BR-01": "T-02.functional_beds; T-04.admission_date; T-04.discharge_date",
    "BR-02": "T-02.ward_id; T-02.functional_beds; T-04.ward_id; T-04.discharge_date",
    "BR-03": "T-04.department_id; T-04.payer_id; T-04.admission_date; T-05.visit_date",
    "BR-04": "\u2014 (no column: interaction design, FR-04/FR-22)",
    "BR-05": "\u2014 (config: src/definitions.py; BRL-14)",
    "BR-06": "T-04.los_days; T-04.department_id; T-04.diagnosis_group",
    "BR-07": "T-04.readmit_30d_flag; T-04.index_episode_id; T-04.patient_key; T-04.department_id",
    "BR-08": "T-03.payer_category; T-04.payer_id; T-04.billed_amount",
    "BR-09": "T-05.visit_date; T-05.department_id",
    "BR-10": "vw_dept_month_kpi (derived from T-02, T-04)",
    "BR-11": "T-04.age_band; T-04.gender; T-05.age_band; T-05.gender",
    "BR-12": "T-04.billed_amount; T-02.functional_beds",
    "BR-13": "T-04.death_flag; T-04.discharge_type; T-04.discharge_date",
    "BR-14": "T-04.payer_id (null); T-04.patient_key (dup); T-04.los_days (negative); T-05.visit_date (coverage)",
    "BR-15": "\u2014 (Won't; no table by design)",
}

wb2 = load_workbook(RTM)
rt = wb2["RTM"]
rt["A2"] = ("v1.1  \u00b7  19-Nov-2026  \u00b7  DB column filled from Step 8 (T-01\u2026T-05)  \u00b7  "
            "Feature still reserved for Step 6 (PRD)  \u00b7  Status remains Identified")
rt["A2"].font = Font(name="Calibri", size=9, italic=True, color="404040")
rt["A2"].fill = PatternFill("solid", fgColor="EAF0F8")
rt["A2"].alignment = Alignment(vertical="center", wrap_text=True, indent=1)

for r in range(4, rt.max_row + 1):
    br = rt.cell(row=r, column=3).value
    if br in DB_MAP:
        cell = rt.cell(row=r, column=9, value=DB_MAP[br])
        cell.font = Font(name="Consolas", size=8)
        cell.alignment = Alignment(vertical="top", wrap_text=True)
        rt.row_dimensions[r].height = 62

cl = wb2["Change Log"]
# idempotent: remove any previously written v1.1 row before re-adding
for r in range(cl.max_row, 4, -1):
    if str(cl.cell(row=r, column=1).value).strip() == "1.1":
        cl.delete_rows(r)

# mark the Step 8 pending row as done exactly once
for r in range(5, cl.max_row + 1):
    if cl.cell(row=r, column=3).value == "Step 8":
        txt = str(cl.cell(row=r, column=4).value or "").replace("[Done in v1.1] ", "")
        cl.cell(row=r, column=4, value="[Done in v1.1] " + txt)
        cl.cell(row=r, column=4).alignment = Alignment(wrap_text=True, vertical="top")

row = cl.max_row + 1
vals = ("1.1", "19-Nov-2026", "Step 8",
        "DB table.column filled for all 16 rows (T-01\u2026T-05). BR-04/BR-05 have no physical column by design "
        "(interaction / config); BR-15 has none (Won't). Feature column still pending Step 6.", "Current")
for j, v in enumerate(vals, start=1):
    cell = cl.cell(row=row, column=j, value=v)
    cell.font = Font(size=9)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    cell.border = border
cl.row_dimensions[row].height = 40

wb2.save(RTM)
print("updated:", RTM)
