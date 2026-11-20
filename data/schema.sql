PRAGMA foreign_keys = ON;
CREATE TABLE dim_department (
  department_id   TEXT    PRIMARY KEY CHECK (length(department_id) = 5),
  department_name TEXT    NOT NULL UNIQUE,
  specialty       TEXT    NOT NULL,
  floor_block     TEXT    NOT NULL,
  is_active       INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0,1)),
  effective_from  TEXT    NOT NULL CHECK (date(effective_from) IS NOT NULL)
) STRICT;
CREATE TABLE dim_ward (
  ward_id         TEXT    PRIMARY KEY,
  ward_name       TEXT    NOT NULL UNIQUE,
  ward_type       TEXT    NOT NULL CHECK (ward_type IN
                    ('General','Semi-private','Private','ICU','HDU','Maternity','Paediatric','Day-care')),
  bed_class       TEXT    NOT NULL CHECK (bed_class IN ('Economy','Twin','Single','Intensive')),
  department_id   TEXT    NOT NULL REFERENCES dim_department(department_id)
                    ON UPDATE CASCADE ON DELETE RESTRICT,
  functional_beds INTEGER NOT NULL CHECK (functional_beds BETWEEN 1 AND 60),
  is_active       INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0,1)),
  effective_from  TEXT    NOT NULL,
  effective_to    TEXT    NULL,
  CHECK (effective_to IS NULL OR effective_to >= effective_from)
) STRICT;
CREATE TABLE dim_payer (
  payer_id        TEXT    PRIMARY KEY,
  payer_name      TEXT    NOT NULL UNIQUE,
  payer_category  TEXT    NOT NULL CHECK (payer_category IN
                    ('Cash','Private insurance','TPA','Government scheme','Corporate')),
  settlement_days INTEGER NOT NULL DEFAULT 0 CHECK (settlement_days BETWEEN 0 AND 180),
  is_active       INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0,1))
) STRICT;
CREATE TABLE fact_inpatient_episode (
  episode_id          INTEGER PRIMARY KEY,
  patient_key         TEXT    NOT NULL CHECK (length(patient_key) = 12),
  department_id       TEXT    NOT NULL REFERENCES dim_department(department_id),
  ward_id             TEXT    NOT NULL REFERENCES dim_ward(ward_id),
  payer_id            TEXT    NOT NULL REFERENCES dim_payer(payer_id),
  admission_date      TEXT    NOT NULL CHECK (date(admission_date) IS NOT NULL),
  discharge_date      TEXT    NOT NULL CHECK (date(discharge_date) IS NOT NULL),
  los_days            INTEGER NOT NULL CHECK (los_days BETWEEN 1 AND 120),
  discharge_type      TEXT    NOT NULL CHECK (discharge_type IN ('Routine','Transfer','LAMA','Expired')),
  death_flag          INTEGER NOT NULL CHECK (death_flag IN (0,1)),
  readmit_30d_flag    INTEGER NOT NULL DEFAULT 0 CHECK (readmit_30d_flag IN (0,1)),
  index_episode_id    INTEGER NULL REFERENCES fact_inpatient_episode(episode_id) ON DELETE SET NULL,
  diagnosis_group     TEXT    NOT NULL,
  is_planned_procedure INTEGER NOT NULL DEFAULT 0 CHECK (is_planned_procedure IN (0,1)),
  age_band            TEXT    NOT NULL CHECK (age_band IN ('0-5','6-17','18-35','36-50','51-65','66+')),
  gender              TEXT    NOT NULL CHECK (gender IN ('M','F','O')),
  billed_amount       REAL    NOT NULL CHECK (billed_amount >= 0),
  bed_days            INTEGER NOT NULL CHECK (bed_days BETWEEN 1 AND 121),
  created_at          TEXT    NOT NULL DEFAULT (datetime('now')),
  CHECK (discharge_date >= admission_date),
  CHECK (death_flag = 0 OR discharge_type = 'Expired'),
  CHECK (discharge_type <> 'Expired' OR death_flag = 1),
  CHECK (readmit_30d_flag = 0 OR index_episode_id IS NOT NULL),
  CHECK (index_episode_id IS NULL OR index_episode_id <> episode_id),
  CHECK (los_days = CAST(julianday(discharge_date) - julianday(admission_date) AS INTEGER) + 1),
  CHECK (bed_days = los_days),
  UNIQUE (patient_key, admission_date)
) STRICT;
CREATE TABLE fact_opd_visit (
  visit_id         INTEGER PRIMARY KEY,
  patient_key      TEXT    NOT NULL CHECK (length(patient_key) = 12),
  visit_date       TEXT    NOT NULL CHECK (date(visit_date) IS NOT NULL),
  department_id    TEXT    NOT NULL REFERENCES dim_department(department_id),
  payer_id         TEXT    NOT NULL REFERENCES dim_payer(payer_id),
  age_band         TEXT    NOT NULL CHECK (age_band IN ('0-5','6-17','18-35','36-50','51-65','66+')),
  gender           TEXT    NOT NULL CHECK (gender IN ('M','F','O')),
  visit_type       TEXT    NOT NULL CHECK (visit_type IN ('New','Follow-up')),
  consultation_fee REAL    NOT NULL CHECK (consultation_fee >= 0),
  created_at       TEXT    NOT NULL DEFAULT (datetime('now')),
  UNIQUE (patient_key, visit_date, department_id)
) STRICT;
-- Indexes tuned to the filter and join paths used by the app
CREATE INDEX idx_ep_dept_adm    ON fact_inpatient_episode(department_id, admission_date);
CREATE INDEX idx_ep_dept_dis    ON fact_inpatient_episode(department_id, discharge_date);
CREATE INDEX idx_ep_discharge   ON fact_inpatient_episode(discharge_date);
CREATE INDEX idx_ep_payer       ON fact_inpatient_episode(payer_id);
CREATE INDEX idx_ep_patient     ON fact_inpatient_episode(patient_key);
CREATE INDEX idx_ep_diag        ON fact_inpatient_episode(diagnosis_group);
CREATE INDEX idx_ep_readmit     ON fact_inpatient_episode(readmit_30d_flag) WHERE readmit_30d_flag = 1;
CREATE INDEX idx_opd_date       ON fact_opd_visit(visit_date);
CREATE INDEX idx_opd_dept_date  ON fact_opd_visit(department_id, visit_date);
CREATE INDEX idx_opd_payer      ON fact_opd_visit(payer_id);
-- Ward must belong to the same department as the episode
CREATE TRIGGER trg_episode_ward_dept
BEFORE INSERT ON fact_inpatient_episode
FOR EACH ROW
WHEN (SELECT department_id FROM dim_ward WHERE ward_id = NEW.ward_id) <> NEW.department_id
BEGIN
  SELECT RAISE(ABORT, 'ward_id does not belong to department_id');
END;

-- vw_analysis_episode : 36 columns, episode grain, one row per episode.
-- week_of_year is strftime('%W') (Monday-based); ISO-8601 week numbers are computed in pandas,
-- because SQLite cannot produce ISO week 53 correctly (BRL-08).
CREATE VIEW vw_analysis_episode AS
SELECT
  e.episode_id, e.patient_key, e.admission_date, e.discharge_date,
  e.los_days, e.discharge_type, e.death_flag, e.readmit_30d_flag,
  e.index_episode_id, e.diagnosis_group, e.is_planned_procedure,
  e.age_band, e.gender, e.billed_amount, e.bed_days,
  e.payer_id, p.payer_name, p.payer_category,
  e.department_id, d.department_name, d.specialty,
  e.ward_id, w.ward_name, w.ward_type, w.bed_class, w.functional_beds,
  CAST(strftime('%Y', e.admission_date) AS INTEGER) AS admission_year,
  strftime('%Y-%m', e.admission_date)                 AS admission_month,
  CAST(strftime('%W', e.admission_date) AS INTEGER)   AS week_of_year,
  CASE strftime('%w', e.admission_date)
    WHEN '0' THEN 'Sunday'   WHEN '1' THEN 'Monday'
    WHEN '2' THEN 'Tuesday'  WHEN '3' THEN 'Wednesday'
    WHEN '4' THEN 'Thursday' WHEN '5' THEN 'Friday'
    ELSE 'Saturday' END                                AS weekday_name,
  'Q' || ((CAST(strftime('%m', e.admission_date) AS INTEGER) - 1) / 3 + 1) AS admission_quarter,
  CASE WHEN strftime('%w', e.admission_date) IN ('0','6') THEN 1 ELSE 0 END AS is_weekend_admission,
  CASE
    WHEN CAST(strftime('%m', e.admission_date) AS INTEGER) IN (12,1,2) THEN 'Winter'
    WHEN CAST(strftime('%m', e.admission_date) AS INTEGER) IN (3,4,5)  THEN 'Summer'
    WHEN CAST(strftime('%m', e.admission_date) AS INTEGER) IN (6,7,8,9) THEN 'Monsoon'
    ELSE 'Post-monsoon' END                            AS season,
  CASE WHEN e.los_days <= 2 THEN '0-2' WHEN e.los_days <= 4 THEN '3-4'
       WHEN e.los_days <= 7 THEN '5-7' ELSE '8+' END   AS los_band,
  CASE WHEN e.billed_amount < 25000  THEN 'Low'
       WHEN e.billed_amount < 75000  THEN 'Medium'
       WHEN e.billed_amount < 200000 THEN 'High'
       ELSE 'Very high' END                            AS revenue_band,
  CASE WHEN COUNT(*) OVER (
         PARTITION BY e.patient_key
         ORDER BY e.admission_date, e.episode_id
         ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING) > 0
       THEN 1 ELSE 0 END                               AS is_repeat_patient
FROM fact_inpatient_episode e
JOIN dim_department d ON d.department_id = e.department_id
JOIN dim_ward       w ON w.ward_id       = e.ward_id
JOIN dim_payer      p ON p.payer_id      = e.payer_id;
-- vw_ward_day_beds : ward-day occupancy, used by BR-02 (next-7-day ward view)
CREATE VIEW vw_ward_day_beds AS
WITH RECURSIVE cal(day) AS (
  SELECT MIN(admission_date) FROM fact_inpatient_episode
  UNION ALL
  SELECT date(day, '+1 day') FROM cal
  WHERE day < (SELECT MAX(discharge_date) FROM fact_inpatient_episode)
),
stays AS (
  SELECT e.ward_id, c.day,
         COUNT(DISTINCT e.episode_id) AS beds_occupied,
         MAX(w.functional_beds)       AS functional_beds,
         MAX(w.ward_name)             AS ward_name,
         MAX(w.ward_type)             AS ward_type,
         MAX(w.department_id)         AS department_id
  FROM fact_inpatient_episode e
  JOIN cal c ON c.day >= e.admission_date
             AND c.day <  date(e.discharge_date, '+1 day')
  JOIN dim_ward w ON w.ward_id = e.ward_id
                 AND c.day BETWEEN w.effective_from
                               AND COALESCE(w.effective_to, '9999-12-31')
  GROUP BY 1,2
)
SELECT day, ward_id, ward_name, ward_type, department_id,
       functional_beds, beds_occupied,
       ROUND(100.0 * beds_occupied / NULLIF(functional_beds,0), 2) AS bor_pct
FROM stays;
-- vw_dept_month_kpi : one row per department-month, all six KPI families pre-aggregated.
-- Requires vw_ward_day_beds to exist first (bed-days are allocated to the month they occur in,
-- not to the discharge month, so occupancy cannot exceed 100% at month boundaries).
CREATE VIEW vw_dept_month_kpi AS
WITH beds AS (
  SELECT department_id,
         strftime('%Y-%m', day) AS month,
         SUM(functional_beds)   AS bed_days_available,
         SUM(beds_occupied)     AS bed_days_occupied
  FROM vw_ward_day_beds
  GROUP BY 1,2
),
ep AS (
  SELECT department_id,
         strftime('%Y-%m', discharge_date) AS month,
         COUNT(*)                           AS discharges,
         SUM(los_days)                      AS inpatient_days,
         SUM(billed_amount)                 AS revenue,
         SUM(death_flag)                    AS deaths,
         SUM(readmit_30d_flag)              AS readmissions
  FROM fact_inpatient_episode
  GROUP BY 1,2
),
opd AS (
  SELECT department_id, strftime('%Y-%m', visit_date) AS month, COUNT(*) AS opd_visits
  FROM fact_opd_visit GROUP BY 1,2
)
SELECT b.department_id, b.month,
       COALESCE(ep.discharges, 0)      AS discharges,
       COALESCE(ep.inpatient_days, 0)  AS inpatient_days,
       b.bed_days_occupied,
       b.bed_days_available,
       ROUND(100.0 * b.bed_days_occupied / NULLIF(b.bed_days_available,0), 2) AS bor_pct,
       ROUND(1.0 * ep.inpatient_days / NULLIF(ep.discharges,0), 2)            AS alos,
       COALESCE(ep.readmissions, 0)    AS readmissions,
       ROUND(100.0 * ep.readmissions / NULLIF(ep.discharges,0), 2)            AS readmission_rate,
       COALESCE(ep.revenue, 0)         AS revenue,
       ROUND(1.0 * ep.revenue / NULLIF(ep.discharges,0), 2)                   AS revenue_per_discharge,
       ROUND(1.0 * ep.revenue / NULLIF(b.bed_days_occupied,0), 2)             AS revenue_per_bed_day,
       COALESCE(ep.deaths, 0)          AS deaths,
       ROUND(100.0 * ep.deaths / NULLIF(ep.discharges,0), 2)                  AS mortality_rate,
       COALESCE(opd.opd_visits, 0)     AS opd_visits,
       CASE WHEN COALESCE(ep.discharges,0) < 30 THEN 1 ELSE 0 END             AS low_confidence
FROM beds b
LEFT JOIN ep  ON ep.department_id  = b.department_id  AND ep.month  = b.month
LEFT JOIN opd ON opd.department_id = b.department_id  AND opd.month = b.month;
