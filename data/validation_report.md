# p107 · Step 10 validation report

Generated from `data/app.db` (seed 42, 2023-01-01 to 2025-12-31).

## Row counts

| Table | Raw | Cleaned | Rejected/removed |
|---|---:|---:|---:|
| dim_department | 8 | 8 | 0 |
| dim_ward | 20 | 20 | 0 |
| dim_payer | 5 | 5 | 0 |
| fact_inpatient_episode | 56,094 | 53,915 | 2,179 |
| fact_opd_visit | 146,328 | 143,139 | 3,189 |

## Checks

| Check | Expectation | Actual | Result |
|---|---|---|---|
| Row counts reconcile | clean + rejected + removed = raw | 197,087 + 5,368 = 202,455 | PASS |
| Primary keys unique | 0 duplicates | 0 duplicate (patient, admission) | PASS |
| Foreign keys | PRAGMA foreign_key_check returns no rows | 0 violation(s) | PASS |
| Load errors | 0 rows rejected by the DDL | 0 row(s) | PASS |
| Longest stay in range | 1 ≤ los ≤ 120 | 1–90 days | PASS |
| Billing non-negative | billed_amount ≥ 0 | ₹5,788 – ₹6,056,692 | PASS |
| Dates inside the window | 2023-01-01 … 2025-12-31 | 2023-01-01 … 2025-12-31 | PASS |
| Occupancy cannot exceed 100% | max ward-day BOR ≤ 100 | 100.0% (ward-day), 92.72% (dept-month) | PASS |
| Bed-days reconcile | Σ bed_days = Σ ward-day occupancy | 237,549 = 237,549 | PASS |
| Analysis view width | 36 columns | 36 columns | PASS |
| No dangling readmission links | every index_episode_id resolves | 3176 links checked | PASS |
| No real personal data | no name/phone/address columns | schema stores pseudonymous keys only | PASS |

## Patterns

**P-01 Winter respiratory surge** — GenMed + Paediatrics + ENT admit 76.7 patients a day in Nov–Feb against 60.1 in Mar–Oct, a 28% winter lift.  
Supports: OQ-02, OQ-06, BR-09
**P-02 Cardiology capacity pressure** — Cardiology BOR 85.7% vs the 85% target; ALOS 5.43 days, the longest of the eight departments. Orthopaedics runs at 67.8%.  
Supports: OQ-01, OQ-03, BR-01, BR-06
**P-03 Readmission hotspots** — General Medicine 8.47%, Oncology 6.34%, Cardiology 5.78% against Orthopaedics 2.61% and Maternity 2.62%. Short stays (≤3 days) readmit at 7.08% vs 3.57% for stays of 8+ days.  
Supports: OQ-04, OQ-07, BR-07, BR-10
**P-04 Payer-mix shift and revenue per bed-day** — Corporate share 6.37% in 2023 → 11.23% in 2025 while self-pay falls 41.35% → 33.49%. Revenue per bed-day: Oncology ₹30,840, Orthopaedics ₹21,809, General Medicine ₹10,614.  
Supports: OQ-05, OQ-09, BR-08, BR-12
**P-05 Oncology mortality drift and the Monday OPD peak** — Oncology mortality rises from 3.52% (2023) to 5.62% (2025); the last half-year reads 6.0%. OPD peak-to-average ratio 1.53 (Monday busiest, Sunday near-closed).  
Supports: OQ-06, OQ-10, BR-09, BR-13

## Injected issues

| Table | Issue | Rows | % | Handling |
|---|---|---:|---:|---|
| fact_inpatient_episode | missing payer_id | 970 | 1.8 | rejected |
| fact_inpatient_episode | missing billed_amount | 377 | 0.7 | rejected |
| fact_inpatient_episode | duplicate episode row | 646 | 1.2 | removed (duplicate) |
| fact_inpatient_episode | negative length of stay | 64 | 0.12 | rejected |
| fact_inpatient_episode | discharge before admission | 53 | 0.1 | rejected |
| fact_inpatient_episode | unknown department (orphan FK) | 43 | 0.08 | rejected |
| fact_inpatient_episode | invalid gender code | 26 | 0.05 | rejected |
| fact_opd_visit | missing payer_id | 2147 | 1.5 | rejected |
| fact_opd_visit | duplicate visit row | 859 | 0.6 | removed (duplicate) |
| fact_opd_visit | zero consultation fee | 143 | 0.1 | rejected |
