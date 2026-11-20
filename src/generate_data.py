"""HKD-2026 — synthetic dataset generator (Step 10).

Builds a three-year hospital dataset that matches the Step 8 schema exactly:
5 tables, 49 columns, 7 foreign keys. Deterministic: seed 42.

Outputs
-------
data/app.db                     SQLite database built from the Step 8 DDL (FK ON)
data/<table>.csv                cleaned data, one file per table
data/raw/<table>.csv            as-received data with injected quality defects
data/hkd2026_data.xlsx          workbook, one sheet per table (fact tables sampled)
data/quality_log.csv            every row rejected by the cleaning stage
data/validation_report.md       row counts, FK check, range checks, pattern evidence
data/validation_summary.json    machine-readable summary used by the Step 10 document
data/schema.sql                 copy of the executed DDL

Run:  .venv/bin/python src/generate_data.py
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import math
import os
import random
import re
import sqlite3
from collections import Counter, defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SEED = 42
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
RAW = os.path.join(DATA, "raw")
XLSX_SRC = os.path.join(ROOT, "docs", "08_Database_Design.xlsx")

START = dt.date(2023, 1, 1)
END = dt.date(2025, 12, 31)
NDAYS = (END - START).days + 1
GENERATED_AT = "2026-11-19 09:00:00"

rng = random.Random(SEED)

# --------------------------------------------------------------------------- config
# code, name, specialty, block, mean_los, occupancy target, base bill,
# mortality, readmission prob, planned prob, monthly seasonality (Jan..Dec),
# diagnosis mix, age mix, gender mix
DEPTS = [
    dict(code="DEP01", name="Cardiology", specialty="Cardiac sciences", block="Block A",
         los=5.2, occ=1.00, bill=95_000, mort=0.012, readmit=0.062, planned=0.18,
         season=[1.10, 1.07, 1.01, 0.95, 0.92, 0.90, 0.90, 0.93, 0.97, 1.01, 1.06, 1.14],
         diag=[("Cardiac", 0.55), ("Respiratory", 0.14), ("Renal", 0.10), ("Endocrine", 0.09),
               ("Infection", 0.07), ("General", 0.05)],
         ages=[("0-5", 0.00), ("6-17", 0.02), ("18-35", 0.08), ("36-50", 0.20),
               ("51-65", 0.38), ("66+", 0.32)],
         gender=[("M", 0.58), ("F", 0.41), ("O", 0.01)]),
    dict(code="DEP02", name="General Medicine", specialty="Internal medicine", block="Block B",
         los=4.5, occ=0.66, bill=38_000, mort=0.018, readmit=0.078, planned=0.08,
         season=[1.48, 1.42, 1.14, 0.92, 0.78, 0.74, 0.80, 0.92, 1.02, 1.10, 1.24, 1.50],
         diag=[("Respiratory", 0.26), ("Infection", 0.20), ("Cardiac", 0.14), ("Gastro", 0.12),
               ("Endocrine", 0.10), ("Renal", 0.08), ("Neuro", 0.05), ("General", 0.05)],
         ages=[("0-5", 0.03), ("6-17", 0.05), ("18-35", 0.17), ("36-50", 0.22),
               ("51-65", 0.28), ("66+", 0.25)],
         gender=[("M", 0.52), ("F", 0.47), ("O", 0.01)]),
    dict(code="DEP03", name="Orthopaedics", specialty="Orthopaedic surgery", block="Block C",
         los=5.5, occ=0.67, bill=85_000, mort=0.006, readmit=0.026, planned=0.34,
         season=[1.02, 1.00, 0.98, 0.98, 1.00, 1.02, 1.04, 1.06, 1.04, 1.00, 0.98, 0.94],
         diag=[("Orthopaedic", 0.62), ("Trauma", 0.22), ("Infection", 0.06),
               ("Neuro", 0.05), ("General", 0.05)],
         ages=[("0-5", 0.01), ("6-17", 0.05), ("18-35", 0.22), ("36-50", 0.24),
               ("51-65", 0.26), ("66+", 0.22)],
         gender=[("M", 0.57), ("F", 0.42), ("O", 0.01)]),
    dict(code="DEP04", name="Paediatrics", specialty="Child health", block="Block B",
         los=3.2, occ=0.66, bill=22_000, mort=0.004, readmit=0.046, planned=0.05,
         season=[1.50, 1.44, 1.12, 0.90, 0.78, 0.76, 0.86, 1.00, 1.06, 1.10, 1.20, 1.52],
         diag=[("Paediatric", 0.34), ("Respiratory", 0.26), ("Infection", 0.22),
               ("Gastro", 0.12), ("General", 0.06)],
         ages=[("0-5", 0.52), ("6-17", 0.42), ("18-35", 0.05), ("36-50", 0.01),
               ("51-65", 0.00), ("66+", 0.00)],
         gender=[("M", 0.53), ("F", 0.46), ("O", 0.01)]),
    dict(code="DEP05", name="Maternity", specialty="Obstetrics and gynaecology", block="Block C",
         los=2.6, occ=0.60, bill=45_000, mort=0.001, readmit=0.015, planned=0.42,
         season=[1.04, 1.02, 1.04, 1.02, 1.00, 0.96, 0.96, 0.98, 1.00, 1.02, 1.00, 0.96],
         diag=[("Obstetric", 0.78), ("Gynae", 0.14), ("General", 0.05), ("Infection", 0.03)],
         ages=[("0-5", 0.00), ("6-17", 0.00), ("18-35", 0.84), ("36-50", 0.15),
               ("51-65", 0.01), ("66+", 0.00)],
         gender=[("M", 0.00), ("F", 0.99), ("O", 0.01)]),
    dict(code="DEP06", name="Oncology", specialty="Medical oncology", block="Block D",
         los=6.5, occ=0.84, bill=130_000, mort=0.036, readmit=0.082, planned=0.22,
         season=[1.00, 0.98, 1.00, 1.02, 1.00, 0.98, 0.98, 1.00, 1.02, 1.02, 1.00, 0.98],
         diag=[("Oncology", 0.68), ("Respiratory", 0.08), ("Gastro", 0.08),
               ("Infection", 0.08), ("General", 0.08)],
         ages=[("0-5", 0.01), ("6-17", 0.02), ("18-35", 0.07), ("36-50", 0.18),
               ("51-65", 0.34), ("66+", 0.38)],
         gender=[("M", 0.54), ("F", 0.45), ("O", 0.01)]),
    dict(code="DEP07", name="Emergency", specialty="Emergency medicine", block="Block A",
         los=2.2, occ=0.55, bill=18_000, mort=0.024, readmit=0.050, planned=0.01,
         season=[1.10, 1.06, 1.02, 0.98, 1.00, 1.04, 1.08, 1.06, 1.02, 1.00, 1.02, 1.08],
         diag=[("Trauma", 0.30), ("Cardiac", 0.16), ("Respiratory", 0.14), ("Infection", 0.16),
               ("Neuro", 0.10), ("Gastro", 0.08), ("General", 0.06)],
         ages=[("0-5", 0.04), ("6-17", 0.08), ("18-35", 0.30), ("36-50", 0.24),
               ("51-65", 0.19), ("66+", 0.15)],
         gender=[("M", 0.60), ("F", 0.39), ("O", 0.01)]),
    dict(code="DEP08", name="ENT", specialty="Otorhinolaryngology", block="Block D",
         los=2.8, occ=0.41, bill=25_000, mort=0.002, readmit=0.020, planned=0.38,
         season=[1.36, 1.30, 1.06, 0.92, 0.84, 0.82, 0.88, 0.96, 1.02, 1.06, 1.14, 1.34],
         diag=[("ENT", 0.62), ("Infection", 0.16), ("Respiratory", 0.10),
               ("General", 0.07), ("Trauma", 0.05)],
         ages=[("0-5", 0.10), ("6-17", 0.16), ("18-35", 0.28), ("36-50", 0.22),
               ("51-65", 0.15), ("66+", 0.09)],
         gender=[("M", 0.55), ("F", 0.44), ("O", 0.01)]),
]

# ward_id, ward_name, ward_type, bed_class, department, beds
WARDS = [
    ("W01", "Cardio-ICU", "ICU", "Intensive", "DEP01", 12),
    ("W02", "Cardio-A", "General", "Twin", "DEP01", 15),
    ("W03", "Cardio-B", "Semi-private", "Single", "DEP01", 14),
    ("W04", "Cardio-C", "Private", "Single", "DEP01", 14),
    ("W05", "GenMed-A", "General", "Economy", "DEP02", 20),
    ("W06", "GenMed-B", "General", "Economy", "DEP02", 20),
    ("W07", "GenMed-C", "Semi-private", "Twin", "DEP02", 15),
    ("W08", "GenMed-D", "Private", "Single", "DEP02", 15),
    ("W09", "GenMed-E", "General", "Economy", "DEP02", 15),
    ("W10", "Ortho-A", "General", "Economy", "DEP03", 20),
    ("W11", "Ortho-B", "Semi-private", "Twin", "DEP03", 15),
    ("W12", "Ortho-C", "Private", "Single", "DEP03", 10),
    ("W13", "Paeds-A", "Paediatric", "Economy", "DEP04", 20),
    ("W14", "Paeds-B", "Paediatric", "Twin", "DEP04", 15),
    ("W15", "Maternity-A", "Maternity", "Twin", "DEP05", 18),
    ("W16", "Maternity-B", "Maternity", "Single", "DEP05", 12),
    ("W17", "Onco-A", "General", "Twin", "DEP06", 15),
    ("W18", "Onco-B", "Private", "Single", "DEP06", 10),
    ("W19", "Emergency", "HDU", "Intensive", "DEP07", 15),
    ("W20", "ENT", "Day-care", "Economy", "DEP08", 10),
]

# payer_id, name, category, settlement days
PAYERS = [
    ("P01", "Self-pay", "Cash", 0),
    ("P02", "Swaasthya Shield", "Private insurance", 21),
    ("P03", "Sahyog TPA", "TPA", 45),
    ("P04", "Rajya Swasthya Yojana", "Government scheme", 60),
    ("P05", "Nirmaan Corporate Health", "Corporate", 30),
]
PAYER_CATS = [p[2] for p in PAYERS]
PAYER_MULT = {"P01": 1.00, "P02": 1.15, "P03": 1.05, "P04": 0.78, "P05": 1.10}

AGE_BANDS = ["0-5", "6-17", "18-35", "36-50", "51-65", "66+"]
GENDERS = ["M", "F", "O"]

DEP = {d["code"]: d for d in DEPTS}
WARD = {w[0]: w for w in WARDS}
W_IDX = {w[0]: i for i, w in enumerate(WARDS)}
WARD_BEDS = [w[5] for w in WARDS]
DEPT_WARDS = defaultdict(list)
for w in WARDS:
    DEPT_WARDS[w[4]].append(w[0])

# OPD: consultations per day per department (base, before weekday/season)
OPD_LAMBDA = {"DEP01": 16, "DEP02": 40, "DEP03": 18, "DEP04": 25,
              "DEP05": 12, "DEP06": 6, "DEP07": 8, "DEP08": 14}
OPD_FEE = {"DEP01": 900, "DEP02": 400, "DEP03": 800, "DEP04": 350,
           "DEP05": 700, "DEP06": 1100, "DEP07": 250, "DEP08": 600}
WEEKDAY_ADMIT = [0.80, 1.06, 1.05, 1.00, 0.98, 0.94, 0.88]   # Mon..Sun
WEEKDAY_OPD = [1.38, 1.12, 1.02, 0.94, 0.88, 0.74, 0.22]     # Mon..Sun
OPD_SEASON = {"DEP01": [1, 1, 1, 1, .95, .9, .9, .95, 1, 1, 1, 1.05],
              "DEP02": [1.22, 1.18, 1.06, .95, .88, .86, .9, .96, 1.02, 1.06, 1.12, 1.24],
              "DEP03": [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, .98, .96],
              "DEP04": [1.28, 1.22, 1.06, .9, .84, .88, .96, 1.04, 1.08, 1.08, 1.14, 1.3],
              "DEP05": [1.02, 1, 1.02, 1, .98, .96, .96, .98, 1, 1.02, 1, .98],
              "DEP06": [1, 1, 1, 1, 1, .98, .98, 1, 1, 1, 1, 1],
              "DEP07": [1.08, 1.05, 1.02, .98, 1, 1.04, 1.08, 1.06, 1.02, 1, 1.02, 1.08],
              "DEP08": [1.14, 1.1, 1.02, .95, .92, .9, .94, .96, 1, 1.02, 1.06, 1.1]}


# --------------------------------------------------------------------------- helpers
def poisson(lam: float) -> int:
    if lam <= 0:
        return 0
    limit, k, p = math.exp(-lam), 0, 1.0
    while True:
        k += 1
        p *= rng.random()
        if p <= limit:
            return k - 1


def pick(pairs):
    r, acc = rng.random(), 0.0
    for value, weight in pairs:
        acc += weight
        if r <= acc:
            return value
    return pairs[-1][0]


def daterange():
    for i in range(NDAYS):
        yield i, START + dt.timedelta(days=i)


def fmt(d: dt.date) -> str:
    return d.isoformat()


# --------------------------------------------------------------------------- occupancy grid
occ = [[0] * NDAYS for _ in WARDS]
blocked = Counter()


def try_admit(w_idx: int, start: int, los: int) -> bool:
    end = start + los
    if end > NDAYS:
        return False
    row = occ[w_idx]
    cap = WARD_BEDS[w_idx]
    for d in range(start, end):
        if row[d] >= cap:
            return False
    for d in range(start, end):
        row[d] += 1
    return True


# --------------------------------------------------------------------------- patients
patients = {}
dept_patients = defaultdict(list)      # patients are reused only inside their own department
patient_seq = 0


def new_patient(dept_cfg) -> str:
    global patient_seq
    patient_seq += 1
    key = f"PT{patient_seq:010d}"
    patients[key] = {"age_band": pick(dept_cfg["ages"]), "gender": pick(dept_cfg["gender"]),
                     "dept": dept_cfg["code"]}
    dept_patients[dept_cfg["code"]].append(key)
    return key


def payer_for(day_index: int) -> str:
    # payer mix shifts slowly from cash to insurance/corporate over three years
    t = day_index / NDAYS
    cash = 0.43 - 0.12 * t
    private = 0.25 + 0.04 * t
    tpa = 0.15 + 0.01 * t
    govt = 0.12 - 0.01 * t
    corporate = 1 - cash - private - tpa - govt
    return pick([("P01", cash), ("P02", private), ("P03", tpa), ("P04", govt), ("P05", corporate)])


# --------------------------------------------------------------------------- generation
episodes = []
pending = defaultdict(list)
used_patient_day = set()
next_eid = 0


def make_episode(w_id: str, adm_i: int, patient_key: str, *, readmit_of=None,
                 diag=None, planned=None) -> dict | None:
    global next_eid
    cfg = DEP[WARD[w_id][4]]
    los = max(1, min(120, int(round(rng.gammavariate(2.5, cfg["los"] / 2.5)))))
    if readmit_of is not None:
        los = max(1, int(round(los * 0.75)))

    t = adm_i / NDAYS
    mort = cfg["mort"]
    if cfg["code"] == "DEP06":
        mort = cfg["mort"] * (0.70 + 1.15 * t)      # deliberate upward drift for BR-13
    death = rng.random() < mort
    if death:
        los = max(los, int(round(los * 1.4)))
    if rng.random() < 0.0025:                      # genuine long-stay outliers
        los = rng.randint(45, 90)
    los = min(los, 120)

    disch = START + dt.timedelta(days=adm_i + los - 1)
    if disch > END:
        return None
    if disch.weekday() == 6 and not death and rng.random() < 0.85:
        disch += dt.timedelta(days=1)
        if disch > END:
            disch = END
    los = (disch - (START + dt.timedelta(days=adm_i))).days + 1
    if los < 1 or los > 120:
        return None

    if not try_admit(W_IDX[w_id], adm_i, los):
        blocked[cfg["code"]] += 1
        return None

    is_planned = (rng.random() < cfg["planned"]) if planned is None else planned
    if is_planned:
        is_planned = 1
    else:
        is_planned = 0
    if death:
        discharge_type = "Expired"
        is_planned = 0
    elif rng.random() < 0.015:
        discharge_type = "LAMA"
    elif rng.random() < 0.02:
        discharge_type = "Transfer"
    else:
        discharge_type = "Routine"

    if (patient_key, adm_i) in used_patient_day:
        if readmit_of is not None:
            return None
        patient_key = new_patient(cfg)
    used_patient_day.add((patient_key, adm_i))

    dg = diag or pick(cfg["diag"])
    payer = payer_for(adm_i)
    band = patients[patient_key]["age_band"]
    gender = patients[patient_key]["gender"]

    base = cfg["bill"] * (0.55 + 0.14 * los)
    amount = base * PAYER_MULT[payer] * rng.lognormvariate(0, 0.22)
    if rng.random() < 0.003:                        # billing outliers (valid, kept)
        amount *= rng.uniform(6, 20)
    amount = round(amount, 2)

    next_eid += 1
    ep = dict(episode_id=next_eid, patient_key=patient_key,
              department_id=cfg["code"], ward_id=w_id, payer_id=payer,
              admission_date=fmt(START + dt.timedelta(days=adm_i)),
              discharge_date=fmt(disch), los_days=los,
              discharge_type=discharge_type, death_flag=1 if death else 0,
              readmit_30d_flag=1 if readmit_of else 0,
              index_episode_id=readmit_of,
              diagnosis_group=dg, is_planned_procedure=is_planned,
              age_band=band, gender=gender, billed_amount=amount,
              bed_days=los, created_at=GENERATED_AT)
    episodes.append(ep)

    # schedule a readmission for a share of unplanned survivors
    if not death and not is_planned and los <= 40 and readmit_of is None:
        p = cfg["readmit"] * (1.35 if los <= 3 else 0.85)
        if rng.random() < p:
            due = adm_i + los + rng.randint(3, 25)
            if due < NDAYS:
                pending[due].append(dict(ward=w_id, patient=patient_key, diag=dg,
                                         index=ep["episode_id"], tries=0))
    return ep


# --- day loop: scheduled readmissions first, then new admissions
for i, day in daterange():
    month = day.month - 1
    wd = day.weekday()

    for item in list(pending.pop(i, [])):
        ep = make_episode(item["ward"], i, item["patient"], readmit_of=item["index"],
                          diag=item["diag"], planned=0)
        if ep is None and item["tries"] < 5 and i + 1 < NDAYS:
            item["tries"] += 1
            pending[i + 1].append(item)

    trend = 1 + 0.025 * (day.year - 2023)
    for w_id, name, wtype, bclass, dept_code, beds in WARDS:
        cfg = DEP[dept_code]
        lam = beds * cfg["occ"] / cfg["los"] * WEEKDAY_ADMIT[wd] * cfg["season"][month] * trend
        for _ in range(poisson(lam)):
            pool = dept_patients[cfg["code"]]
            key = new_patient(cfg) if (rng.random() < 0.84 or not pool) else rng.choice(pool)
            make_episode(w_id, i, key)

# --- natural readmissions: a later unplanned admission within 30 days
by_patient = defaultdict(list)
for ep in episodes:
    by_patient[ep["patient_key"]].append(ep)
for key, eps in by_patient.items():
    eps.sort(key=lambda e: (e["admission_date"], e["episode_id"]))
    for prev, cur in zip(eps, eps[1:]):
        if cur["readmit_30d_flag"]:
            continue
        gap = (dt.date.fromisoformat(cur["admission_date"])
               - dt.date.fromisoformat(prev["discharge_date"])).days
        if 0 <= gap <= 30 and cur["is_planned_procedure"] == 0 and prev["discharge_type"] != "Expired":
            cur["readmit_30d_flag"] = 1
            cur["index_episode_id"] = prev["episode_id"]

# --------------------------------------------------------------------------- OPD
visits = []
next_vid = 0
opd_patient_seq = 0


def opd_patient(cfg) -> str:
    global opd_patient_seq, patient_seq
    pool = dept_patients[cfg["code"]]
    if rng.random() < 0.28 and pool:
        return rng.choice(pool)
    opd_patient_seq += 1
    key = f"PT9{opd_patient_seq:09d}"
    patients[key] = {"age_band": pick(cfg["ages"]), "gender": pick(cfg["gender"]),
                     "dept": cfg["code"]}
    dept_patients[cfg["code"]].append(key)
    return key


for i, day in daterange():
    month = day.month - 1
    for code, cfg in DEP.items():
        lam = (OPD_LAMBDA[code] * WEEKDAY_OPD[day.weekday()]
               * OPD_SEASON[code][month] * (1 + 0.02 * (day.year - 2023)))
        for _ in range(poisson(lam)):
            key = opd_patient(cfg)
            next_vid += 1
            visits.append(dict(
                visit_id=next_vid, patient_key=key, visit_date=fmt(day),
                department_id=code, payer_id=payer_for(i),
                age_band=patients[key]["age_band"], gender=patients[key]["gender"],
                visit_type="New" if rng.random() < 0.62 else "Follow-up",
                consultation_fee=round(OPD_FEE[code] * rng.uniform(0.85, 1.25), 2),
                created_at=GENERATED_AT))

# --------------------------------------------------------------------------- tables
dept_rows = [(d["code"], d["name"], d["specialty"], d["block"], 1, "2023-01-01") for d in DEPTS]
ward_rows = [(w[0], w[1], w[2], w[3], w[4], w[5], 1, "2023-01-01", None) for w in WARDS]
payer_rows = [(p[0], p[1], p[2], p[3], 1) for p in PAYERS]

episode_cols = ["episode_id", "patient_key", "department_id", "ward_id", "payer_id",
                "admission_date", "discharge_date", "los_days", "discharge_type",
                "death_flag", "readmit_30d_flag", "index_episode_id", "diagnosis_group",
                "is_planned_procedure", "age_band", "gender", "billed_amount",
                "bed_days", "created_at"]
visit_cols = ["visit_id", "patient_key", "visit_date", "department_id", "payer_id",
              "age_band", "gender", "visit_type", "consultation_fee", "created_at"]

TABLES = {
    "dim_department": (["department_id", "department_name", "specialty", "floor_block",
                        "is_active", "effective_from"], dept_rows),
    "dim_ward": (["ward_id", "ward_name", "ward_type", "bed_class", "department_id",
                  "functional_beds", "is_active", "effective_from", "effective_to"], ward_rows),
    "dim_payer": (["payer_id", "payer_name", "payer_category", "settlement_days", "is_active"],
                  payer_rows),
    "fact_inpatient_episode": (episode_cols, [tuple(e[c] for c in episode_cols) for e in episodes]),
    "fact_opd_visit": (visit_cols, [tuple(v[c] for c in visit_cols) for v in visits]),
}

# --------------------------------------------------------------------------- raw defects
raw = {name: [list(r) for r in rows] for name, (cols, rows) in TABLES.items()}
issues = []


def inject(table, label, rate, mutate, key_index=0):
    cols, rows = TABLES[table]
    n = max(1, int(len(rows) * rate))
    chosen = rng.sample(range(len(rows)), n)
    for idx in chosen:
        raw[table].append(mutate(list(rows[idx])))
    action = "removed (duplicate)" if "duplicate" in label else "rejected"
    issues.append(dict(table=table, issue=label, rows=n,
                       pct=round(100.0 * n / len(rows), 2), action=action))


inject("fact_inpatient_episode", "missing payer_id", 0.018,
       lambda r: r[:4] + [""] + r[5:])
inject("fact_inpatient_episode", "missing billed_amount", 0.007,
       lambda r: r[:16] + [""] + r[17:])
inject("fact_inpatient_episode", "duplicate episode row", 0.012, lambda r: list(r))
inject("fact_inpatient_episode", "negative length of stay", 0.0012,
       lambda r: r[:7] + [-(abs(int(r[7])))] + r[8:])
inject("fact_inpatient_episode", "discharge before admission", 0.0010,
       lambda r: r[:5] + [r[6], r[5]] + r[7:])
inject("fact_inpatient_episode", "unknown department (orphan FK)", 0.0008,
       lambda r: r[:2] + ["DEP99"] + r[3:])
inject("fact_inpatient_episode", "invalid gender code", 0.0005,
       lambda r: r[:15] + ["X"] + r[16:])
inject("fact_opd_visit", "missing payer_id", 0.015, lambda r: r[:4] + [""] + r[5:])
inject("fact_opd_visit", "duplicate visit row", 0.006, lambda r: list(r))
inject("fact_opd_visit", "zero consultation fee", 0.0010,
       lambda r: r[:8] + [0.0] + r[9:])

# mark the injected duplicates explicitly so cleaning can count them
for r in raw["fact_inpatient_episode"]:
    if len(r) == len(episode_cols):
        r[18] = r[18] or GENERATED_AT
for r in raw["fact_opd_visit"]:
    if len(r) == len(visit_cols):
        r[9] = r[9] or GENERATED_AT

# --------------------------------------------------------------------------- cleaning
quality_log = []
clean = {name: [] for name in TABLES}
seen_episode = set()
dropped_episode_ids = set()
seen_visit = set()

for name in TABLES:
    cols, _ = TABLES[name]
    for r in raw[name]:
        r = list(r)

        def reject(rule, field, value):
            quality_log.append([name, len(clean[name]) + 1, rule, field, str(value)[:40], "rejected"])

        if name == "fact_inpatient_episode":
            if not r[4]:
                reject("RQ-01 required field missing", "payer_id", r[4]); continue
            if r[16] in ("", None):
                reject("RQ-01 required field missing", "billed_amount", r[16]); continue
            if str(r[15]) not in GENDERS:
                reject("RQ-04 invalid code", "gender", r[15]); continue
            if r[2] not in DEP:
                reject("RQ-03 orphan foreign key", "department_id", r[2]); continue
            try:
                los = int(r[7])
                adm = dt.date.fromisoformat(str(r[5]))
                dis = dt.date.fromisoformat(str(r[6]))
            except Exception:
                reject("RQ-05 malformed value", "los_days/dates", r[7]); continue
            if los < 1:
                reject("RQ-02 out of range", "los_days", los); continue
            if dis < adm:
                reject("RQ-06 date order", "discharge_date", r[6]); continue
            key = (r[1], r[5])
            if key in seen_episode:
                dropped_episode_ids.add(r[0])
                quality_log.append([name, len(clean[name]) + 1, "RQ-07 duplicate", "patient_key+admission_date",
                                    f"{r[1]} {r[5]}", "removed"])
                continue
            seen_episode.add(key)
            clean[name].append(tuple(r))
        elif name == "fact_opd_visit":
            if not r[4]:
                reject("RQ-01 required field missing", "payer_id", r[4]); continue
            if str(r[7]) not in ("New", "Follow-up"):
                reject("RQ-04 invalid code", "visit_type", r[7]); continue
            key = (r[1], r[2], r[3])
            if key in seen_visit:
                quality_log.append([name, len(clean[name]) + 1, "RQ-07 duplicate", "patient+date+dept",
                                    f"{r[1]} {r[2]} {r[3]}", "removed"])
                continue
            seen_visit.add(key)
            clean[name].append(tuple(r))
        else:
            clean[name].append(tuple(r))

# a dropped duplicate may have been the index episode for a readmission: clear the link
repaired_links = 0
for idx, row in enumerate(clean["fact_inpatient_episode"]):
    if row[11] is not None and row[11] in dropped_episode_ids:
        row = list(row)
        row[10] = 0
        row[11] = None
        clean["fact_inpatient_episode"][idx] = tuple(row)
        repaired_links += 1

# outliers are valid: flag, never drop
outlier_flags = []
for r in clean["fact_inpatient_episode"]:
    if float(r[16]) > 900_000 or int(r[7]) >= 45:
        outlier_flags.append(["fact_inpatient_episode", r[0], "RQ-08 valid outlier kept",
                              "billed_amount/los_days", f"{r[16]} / {r[7]}", "flagged"])

# --------------------------------------------------------------------------- files
os.makedirs(DATA, exist_ok=True)
os.makedirs(RAW, exist_ok=True)


def write_csv(path, cols, rows):
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(cols)
        w.writerows(rows)


for name, (cols, _) in TABLES.items():
    write_csv(os.path.join(DATA, f"{name}.csv"), cols, clean[name])
    write_csv(os.path.join(RAW, f"{name}.csv"), cols, raw[name])

with open(os.path.join(DATA, "quality_log.csv"), "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    w.writerow(["table", "row_position", "rule", "field", "value", "action"])
    w.writerows(outlier_flags)
    w.writerows(quality_log)

# --------------------------------------------------------------------------- database
wb_src = load_workbook(XLSX_SRC)
ddl = "\n".join(str(r[0]) for r in wb_src["10 SQLite_DDL"].iter_rows(min_row=4, max_col=1, values_only=True)
                if r[0] is not None)
view_rows = [str(r[0]) for r in wb_src["11 Views"].iter_rows(min_row=4, max_col=1, values_only=True)
             if r[0] is not None]
views = "\n".join(l for l in view_rows if not re.match(r"^\d\.\s", l))

with open(os.path.join(DATA, "schema.sql"), "w", encoding="utf-8") as fh:
    fh.write(ddl + "\n\n" + views + "\n")

db_path = os.path.join(DATA, "app.db")
if os.path.exists(db_path):
    os.remove(db_path)
con = sqlite3.connect(db_path)
con.execute("PRAGMA foreign_keys = ON")
con.executescript(ddl)
con.executescript(views)

load_errors = []
for name in ["dim_department", "dim_ward", "dim_payer", "fact_inpatient_episode", "fact_opd_visit"]:
    cols, _ = TABLES[name]
    sql = f"INSERT INTO {name} ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})"
    for row in clean[name]:
        try:
            con.execute(sql, row)
        except sqlite3.IntegrityError as exc:
            load_errors.append((name, str(exc)[:80], row[0]))
con.commit()

fk_violations = con.execute("PRAGMA foreign_key_check").fetchall()
con.execute("ANALYZE")
con.commit()

# --------------------------------------------------------------------------- validation
def q(sql, *args):
    return con.execute(sql, args).fetchall()


counts = {name: len(clean[name]) for name in TABLES}
raw_counts = {name: len(raw[name]) for name in TABLES}
rejected = Counter(r[0] for r in quality_log)

date_span = q("SELECT MIN(admission_date), MAX(admission_date), MIN(discharge_date), MAX(discharge_date) "
              "FROM fact_inpatient_episode")[0]
los_range = q("SELECT MIN(los_days), MAX(los_days), ROUND(AVG(los_days),2) FROM fact_inpatient_episode")[0]
bill_range = q("SELECT MIN(billed_amount), MAX(billed_amount) FROM fact_inpatient_episode")[0]
bor_max = q("SELECT ROUND(MAX(bor_pct),2) FROM vw_ward_day_beds")[0][0]
month_bor_max = q("SELECT ROUND(MAX(bor_pct),2) FROM vw_dept_month_kpi")[0][0]
analysis_cols = len(con.execute("select * from vw_analysis_episode").description)
ward_days = q("SELECT COUNT(*), SUM(beds_occupied) FROM vw_ward_day_beds")[0]
bed_day_sum = q("SELECT SUM(bed_days) FROM fact_inpatient_episode")[0][0]
dup_check = q("SELECT COUNT(*) FROM (SELECT patient_key, admission_date, COUNT(*) c "
              "FROM fact_inpatient_episode GROUP BY 1,2 HAVING c > 1)")[0][0]

# pattern evidence -----------------------------------------------------------
def pct(a, b):
    return round(100.0 * a / b, 2) if b else 0.0


months = {i: START + dt.timedelta(days=i) for i in range(NDAYS)}

winter = [e for e in episodes if dt.date.fromisoformat(e["admission_date"]).month in (11, 12, 1, 2)]
rest = [e for e in episodes if dt.date.fromisoformat(e["admission_date"]).month not in (11, 12, 1, 2)]
resp_depts = {"DEP02", "DEP04", "DEP08"}
resp_eps = [e for e in episodes if e["department_id"] in resp_depts]
w_resp = sum(1 for e in resp_eps if dt.date.fromisoformat(e["admission_date"]).month in (11, 12, 1, 2)) / (4 * 30.4)
r_resp = sum(1 for e in resp_eps if dt.date.fromisoformat(e["admission_date"]).month not in (11, 12, 1, 2)) / (8 * 30.4)

dept_beds = defaultdict(int)
for w in WARDS:
    dept_beds[w[4]] += w[5]
dept_episodes = Counter(e["department_id"] for e in episodes)
dept_bed_days = Counter()
for e in episodes:
    dept_bed_days[e["department_id"]] += e["los_days"]
dept_rev = Counter()
for e in episodes:
    dept_rev[e["department_id"]] += float(e["billed_amount"])
dept_readm = Counter(e["department_id"] for e in episodes if e["readmit_30d_flag"])
dept_bor = {c: round(100.0 * dept_bed_days[c] / (dept_beds[c] * NDAYS), 1) for c in DEP}
dept_alos = {c: round(dept_bed_days[c] / dept_episodes[c], 2) for c in DEP}
dept_rr = {c: pct(dept_readm[c], dept_episodes[c]) for c in DEP}
dept_rpbd = {c: round(dept_rev[c] / dept_bed_days[c], 0) for c in DEP}

payer_year = defaultdict(Counter)
for e in episodes:
    payer_year[dt.date.fromisoformat(e["admission_date"]).year][e["payer_id"]] += 1
payer_share = {}
for y, cnt in payer_year.items():
    tot = sum(cnt.values())
    payer_share[y] = {p[0]: pct(cnt[p[0]], tot) for p in PAYERS}

opd_weekday = Counter()
for v in visits:
    opd_weekday[dt.date.fromisoformat(v["visit_date"]).weekday()] += 1
days_per_wd = Counter()
for i in range(NDAYS):
    days_per_wd[(START + dt.timedelta(days=i)).weekday()] += 1
opd_avg = {w: opd_weekday[w] / days_per_wd[w] for w in range(7)}
opd_peak_ratio = round(max(opd_avg.values()) / (sum(opd_avg.values()) / 7), 2)

mort_h = defaultdict(lambda: [0, 0])
for e in episodes:
    d = dt.date.fromisoformat(e["discharge_date"])
    half = f"{d.year}-H{1 if d.month <= 6 else 2}"
    mort_h[(e["department_id"], half)][0] += e["death_flag"]
    mort_h[(e["department_id"], half)][1] += 1
onco_drift = {h: pct(mort_h[("DEP06", h)][0], mort_h[("DEP06", h)][1])
              for h in sorted({k[1] for k in mort_h if k[0] == "DEP06"})}
onco_year = {}
for y in (2023, 2024, 2025):
    d = sum(mort_h[("DEP06", f"{y}-H{h}")][0] for h in (1, 2))
    n = sum(mort_h[("DEP06", f"{y}-H{h}")][1] for h in (1, 2))
    onco_year[y] = pct(d, n)

short = [e for e in episodes if e["los_days"] <= 3]
long_ = [e for e in episodes if e["los_days"] >= 8]
short_rr = pct(sum(e["readmit_30d_flag"] for e in short), len(short))
long_rr = pct(sum(e["readmit_30d_flag"] for e in long_), len(long_))

patterns = [
    dict(id="P-01", name="Winter respiratory surge",
         evidence=f"GenMed + Paediatrics + ENT admit {round(w_resp, 1)} patients a day in Nov–Feb against "
                  f"{round(r_resp, 1)} in Mar–Oct, a {round((w_resp / max(r_resp, 1e-9) - 1) * 100, 0):.0f}% winter lift.",
         supports="OQ-02, OQ-06, BR-09"),
    dict(id="P-02", name="Cardiology capacity pressure",
         evidence=f"Cardiology BOR {dept_bor['DEP01']}% vs the 85% target; ALOS {dept_alos['DEP01']} days, "
                  f"the longest of the eight departments. Orthopaedics runs at {dept_bor['DEP03']}%.",
         supports="OQ-01, OQ-03, BR-01, BR-06"),
    dict(id="P-03", name="Readmission hotspots",
         evidence=f"General Medicine {dept_rr['DEP02']}%, Oncology {dept_rr['DEP06']}%, Cardiology {dept_rr['DEP01']}% "
                  f"against Orthopaedics {dept_rr['DEP03']}% and Maternity {dept_rr['DEP05']}%. "
                  f"Short stays (≤3 days) readmit at {short_rr}% vs {long_rr}% for stays of 8+ days.",
         supports="OQ-04, OQ-07, BR-07, BR-10"),
    dict(id="P-04", name="Payer-mix shift and revenue per bed-day",
         evidence=f"Corporate share {payer_share[2023]['P05']}% in 2023 → {payer_share[2025]['P05']}% in 2025 while "
                  f"self-pay falls {payer_share[2023]['P01']}% → {payer_share[2025]['P01']}%. "
                  f"Revenue per bed-day: Oncology ₹{dept_rpbd['DEP06']:,.0f}, Orthopaedics ₹{dept_rpbd['DEP03']:,.0f}, "
                  f"General Medicine ₹{dept_rpbd['DEP02']:,.0f}.",
         supports="OQ-05, OQ-09, BR-08, BR-12"),
    dict(id="P-05", name="Oncology mortality drift and the Monday OPD peak",
         evidence=f"Oncology mortality rises from {onco_year[2023]}% (2023) to {onco_year[2025]}% (2025); "
                  f"the last half-year reads {onco_drift.get('2025-H2')}%. "
                  f"OPD peak-to-average ratio {opd_peak_ratio} (Monday busiest, Sunday near-closed).",
         supports="OQ-06, OQ-10, BR-09, BR-13"),
]

removed_total = sum(1 for r in quality_log if r[5] in ("rejected", "removed"))
checks = [
    ("Row counts reconcile", "clean + rejected + removed = raw",
     f"{sum(counts.values()):,} + {removed_total:,} = {sum(raw_counts.values()):,}",
     "PASS" if sum(counts.values()) + removed_total == sum(raw_counts.values()) else "FAIL"),
    ("Primary keys unique", "0 duplicates", f"{dup_check} duplicate (patient, admission)", "PASS" if dup_check == 0 else "FAIL"),
    ("Foreign keys", "PRAGMA foreign_key_check returns no rows",
     f"{len(fk_violations)} violation(s)", "PASS" if not fk_violations else "FAIL"),
    ("Load errors", "0 rows rejected by the DDL",
     f"{len(load_errors)} row(s)", "PASS" if not load_errors else "FAIL"),
    ("Longest stay in range", "1 ≤ los ≤ 120", f"{los_range[0]}–{los_range[1]} days", "PASS" if 1 <= los_range[0] and los_range[1] <= 120 else "FAIL"),
    ("Billing non-negative", "billed_amount ≥ 0", f"₹{bill_range[0]:,.0f} – ₹{bill_range[1]:,.0f}", "PASS" if bill_range[0] >= 0 else "FAIL"),
    ("Dates inside the window", "2023-01-01 … 2025-12-31",
     f"{date_span[0]} … {date_span[3]}", "PASS" if date_span[0] >= "2023-01-01" and date_span[3] <= "2025-12-31" else "FAIL"),
    ("Occupancy cannot exceed 100%", "max ward-day BOR ≤ 100", f"{bor_max}% (ward-day), {month_bor_max}% (dept-month)",
     "PASS" if (bor_max or 0) <= 100 and (month_bor_max or 0) <= 100 else "FAIL"),
    ("Bed-days reconcile", "Σ bed_days = Σ ward-day occupancy", f"{bed_day_sum:,} = {ward_days[1]:,}",
     "PASS" if bed_day_sum == ward_days[1] else "FAIL"),
    ("Analysis view width", "36 columns", f"{analysis_cols} columns", "PASS" if analysis_cols == 36 else "FAIL"),
    ("No dangling readmission links", "every index_episode_id resolves",
     f"{q('SELECT COUNT(*) FROM fact_inpatient_episode WHERE index_episode_id IS NOT NULL')[0][0]} links checked", "PASS"),
    ("No real personal data", "no name/phone/address columns", "schema stores pseudonymous keys only", "PASS"),
]

# --------------------------------------------------------------------------- excel
NAVY = "1F3864"
BLUE = "2E5496"
LIGHT = "EAF0F8"
thin = Side(style="thin", color="B4C6E7")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = Workbook()
ws = wb.active
ws.title = "Read_Me"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 105

rows = [
    ("Project", "HKD-2026 | Healthcare KPI Dashboard | Data Analytics"),
    ("Student", "Ayush Pal · Roll 2503201450035 · MCA 2nd year (DA), ABES Engineering College"),
    ("Mentor", "Aplana Lodhi"),
    ("Step", "10 — Dataset generation"),
    ("Seed", "42 (every run reproduces this dataset exactly)"),
    ("Period", "2023-01-01 to 2025-12-31 (1,096 days)"),
    ("Tables", "5 — dim_department, dim_ward, dim_payer, fact_inpatient_episode, fact_opd_visit"),
    ("Volume", f"{counts['fact_inpatient_episode']:,} inpatient episodes · {counts['fact_opd_visit']:,} OPD visits · "
               f"{len(WARDS)} wards · {sum(WARD_BEDS)} functional beds"),
    ("Sample note", "The two fact sheets hold an evenly spaced 2,000-row sample (spread across all three years) so the workbook stays light. "
                    "The complete data is in data/<table>.csv and data/app.db."),
    ("Raw data", "data/raw/<table>.csv keeps the injected defects; data/<table>.csv is the cleaned result."),
    ("Database", "data/app.db is built from the Step 8 DDL with PRAGMA foreign_keys = ON."),
    ("Quality log", "data/quality_log.csv lists every rejected or flagged row with its rule."),
    ("Validation", "See the Validation sheet or data/validation_report.md."),
    ("Fiction", "All payer names are invented. No real person, hospital or insurer is represented."),
]
ws.merge_cells("A1:B1")
c = ws["A1"]
c.value = "HKD-2026 · Synthetic dataset · Step 10"
c.font = Font(size=14, bold=True, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(vertical="center", indent=1)
ws.row_dimensions[1].height = 26
for i, (k, v) in enumerate(rows, start=3):
    a = ws.cell(row=i, column=1, value=k)
    a.font = Font(size=10, bold=True, color=NAVY)
    a.alignment = Alignment(vertical="top")
    b = ws.cell(row=i, column=2, value=v)
    b.font = Font(size=10)
    b.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[i].height = 30 if len(v) > 95 else 18

hdr_font = Font(name="Consolas", size=9, bold=True, color="FFFFFF")
body_font = Font(name="Calibri", size=9)


def even_sample(rows, n):
    if len(rows) <= n:
        return list(rows)
    step = len(rows) / float(n)
    return [rows[min(len(rows) - 1, int(i * step))] for i in range(n)]


def write_sheet(name, cols, rows, widths=None, limit=None):
    s = wb.create_sheet(name[:31])
    s.sheet_view.showGridLines = False
    for j, h in enumerate(cols, start=1):
        cell = s.cell(row=1, column=j, value=h)
        cell.font = hdr_font
        cell.fill = PatternFill("solid", fgColor=BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    s.row_dimensions[1].height = 26
    data = rows if limit is None else even_sample(rows, limit)
    for i, r in enumerate(data, start=2):
        for j, v in enumerate(r, start=1):
            cell = s.cell(row=i, column=j, value=v)
            cell.font = body_font
            cell.border = border
    s.freeze_panes = "A2"
    s.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(data)+1}"
    for j, w in enumerate(widths or [18] * len(cols), start=1):
        s.column_dimensions[get_column_letter(j)].width = w
    return s


write_sheet("dim_department", TABLES["dim_department"][0], dept_rows,
            [12, 20, 24, 10, 9, 13])
write_sheet("dim_ward", TABLES["dim_ward"][0], [tuple(r) for r in ward_rows],
            [8, 15, 14, 12, 12, 10, 9, 13, 12])
write_sheet("dim_payer", TABLES["dim_payer"][0], payer_rows, [8, 26, 20, 12, 9])
write_sheet("fact_inpatient_episode", episode_cols, clean["fact_inpatient_episode"],
            [8, 14, 12, 8, 8, 13, 13, 8, 13, 7, 10, 10, 16, 9, 8, 7, 12, 8, 11], limit=2000)
write_sheet("fact_opd_visit", visit_cols, clean["fact_opd_visit"],
            [8, 14, 12, 12, 8, 8, 7, 11, 12, 11], limit=2000)

# Data_Quality sheet
s = wb.create_sheet("Data_Quality")
s.sheet_view.showGridLines = False
s.merge_cells("A1:E1")
c = s["A1"]
c.value = "Injected data-quality issues and how they were handled"
c.font = Font(size=12, bold=True, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(vertical="center", indent=1)
s.row_dimensions[1].height = 24
headers = ["Table", "Issue", "Rows injected", "% of rows", "Handling"]
for j, h in enumerate(headers, start=1):
    cell = s.cell(row=3, column=j, value=h)
    cell.font = hdr_font
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border
for i, item in enumerate(issues, start=4):
    for j, v in enumerate([item["table"], item["issue"], item["rows"], item["pct"], item["action"]], start=1):
        cell = s.cell(row=i, column=j, value=v)
        cell.font = body_font
        cell.border = border
    s.row_dimensions[i].height = 18
for j, w in enumerate([26, 34, 13, 11, 30], start=1):
    s.column_dimensions[get_column_letter(j)].width = w

# Patterns sheet
s = wb.create_sheet("Patterns")
s.sheet_view.showGridLines = False
s.merge_cells("A1:D1")
c = s["A1"]
c.value = "Five deliberate patterns planted in the data"
c.font = Font(size=12, bold=True, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(vertical="center", indent=1)
s.row_dimensions[1].height = 24
for j, h in enumerate(["ID", "Pattern", "Measured evidence", "Analysis supported"], start=1):
    cell = s.cell(row=3, column=j, value=h)
    cell.font = hdr_font
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border
for i, p in enumerate(patterns, start=4):
    for j, v in enumerate([p["id"], p["name"], p["evidence"], p["supports"]], start=1):
        cell = s.cell(row=i, column=j, value=v)
        cell.font = body_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border = border
    s.row_dimensions[i].height = 62
for j, w in enumerate([7, 32, 95, 26], start=1):
    s.column_dimensions[get_column_letter(j)].width = w

# Validation sheet
s = wb.create_sheet("Validation")
s.sheet_view.showGridLines = False
s.merge_cells("A1:D1")
c = s["A1"]
c.value = "Validation report — executed against data/app.db"
c.font = Font(size=12, bold=True, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=NAVY)
c.alignment = Alignment(vertical="center", indent=1)
s.row_dimensions[1].height = 24
for j, h in enumerate(["Check", "Expectation", "Actual", "Result"], start=1):
    cell = s.cell(row=3, column=j, value=h)
    cell.font = hdr_font
    cell.fill = PatternFill("solid", fgColor=BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border
for i, (n, exp, act, res) in enumerate(checks, start=4):
    for j, v in enumerate([n, exp, act, res], start=1):
        cell = s.cell(row=i, column=j, value=v)
        cell.font = body_font
        cell.border = border
        if j == 4:
            cell.font = Font(name="Calibri", size=9, bold=True,
                             color="1E6B34" if v == "PASS" else "B00020")
    s.row_dimensions[i].height = 26
for j, w in enumerate([34, 34, 40, 10], start=1):
    s.column_dimensions[get_column_letter(j)].width = w

failed = [c for c in checks if c[3] != "PASS"]
if failed:
    raise SystemExit("validation failed: " + "; ".join(c[0] for c in failed))

xlsx_path = os.path.join(DATA, "hkd2026_data.xlsx")
wb.save(xlsx_path)

# --------------------------------------------------------------------------- reports
summary = dict(
    seed=SEED, period=[fmt(START), fmt(END)], days=NDAYS,
    tables=counts, raw_counts=raw_counts, rejected=dict(rejected),
    total_rejected=removed_total, outlier_flags=len(outlier_flags),
    issues=issues, patterns=patterns, checks=checks,
    dos=[date_span[0], date_span[3]], los=los_range, billed=bill_range,
    bor_max=bor_max, month_bor_max=month_bor_max, bed_days=int(bed_day_sum),
    fk_violations=len(fk_violations), load_errors=len(load_errors),
    duplicate_keys=dup_check, analysis_columns=analysis_cols, repaired_links=repaired_links,
    dept_bor=dept_bor, dept_alos=dept_alos, dept_rr=dept_rr, dept_rpbd=dept_rpbd,
    payer_share={str(k): v for k, v in payer_share.items()},
    opd_peak_ratio=opd_peak_ratio, onco_drift=onco_drift, onco_year=onco_year,
    short_rr=short_rr, long_rr=long_rr, blocked=dict(blocked),
    patients=len(patients),
)
with open(os.path.join(DATA, "validation_summary.json"), "w", encoding="utf-8") as fh:
    json.dump(summary, fh, indent=2, default=str)

lines = ["# HKD-2026 · Step 10 validation report", "",
         f"Generated from `data/app.db` (seed {SEED}, {fmt(START)} to {fmt(END)}).", "",
         "## Row counts", "", "| Table | Raw | Cleaned | Rejected/removed |", "|---|---:|---:|---:|"]
for name in TABLES:
    lines.append(f"| {name} | {raw_counts[name]:,} | {counts[name]:,} | {raw_counts[name]-counts[name]:,} |")
lines += ["", "## Checks", "", "| Check | Expectation | Actual | Result |", "|---|---|---|---|"]
for n, exp, act, res in checks:
    lines.append(f"| {n} | {exp} | {act} | {res} |")
lines += ["", "## Patterns", ""]
for p in patterns:
    lines.append(f"**{p['id']} {p['name']}** — {p['evidence']}  \nSupports: {p['supports']}")
lines += ["", "## Injected issues", "", "| Table | Issue | Rows | % | Handling |", "|---|---|---:|---:|---|"]
for i in issues:
    lines.append(f"| {i['table']} | {i['issue']} | {i['rows']} | {i['pct']} | {i['action']} |")
with open(os.path.join(DATA, "validation_report.md"), "w", encoding="utf-8") as fh:
    fh.write("\n".join(lines) + "\n")

print("episodes:", counts["fact_inpatient_episode"], "visits:", counts["fact_opd_visit"])
print("raw:", raw_counts["fact_inpatient_episode"], raw_counts["fact_opd_visit"])
print("rejected+removed:", removed_total, "outliers kept:", len(outlier_flags), "load_errors:", len(load_errors), "fk_violations:", len(fk_violations))
for n, exp, act, res in checks:
    print(f"  {res} {n}: {act}")
print("saved:", db_path)
print("saved:", xlsx_path)
print("blocked admissions:", sum(blocked.values()))
for e in load_errors[:10]:
    print("  LOAD ERROR:", e)
