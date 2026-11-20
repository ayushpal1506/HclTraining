"""Data-layer tests — TC-T-01 … TC-T-20.

These run against the shipped warehouse with the standard library only, so they
are green the moment data/app.db exists. They are the cheapest tests in the
suite and the ones that catch a bad regeneration before any page renders.

    .venv/bin/python -m pytest tests/test_data.py -q
"""
from __future__ import annotations

import csv
import sqlite3
from pathlib import Path

import pytest

DB_MB_MIN, DB_MB_MAX = 30, 80


def test_TC_T_01_db_present_and_read_only(db, root):
    """TC-T-01 | data/app.db exists and opens read-only | file is 30–80 MB and an INSERT raises OperationalError"""
    size_mb = (root / "data" / "app.db").stat().st_size / 1_000_000
    assert DB_MB_MIN <= size_mb <= DB_MB_MAX, f"database is {size_mb:.1f} MB"
    with pytest.raises(sqlite3.OperationalError):
        db.execute("INSERT INTO dim_payer (payer_id, payer_name, payer_category, settlement_days, is_active) "
                   "VALUES (99, 'x', 'Cash', 1, 1)")


def test_TC_T_02_schema_objects(db):
    """TC-T-02 | schema contains 5 tables and 3 views | exactly the objects listed in 08_Database_Design"""
    tables = {r["name"] for r in db.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite%'")}
    views = {r["name"] for r in db.execute("SELECT name FROM sqlite_master WHERE type='view'")}
    assert tables == {"dim_department", "dim_ward", "dim_payer",
                      "fact_inpatient_episode", "fact_opd_visit"}, tables
    assert views == {"vw_analysis_episode", "vw_ward_day_beds", "vw_dept_month_kpi"}, views


@pytest.mark.parametrize("table,expected", [
    ("dim_department", 8), ("dim_ward", 20), ("dim_payer", 5),
    ("fact_inpatient_episode", 53_915), ("fact_opd_visit", 143_139)])
def test_TC_T_03_row_counts_match_validated_snapshot(db, summary, table, expected):
    """TC-T-03 | row count per table | equals the seed-42 count in validation_summary.json"""
    assert summary["tables"][table] == expected
    actual = db.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
    assert actual == expected, f"{table}: {actual} rows, expected {expected}"


def test_TC_T_04_primary_keys_not_null(db, tables):
    """TC-T-04 | every primary key is NOT NULL | zero null keys in all five tables"""
    keys = {"dim_department": "department_id", "dim_ward": "ward_id", "dim_payer": "payer_id",
            "fact_inpatient_episode": "episode_id", "fact_opd_visit": "visit_id"}
    for table, pk in keys.items():
        nulls = db.execute(f"SELECT COUNT(*) FROM {table} WHERE {pk} IS NULL").fetchone()[0]
        assert nulls == 0, f"{table}.{pk} has {nulls} nulls"


def test_TC_T_05_primary_keys_unique(db, tables):
    """TC-T-05 | every primary key is unique | COUNT(*) == COUNT(DISTINCT pk)"""
    keys = {"dim_department": "department_id", "dim_ward": "ward_id", "dim_payer": "payer_id",
            "fact_inpatient_episode": "episode_id", "fact_opd_visit": "visit_id"}
    for table, pk in keys.items():
        total = db.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
        distinct = db.execute(f"SELECT COUNT(DISTINCT {pk}) FROM {table}").fetchone()[0]
        assert total == distinct, f"{table}.{pk}: {total} rows, {distinct} distinct"


def test_TC_T_06_foreign_keys_resolve(db):
    """TC-T-06 | no orphan foreign keys | PRAGMA foreign_key_check returns zero rows"""
    assert db.execute("PRAGMA foreign_key_check").fetchall() == []


def test_TC_T_07_stay_arithmetic(db):
    """TC-T-07 | los_days and bed_days agree with the dates | los = discharge − admission + 1, bed_days = los_days, none negative"""
    bad_dates = db.execute(
        "SELECT COUNT(*) FROM fact_inpatient_episode WHERE discharge_date < admission_date").fetchone()[0]
    bad_los = db.execute(
        "SELECT COUNT(*) FROM fact_inpatient_episode "
        "WHERE los_days <> CAST(julianday(discharge_date) - julianday(admission_date) AS INT) + 1").fetchone()[0]
    bad_bed = db.execute(
        "SELECT COUNT(*) FROM fact_inpatient_episode WHERE bed_days <> los_days").fetchone()[0]
    assert (bad_dates, bad_los, bad_bed) == (0, 0, 0), (bad_dates, bad_los, bad_bed)


def test_TC_T_08_los_distribution(db, summary):
    """TC-T-08 | length of stay stays in the documented range | 1 ≤ los ≤ 90 and mean 4.41 days"""
    lo, hi, mean = db.execute(
        "SELECT MIN(los_days), MAX(los_days), ROUND(AVG(los_days), 2) FROM fact_inpatient_episode").fetchone()
    assert (lo, hi) == (1, 90), (lo, hi)
    assert [lo, hi] == summary["los"][:2]
    assert mean == pytest.approx(summary["los"][2], abs=0.01)
    assert mean == pytest.approx(4.41, abs=0.01)


def test_TC_T_09_billing_range(db, summary):
    """TC-T-09 | billed_amount is present and plausible | ₹5,788.15 ≤ amount ≤ ₹60,56,691.77, no null or zero"""
    lo, hi = db.execute("SELECT MIN(billed_amount), MAX(billed_amount) FROM fact_inpatient_episode").fetchone()
    assert lo == pytest.approx(5788.15, abs=0.01)
    assert hi == pytest.approx(6_056_691.77, abs=0.01)
    assert db.execute("SELECT COUNT(*) FROM fact_inpatient_episode "
                      "WHERE billed_amount IS NULL OR billed_amount <= 0").fetchone()[0] == 0


def test_TC_T_10_death_flag_consistency(db):
    """TC-T-10 | death_flag and discharge_type never disagree | death_flag = 1 ⇔ discharge_type = 'Expired'"""
    mismatches = db.execute(
        "SELECT COUNT(*) FROM fact_inpatient_episode "
        "WHERE (death_flag = 1) <> (discharge_type = 'Expired')").fetchone()[0]
    assert mismatches == 0


def test_TC_T_11_readmission_links(db):
    """TC-T-11 | every readmission resolves to one earlier index within 30 days | no dangling index, no planned index, gap ≤ 30"""
    dangling = db.execute(
        "SELECT COUNT(*) FROM fact_inpatient_episode a WHERE index_episode_id IS NOT NULL "
        "AND NOT EXISTS (SELECT 1 FROM fact_inpatient_episode b WHERE b.episode_id = a.index_episode_id)"
    ).fetchone()[0]
    missing = db.execute(
        "SELECT COUNT(*) FROM fact_inpatient_episode "
        "WHERE readmit_30d_flag = 1 AND index_episode_id IS NULL").fetchone()[0]
    planned = db.execute(
        "SELECT COUNT(*) FROM fact_inpatient_episode "
        "WHERE readmit_30d_flag = 1 AND is_planned_procedure = 1").fetchone()[0]
    stale = db.execute(
        "SELECT COUNT(*) FROM fact_inpatient_episode a JOIN fact_inpatient_episode b "
        "ON a.index_episode_id = b.episode_id "
        "WHERE a.readmit_30d_flag = 1 "
        "AND julianday(a.admission_date) - julianday(b.discharge_date) > 30").fetchone()[0]
    assert (dangling, missing, planned, stale) == (0, 0, 0, 0), (dangling, missing, planned, stale)


def test_TC_T_12_coded_domains(db):
    """TC-T-12 | coded columns use only documented values | 6 age bands, 3 genders, 4 discharge types, 2 visit types, 5 payer categories"""
    def values(table, column):
        return {r[0] for r in db.execute(f"SELECT DISTINCT {column} FROM {table}")}

    assert values("fact_inpatient_episode", "age_band") == {"0-5", "6-17", "18-35", "36-50", "51-65", "66+"}
    assert values("fact_inpatient_episode", "gender") == {"F", "M", "O"}
    assert values("fact_inpatient_episode", "discharge_type") == {"Routine", "Expired", "LAMA", "Transfer"}
    assert values("fact_opd_visit", "visit_type") == {"New", "Follow-up"}
    assert values("dim_payer", "payer_category") == {
        "Cash", "Private insurance", "TPA", "Government scheme", "Corporate"}


def test_TC_T_13_date_coverage(db, summary):
    """TC-T-13 | both facts span the documented window | 2023-01-01 … 2025-12-31, 1,096 days"""
    assert tuple(summary["period"]) == ("2023-01-01", "2025-12-31")
    for table, col in (("fact_inpatient_episode", "admission_date"),
                       ("fact_opd_visit", "visit_date")):
        lo, hi = db.execute(f"SELECT MIN({col}), MAX({col}) FROM {table}").fetchone()
        assert (lo, hi) == ("2023-01-01", "2025-12-31"), (table, lo, hi)


def test_TC_T_14_bed_capacity(db):
    """TC-T-14 | ward capacity is coherent | 300 functional beds across 20 wards, every ward in an active department"""
    assert db.execute("SELECT SUM(functional_beds) FROM dim_ward").fetchone()[0] == 300
    orphans = db.execute(
        "SELECT COUNT(*) FROM dim_ward w LEFT JOIN dim_department d "
        "ON d.department_id = w.department_id WHERE d.department_id IS NULL OR d.is_active = 0"
    ).fetchone()[0]
    assert orphans == 0


def test_TC_T_15_analysis_view_contract(db):
    """TC-T-15 | vw_analysis_episode exposes 36 columns and one row per episode | column count = 36, rows = 53,915"""
    cols = db.execute("PRAGMA table_info(vw_analysis_episode)").fetchall()
    assert len(cols) == 36, [c[1] for c in cols]
    assert db.execute("SELECT COUNT(*) FROM vw_analysis_episode").fetchone()[0] == 53_915


@pytest.mark.slow
def test_TC_T_16_ward_day_bor_bounds(db):
    """TC-T-16 | vw_ward_day_beds never reports impossible occupancy | 0 % ≤ bor_pct ≤ 100 %, 8 columns"""
    assert len(db.execute("PRAGMA table_info(vw_ward_day_beds)").fetchall()) == 8
    lo, hi = db.execute("SELECT MIN(bor_pct), MAX(bor_pct) FROM vw_ward_day_beds").fetchone()
    assert 0 <= lo and hi <= 100, (lo, hi)
    assert hi == 100.0


@pytest.mark.slow
def test_TC_T_17_dept_month_grain(db):
    """TC-T-17 | vw_dept_month_kpi holds one row per department-month | 8 × 36 = 288 rows, no duplicate grain"""
    rows = db.execute("SELECT COUNT(*) FROM vw_dept_month_kpi").fetchone()[0]
    assert rows == 288
    dupes = db.execute(
        "SELECT COUNT(*) FROM (SELECT department_id, month, COUNT(*) c "
        "FROM vw_dept_month_kpi GROUP BY 1, 2 HAVING c > 1)").fetchone()[0]
    assert dupes == 0
    assert db.execute(
        "SELECT COUNT(DISTINCT month) FROM vw_dept_month_kpi").fetchone()[0] == 36


def test_TC_T_18_quality_log_reconciles(root, summary):
    """TC-T-18 | the reject log accounts for every removed row | 5,527 entries = 5,368 rejected + 159 flagged outliers"""
    path = root / "data" / "quality_log.csv"
    with path.open(encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    assert len(rows) == 5_527
    assert sum(summary["rejected"].values()) == 5_368
    assert summary["outlier_flags"] == 159
    assert len(rows) == sum(summary["rejected"].values()) + summary["outlier_flags"]
    assert {r["action"] for r in rows} <= {"rejected", "removed", "flagged", "repaired"}


def test_TC_T_19_dimensions_are_used(db):
    """TC-T-19 | no orphan dimension rows | all 8 departments and all 5 payer categories appear in the facts"""
    depts = db.execute("SELECT COUNT(DISTINCT department_id) FROM fact_inpatient_episode").fetchone()[0]
    payers = db.execute(
        "SELECT COUNT(DISTINCT p.payer_category) FROM fact_inpatient_episode f "
        "JOIN dim_payer p ON p.payer_id = f.payer_id").fetchone()[0]
    assert depts == 8 and payers == 5


def test_TC_T_20_cleaning_removed_rows(root, summary):
    """TC-T-20 | cleaning actually reduced the raw layer | cleaned < raw for both facts, seed = 42"""
    assert summary["seed"] == 42
    for table in ("fact_inpatient_episode", "fact_opd_visit"):
        assert summary["tables"][table] < summary["raw_counts"][table], table
    assert summary["total_rejected"] == sum(summary["rejected"].values())
    assert (root / "data" / "schema.sql").exists()
