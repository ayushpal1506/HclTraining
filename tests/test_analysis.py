"""Analysis tests — TC-T-21 … TC-T-40.

Two kinds of check, both runnable before pandas, NumPy, SciPy, Plotly or
Streamlit are installed:

* TC-T-21 … TC-T-30 pin the KPI mathematics to hand arithmetic on a five-row
  warehouse built in memory. If a formula changes, this file fails, not a chart.
* TC-T-31 … TC-T-40 pin the statistics — Wilson interval, rank correlation,
  t and chi-square p-values, drift flag, least squares — to published values.

    .venv/bin/python -m pytest tests/test_analysis.py -q
"""
from __future__ import annotations

import sqlite3

import pytest

from src import kpi, stats

TINY_DDL = """
CREATE TABLE fact_inpatient_episode (
    episode_id      TEXT PRIMARY KEY,
    department_id   TEXT NOT NULL,
    admission_date  TEXT NOT NULL,
    discharge_date  TEXT NOT NULL,
    los_days        INTEGER NOT NULL,
    bed_days        INTEGER NOT NULL,
    billed_amount   REAL NOT NULL,
    death_flag      INTEGER NOT NULL,
    readmit_30d_flag INTEGER NOT NULL,
    index_episode_id TEXT,
    is_planned_procedure INTEGER NOT NULL
);
"""

TINY_ROWS = [
    # id,  dept,   admit,       discharge,   los, bed, billed, death, readmit, index, planned
    ("E1", "DEP01", "2025-01-01", "2025-01-05", 5, 5, 50_000.0, 0, 0, None, 0),
    ("E2", "DEP01", "2025-01-10", "2025-01-14", 5, 5, 30_000.0, 1, 0, None, 0),
    ("E3", "DEP01", "2025-01-24", "2025-01-26", 3, 3, 20_000.0, 0, 1, "E1", 0),
    ("E4", "DEP02", "2025-02-01", "2025-02-09", 9, 9, 90_000.0, 0, 0, None, 1),
    ("E5", "DEP02", "2025-02-10", "2025-02-10", 1, 1, 10_000.0, 0, 0, None, 0),
]


@pytest.fixture(scope="module")
def tiny() -> sqlite3.Connection:
    con = sqlite3.connect(":memory:")
    con.row_factory = sqlite3.Row
    con.executescript(TINY_DDL)
    con.executemany(
        "INSERT INTO fact_inpatient_episode VALUES (?,?,?,?,?,?,?,?,?,?,?)", TINY_ROWS)
    con.commit()
    yield con
    con.close()


def tiny_aggregate(con, dept=None):
    """The same aggregate the warehouse view publishes, for one department."""
    where, args = ("WHERE department_id = ?", [dept]) if dept else ("", [])
    return con.execute(
        f"""SELECT COUNT(*) AS discharges, SUM(los_days) AS inpatient_days,
                   SUM(bed_days) AS bed_days, SUM(billed_amount) AS revenue,
                   SUM(death_flag) AS deaths, SUM(readmit_30d_flag) AS readmissions
            FROM fact_inpatient_episode {where}""", args).fetchone()


# --------------------------------------------------------------- KPI formulas
def test_TC_T_21_bor_and_gap():
    """TC-T-21 | bor_pct / occupancy_gap_pp | 10 of 20 bed-days = 50.0 %, 35.0 pp below the 85 % target"""
    assert kpi.bor_pct(10, 20) == 50.0
    assert kpi.occupancy_gap_pp(50.0) == -35.0
    assert kpi.occupancy_gap_pp(86.12) == 1.12
    assert kpi.bor_pct(0, 0) is None


def test_TC_T_22_beds_to_target():
    """TC-T-22 | beds_to_target | 50 % on 310 bed-days over 31 days needs 3.5 more occupied beds at 85 %"""
    assert kpi.beds_to_target(50.0, 310, 31) == pytest.approx(3.5, abs=0.01)
    assert kpi.beds_to_target(85.0, 310, 31) == pytest.approx(0.0, abs=0.01)
    assert kpi.beds_to_target(50.0, 310, 0) is None


def test_TC_T_23_alos():
    """TC-T-23 | alos_days | 23 inpatient days over 5 discharges = 4.60 days"""
    assert kpi.alos_days(23, 5) == 4.6
    assert kpi.alos_days(0, 5) == 0.0
    assert kpi.alos_days(23, 0) is None


def test_TC_T_24_mortality():
    """TC-T-24 | mortality_pct | 1 death in 5 discharges = 20.0 %"""
    assert kpi.mortality_pct(1, 5) == 20.0
    assert kpi.mortality_pct(0, 5) == 0.0
    assert kpi.mortality_pct(2, 0) is None


def test_TC_T_25_readmission_uses_live_discharges():
    """TC-T-25 | eligibility / readmission_rate_pct | 1 readmission, 5 discharges, 1 death → 1/4 = 25.0 %"""
    assert kpi.eligibility(5, 1) == 4
    assert kpi.eligibility(2, 5) == 0
    assert kpi.readmission_rate_pct(1, kpi.eligibility(5, 1)) == 25.0
    assert kpi.readmission_rate_pct(1, 0) is None


def test_TC_T_26_revenue_per_patient_and_bed_day():
    """TC-T-26 | revenue_per_patient / revenue_per_bed_day | ₹2,00,000 over 5 patients and 23 bed-days"""
    assert kpi.revenue_per_patient(200_000, 5) == 40_000.0
    assert kpi.revenue_per_bed_day(200_000, 23) == pytest.approx(8_695.65, abs=0.01)


def test_TC_T_27_share_pct():
    """TC-T-27 | share_pct | 7 of 28 episodes = 25.0 %; zero whole → None, never nan"""
    assert kpi.share_pct(7, 28) == 25.0
    assert kpi.share_pct(0, 28) == 0.0
    assert kpi.share_pct(7, 0) is None


def test_TC_T_28_readmission_window_boundary():
    """TC-T-28 | readmitted_within | day 0 and day 30 count, day 31 and earlier-than-discharge do not"""
    assert kpi.readmitted_within("2025-01-01", "2025-01-01") is True
    assert kpi.readmitted_within("2025-01-01", "2025-01-31") is True
    assert kpi.readmitted_within("2025-01-01", "2025-02-01") is False
    assert kpi.readmitted_within("2025-01-10", "2025-01-09") is False


def test_TC_T_29_low_confidence_rule():
    """TC-T-29 | low_confidence | 29 is flagged, 30 is not — the published n < 30 rule (US-18)"""
    assert kpi.low_confidence(29) is True
    assert kpi.low_confidence(30) is False
    assert kpi.low_confidence(0) is True
    assert kpi.low_confidence(None) is False


def test_TC_T_30_kpi_kernel_matches_warehouse_aggregate(tiny):
    """TC-T-30 | formulas applied to the SQL aggregate equal hand arithmetic and the view's own arithmetic"""
    total = tiny_aggregate(tiny)
    assert total["discharges"] == 5
    assert kpi.alos_days(total["inpatient_days"], total["discharges"]) == 4.6
    assert kpi.mortality_pct(total["deaths"], total["discharges"]) == 20.0
    assert kpi.revenue_per_patient(total["revenue"], total["discharges"]) == 40_000.0
    assert kpi.readmission_rate_pct(
        total["readmissions"], kpi.eligibility(total["discharges"], total["deaths"])) == 25.0

    dep01 = tiny_aggregate(tiny, "DEP01")
    assert dep01["discharges"] == 3
    assert kpi.alos_days(dep01["inpatient_days"], dep01["discharges"]) == pytest.approx(4.33, abs=0.01)
    assert kpi.mortality_pct(dep01["deaths"], dep01["discharges"]) == pytest.approx(33.33, abs=0.01)
    assert kpi.readmission_rate_pct(
        dep01["readmissions"], kpi.eligibility(dep01["discharges"], dep01["deaths"])) == 50.0


def test_TC_T_30b_decomposition_conserves_the_total():
    """TC-T-30b | decompose | contributors sum to the total movement; the residual is printed, not absorbed"""
    out = kpi.decompose({"Cardiology": 100.0, "Oncology": 50.0, "ENT": 10.0},
                        {"Cardiology": 114.0, "Oncology": 54.0, "ENT": 9.0})
    assert out["total_a"] == 160.0 and out["total_b"] == 177.0
    assert out["total_delta"] == 17.0
    assert out["residual"] == 0.0
    assert [r[0] for r in out["rows"]] == ["Cardiology", "Oncology", "ENT"]
    assert out["rows"][0][1] == 14.0

    opened = kpi.decompose({"A": 10.0}, {"A": 10.0, "B": 5.0})
    assert opened["total_delta"] == 5.0
    assert dict(opened["rows"])["B"] == 5.0


# ----------------------------------------------------------------- statistics
def test_TC_T_31_wilson_known_values():
    """TC-T-31 | wilson_interval | 50/100 → 40.38–59.62 %; 0/10 → 0–27.75 % (upper bound z²/(n+z²))"""
    lo, hi = stats.wilson_interval(50, 100)
    assert (lo, hi) == pytest.approx((40.38, 59.62), abs=0.02)
    lo, hi = stats.wilson_interval(0, 10)
    assert lo == 0.0
    assert hi == pytest.approx(27.75, abs=0.05)
    lo, hi = stats.wilson_interval(10, 10)
    assert hi == 100.0 and lo == pytest.approx(72.25, abs=0.05)
    assert stats.wilson_interval(0, 0) is None


def test_TC_T_32_wilson_stays_inside_zero_to_hundred():
    """TC-T-32 | wilson_interval | every k of every n from 0 to 60 stays inside 0–100 and contains p̂"""
    for n in range(1, 61):
        for k in range(n + 1):
            lo, hi = stats.wilson_interval(k, n)
            p_hat = 100.0 * k / n
            assert 0.0 <= lo <= p_hat + 0.01, (k, n, lo, p_hat)
            assert p_hat - 0.01 <= hi <= 100.0, (k, n, p_hat, hi)


def test_TC_T_33_pearson():
    """TC-T-33 | pearson | a perfect line gives 1.0, a reversed line −1.0, a constant None"""
    assert stats.pearson([1, 2, 3], [2, 4, 6]) == pytest.approx(1.0)
    assert stats.pearson([1, 2, 3], [6, 4, 2]) == pytest.approx(-1.0)
    assert stats.pearson([1, 2, 3], [5, 5, 5]) is None
    assert stats.pearson([1, 2], [1, 2]) is None


def test_TC_T_34_spearman_handles_ties_and_monotone_curves():
    """TC-T-34 | spearman | monotone but non-linear still gives 1.0; ties get average ranks"""
    assert stats.spearman([1, 2, 3, 4], [1, 10, 100, 1000]) == pytest.approx(1.0)
    assert stats.spearman([1, 2, 3, 4], [1, 3, 2, 4]) == pytest.approx(0.8)
    assert stats.spearman([1, 1, 1, 2], [1, 2, 3, 4]) is not None


def test_TC_T_35_rank_averages_ties():
    """TC-T-35 | _ranks | values 10, 20, 20, 30 rank 1, 2.5, 2.5, 4"""
    assert stats._ranks([10, 20, 20, 30]) == [1.0, 2.5, 2.5, 4.0]


def test_TC_T_36_student_t_published_values():
    """TC-T-36 | student_t_p | t = 1.96 at df 1e6 → 0.050; t = 2.228 at df 10 → 0.050; t = 0 → 1.0"""
    assert stats.student_t_p(1.959964, 1_000_000) == pytest.approx(0.05, abs=0.0005)
    assert stats.student_t_p(2.228139, 10) == pytest.approx(0.05, abs=0.001)
    assert stats.student_t_p(0.0, 10) == 1.0
    assert stats.student_t_p(5.0, 3) < 0.02


def test_TC_T_37_chi_square_2x2():
    """TC-T-37 | chi_square_2x2 | 10/20 vs 30/40 → χ² 0.794, p ≈ 0.373; a degenerate table → None"""
    chi2, p = stats.chi_square_2x2(10, 20, 30, 40)
    assert chi2 == pytest.approx(0.7937, abs=0.001)
    assert p == pytest.approx(0.373, abs=0.005)
    assert stats.chi_square_2x2(0, 0, 0, 0) is None
    assert stats.chi_square_2x2(0, 0, 5, 5) is None


def test_TC_T_38_correlation_test_contract():
    """TC-T-38 | correlation_test | prints rho, n, t and p; a perfect monotone series has p = 0"""
    x = [1, 2, 3, 4, 5, 6, 7, 8]
    y = [2, 1, 4, 3, 6, 5, 8, 7]
    out = stats.correlation_test(x, y)
    assert out["method"] == "spearman"
    assert out["n"] == 8
    assert out["rho"] == pytest.approx(0.9048, abs=0.001)
    assert 0.0 < out["p"] < 1.0
    assert stats.correlation_test([1, 2, 3], [3, 2, 1])["p"] == 0.0
    assert stats.correlation_test([1], [2])["p"] is None


def test_TC_T_39_drift_flag():
    """TC-T-39 | drift_flag | a 4.0 month after a 3.0 ± 0.1 baseline is flagged; a 3.05 month is not"""
    series = [3.0, 3.1, 2.9, 3.0, 3.05, 2.95, 3.0, 3.1, 2.9, 3.0, 3.05, 2.95, 3.05, 4.0]
    out = stats.drift_flag(series, baseline_n=12, z=2.0)
    assert out["first"] == 13
    assert out["flags"] == [13]
    assert out["mean"] == pytest.approx(3.0, abs=0.01)
    short = stats.drift_flag([3.0, 3.1], baseline_n=12)
    assert short["flags"] == []


def test_TC_T_40_ols_line():
    """TC-T-40 | ols_line | y = 2x + 1 recovers slope 2 and intercept 1"""
    slope, intercept = stats.ols_line([0, 1, 2, 3], [1, 3, 5, 7])
    assert slope == pytest.approx(2.0)
    assert intercept == pytest.approx(1.0)
    assert stats.ols_line([5, 5, 5], [1, 2, 3]) is None
