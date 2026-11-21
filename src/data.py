"""p107 Data access, caching and filtering module (BUILD-02).

Read-only interface to data/app.db.
Provides @st.cache_data query functions and multi-dimensional cohort slicing.
Enforces NFR-01 (< 3.0 s latency) and NFR-09 (zero PII exposure).
"""
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Sequence

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent.parent
DB_PATH = ROOT / "data" / "app.db"


def get_connection() -> sqlite3.Connection:
    """Open read-only SQLite connection in URI mode."""
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Database not found at {DB_PATH}")
    return sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)


@st.cache_data(show_spinner=False)
def load_episodes() -> pd.DataFrame:
    """Load the 53,915 validated inpatient episodes from vw_analysis_episode."""
    con = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM vw_analysis_episode", con)
    finally:
        con.close()

    df["admission_date"] = pd.to_datetime(df["admission_date"])
    df["discharge_date"] = pd.to_datetime(df["discharge_date"])
    df["admission_year"] = pd.to_numeric(df["admission_year"], errors="coerce").fillna(0).astype(int)
    assert len(df) == 53_915, f"Expected 53,915 episodes, got {len(df)}"
    return df


@st.cache_data(show_spinner=False)
def load_ward_day_beds() -> pd.DataFrame:
    """Load daily ward census and capacity from vw_ward_day_beds."""
    con = get_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM vw_ward_day_beds", con)
    finally:
        con.close()

    df["day"] = pd.to_datetime(df["day"])
    df["year"] = df["day"].dt.year
    df["month"] = df["day"].dt.month
    return df


@st.cache_data(show_spinner=False)
def load_dept_month_kpi() -> pd.DataFrame:
    """Load department-month aggregated KPIs from vw_dept_month_kpi."""
    con = get_connection()
    try:
        df = pd.read_sql_query(
            "SELECT k.*, d.department_name "
            "FROM vw_dept_month_kpi k "
            "JOIN dim_department d ON k.department_id = d.department_id",
            con
        )
    finally:
        con.close()
    return df


@st.cache_data(show_spinner=False)
def load_opd_visits() -> pd.DataFrame:
    """Load the 143,139 OPD visits from fact_opd_visit."""
    con = get_connection()
    try:
        df = pd.read_sql_query(
            "SELECT v.*, d.department_name, p.payer_name, p.payer_category "
            "FROM fact_opd_visit v "
            "JOIN dim_department d ON v.department_id = d.department_id "
            "JOIN dim_payer p ON v.payer_id = p.payer_id",
            con
        )
    finally:
        con.close()

    df["visit_date"] = pd.to_datetime(df["visit_date"])
    df["visit_year"] = df["visit_date"].dt.year
    df["visit_month"] = df["visit_date"].dt.month
    df["weekday_name"] = df["visit_date"].dt.day_name()
    assert len(df) == 143_139, f"Expected 143,139 OPD visits, got {len(df)}"
    return df


@st.cache_data(show_spinner=False)
def load_dimensions() -> dict[str, pd.DataFrame]:
    """Load reference dimensions."""
    con = get_connection()
    try:
        depts = pd.read_sql_query("SELECT * FROM dim_department ORDER BY department_name", con)
        wards = pd.read_sql_query("SELECT * FROM dim_ward ORDER BY ward_id", con)
        payers = pd.read_sql_query("SELECT * FROM dim_payer ORDER BY payer_id", con)
    finally:
        con.close()
    return {"department": depts, "ward": wards, "payer": payers}


def filter_episodes(
    df: pd.DataFrame,
    period: int | str | None = 2025,
    departments: Sequence[str] | None = None,
    wards: Sequence[str] | None = None,
    payers: Sequence[str] | None = None,
    gender: str | None = None,
    age_bands: Sequence[str] | None = None,
) -> pd.DataFrame:
    """Slice inpatient episodes by multi-dimensional criteria."""
    filtered = df

    if period is not None and str(period).lower() != "all":
        try:
            year = int(period)
            filtered = filtered[filtered["admission_year"] == year]
        except (ValueError, TypeError):
            pass

    if departments:
        filtered = filtered[filtered["department_name"].isin(departments)]

    if wards:
        filtered = filtered[filtered["ward_name"].isin(wards)]

    if payers:
        filtered = filtered[filtered["payer_category"].isin(payers)]

    if gender and gender.upper() != "ALL":
        filtered = filtered[filtered["gender"] == gender.upper()]

    if age_bands:
        filtered = filtered[filtered["age_band"].isin(age_bands)]

    return filtered
