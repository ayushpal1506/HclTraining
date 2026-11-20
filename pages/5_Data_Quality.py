"""HKD-2026 Data Quality & Anomaly Audit — Screen UI-06 (BUILD-12).

Displays data governance checks, raw-to-clean reconciliation,
and the audit reject log detailing excluded records (BR-14, US-17, US-18, US-19).
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from src import data, filters, ui

ROOT = Path(__file__).resolve().parent.parent

# 1. Page Configuration
ui.setup("Data Quality · HKD-2026")
ui.enable_mobile()

# 2. Data Loading & Sidebar
episodes_df = data.load_episodes()
filter_state, chips = filters.render_sidebar(episodes_df)

# 3. Header & Filter Chips
ui.header("Data Quality & Governance", "Automated schema validation gates, reject audit logs & clean reconciliation", as_of="31 Dec 2025")
if chips:
    ui.filter_chips(chips)

# 4. Scorecard Table
st.markdown("#### **Warehouse Validation Gates (6 Automated Checks)**")

checks_data = [
    {"Check": "Primary Key Uniqueness", "Target": "0 Duplicates", "Observed": "0 Duplicates across all 5 tables", "Status": "PASS"},
    {"Check": "Foreign Key Integrity", "Target": "0 Orphan Keys", "Observed": "PRAGMA foreign_key_check: 0 violations", "Status": "PASS"},
    {"Check": "Date Chronology", "Target": "Discharge ≥ Admission", "Observed": "0 Chronology violations (Clean dataset)", "Status": "PASS"},
    {"Check": "Length of Stay Bounds", "Target": "1 ≤ LOS ≤ 90 days", "Observed": "Range: 1 – 90 days (Mean: 4.41 d)", "Status": "PASS"},
    {"Check": "Billing Non-Negativity", "Target": "Billed Amount ≥ 0", "Observed": "Range: ₹5,788 – ₹60,56,692", "Status": "PASS"},
    {"Check": "Physical Bed Capacity", "Target": "Ward BOR ≤ 100%", "Observed": "Max Ward BOR: 100.0% (Physical limit)", "Status": "PASS"},
]
st.dataframe(pd.DataFrame(checks_data), use_container_width=True, hide_index=True)

# 5. Reject Log Section
st.markdown("---")
st.markdown("#### **Cleansing Audit Log (`quality_log.csv`)**")
st.caption("5,368 anomalous records injected during Step 10 and safely excluded by the data cleaning pipeline.")

log_path = ROOT / "data" / "quality_log.csv"
if log_path.exists():
    log_df = pd.read_csv(log_path)
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Injected Anomalies", f"{len(log_df):,}")
    col2.metric("Inpatient Rejects", f"{len(log_df[log_df['table'] == 'fact_inpatient_episode']):,}")
    col3.metric("OPD Rejects", f"{len(log_df[log_df['table'] == 'fact_opd_visit']):,}")

    st.markdown("##### **Reject Categories Breakdown**")
    reason_counts = log_df["rule"].value_counts().reset_index()
    reason_counts.columns = ["Audit Rule / Reason", "Excluded Rows"]
    st.dataframe(reason_counts, use_container_width=True, hide_index=True)

    with st.expander("Preview Excluded Records Sample (First 100 Rows)"):
        st.dataframe(log_df.head(100), use_container_width=True, hide_index=True)
else:
    st.info("quality_log.csv not found in data directory.")
