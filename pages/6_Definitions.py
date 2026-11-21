"""p107 Definitions, KPI Registry & Governance — Screen UI-07 (BUILD-13).

Official metric definitions, clinical numerator/denominator formulas,
governance standards, benchmarks, and data stewardship catalog (FR-16, BR-05, US-16).
"""
from __future__ import annotations

import pandas as pd
import streamlit as st

from src import data, filters, ui

# 1. Page Configuration
ui.setup("Definitions · p107")
ui.enable_mobile()

# 2. Data Loading & Sidebar
episodes_df = data.load_episodes()
filter_state, chips = filters.render_sidebar(episodes_df)

# 3. Header & Filter Chips
ui.header("KPI Definitions & Governance", "Official healthcare metrics registry, clinical formulas & stewardship directory", as_of="31 Dec 2025")
if chips:
    ui.filter_chips(chips)

st.markdown("#### **Master Healthcare Metric Registry (v1.0)**")
st.caption("Single source of truth for clinical performance and financial accounting definitions.")

registry = [
    {
        "Metric ID": "KPI-01",
        "Metric Name": "Bed Occupancy Rate (BOR)",
        "Clinical Definition": "Percentage of operational inpatient beds occupied over a given period.",
        "Formula / Basis": "(Occupied Bed-Days ÷ Available Bed-Days) × 100",
        "Benchmark / Target": "85.0% (NABH / NHS standard)",
        "Steward": "Operations Director",
        "Refresh": "Daily midnight census"
    },
    {
        "Metric ID": "KPI-02",
        "Metric Name": "Average Length of Stay (ALOS)",
        "Clinical Definition": "Average days an admitted inpatient occupies an acute care bed.",
        "Formula / Basis": "Σ Inpatient Stay Days ÷ Total Discharges (Inclusive: discharge - admit + 1)",
        "Benchmark / Target": "4.00 Days",
        "Steward": "Medical Superintendent",
        "Refresh": "Monthly closed chart"
    },
    {
        "Metric ID": "KPI-03",
        "Metric Name": "30-Day Unplanned Readmissions",
        "Clinical Definition": "Unplanned return to acute inpatient care within 30 days of prior live discharge.",
        "Formula / Basis": "(Unplanned 30d Readmits ÷ (Discharges − Deaths)) × 100",
        "Benchmark / Target": "< 5.00%",
        "Steward": "Clinical Quality Head",
        "Refresh": "Monthly closed chart"
    },
    {
        "Metric ID": "KPI-04",
        "Metric Name": "In-Hospital Mortality Rate",
        "Clinical Definition": "Proportion of admitted inpatient episodes that result in clinical death.",
        "Formula / Basis": "(Recorded Inpatient Deaths ÷ Total Discharges) × 100",
        "Benchmark / Target": "< 2.00% (Risk-adjusted)",
        "Steward": "Clinical Quality Head",
        "Refresh": "Weekly review"
    },
    {
        "Metric ID": "KPI-05",
        "Metric Name": "Revenue per Discharged Patient",
        "Clinical Definition": "Mean total billed gross revenue generated per completed inpatient stay.",
        "Formula / Basis": "Total Billed Gross Revenue ÷ Total Discharges",
        "Benchmark / Target": "Specialty-dependent",
        "Steward": "Chief Financial Officer",
        "Refresh": "Monthly audited billing"
    },
    {
        "Metric ID": "KPI-06",
        "Metric Name": "Revenue per Occupied Bed-Day",
        "Clinical Definition": "Average financial yield produced by one occupied bed across 24 hours.",
        "Formula / Basis": "Total Billed Inpatient Revenue ÷ Episode Bed-Days",
        "Benchmark / Target": "Specialty-dependent",
        "Steward": "Chief Financial Officer",
        "Refresh": "Monthly audited billing"
    },
]

st.dataframe(pd.DataFrame(registry), use_container_width=True, hide_index=True)

st.markdown("---")
st.markdown("#### **Governance Rules & Precision Policies**")
col1, col2 = st.columns(2)

with col1:
    st.markdown("##### **Division-by-Zero Resilience**")
    st.write(
        "In accordance with NFR-10, any KPI whose denominator resolves to zero or null "
        "(e.g., a specialty with zero discharges in a given month) must return `None` and "
        "render on screen as an em-dash (`—`). The system strictly forbids raising unhandled exceptions or emitting NaN."
    )

with col2:
    st.markdown("##### **Low Confidence Threshold (n < 30)**")
    st.write(
        "Any cohort or subgroup containing fewer than 30 observations is automatically tagged "
        "with a 'low confidence' badge. Wilson confidence intervals and statistical tests are "
        "annotated with sample size caveats to prevent erroneous operational conclusions."
    )
