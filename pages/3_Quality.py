"""p107 Clinical Quality, Mortality & Drift Detection — Screen UI-04 (BUILD-10).

Monitors in-hospital mortality, statistical process control drift detection,
and discharge disposition case-mix breakdowns (OQ-10, BR-13, US-13).
"""
from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from src import data, filters, kpi, stats, ui

# 1. Page Configuration
ui.setup("Quality · p107")
ui.enable_mobile()

# 2. Data Loading & Sidebar
episodes_df = data.load_episodes()
dept_month_df = data.load_dept_month_kpi()
filter_state, chips = filters.render_sidebar(episodes_df)

# 3. Header & Filter Chips
ui.header("Clinical Quality & Safety", "In-hospital mortality monitoring, statistical drift flags & discharge outcomes", as_of="31 Dec 2025")
if chips:
    ui.filter_chips(chips)

# 4. Filter Cohort
cohort = data.filter_episodes(
    episodes_df,
    period=filter_state.period_int if filter_state.period != "All" else None,
    departments=filter_state.departments,
    wards=filter_state.wards,
    payers=filter_state.payers,
    gender=filter_state.gender,
    age_bands=filter_state.age_bands,
)

if len(cohort) == 0:
    ui.empty_state("No patient records match the selected quality filters.", chips)
    st.stop()

# 5. Top Quality KPIs
discharges = len(cohort)
deaths = cohort["death_flag"].sum()
mort_rate = kpi.mortality_pct(deaths, discharges)
lama_count = (cohort["discharge_type"] == "LAMA").sum()
transfer_count = (cohort["discharge_type"] == "Transfer").sum()

ui.kpi_row([
    ui.card("In-Hospital Mortality", f"{mort_rate:.2f}%" if mort_rate else "—", "National NABH norm: < 2.0%", "neutral", "ok" if (mort_rate and mort_rate < 2.0) else "bad", n=discharges),
    ui.card("Total Recorded Deaths", f"{deaths:,}", f"From {discharges:,} discharges", "neutral", "neutral"),
    ui.card("LAMA (Against Advice)", f"{lama_count:,}", f"{(lama_count/discharges*100):.2f}% of cohort", "neutral", "neutral"),
    ui.card("Transfers Out", f"{transfer_count:,}", f"{(transfer_count/discharges*100):.2f}% of cohort", "neutral", "neutral"),
])

col1, col2 = st.columns(2)

with col1:
    st.markdown("#### **Statistical Drift Monitor (Monthly Mortality)**")
    st.caption("P-05: Monthly mortality with baseline mean ± 2.0 SD control limits (BR-13, US-13).")

    # Aggregate mortality by month across the selected department(s)
    monthly_mort = episodes_df.groupby("admission_month").agg(
        discharges=("episode_id", "count"),
        deaths=("death_flag", "sum")
    ).reset_index()
    monthly_mort["mortality_pct"] = (monthly_mort["deaths"] / monthly_mort["discharges"] * 100).round(2)
    monthly_mort = monthly_mort.sort_values("admission_month")

    # Run drift flag calculation using 2023 as 12-month baseline
    mort_series = monthly_mort["mortality_pct"].tolist()
    drift_res = stats.drift_flag(mort_series, baseline_n=12, z=2.0)

    fig_drift = go.Figure()

    # Monthly line
    fig_drift.add_trace(go.Scatter(
        x=monthly_mort["admission_month"],
        y=monthly_mort["mortality_pct"],
        mode="lines+markers",
        line=dict(color="#1F3864", width=2),
        name="Monthly Mortality %"
    ))

    # Baseline & bands
    if drift_res["mean"] is not None:
        mean_val = drift_res["mean"]
        lo_band, hi_band = drift_res["band"]
        fig_drift.add_hline(y=mean_val, line_dash="solid", line_color="#0F766E", annotation_text="Baseline Mean (1.3%)")
        fig_drift.add_hline(y=hi_band, line_dash="dash", line_color="#B91C1C", annotation_text="+2σ Threshold")

    # Mark flagged points
    flagged_idx = drift_res["flags"]
    if flagged_idx:
        flagged_df = monthly_mort.iloc[flagged_idx]
        fig_drift.add_trace(go.Scatter(
            x=flagged_df["admission_month"],
            y=flagged_df["mortality_pct"],
            mode="markers",
            marker=dict(color="#B91C1C", size=10, symbol="x"),
            name="Drift Alert (> +2σ)"
        ))

    fig_drift.update_layout(xaxis_title="Month (YYYY-MM)", yaxis_title="Mortality Rate (%)", height=380, margin=dict(l=10, r=10, t=10, b=10))
    ui.chart_card("Mortality Drift Control Chart", fig_drift)

with col2:
    st.markdown("#### **Discharge Disposition Breakdown**")
    st.caption("Distribution of clinical outcomes across discharges.")

    disp_counts = cohort["discharge_type"].value_counts().reset_index()
    disp_counts.columns = ["discharge_type", "count"]

    fig_disp = px.pie(
        disp_counts,
        names="discharge_type",
        values="count",
        color="discharge_type",
        color_discrete_map={"Normal": "#0F766E", "Transfer": "#1F3864", "LAMA": "#B45309", "Expired": "#B91C1C"},
        hole=0.4
    )
    fig_disp.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10))
    ui.chart_card("Discharge Disposition Mix", fig_disp)

ui.insight_box(
    "Mortality drift detection is a statistical review trigger to prompt clinical audit, not evidence of negligence.",
    level="warn",
    action="Review Clinical Protocols",
    source="TRD §5.5 · Quality Governance"
)
