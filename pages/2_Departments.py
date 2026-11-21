"""p107 Clinical Departments, ALOS & Readmission — Screen UI-03 (BUILD-09).

Department-level benchmarking: ALOS dumbbell chart against 4.0-day standard,
30-day readmission rates with Wilson score 95% confidence intervals (OQ-03, OQ-04, OQ-08).
"""
from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from src import data, filters, kpi, stats, ui

# 1. Page Configuration
ui.setup("Departments · p107")
ui.enable_mobile()

# 2. Data Loading & Sidebar
episodes_df = data.load_episodes()
filter_state, chips = filters.render_sidebar(episodes_df)

# 3. Header & Filter Chips
ui.header("Clinical Departments", "ALOS benchmarking, 30-day readmissions & case-mix profiles", as_of="31 Dec 2025")
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
    ui.empty_state("No patient records match the selected departmental filters.", chips)
    st.stop()

# 5. Department Summary Metrics
dept_agg = cohort.groupby("department_name").agg(
    discharges=("episode_id", "count"),
    bed_days=("bed_days", "sum"),
    deaths=("death_flag", "sum"),
    readmissions=("readmit_30d_flag", "sum"),
    billed_amount=("billed_amount", "sum"),
).reset_index()

dept_agg["alos"] = (dept_agg["bed_days"] / dept_agg["discharges"]).round(2)
dept_agg["eligible_readmit"] = (dept_agg["discharges"] - dept_agg["deaths"]).clip(lower=0)
dept_agg["readmission_rate"] = (dept_agg["readmissions"] / dept_agg["eligible_readmit"] * 100).round(2)

# Compute Wilson 95% intervals
wilson_lo, wilson_hi = [], []
for _, row in dept_agg.iterrows():
    k, n = int(row["readmissions"]), int(row["eligible_readmit"])
    res = stats.wilson_interval(k, n)
    if res:
        wilson_lo.append(res[0])
        wilson_hi.append(res[1])
    else:
        wilson_lo.append(0.0)
        wilson_hi.append(0.0)
dept_agg["wilson_lo"] = wilson_lo
dept_agg["wilson_hi"] = wilson_hi
dept_agg["error_plus"] = (dept_agg["wilson_hi"] - dept_agg["readmission_rate"]).clip(lower=0)
dept_agg["error_minus"] = (dept_agg["readmission_rate"] - dept_agg["wilson_lo"]).clip(lower=0)
dept_agg["is_low_n"] = dept_agg["discharges"] < 30

col1, col2 = st.columns(2)

with col1:
    st.markdown("#### **ALOS vs 4.0-Day Benchmark (Dumbbell Chart)**")
    st.caption("Benchmark: 4.0 days. Oncology has highest ALOS (6.49 d); Maternity has lowest (2.76 d).")

    dept_sorted = dept_agg.sort_values("alos", ascending=True)
    fig_alo = go.Figure()

    # Connector lines
    for _, row in dept_sorted.iterrows():
        fig_alo.add_trace(go.Scatter(
            x=[4.0, row["alos"]],
            y=[row["department_name"], row["department_name"]],
            mode="lines",
            line=dict(color="#CBD5E1", width=3),
            showlegend=False,
            hoverinfo="none"
        ))

    # Benchmark point
    fig_alo.add_trace(go.Scatter(
        x=[4.0] * len(dept_sorted),
        y=dept_sorted["department_name"],
        mode="markers",
        marker=dict(color="#64748B", size=8),
        name="Benchmark (4.0 d)"
    ))

    # Actual ALOS point
    fig_alo.add_trace(go.Scatter(
        x=dept_sorted["alos"],
        y=dept_sorted["department_name"],
        mode="markers+text",
        marker=dict(color=["#B91C1C" if a > 5.0 else "#0F766E" for a in dept_sorted["alos"]], size=12),
        text=[f"{a:.2f} d" for a in dept_sorted["alos"]],
        textposition="top right",
        name="Actual ALOS"
    ))

    fig_alo.add_vline(x=4.0, line_dash="dash", line_color="#94A3B8")
    fig_alo.update_layout(xaxis_title="Length of Stay (days)", yaxis_title="", height=380, margin=dict(l=10, r=10, t=10, b=10))
    ui.chart_card("ALOS Dumbbell Benchmark", fig_alo)

with col2:
    st.markdown("#### **30-Day Unplanned Readmissions (Wilson 95% CI)**")
    st.caption("General Medicine, Oncology, and Cardiology show highest clinical readmission pressure.")

    dept_readm = dept_agg.sort_values("readmission_rate", ascending=False)
    fig_readm = go.Figure()

    fig_readm.add_trace(go.Bar(
        x=dept_readm["department_name"],
        y=dept_readm["readmission_rate"],
        error_y=dict(
            type="data",
            symmetric=False,
            array=dept_readm["error_plus"],
            arrayminus=dept_readm["error_minus"],
            color="#1F3864"
        ),
        marker_color=["#B91C1C" if r > 6.0 else "#1F3864" for r in dept_readm["readmission_rate"]],
        text=[f"{r:.2f}%<br>(n={d})" for r, d in zip(dept_readm["readmission_rate"], dept_readm["discharges"])],
        textposition="outside"
    ))
    fig_readm.update_layout(xaxis_title="", yaxis_title="30-Day Readmission Rate (%)", height=380, margin=dict(l=10, r=10, t=10, b=10))
    ui.chart_card("30-Day Readmissions with Confidence Intervals", fig_readm)

# 6. Demographics breakdown
st.markdown("#### **Patient Demographics & Age-Band Distribution**")
age_order = ["0-5", "6-17", "18-35", "36-50", "51-65", "65+"]
age_counts = cohort.groupby(["age_band", "gender"]).size().reset_index(name="count")
fig_demo = px.bar(
    age_counts,
    x="age_band",
    y="count",
    color="gender",
    category_orders={"age_band": age_order},
    barmode="group",
    color_discrete_map={"M": "#1F3864", "F": "#0F766E", "O": "#B45309"},
    labels={"count": "Inpatient Discharges", "age_band": "Age Category", "gender": "Gender"}
)
fig_demo.update_layout(height=340, margin=dict(l=10, r=10, t=10, b=10))
ui.chart_card("Demographics: Discharges by Age Band and Gender", fig_demo)
