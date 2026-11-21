"""p107 Operations & Capacity Bottlenecks — Screen UI-02 (BUILD-08).

Analyzes bed occupancy against the 85.0% clinical target, 7-day ward projections,
and outpatient volume seasonality patterns (OQ-01, OQ-02, OQ-06).
"""
from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from src import data, filters, kpi, ui

# 1. Page Configuration
ui.setup("Operations · p107")
ui.enable_mobile()

# 2. Data Loading & Sidebar
episodes_df = data.load_episodes()
ward_df = data.load_ward_day_beds()
dept_month_df = data.load_dept_month_kpi()
opd_df = data.load_opd_visits()
filter_state, chips = filters.render_sidebar(episodes_df)

# 3. Header & Filter Chips
ui.header("Operations & Capacity", "Bed occupancy bottlenecks, forward ward demand & OPD seasonality", as_of="31 Dec 2025")
if chips:
    ui.filter_chips(chips)

# 4. Filter Ward Data
if filter_state.period_int:
    ward_filtered = ward_df[ward_df["year"] == filter_state.period_int]
else:
    ward_filtered = ward_df

if filter_state.departments:
    dims = data.load_dimensions()
    dept_ids = dims["department"][dims["department"]["department_name"].isin(filter_state.departments)]["department_id"]
    ward_filtered = ward_filtered[ward_filtered["department_id"].isin(dept_ids)]

if len(ward_filtered) == 0:
    ui.empty_state("No operational census records match the selected filters.", chips)
    st.stop()

# 5. Top KPI Row for Operations
occ_beds = ward_filtered["beds_occupied"].sum()
avail_beds = ward_filtered["functional_beds"].sum()
overall_bor = kpi.bor_pct(occ_beds, avail_beds)
bor_gap = kpi.occupancy_gap_pp(overall_bor, 85.0)

ui.kpi_row([
    ui.card("Overall Bed Occupancy", f"{overall_bor:.1f}%" if overall_bor else "—", "Target: 85.0%", "neutral", "warn" if (overall_bor and overall_bor > 85) else "ok", n=occ_beds),
    ui.card("Occupancy Gap to Target", f"{'+' if (bor_gap and bor_gap >= 0) else ''}{bor_gap:.2f} pp" if bor_gap is not None else "—", "85% clinical threshold", "neutral", "neutral"),
    ui.card("Total Occupied Bed-Days", f"{occ_beds:,}", f"Available: {avail_beds:,}", "neutral", "ok", n=occ_beds),
])

tab1, tab2, tab3 = st.tabs(["📊 Department Occupancy vs 85%", "📅 7-Day Ward Projections", "📈 OPD Volume & Seasonality"])

with tab1:
    st.markdown("#### **Departmental Bed Occupancy Rate (BOR) vs 85.0% Benchmark**")
    st.caption("Cardiology operates above 85% safe capacity limit (P-02). General Medicine and Orthopaedics operate with headroom.")

    # Calculate BOR per department
    dims = data.load_dimensions()
    dept_map = dict(zip(dims["department"]["department_id"], dims["department"]["department_name"]))
    dept_summary = ward_filtered.groupby("department_id").agg({
        "beds_occupied": "sum",
        "functional_beds": "sum"
    }).reset_index()
    dept_summary["department_name"] = dept_summary["department_id"].map(dept_map)
    dept_summary["bor_pct"] = (dept_summary["beds_occupied"] / dept_summary["functional_beds"] * 100).round(2)
    dept_summary["gap_pp"] = dept_summary["bor_pct"] - 85.0
    dept_summary = dept_summary.sort_values("bor_pct", ascending=True)

    fig_bor = go.Figure()
    colors = ["#B91C1C" if bor > 85.0 else "#1F3864" for bor in dept_summary["bor_pct"]]

    fig_bor.add_trace(go.Bar(
        y=dept_summary["department_name"],
        x=dept_summary["bor_pct"],
        orientation="h",
        marker_color=colors,
        text=[f"{b:.1f}% ({'+' if g >= 0 else ''}{g:.1f} pp)" for b, g in zip(dept_summary["bor_pct"], dept_summary["gap_pp"])],
        textposition="auto"
    ))
    fig_bor.add_vline(x=85.0, line_dash="dash", line_color="#B91C1C", annotation_text="85.0% Safe Capacity Target")
    fig_bor.update_layout(
        xaxis_title="Bed Occupancy Rate (%)",
        yaxis_title="",
        xaxis_range=[0, 100],
        margin=dict(l=20, r=20, t=20, b=20),
        height=380,
    )
    ui.chart_card("Department Bed Occupancy", fig_bor, note="Source: vw_ward_day_beds aggregated by clinical department.")

with tab2:
    st.markdown("#### **Ward Forward Demand & Occupancy Projection**")
    st.caption("Rolling 7-day average occupancy per ward showing capacity stress points (US-05).")

    # Latest 7-day slice
    max_day = ward_filtered["day"].max()
    min_7d = max_day - pd.Timedelta(days=7)
    recent_ward = ward_filtered[ward_filtered["day"] >= min_7d]
    ward_agg = recent_ward.groupby(["ward_name", "ward_type"]).agg({
        "beds_occupied": "mean",
        "functional_beds": "mean"
    }).reset_index()
    ward_agg["bor_pct"] = (ward_agg["beds_occupied"] / ward_agg["functional_beds"] * 100).round(1)
    ward_agg = ward_agg.sort_values("bor_pct", ascending=False)

    fig_ward = px.bar(
        ward_agg,
        x="ward_name",
        y="bor_pct",
        color="ward_type",
        color_discrete_sequence=["#1F3864", "#0F766E", "#B45309", "#475569"],
        labels={"bor_pct": "Projected BOR (%)", "ward_name": "Ward", "ward_type": "Ward Class"},
    )
    fig_ward.add_hline(y=85.0, line_dash="dash", line_color="#B91C1C", annotation_text="85% Target")
    fig_ward.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=380)
    ui.chart_card("7-Day Forward Ward Utilization", fig_ward)

with tab3:
    st.markdown("#### **OPD Volume Demand & Seasonality Patterns**")
    st.caption("P-01 Winter respiratory surge (Nov–Feb +28%) and P-05 Monday clinic rush (ratio 1.52).")

    # Filter OPD by year
    if filter_state.period_int:
        opd_filtered = opd_df[opd_df["visit_year"] == filter_state.period_int]
    else:
        opd_filtered = opd_df

    if filter_state.departments:
        opd_filtered = opd_filtered[opd_filtered["department_name"].isin(filter_state.departments)]

    # Heatmap of Month x Weekday
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    opd_heat = opd_filtered.groupby(["visit_month", "weekday_name"]).size().reset_index(name="visits")
    heat_pivot = opd_heat.pivot(index="visit_month", columns="weekday_name", values="visits").reindex(columns=day_order).fillna(0)

    fig_heat = px.imshow(
        heat_pivot,
        labels=dict(x="Day of Week", y="Month", color="OPD Visits"),
        x=day_order,
        y=[f"Month {m}" for m in heat_pivot.index],
        color_continuous_scale="Blues",
        aspect="auto"
    )
    fig_heat.update_layout(margin=dict(l=20, r=20, t=20, b=20), height=380)
    ui.chart_card("OPD Heatmap: Arrivals by Month and Weekday", fig_heat, note="Notice heavy Monday clinic loading and Sunday minimal staffing.")
