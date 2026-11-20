"""HKD-2026 Hospital Finance, Billing & Payer Mix — Screen UI-05 (BUILD-11).

Tracks payer market share transitions, revenue per discharged patient,
and ranked revenue per occupied bed-day across clinical specialties (OQ-05, OQ-09).
"""
from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from src import data, filters, kpi, ui

# 1. Page Configuration
ui.setup("Finance · HKD-2026")
ui.enable_mobile()

# 2. Data Loading & Sidebar
episodes_df = data.load_episodes()
filter_state, chips = filters.render_sidebar(episodes_df)

# 3. Header & Filter Chips
ui.header("Hospital Financial Intelligence", "Payer mix migration, revenue per patient & yield per occupied bed-day", as_of="31 Dec 2025")
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
    ui.empty_state("No financial billing records match the selected criteria.", chips)
    st.stop()

# 5. Top Financial KPIs
total_discharges = len(cohort)
total_rev = cohort["billed_amount"].sum()
total_bed_days = cohort["bed_days"].sum()

rev_per_patient = kpi.revenue_per_patient(total_rev, total_discharges)
rev_per_bed_day = kpi.revenue_per_bed_day(total_rev, total_bed_days)
rev_cr = total_rev / 10_000_000.0

ui.kpi_row([
    ui.card("Total Billed Revenue", f"₹{rev_cr:.2f} cr", f"{total_discharges:,} discharges", "neutral", "ok", n=total_discharges),
    ui.card("Revenue per Patient", f"₹{rev_per_patient:,.0f}" if rev_per_patient else "—", "Average billing yield", "neutral", "ok", n=total_discharges),
    ui.card("Revenue per Bed-Day", f"₹{rev_per_bed_day:,.0f}" if rev_per_bed_day else "—", f"{total_bed_days:,} bed-days", "neutral", "ok", n=total_bed_days),
])

col1, col2 = st.columns(2)

with col1:
    st.markdown("#### **Payer Mix Evolution (2023–2025)**")
    st.caption("P-04: Corporate volume expands (6.37% → 11.23%) as Self-Pay contracts (41.35% → 33.49%).")

    # Group by admission_year and payer_category
    payer_trend = episodes_df.groupby(["admission_year", "payer_category"]).size().reset_index(name="count")
    total_per_yr = payer_trend.groupby("admission_year")["count"].transform("sum")
    payer_trend["share_pct"] = (payer_trend["count"] / total_per_yr * 100).round(2)

    fig_payer = px.bar(
        payer_trend,
        x="admission_year",
        y="share_pct",
        color="payer_category",
        barmode="stack",
        color_discrete_sequence=["#1F3864", "#0F766E", "#B45309", "#64748B", "#CBD5E1"],
        labels={"share_pct": "Share of Volume (%)", "admission_year": "Year", "payer_category": "Payer"}
    )
    fig_payer.update_layout(yaxis_range=[0, 100], height=380, margin=dict(l=10, r=10, t=10, b=10))
    ui.chart_card("Payer Mix Transition Stack", fig_payer)

with col2:
    st.markdown("#### **Specialty Ranking: Revenue per Occupied Bed-Day**")
    st.caption("Ranked yield against #1 leader (Oncology ₹31,232/day) with gap lines (BR-12).")

    dept_fin = cohort.groupby("department_name").agg(
        rev=("billed_amount", "sum"),
        days=("bed_days", "sum")
    ).reset_index()
    dept_fin["yield_per_day"] = (dept_fin["rev"] / dept_fin["days"]).round(0)
    dept_fin = dept_fin.sort_values("yield_per_day", ascending=True)

    max_yield = dept_fin["yield_per_day"].max()

    fig_rank = go.Figure()

    # Stem lines to leader
    for _, row in dept_fin.iterrows():
        fig_rank.add_trace(go.Scatter(
            x=[row["yield_per_day"], max_yield],
            y=[row["department_name"], row["department_name"]],
            mode="lines",
            line=dict(color="#CBD5E1", width=2),
            showlegend=False,
            hoverinfo="none"
        ))

    # Dot markers
    fig_rank.add_trace(go.Scatter(
        x=dept_fin["yield_per_day"],
        y=dept_fin["department_name"],
        mode="markers+text",
        marker=dict(color="#1F3864", size=10),
        text=[f"₹{y:,.0f}" for y in dept_fin["yield_per_day"]],
        textposition="top right",
        name="Yield / Day"
    ))

    fig_rank.update_layout(xaxis_title="Revenue / Bed-Day (₹)", yaxis_title="", height=380, margin=dict(l=10, r=10, t=10, b=10))
    ui.chart_card("Revenue per Bed-Day Ranking", fig_rank)
