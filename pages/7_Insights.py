"""p107 Advanced Insights & Decomposition — Screen UI-08 (BUILD-14).

Mathematical waterfall decomposition of KPI shifts into departmental contributors,
and formal statistical hypothesis testing (Spearman rank correlation) (FR-14, FR-15, BR-03, BR-10).
"""
from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from src import data, filters, kpi, stats, ui

# 1. Page Configuration
ui.setup("Insights · p107")
ui.enable_mobile()

# 2. Data Loading & Sidebar
episodes_df = data.load_episodes()
dept_month_df = data.load_dept_month_kpi()
filter_state, chips = filters.render_sidebar(episodes_df)

# 3. Header & Filter Chips
ui.header("Advanced Analytics & Insights", "Mathematical KPI variance decomposition & statistical hypothesis testing", as_of="31 Dec 2025")
if chips:
    ui.filter_chips(chips)

tab1, tab2 = st.tabs(["💧 Revenue Movement Waterfall Decomposition", "📈 Association Hypothesis Testing (OQ-07)"])

with tab1:
    st.markdown("#### **Departmental Decomposition of Annual Revenue Shift**")
    st.caption("Decomposes movement between baseline and reporting period (e.g., 2024 to 2025: +₹2.16 cr) with exact printed residual (US-14).")

    # Baseline 2024 vs 2025 revenue per department
    rev_2024 = episodes_df[episodes_df["admission_year"] == 2024].groupby("department_name")["billed_amount"].sum() / 10_000_000.0
    rev_2025 = episodes_df[episodes_df["admission_year"] == 2025].groupby("department_name")["billed_amount"].sum() / 10_000_000.0

    decomp = kpi.decompose(dict(rev_2024), dict(rev_2025))

    names = ["2024 Total"] + [r[0] for r in decomp["rows"]] + ["2025 Total"]
    measures = ["absolute"] + ["relative"] * len(decomp["rows"]) + ["total"]
    vals = [decomp["total_a"]] + [r[1] for r in decomp["rows"]] + [decomp["total_b"]]

    fig_water = go.Figure(go.Waterfall(
        name="Revenue Decomposition",
        orientation="v",
        measure=measures,
        x=names,
        textposition="outside",
        text=[f"₹{v:+.2f} cr" if m == "relative" else f"₹{v:.2f} cr" for v, m in zip(vals, measures)],
        y=vals,
        connector={"line": {"color": "#94A3B8"}},
        decreasing={"marker": {"color": "#B91C1C"}},
        increasing={"marker": {"color": "#0F766E"}},
        totals={"marker": {"color": "#1F3864"}}
    ))
    fig_water.update_layout(yaxis_title="Inpatient Revenue (₹ Crores)", height=400, margin=dict(l=10, r=10, t=10, b=10))
    ui.chart_card("Revenue Movement Waterfall (2024 → 2025)", fig_water, note=f"Total Delta: ₹{decomp['total_delta']:+.2f} cr | Calculated Residual: ₹{decomp['residual']:.2f} cr")

with tab2:
    st.markdown("#### **Hypothesis Test: Bed Occupancy Rate vs 30-Day Readmission Rate**")
    st.caption("OQ-07: Testing whether high operational bed occupancy correlates with elevated readmission pressure across 288 department-months.")

    # Calculate monthly BOR and readmission rate per department
    clean_kpi = dept_month_df.dropna(subset=["bor_pct", "readmission_rate"])
    x_bor = clean_kpi["bor_pct"].tolist()
    y_readm = clean_kpi["readmission_rate"].tolist()

    test_res = stats.correlation_test(x_bor, y_readm)

    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("##### **Statistical Test Results**")
        st.write(f"**Method:** Spearman Rank Correlation (`{test_res['method']}`)")
        st.write(f"**Sample Size (n):** `{test_res['n']}` department-months")
        st.write(f"**Correlation (rho):** `{test_res['rho']}`")
        st.write(f"**t-Statistic:** `{test_res['t']}`")
        st.write(f"**Two-Sided p-Value:** `{test_res['p']}`")
        if test_res["p"] is not None and test_res["p"] < 0.05:
            st.success("Statistically significant association detected (p < 0.05).")
        else:
            st.info("No statistically significant monotonic association (p ≥ 0.05).")

        ui.insight_box(
            "Wording Note: This confirms statistical association between capacity strain and patient rebound, NOT clinical causation.",
            level="warn",
            action="Investigate Discharge Quality",
            source="BR-10 · Statistical Protocol"
        )

    with col2:
        ols = stats.ols_line(x_bor, y_readm)
        fig_scat = px.scatter(
            clean_kpi,
            x="bor_pct",
            y="readmission_rate",
            color="department_name",
            labels={"bor_pct": "Monthly Bed Occupancy (%)", "readmission_rate": "30-Day Readmission Rate (%)", "department_name": "Specialty"},
        )
        if ols:
            slope, intercept = ols
            line_x = [min(x_bor), max(x_bor)]
            line_y = [slope * x + intercept for x in line_x]
            fig_scat.add_trace(go.Scatter(
                x=line_x,
                y=line_y,
                mode="lines",
                line=dict(color="#B91C1C", dash="dash"),
                name="OLS Trendline"
            ))
        fig_scat.update_layout(height=380, margin=dict(l=10, r=10, t=10, b=10))
        ui.chart_card("Scatter Plot: Occupancy vs Readmissions", fig_scat)
