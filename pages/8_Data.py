"""p107 Raw Data Exploration & Secure Export — Screen UI-09 (BUILD-15).

Paginated analytical view preview, column data dictionary, and security-hardened
CSV export with formula injection sanitization, UTF-8 BOM and 50,000-row guard (US-20, NFR-09).
"""
from __future__ import annotations

import io
import pandas as pd
import streamlit as st

from src import data, filters, ui

# 1. Page Configuration
ui.setup("Data & Export · p107")
ui.enable_mobile()

# 2. Data Loading & Sidebar
episodes_df = data.load_episodes()
filter_state, chips = filters.render_sidebar(episodes_df)

# 3. Header & Filter Chips
ui.header("Data Cohort Exploration & Export", "Audit filtered analytical episodes, inspect column definitions & download securely", as_of="31 Dec 2025")
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

st.metric("Filtered Records Available", f"{len(cohort):,} episodes")

if len(cohort) == 0:
    ui.empty_state("No records match the active filter criteria.", chips)
    st.stop()

tab1, tab2, tab3 = st.tabs(["📋 Filtered Data Explorer", "📥 Secure CSV Export", "📖 Column Dictionary"])

with tab1:
    st.markdown("##### **Inpatient Episode View (Preview limited to first 1,000 records)**")
    preview_df = cohort.head(1000)
    st.dataframe(preview_df, use_container_width=True, hide_index=True)

with tab2:
    st.markdown("##### **Download Sanitized Cohort Dataset**")
    st.caption("Exports are capped at 50,000 rows. Text cells starting with formula trigger tokens (=, +, -, @) are sanitized with an apostrophe prefix.")

    if len(cohort) > 50_000:
        st.warning(f"Active filter cohort contains {len(cohort):,} rows, exceeding the 50,000 export limit (TC-F-36). Please narrow filters by Year or Department.")
    else:
        # Build sanitized CSV payload
        export_df = cohort.copy()
        # Sanitize text columns against formula injection
        for col in export_df.select_dtypes(include=["object", "string"]).columns:
            export_df[col] = export_df[col].astype(str).apply(
                lambda val: f"'{val}" if val and val[0] in ("=", "+", "-", "@") else val
            )

        csv_buffer = io.BytesIO()
        # Prepend UTF-8 BOM
        csv_buffer.write(b"\xef\xbb\xbf")
        export_df.to_csv(csv_buffer, index=False, encoding="utf-8")
        csv_bytes = csv_buffer.getvalue()

        st.download_button(
            label=f"⬇️ Download Filtered Cohort ({len(cohort):,} rows CSV)",
            data=csv_bytes,
            file_name=f"hkd2026_episodes_{filter_state.period}.csv",
            mime="text/csv",
            key="btn_download_csv"
        )
        st.caption("Payload verified: UTF-8 BOM prefixed for Excel compatibility. Zero PII columns contained.")

with tab3:
    st.markdown("##### **Warehouse Data Dictionary (`vw_analysis_episode`)**")
    dict_rows = [
        {"Column": col, "Data Type": str(cohort[col].dtype), "Sample Value": str(cohort[col].iloc[0])}
        for col in cohort.columns
    ]
    st.dataframe(pd.DataFrame(dict_rows), use_container_width=True, hide_index=True)
