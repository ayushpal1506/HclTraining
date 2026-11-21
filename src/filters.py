"""p107 Shared sidebar filters and session state manager (BUILD-06).

Maintains global filter state across all 9 screens.
Cascades wards based on selected departments.
Provides reset capabilities and produces active filter chip summaries.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Sequence

import pandas as pd
import streamlit as st


@dataclass
class FilterState:
    period: str = "2025"
    compare_to: str | None = "2024"
    departments: list[str] = field(default_factory=list)
    wards: list[str] = field(default_factory=list)
    payers: list[str] = field(default_factory=list)
    gender: str = "All"
    age_bands: list[str] = field(default_factory=list)

    @property
    def period_int(self) -> int | None:
        try:
            return int(self.period)
        except (ValueError, TypeError):
            return None

    @property
    def compare_int(self) -> int | None:
        try:
            return int(self.compare_to) if self.compare_to and self.compare_to != "None" else None
        except (ValueError, TypeError):
            return None


def init_session_state() -> None:
    """Ensure filter state keys exist in st.session_state."""
    defaults = {
        "flt_period": "2025",
        "flt_compare": "2024",
        "flt_dept": [],
        "flt_ward": [],
        "flt_payer": [],
        "flt_gender": "All",
        "flt_age": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def reset_filters() -> None:
    """Reset all filters to default state."""
    st.session_state["flt_period"] = "2025"
    st.session_state["flt_compare"] = "2024"
    st.session_state["flt_dept"] = []
    st.session_state["flt_ward"] = []
    st.session_state["flt_payer"] = []
    st.session_state["flt_gender"] = "All"
    st.session_state["flt_age"] = []


def render_sidebar(episodes_df: pd.DataFrame) -> tuple[FilterState, list[str]]:
    """Render the standard sidebar controls and return the active FilterState."""
    init_session_state()

    with st.sidebar:
        st.markdown("### 🏥 **p107 Filters**")
        st.caption("Filters apply across all pages via session state.")

        # 1. Period Selector (includes 2026 for empty state testing)
        period_options = ["2025", "2024", "2023", "2026", "All"]
        idx_period = period_options.index(st.session_state["flt_period"]) if st.session_state["flt_period"] in period_options else 0
        st.selectbox(
            "Reporting Period",
            options=period_options,
            index=idx_period,
            key="flt_period",
            help="Primary observation calendar year."
        )

        # 2. Comparison Period
        compare_options = ["None", "2024", "2023"]
        idx_comp = compare_options.index(st.session_state["flt_compare"]) if st.session_state["flt_compare"] in compare_options else 1
        st.selectbox(
            "Compare To (Baseline)",
            options=compare_options,
            index=idx_comp,
            key="flt_compare",
            help="Baseline period for calculating delta changes."
        )

        st.markdown("---")

        # 3. Department Multiselect
        all_depts = sorted(episodes_df["department_name"].dropna().unique().tolist())
        st.multiselect(
            "Department",
            options=all_depts,
            key="flt_dept",
            placeholder="All Departments"
        )

        # 4. Ward Multiselect (cascaded by department)
        if st.session_state["flt_dept"]:
            available_wards = sorted(
                episodes_df[episodes_df["department_name"].isin(st.session_state["flt_dept"])]["ward_name"]
                .dropna().unique().tolist()
            )
        else:
            available_wards = sorted(episodes_df["ward_name"].dropna().unique().tolist())

        # Prune any selected wards not in available
        st.session_state["flt_ward"] = [w for w in st.session_state["flt_ward"] if w in available_wards]
        st.multiselect(
            "Ward",
            options=available_wards,
            key="flt_ward",
            placeholder="All Wards"
        )

        # 5. Payer Multiselect
        all_payers = sorted(episodes_df["payer_category"].dropna().unique().tolist())
        st.multiselect(
            "Payer Category",
            options=all_payers,
            key="flt_payer",
            placeholder="All Payers"
        )

        # 6. Gender & Age
        with st.expander("Demographics Filter", expanded=False):
            gender_options = ["All", "M", "F", "O"]
            idx_g = gender_options.index(st.session_state["flt_gender"]) if st.session_state["flt_gender"] in gender_options else 0
            st.radio(
                "Gender",
                options=gender_options,
                index=idx_g,
                key="flt_gender",
                horizontal=True
            )

            all_ages = ["0-5", "6-17", "18-35", "36-50", "51-65", "65+"]
            st.multiselect(
                "Age Band",
                options=all_ages,
                key="flt_age",
                placeholder="All Ages"
            )

        st.markdown("---")
        if st.button("↺ Reset All Filters", use_container_width=True):
            reset_filters()
            st.rerun()

    # Build FilterState
    state = FilterState(
        period=st.session_state["flt_period"],
        compare_to=st.session_state["flt_compare"] if st.session_state["flt_compare"] != "None" else None,
        departments=list(st.session_state["flt_dept"]),
        wards=list(st.session_state["flt_ward"]),
        payers=list(st.session_state["flt_payer"]),
        gender=st.session_state["flt_gender"],
        age_bands=list(st.session_state["flt_age"]),
    )

    # Build active chips list
    chips = []
    if state.period:
        chips.append(f"Period: {state.period}")
    if state.compare_to:
        chips.append(f"Compare: {state.compare_to}")
    if state.departments:
        chips.extend(state.departments)
    if state.wards:
        chips.extend(state.wards)
    if state.payers:
        chips.extend(state.payers)
    if state.gender != "All":
        chips.append(f"Gender: {state.gender}")
    if state.age_bands:
        chips.extend(state.age_bands)

    return state, chips
