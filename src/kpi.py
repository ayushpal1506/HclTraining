"""KPI mathematics — the single place a published figure is computed.

Design rule (NFR-10): no page and no chart computes a KPI. A page calls one of
these functions and formats the result. That is what makes TC-T-21…TC-T-30
possible: the maths is a pure function of numbers, so it can be checked against
hand arithmetic before any dataframe, chart or browser exists.

Every function returns ``None`` when the denominator is zero rather than raising
or emitting ``nan``. A missing denominator is a real state in this dataset —
a department-month with no eligible discharge — and the UI must render "—", not
``ZeroDivisionError`` (US-18, UIF-03).
"""
from __future__ import annotations

from datetime import date

# Targets and thresholds live here, not in a page (BR-01, BR-06, BR-07, BR-10).
TARGET_BOR_PCT = 85.0
BENCHMARK_ALOS_DAYS = 4.0
READMISSION_WINDOW_DAYS = 30
LOW_CONFIDENCE_N = 30
DRIFT_Z = 2.0


def _div(numerator: float, denominator: float, places: int | None = None) -> float | None:
    if denominator in (0, None):
        return None
    value = numerator / denominator
    return round(value, places) if places is not None else value


def bor_pct(bed_days_occupied: float, bed_days_available: float) -> float | None:
    """Bed occupancy rate: occupied ÷ available bed-days × 100 (BR-01, OQ-01)."""
    rate = _div(bed_days_occupied, bed_days_available)
    return None if rate is None else round(rate * 100, 2)


def alos_days(inpatient_days: float, discharges: float) -> float | None:
    """Average length of stay: Σ los_days ÷ discharges (BR-06, OQ-03).

    Deaths and LAMA are included; the basis is stated on screen (US-06 AC2).
    """
    return _div(inpatient_days, discharges, 2)


def mortality_pct(deaths: float, discharges: float) -> float | None:
    """In-hospital mortality: deaths ÷ discharges × 100 (BR-13, OQ-10)."""
    rate = _div(deaths, discharges)
    return None if rate is None else round(rate * 100, 2)


def readmission_rate_pct(readmissions: float, eligible_discharges: float) -> float | None:
    """30-day unplanned readmission rate (BR-07, OQ-04).

    The denominator is discharges that could be readmitted — deaths are removed
    because a deceased patient cannot return. Planned readmissions are already
    excluded upstream (DQ rule, T-04).
    """
    rate = _div(readmissions, eligible_discharges)
    return None if rate is None else round(rate * 100, 2)


def revenue_per_patient(revenue: float, discharges: float) -> float | None:
    """Revenue per discharged patient (BR-08, OQ-05)."""
    return _div(revenue, discharges, 2)


def revenue_per_bed_day(revenue: float, bed_days: float) -> float | None:
    """Revenue per occupied bed-day (BR-12, OQ-09)."""
    return _div(revenue, bed_days, 2)


def occupancy_gap_pp(bor: float | None, target: float = TARGET_BOR_PCT) -> float | None:
    """Gap to the occupancy target in percentage points; negative = below target."""
    if bor is None:
        return None
    return round(bor - target, 2)


def beds_to_target(bor: float | None, bed_days_available: float, days: int,
                   target: float = TARGET_BOR_PCT) -> float | None:
    """Occupied beds needed to reach the target, averaged over *days*.

    This is the number an operations manager acts on, so the gap is published in
    beds as well as points (US-04 AC3). Returns None when bed-days or days is 0.
    """
    if bor is None or days <= 0:
        return None
    target_occupied = bed_days_available * target / 100.0
    current_occupied = bed_days_available * bor / 100.0
    return round((target_occupied - current_occupied) / days, 2)


def eligibility(discharges: float, deaths: float) -> float:
    """Discharges that could be readmitted: everyone who went home alive."""
    return max(discharges - deaths, 0)


def readmitted_within(index_discharge: date | str, next_admission: date | str,
                      window: int = READMISSION_WINDOW_DAYS) -> bool:
    """True when the next admission falls inside the 30-day readmission window.

    Both arguments accept ``date`` or ISO ``YYYY-MM-DD``. An admission on the
    same day as discharge counts (gap 0) and an admission on day 30 counts; day
    31 does not (BR-07 boundary, US-09 AC3).
    """
    end = _as_date(index_discharge)
    start = _as_date(next_admission)
    return 0 <= (start - end).days <= window


def low_confidence(n: float | None) -> bool:
    """The published rule: fewer than 30 observations is low confidence (US-18)."""
    return n is not None and n < LOW_CONFIDENCE_N


def share_pct(part: float, whole: float) -> float | None:
    """Compositional share, used by the payer-mix and discharge-mix charts."""
    rate = _div(part, whole)
    return None if rate is None else round(rate * 100, 2)


def contribution(part_a: float, part_b: float) -> dict[str, float]:
    """One department's contribution to a total movement (US-14, OQ-09).

    Returns the movement, its share of the total, and leaves the residual to the
    caller: unattributed movement is shown, never merged into "other".
    """
    delta = part_b - part_a
    return {"a": part_a, "b": part_b, "delta": round(delta, 2)}


def decompose(parts_a: dict[str, float], parts_b: dict[str, float]) -> dict[str, object]:
    """Decompose a total movement into named contributors plus a printed residual.

    ``parts_a`` and ``parts_b`` are keyed by the same department labels. Keys
    present in only one side are treated as a 0 baseline (a department that
    opened or closed), which is a fact, not an error.
    """
    keys = sorted(set(parts_a) | set(parts_b))
    rows = []
    for k in keys:
        a, b = parts_a.get(k, 0.0), parts_b.get(k, 0.0)
        rows.append((k, round(b - a, 2)))
    total_a = sum(parts_a.values())
    total_b = sum(parts_b.values())
    total_delta = round(total_b - total_a, 2)
    attributed = round(sum(d for _, d in rows), 2)
    return {
        "rows": sorted(rows, key=lambda r: abs(r[1]), reverse=True),
        "total_a": round(total_a, 2),
        "total_b": round(total_b, 2),
        "total_delta": total_delta,
        "residual": round(total_delta - attributed, 2),
    }


def _as_date(value: date | str) -> date:
    if isinstance(value, date):
        return value
    return date.fromisoformat(str(value)[:10])
