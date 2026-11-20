"""Statistics used by UI-08, in closed form.

SciPy/statsmodels is the intended library (TRD M4) and is installed at Step 14.
The two functions that decide what a reader sees — the Wilson interval behind
every rate and the rank correlation behind OQ-07 — are implemented here in
closed form instead, for one reason: they can be unit-tested before the runtime
stack exists, and Step 15 asserts these values against SciPy to 1e-9.

Closed form is not a shortcut for the hard parts. Anything that needs a real
solver (multiple regression, ANOVA) stays a SciPy call in ``src/stats.py`` when
the stack lands; nothing here pretends to replace it.
"""
from __future__ import annotations

import math

Z_95 = 1.959963984540054


def wilson_interval(successes: int, n: int, z: float = Z_95) -> tuple[float, float] | None:
    """Wilson score interval for a proportion, returned as percentages 0–100.

    Chosen over the normal approximation because it stays inside 0–100 for small
    n and for rates near 0 — which is exactly this dataset's failure mode (a
    5.5% readmission rate on 34 discharges). US-09 AC2, UIF-03.
    """
    if n <= 0:
        return None
    p = successes / n
    denom = 1 + z * z / n
    centre = (p + z * z / (2 * n)) / denom
    half = (z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))) / denom
    lo = max(0.0, centre - half)
    hi = min(1.0, centre + half)
    return round(lo * 100, 2), round(hi * 100, 2)


def _ranks(values: list[float]) -> list[float]:
    """Average ranks, so tied observations share the mean of their positions."""
    order = sorted(range(len(values)), key=lambda i: values[i])
    ranks = [0.0] * len(values)
    i = 0
    while i < len(order):
        j = i
        while j + 1 < len(order) and values[order[j + 1]] == values[order[i]]:
            j += 1
        avg = (i + j) / 2 + 1
        for k in range(i, j + 1):
            ranks[order[k]] = avg
        i = j + 1
    return ranks


def pearson(x: list[float], y: list[float]) -> float | None:
    """Pearson product-moment correlation; None when a series has no variance."""
    if len(x) != len(y) or len(x) < 3:
        return None
    n = len(x)
    mx, my = sum(x) / n, sum(y) / n
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    sxx = sum((a - mx) ** 2 for a in x)
    syy = sum((b - my) ** 2 for b in y)
    if sxx == 0 or syy == 0:
        return None
    return max(-1.0, min(1.0, sxy / math.sqrt(sxx * syy)))


def spearman(x: list[float], y: list[float]) -> float | None:
    """Spearman rank correlation — Pearson on average ranks (OQ-07).

    Rank-based because the association claim is about order, not linearity, and
    because a single 90-day outlier must not decide the answer.
    """
    if len(x) != len(y) or len(x) < 3:
        return None
    return pearson(_ranks(x), _ranks(y))


def _betacf(a: float, b: float, x: float, max_iter: int = 200, eps: float = 3e-14) -> float:
    """Continued fraction for the incomplete beta (Numerical Recipes 6.4)."""
    tiny = 1e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c = 1.0
    d = 1.0 - qab * x / qap
    if abs(d) < tiny:
        d = tiny
    d = 1.0 / d
    h = d
    for m in range(1, max_iter + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < tiny:
            d = tiny
        c = 1.0 + aa / c
        if abs(c) < tiny:
            c = tiny
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < tiny:
            d = tiny
        c = 1.0 + aa / c
        if abs(c) < tiny:
            c = tiny
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < eps:
            break
    return h


def _betai(a: float, b: float, x: float) -> float:
    """Regularised incomplete beta I_x(a, b)."""
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    lbeta = math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
    front = math.exp(lbeta + a * math.log(x) + b * math.log1p(-x))
    if x < (a + 1) / (a + b + 2):
        return front * _betacf(a, b, x) / a
    return 1.0 - front * _betacf(b, a, 1.0 - x) / b


def student_t_p(t: float, df: int) -> float:
    """Two-sided p-value for a t statistic; exact, no SciPy."""
    if df <= 0:
        return float("nan")
    return _betai(df / 2.0, 0.5, df / (df + t * t))


def correlation_test(x: list[float], y: list[float]) -> dict[str, object]:
    """Spearman rho with n, t and a two-sided p — the numbers printed on UI-08."""
    rho = spearman(x, y)
    n = len(x)
    if rho is None or n < 3:
        return {"rho": rho, "n": n, "t": None, "p": None, "method": "spearman",
                "note": "insufficient variation or n < 3"}
    if abs(rho) >= 1.0:
        return {"rho": rho, "n": n, "t": None, "p": 0.0, "method": "spearman",
                "note": "perfect monotone association"}
    t = rho * math.sqrt((n - 2) / (1 - rho * rho))
    return {"rho": round(rho, 4), "n": n, "t": round(t, 4),
            "p": round(student_t_p(t, n - 2), 4), "method": "spearman",
            "note": "association, not causation"}


def chi_square_2x2(a: int, b: int, c: int, d: int) -> tuple[float, float] | None:
    """Pearson chi-square and its p-value for a 2×2 table (df = 1).

    At df = 1 the survival function is exactly ``erfc(sqrt(chi2 / 2))``, so no
    lookup table or library is required.
    """
    n = a + b + c + d
    if n == 0:
        return None
    denom = (a + b) * (c + d) * (a + c) * (b + d)
    if denom == 0:
        return None
    chi2 = n * (a * d - b * c) ** 2 / denom
    p = math.erfc(math.sqrt(chi2 / 2.0))
    return round(chi2, 4), round(p, 4)


def ols_line(x: list[float], y: list[float]) -> tuple[float, float] | None:
    """Least-squares slope and intercept for the scatter's trend line."""
    if len(x) != len(y) or len(x) < 2:
        return None
    n = len(x)
    mx, my = sum(x) / n, sum(y) / n
    sxx = sum((a - mx) ** 2 for a in x)
    if sxx == 0:
        return None
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    slope = sxy / sxx
    return slope, my - slope * mx


def drift_flag(series: list[float], baseline_n: int = 12,
               z: float = 2.0) -> dict[str, object]:
    """Flag months outside baseline mean ± z·SD (BR-13, OQ-10).

    The first ``baseline_n`` points define the baseline; later points are tested
    against it. This is a *review trigger*, not a significance test — the screen
    must say so (US-13 AC3).
    """
    if len(series) <= baseline_n or baseline_n < 2:
        return {"flags": [], "mean": None, "sd": None, "first": None}
    base = series[:baseline_n]
    mean = sum(base) / len(base)
    sd = math.sqrt(sum((v - mean) ** 2 for v in base) / (len(base) - 1))
    if sd == 0:
        return {"flags": [], "mean": round(mean, 4), "sd": 0.0, "first": None}
    low, high = mean - z * sd, mean + z * sd
    flags = [i for i, v in enumerate(series) if i >= baseline_n and (v < low or v > high)]
    return {"flags": flags, "mean": round(mean, 4), "sd": round(sd, 4),
            "first": flags[0] if flags else None,
            "band": (round(low, 4), round(high, 4))}
