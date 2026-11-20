"""Static + behavioural checks for src/ui.py without installing Streamlit.

Streamlit is not part of the Step 12 toolchain (it is installed at Step 14), so
this script injects a minimal stub for ``streamlit`` and
``streamlit.components.v1``, imports the real ui module and inspects the HTML
that would be sent to the browser. It proves the parts that matter here:
escaping, no stray anchors, INR grouping, fixed heights and theme-token parity.

Run:  .venv/bin/python scripts/check_ui.py
"""
from __future__ import annotations

import sys
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

CALLS: list[tuple] = []


class _Ctx:
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def _rec(kind):
    def inner(*args, **kwargs):
        CALLS.append((kind, args[0] if args else kwargs))
        return None
    return inner


st = types.ModuleType("streamlit")
st.set_page_config = _rec("page_config")
st.markdown = _rec("markdown")
st.caption = _rec("caption")
st.title = _rec("title")
st.metric = _rec("metric")
st.info = _rec("info")
st.warning = _rec("warning")
st.error = _rec("error")
st.success = _rec("success")
st.plotly_chart = _rec("plotly_chart")
st.container = lambda **kw: _Ctx()
st.columns = lambda n, gap=None: [_Ctx() for _ in range(n)]

components = types.ModuleType("streamlit.components")
components_v1 = types.ModuleType("streamlit.components.v1")
components_v1.html = lambda body, height=None, scrolling=False: CALLS.append(("html", body, height))
components.v1 = components_v1
st.components = components

sys.modules["streamlit"] = st
sys.modules["streamlit.components"] = components
sys.modules["streamlit.components.v1"] = components_v1

from src import ui  # noqa: E402

checks: list[tuple[str, bool, str]] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    checks.append((name, bool(ok), detail))


def frames() -> list[tuple[str, int]]:
    return [(b, h) for k, b, h in [(c[0], c[1], c[2]) for c in CALLS if c[0] == "html"]]


# --------------------------------------------------------------------- 1. components
CALLS.clear()
ui.header("Operations", "Capacity and demand · 2023–2025", as_of="31 Dec 2025")
check("header renders one iframe", len(frames()) == 1, f"{len(frames())} frames")
check("header height is HEADER_H", frames()[0][1] == ui.HEADER_H, f"{frames()[0][1]}px")

CALLS.clear()
ui.kpi_row([
    ui.card("Bed occupancy", "73.2%", "+1.2 pp vs 2024", "up", "warn", n=109_420, sub="8 departments"),
    ui.card("ALOS", "4.36 d", "−0.05 d", "down", "ok", n=18_142),
    ui.card("Readmission", "4.98%", "+0.3 pp", "up", "bad", low_n=True),
    ui.card("Revenue", ui.fmt_inr(13_526_00_000), "₹135.26 cr", "up", "neutral"),
])
check("kpi_row renders 4 iframes", len(frames()) == 4, f"{len(frames())}")
check("kpi frame height is CARD_H", all(h == ui.CARD_H for _, h in frames()))
html4 = "".join(b for b, _ in frames())
check("low-confidence badge present", "low confidence" in html4)
check("status colours bound to status", "bg-amber-600" in frames()[0][0] and "bg-red-600" in frames()[2][0])
check("no 'None' leaks into a card", "None" not in html4 and "nan" not in html4.lower())

# --------------------------------------------------------------------- 2. escaping
CALLS.clear()
evil = '<script>alert("x")</script>'
ui.header(evil, "sub & <b>bold</b>", as_of=evil)
body = frames()[0][0]
check("script tag escaped", "<script>alert" not in body and "&lt;script&gt;" in body)
check("attribute quotes escaped", 'alert(&quot;x&quot;)' in body or "&quot;" in body)
ui.insight_box("Cardiology is 1.1 pp above target — <img src=x onerror=1>", "warn")
body = frames()[-1][0]
check("insight text escaped", "<img" not in body and "&lt;img" in body)

# --------------------------------------------------------------------- 3. no anchors
CALLS.clear()
ui.header("t", "s")
ui.kpi_row([ui.card("a", "1")])
ui.chart_card("Occupancy", object(), note="Σ occupied ÷ Σ available")   # native, no iframe
ui.insight_box("x", "info", "act", "src")
ui.filter_chips(["Period 2025", "Dept All"])
ui.empty_state("No rows", ["Period 2025"])
ui.dq_banner(6, 0, "2025")
allhtml = "".join(b for b, _ in frames())
check("no anchor inside any iframe", "<a " not in allhtml and "href=" not in allhtml)
check("chart_card is native, not an iframe", "plotly_chart" in [c[0] for c in CALLS])
check("all heights positive", all(h and h > 0 for _, h in frames()))

# --------------------------------------------------------------------- 4. formatting
cases = [
    (1_234_567, "₹12,34,567"),
    (6_056_692, "₹60,56,692"),
    (82_544, "₹82,544"),
    (135_260_000, "₹13,52,60,000"),
    (0, "₹0"),
]
check("INR grouping (Indian lakh/crore)", all(ui.fmt_inr(v) == e for v, e in cases),
      "; ".join(f"{v}->{ui.fmt_inr(v)}" for v, _ in cases))
check("negative INR keeps the sign", ui.fmt_inr(-1234567) == "-₹12,34,567", ui.fmt_inr(-1234567))
check("delta uses the minus sign, not a hyphen", ui.fmt_delta(-1.1) == "−1.1 pp", ui.fmt_delta(-1.1))
check("delta uses the plus sign", ui.fmt_delta(1.1) == "+1.1 pp", ui.fmt_delta(1.1))
check("pct formatting", ui.fmt_pct(73.245) == "73.2%" and ui.fmt_pct(1.29, 2) == "1.29%")

# --------------------------------------------------------------------- 5. fallback
CALLS.clear()
ui.USE_TAILWIND = False
ui.header("Operations", "sub")
ui.kpi_row([ui.card("a", "1"), ui.card("b", "2")])
ui.insight_box("text", "warn")
ui.empty_state("No rows")
ui.dq_banner(5, 1, "2025")
check("fallback draws no iframe", len(frames()) == 0, f"{len(frames())}")
kinds = [c[0] for c in CALLS]
check("fallback uses native widgets",
      all(k in kinds for k in ("title", "metric", "warning", "info")), str(kinds))
ui.USE_TAILWIND = True

# --------------------------------------------------------------------- 6. tokens vs config
try:
    import tomllib
    cfg = tomllib.loads((ROOT / ".streamlit" / "config.toml").read_text())
    theme = cfg["theme"]
    pairs = [("primaryColor", "navy"), ("backgroundColor", "page"),
             ("secondaryBackgroundColor", "surface"), ("textColor", "text")]
    ok = all(theme[k].upper() == ui.THEME[v].upper() for k, v in pairs)
    check("config.toml matches ui.THEME", ok, str({k: theme[k] for k, _ in pairs}))
except Exception as exc:  # noqa: BLE001
    check("config.toml matches ui.THEME", False, repr(exc))

# --------------------------------------------------------------------- 7. contrast
def _lum(hex6: str) -> float:
    h = hex6.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4  # noqa: E731
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _cr(a: str, b: str) -> float:
    la, lb = _lum(a), _lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


ratios = {k: round(_cr(ui.THEME[k], "#FFFFFF"), 2) for k in ("navy", "teal", "amber", "red", "text", "muted")}
check("all text tokens >= 4.5:1 on white", all(r >= 4.5 for r in ratios.values()), str(ratios))

# --------------------------------------------------------------------- report
width = max(len(n) for n, _, _ in checks)
print(f"{'CHECK'.ljust(width)}  RESULT  DETAIL")
print("-" * (width + 40))
for name, ok, detail in checks:
    print(f"{name.ljust(width)}  {'PASS' if ok else 'FAIL':6}  {detail[:70]}")
failed = [c for c in checks if not c[1]]
print(f"\n{len(checks) - len(failed)}/{len(checks)} checks pass")
sys.exit(1 if failed else 0)
