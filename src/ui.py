"""HKD-2026 — Tailwind UI primitives (Step 12 starter).

Only the page header, the KPI cards, the filter chips, the insight boxes, the
data-quality banner and the empty state are rendered through an iframe.
Charts and tables stay native Streamlit widgets, so Plotly hover/zoom and
``st.dataframe`` sorting keep working (see docs/12_UI_Framework.docx §4–§5).

Rules enforced here
-------------------
* Every dynamic value is escaped with ``html.escape`` before it reaches the HTML.
* Numbers are formatted on the server; the client does no arithmetic and no fetch.
* Each component is display-only — no widget, no ``<a href>``. The iframe is a
  separate document, so a link inside it would navigate inside the iframe and
  strand the user. All navigation stays in Streamlit (sidebar / st.page_link).
* Heights are fixed per component (``*_H``) because ``components.html`` cannot
  auto-size; long strings are truncated with a CSS ellipsis, never wrapped.
* ``USE_TAILWIND = False`` degrades to native Streamlit components if the CDN is
  unreachable (AS-29). The page still renders — it just loses the card styling.

Usage
-----
    import streamlit as st
    from src import ui

    ui.setup("Operations · HKD-2026")          # st.set_page_config, once per page
    ui.enable_mobile()                         # 375 px reflow (NFR-08)
    ui.header("Operations", "Capacity and demand · 2023–2025", as_of="31 Dec 2025")
    ui.kpi_row([
        ui.card("Bed occupancy", "73.2%", "+1.2 pp vs 2024", "up", "warn", n=109_420),
        ui.card("ALOS", "4.36 d", "−0.05 d", "down", "ok"),
    ])
"""
from __future__ import annotations

import html
from typing import Iterable, Sequence

import streamlit as st
import streamlit.components.v1 as components

# --------------------------------------------------------------------------- config
TAILWIND_CDN = "https://cdn.tailwindcss.com"
USE_TAILWIND = True

# Single source of truth for colour. .streamlit/config.toml repeats the first
# four values because Streamlit cannot read Python; tests/test_ui_tokens.py
# asserts the two agree (UIF-01).
THEME = {
    "navy": "#1F3864",     # primary — actions, selected filters
    "teal": "#0F766E",     # on target / positive
    "amber": "#B45309",    # warning / low confidence / in progress
    "red": "#B91C1C",      # breach / failed check
    "text": "#0F172A",     # body text (17.8:1 on white)
    "muted": "#475569",    # secondary text (7.6:1 on white)
    "page": "#FFFFFF",
    "surface": "#F4F6FA",
    "border": "#E2E8F0",
}

# Fixed iframe heights in CSS px. Changing one of these changes every screen.
HEADER_H = 78
CARD_H = 118
CHIPS_H = 36
INSIGHT_H = 98
BANNER_H = 58
EMPTY_H = 138

# Test seam: the last value rendered per card label. components.html output is
# invisible to Streamlit AppTest, so tests/test_app.py asserts against this
# dictionary after a page run. header() clears it at the start of every page.
LAST_RENDER: dict[str, str] = {}

STATUS = {
    "ok":      {"dot": "bg-teal-600",  "text": "text-teal-700",  "tint": "bg-teal-50",
                "ring": "border-teal-200",  "edge": "border-l-teal-600"},
    "warn":    {"dot": "bg-amber-600", "text": "text-amber-700", "tint": "bg-amber-50",
                "ring": "border-amber-200", "edge": "border-l-amber-600"},
    "bad":     {"dot": "bg-red-600",   "text": "text-red-700",   "tint": "bg-red-50",
                "ring": "border-red-200",   "edge": "border-l-red-600"},
    "neutral": {"dot": "bg-slate-400", "text": "text-slate-600", "tint": "bg-slate-50",
                "ring": "border-slate-200", "edge": "border-l-slate-400"},
}

_HEAD = """<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<script src="__CDN__"></script>
<script>
  tailwind.config = {
    theme: { extend: {
      colors: { navy: '#1F3864', teal: '#0F766E', amber: '#B45309' },
      fontFamily: { sans: ['Inter', 'ui-sans-serif', 'system-ui', 'sans-serif'] }
    } }
  };
</script>
<style>
  html, body { margin: 0; background: transparent; }
  body { font-family: Inter, ui-sans-serif, system-ui, sans-serif;
         -webkit-font-smoothing: antialiased; overflow: hidden; }
</style></head><body>"""


def _esc(value: object) -> str:
    """Escape anything that reaches the iframe. Numbers become strings."""
    return html.escape(str(value), quote=True)


def _frame(body: str, height: int) -> None:
    components.html(_HEAD.replace("__CDN__", TAILWIND_CDN) + body + "</body></html>",
                    height=height, scrolling=False)


# --------------------------------------------------------------------------- page setup
def setup(page_title: str, layout: str = "wide") -> None:
    """st.set_page_config wrapper — must be the first Streamlit call on a page."""
    st.set_page_config(page_title=page_title, page_icon="🏥", layout=layout,
                       initial_sidebar_state="expanded")


MOBILE_CSS = """<style>
        @media (max-width: 640px) {
          div[data-testid="stHorizontalBlock"] { flex-wrap: wrap !important; gap: .5rem !important; }
          div[data-testid="column"] { min-width: 100% !important; flex: 1 1 100% !important; }
          div.block-container { padding-left: .75rem !important; padding-right: .75rem !important; }
          h1, h2, h3 { font-size: 1.1rem !important; }
        }
        </style>"""


def enable_mobile(max_width: int = 640) -> None:
    """Make st.columns stack below ``max_width`` (NFR-08, US-08 AC2).

    Streamlit keeps columns side by side at 375 px unless they are told to wrap.
    The selectors are Streamlit internals, so this is the one place that may need
    revising on a Streamlit upgrade (UIF-04); the Step 15 screenshot test at
    375 / 768 / 1440 px is what catches a regression.
    """
    css = MOBILE_CSS.replace("640px", f"{max_width}px") if max_width != 640 else MOBILE_CSS
    st.markdown(css, unsafe_allow_html=True)


# --------------------------------------------------------------------------- formatting
def fmt_inr(value: float, decimals: int = 0) -> str:
    """Indian digit grouping: 1234567 -> '₹12,34,567'.

    Formatting happens without grouping first — asking Python for ``,.0f`` would
    insert western commas that the grouping loop would then treat as digits.
    """
    text = f"{value:.{decimals}f}"
    whole, _, frac = text.partition(".")
    sign = ""
    if whole.startswith("-"):
        sign, whole = "-", whole[1:]
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        whole = ",".join(groups) + "," + tail
    return sign + "₹" + whole + (("." + frac) if frac else "")


def fmt_pct(value: float, decimals: int = 1) -> str:
    return f"{value:.{decimals}f}%"


def fmt_delta(value: float, unit: str = "pp", decimals: int = 1) -> str:
    """Signed delta with the correct minus sign (U+2212), never a hyphen."""
    sign = "+" if value > 0 else ("−" if value < 0 else "±")
    return f"{sign}{abs(value):.{decimals}f} {unit}"


# --------------------------------------------------------------------------- components
def header(title: str, subtitle: str = "", as_of: str | None = None,
           version: str = "v1.0") -> None:
    """Page header: title, subtitle, data-as-of chip and definition version."""
    LAST_RENDER.clear()
    st.markdown(f'<span style="display:none">HKD-2026 data to {as_of} defs {version}</span>', unsafe_allow_html=True)
    if not USE_TAILWIND:
        st.title(title)
        if subtitle:
            st.caption(subtitle)
        return
    chips = [f"data to {as_of}" if as_of else None, f"definitions {version}"]
    meta = "".join(
        f'<span class="rounded-full bg-slate-100 px-2.5 py-1 text-[11px] font-medium '
        f'text-slate-600 whitespace-nowrap">{_esc(c)}</span>' for c in chips if c
    )
    body = f"""
<div class="flex flex-wrap items-end justify-between gap-2 border-b border-slate-200 pb-2">
  <div class="min-w-0">
    <h1 class="truncate text-xl font-semibold tracking-tight text-slate-900">{_esc(title)}</h1>
    <p class="mt-0.5 truncate text-xs text-slate-500">{_esc(subtitle)}</p>
  </div>
  <div class="flex flex-wrap items-center gap-1.5">{meta}</div>
</div>"""
    _frame(body, HEADER_H)


def card(label: str, value: str, delta: str = "", direction: str = "flat",
         status: str = "neutral", sub: str = "", n: int | None = None,
         low_n: bool = False) -> dict:
    """Build the dict that :func:`kpi_card` renders. Keeps pages declarative."""
    return {"label": label, "value": value, "delta": delta, "direction": direction,
            "status": status, "sub": sub, "n": n, "low_n": low_n}


def kpi_card(spec: dict, height: int = CARD_H) -> None:
    """One KPI card. Pass a dict from :func:`card`.

    status drives the dot and the delta colour; ``low_n`` adds the BR-14 badge
    so a two-case month can never look like a stable rate (US-18).
    """
    LAST_RENDER[spec.get("label", "")] = spec.get("value", "—")
    if not USE_TAILWIND:
        st.metric(spec.get("label", ""), spec.get("value", "—"), spec.get("delta") or None)
        return
    palette = STATUS.get(spec.get("status", "neutral"), STATUS["neutral"])
    arrow = {"up": "▲", "down": "▼", "flat": "●"}.get(spec.get("direction", "flat"), "●")
    badge = ('<span class="ml-1 rounded bg-slate-200 px-1.5 py-0.5 text-[10px] font-medium '
             'text-slate-700 whitespace-nowrap">low confidence</span>') if spec.get("low_n") else ""
    body = f"""
<div class="h-full rounded-xl border border-slate-200 bg-white p-3 shadow-sm">
  <div class="flex items-start justify-between gap-2">
    <p class="truncate text-[11px] font-medium uppercase tracking-wide text-slate-500">{_esc(spec.get("label", ""))}</p>
    <span class="mt-1 inline-block h-2 w-2 shrink-0 rounded-full {palette['dot']}"></span>
  </div>
  <p class="mt-1 truncate text-[26px] font-semibold leading-none tabular-nums text-slate-900">{_esc(spec.get("value", "—"))}</p>
  <p class="mt-1.5 truncate text-[11px] tabular-nums {palette['text']}">{arrow} {_esc(spec.get("delta", ""))}{badge}</p>
  <p class="mt-0.5 truncate text-[10px] text-slate-500">{_esc(spec.get("sub", ""))}</p>
</div>"""
    _frame(body, height)


def kpi_row(cards: Sequence[dict]) -> None:
    """Lay out cards in Streamlit columns; enable_mobile() stacks them on a phone."""
    if not cards:
        return
    cols = st.columns(len(cards), gap="small")
    for col, spec in zip(cols, cards):
        with col:
            kpi_card(spec)


def chart_card(title: str, fig, note: str | None = None, key: str | None = None) -> None:
    """Native Streamlit card around a Plotly figure (keeps hover, zoom, download)."""
    with st.container(border=True):
        st.markdown(f"**{title}**")
        if note:
            st.caption(note)
        if fig is None:
            st.info("No rows match the current filters.")
            return
        st.plotly_chart(fig, width="stretch", key=key,
                        config={"displayModeBar": False, "responsive": True})


def insight_box(text: str, level: str = "info", action: str | None = None,
                source: str | None = None) -> None:
    """The 'so what' box: one sentence, one action, one source. Max three per page."""
    if not USE_TAILWIND:
        {"warn": st.warning, "bad": st.error}.get(level, st.info)(text)
        return
    palette = STATUS.get(level, STATUS["neutral"])
    tail = " · ".join(x for x in (action, source) if x)
    body = f"""
<div class="rounded-lg border border-l-4 {palette['ring']} {palette['edge']} {palette['tint']} p-3">
  <p class="line-clamp-2 text-[13px] leading-snug text-slate-800">{_esc(text)}</p>
  <p class="mt-1 truncate text-[11px] text-slate-600">{_esc(tail)}</p>
</div>"""
    _frame(body, INSIGHT_H)


def filter_chips(chips: Iterable[str]) -> None:
    """Read-only summary of the active filters, above the KPI band (US-02 AC1)."""
    chips = list(chips)
    if not USE_TAILWIND:
        st.caption(" · ".join(chips))
        return
    if not chips:
        return
    items = "".join(
        f'<span class="rounded-full bg-slate-100 px-2.5 py-1 text-[11px] font-medium '
        f'text-slate-700 whitespace-nowrap">{_esc(c)}</span>' for c in chips
    )
    _frame(f'<div class="flex flex-wrap items-center gap-1.5">{items}</div>', CHIPS_H)


def empty_state(message: str, filters: Sequence[str] = ()) -> None:
    """Empty state that names the active filters instead of drawing zero cards."""
    st.info(f"No match: {message}")
    if not USE_TAILWIND:
        return
    chips = "".join(
        f'<span class="rounded-full bg-white px-2 py-0.5 text-[11px] text-slate-600 '
        f'ring-1 ring-slate-200 whitespace-nowrap">{_esc(f)}</span>' for f in filters
    )
    body = f"""
<div class="rounded-xl border border-dashed border-slate-300 bg-slate-50 p-5 text-center">
  <p class="text-sm font-medium text-slate-800">{_esc(message)}</p>
  <div class="mt-2 flex flex-wrap items-center justify-center gap-1.5">{chips}</div>
  <p class="mt-2 text-[11px] text-slate-500">Reset one filter to see data.</p>
</div>"""
    _frame(body, EMPTY_H)


def dq_banner(passed: int, failed: int, period: str) -> None:
    """Data-quality strip under the header on every KPI page (BR-14, US-17).

    No link: navigation lives in the sidebar, because an anchor inside the
    iframe would open the app inside the iframe (UIF-06).
    """
    if not USE_TAILWIND:
        (st.success if failed == 0 else st.warning)(
            f"{passed} of {passed + failed} data-quality checks pass for {period}.")
        return
    palette = STATUS["ok" if failed == 0 else "warn"]
    body = f"""
<div class="flex flex-wrap items-center gap-2 rounded-lg border {palette['ring']} {palette['tint']} px-3 py-2">
  <span class="text-[10px] font-semibold uppercase tracking-wide text-slate-700">Data quality</span>
  <span class="text-[12px] tabular-nums text-slate-700">{passed} of {passed + failed} checks pass</span>
  <span class="text-[11px] text-slate-500">· {_esc(period)} · full log in Data Quality</span>
</div>"""
    _frame(body, BANNER_H)
