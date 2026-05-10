"""
report_generator.py — pull a company's bs&pl + observations + benchmarks
from Airtable and render a self-contained HTML financial diagnostic report.

The output is one long HTML file the user can:
  - view in a browser
  - print to PDF (Cmd/Ctrl-P -> Save as PDF)
  - email or upload anywhere

No external CSS/JS or fonts. All charts are inline SVG.

Public API:
  build_report_for_company(company_id, fy_iso=None) -> {"html": str, "data": dict}

CLI:
  python modules/report_generator.py --company-id recXXX --out report.html
  python modules/report_generator.py --company-name "Aion Tech" --out report.html
"""
from __future__ import annotations

import argparse
import html
import json
import math
import sys
from pathlib import Path
from typing import Any, Optional

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent))

from observations import (
    METRICS, TBL_BSPL, TBL_OBS, TBL_COMPANY, TBL_BASE,
    at_list_all, at_find_one, get_financial_year, _creds,
)
from benchmarks import (
    DEFAULTS, status_for, zone_bounds, override_from_airtable,
    HEALTHY, WATCH, ALERT, NEUTRAL,
)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Data assembly
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

PILLARS = [
    ("Growth",        "growth",        ["Revenue Growth", "Profit Growth",
                                         "Gross Margin Expansion", "Net Margin Expansion"]),
    ("Profitability", "profitability", ["Gross Profit Margin", "Net Profit Margin",
                                         "Return on Total Assets", "Return on Equity"]),
    ("Cash & Liquidity", "liquidity",  ["Current Ratio", "Quick Ratio", "Cash Ratio",
                                         "Cash Burn",
                                         "Days of Receivable Outstanding",
                                         "Days of Inventory Outstanding",
                                         "Days of Payable Outstanding",
                                         "Working Capital Cycle",
                                         "Cash Conversion Cycle"]),
    ("Efficiency",    "turnover",      ["Receivables Turnover", "Inventory Turnover",
                                         "Payables Turnover", "Asset Turnover",
                                         "Equity Turnover"]),
    ("Risk & Stability", "leverage",   ["Operating Leverage", "Financial Leverage",
                                         "Total Leverage", "Equity Ratio",
                                         "Interest Coverage", "Break Even Sales",
                                         "Margin of Safety"]),
]


def _resolve_company(token: str, base_id: str,
                     company_id: Optional[str], company_name: Optional[str]
                     ) -> Optional[dict]:
    if company_id:
        from observations import at_find_one as _find
        rec = _find(TBL_COMPANY, f"RECORD_ID() = '{company_id}'", token, base_id)
        return rec
    if company_name:
        # Look up by company_name field
        safe = company_name.replace("'", "\\'")
        rec = at_find_one(TBL_COMPANY, f"{{company_name}} = '{safe}'", token, base_id)
        return rec
    return None


def _gather_data(company_id: str, fy_iso: Optional[str] = None) -> dict[str, Any]:
    token, _, base_id = _creds()

    # 1. company
    company_rec = _resolve_company(token, base_id, company_id, None)
    if not company_rec:
        raise RuntimeError(f"Company not found: {company_id}")
    company_name = company_rec.get("fields", {}).get("company_name", company_id)

    # 2. all bs&pl rows for this company
    bspl_all = at_list_all(TBL_BSPL, token, base_id)
    bspl_rows = [r for r in bspl_all
                 if company_rec["id"] in (r.get("fields", {}).get("company") or [])]
    if not bspl_rows:
        raise RuntimeError(f"No bs&pl records for {company_name}")
    # Sort by end_date ascending
    bspl_rows.sort(key=lambda r: r.get("fields", {}).get("end_date") or "")

    # 3. select the focus FY (latest if not specified)
    if fy_iso:
        focus = next((r for r in bspl_rows
                      if r.get("fields", {}).get("end_date") == fy_iso), None)
        if not focus:
            raise RuntimeError(f"No bs&pl record for FY ending {fy_iso}")
    else:
        focus = bspl_rows[-1]

    focus_id_map = {r["id"]: r for r in bspl_rows}
    focus_fields = focus.get("fields", {})
    focus_fy_iso, focus_fy_label = get_financial_year(focus_fields)

    # 4. observations row for this (company, fy)
    # Strict pass: company + financial-year exact match. This is the right
    # answer when our pipeline created the row.
    obs_rec = None
    obs_match_quality = "none"   # "strict" | "loose_recent" | "loose_legacy" | "none"
    if focus_fy_iso:
        formula = (f"AND(FIND('{company_rec['id']}', ARRAYJOIN({{company}})),"
                   f"IS_SAME({{financial year}}, '{focus_fy_iso}', 'day'))")
        obs_rec = at_find_one(TBL_OBS, formula, token, base_id)
        if obs_rec:
            obs_match_quality = "strict"

    # Fallback: 88 of 170 observation rows on the live base have no
    # 'financial year' set (legacy rows from before our pipeline). Without
    # a fallback the report shows "(not generated yet)" for every metric
    # even though the data IS there. Pull every obs row for this company,
    # prefer ones with non-empty Revenue Growth Obs, and pick the most
    # recent by created time.
    if obs_rec is None:
        obs_all = at_list_all(TBL_OBS, token, base_id)
        candidates = [r for r in obs_all
                      if company_rec["id"] in (r.get("fields", {}).get("company") or [])]
        # Sort: rows WITH content first, then by financial year desc (or created desc)
        def _sort_key(r):
            f = r.get("fields", {})
            has_content = bool((f.get("Revenue Growth Obs") or "").strip())
            fy = f.get("financial year") or ""
            created = r.get("createdTime") or ""
            return (has_content, fy, created)
        candidates.sort(key=_sort_key, reverse=True)
        for r in candidates:
            f = r.get("fields", {})
            if (f.get("Revenue Growth Obs") or "").strip():
                obs_rec = r
                obs_match_quality = ("loose_recent" if f.get("financial year")
                                     else "loose_legacy")
                break
    obs_fields = (obs_rec or {}).get("fields", {})

    # 5. base / industry benchmarks for this company
    base_all = at_list_all(TBL_BASE, token, base_id)
    base_rec = next((r for r in base_all
                     if company_rec["id"] in (r.get("fields", {}).get("company") or [])), None)
    base_fields = (base_rec or {}).get("fields", {})

    # 6. assemble per-metric dossier
    metrics_data = []
    for m in METRICS:
        # full historical series for the trend chart
        series = []
        for r in bspl_rows:
            f = r.get("fields", {})
            v = f.get(m.bspl_field)
            label = get_financial_year(f)[1]
            try:
                v_num = float(v) if v is not None else None
            except (TypeError, ValueError):
                v_num = None
            series.append({"fy": label, "value": v_num})

        curr_val = focus_fields.get(m.bspl_field)
        prior_val = None
        prev_link = focus_fields.get("previous_bs&pl") or []
        if prev_link:
            prev = focus_id_map.get(prev_link[0])
            if prev:
                prior_val = prev.get("fields", {}).get(m.bspl_field)
        try:
            curr_val_num = float(curr_val) if curr_val is not None else None
        except (TypeError, ValueError):
            curr_val_num = None
        try:
            prior_val_num = float(prior_val) if prior_val is not None else None
        except (TypeError, ValueError):
            prior_val_num = None

        # benchmark — per-company override if base table has it, else default
        override = None
        if m.base_field:
            override = override_from_airtable(m.name, base_fields.get(m.base_field))
        bm = override or DEFAULTS.get(m.name)
        status = status_for(m.name, curr_val_num, override=override)

        metrics_data.append({
            "name":       m.name,
            "category":   m.category,
            "unit":       m.unit,
            "guidance":   m.guidance,
            "direction":  m.direction,
            "current":    curr_val_num,
            "prior":      prior_val_num,
            "series":     series,
            "status":     status,
            "benchmark":  bm,
            "zones":      zone_bounds(m.name, override=override),
            "obs":        obs_fields.get(m.obs_field, ""),
            "recom":      obs_fields.get(m.recom_field, ""),
            "future":     obs_fields.get(m.future_field, ""),
        })

    # 7. health score + critical insights (from observations row)
    raw_score = obs_fields.get("health_score")
    try:
        # Stored as 0..1 fraction in Airtable per the pipeline; renormalize to 0..100
        score = int(round(float(raw_score) * 100)) if raw_score is not None else None
    except (TypeError, ValueError):
        score = None
    insights = [obs_fields.get(f"ci_{n}", "") for n in range(1, 6)]
    insights = [i for i in insights if i]

    obs_meta = {
        "match_quality": obs_match_quality,
        "obs_record_id": obs_rec.get("id") if obs_rec else None,
        "obs_fy_iso":    obs_fields.get("financial year"),
    }

    return {
        "company_name":   company_name,
        "company_id":     company_rec["id"],
        "fy_label":       focus_fy_label,
        "fy_iso":         focus_fy_iso,
        "all_periods":    [get_financial_year(r.get("fields", {}))[1] for r in bspl_rows],
        "metrics":        metrics_data,
        "pillars":        PILLARS,
        "health_score":   score,
        "critical_insights": insights,
        "obs_meta":       obs_meta,
        "generated_at":   __import__("datetime").datetime.now().isoformat(timespec="seconds"),
    }


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SVG chart helpers — pure, no deps
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _fmt(v: Optional[float], unit: str) -> str:
    if v is None:
        return "—"
    if unit == "%" or unit == "pp":
        return f"{v:+.2f}{unit}" if unit == "pp" else f"{v:.2f}%"
    if unit == "x":
        return f"{v:.2f}×"
    if unit == "days":
        return f"{int(round(v))}d"
    if unit == "months":
        return f"{v:.1f}mo"
    if unit == "INR":
        return f"₹{v:,.0f}"
    return f"{v:.2f}"


def _color_for_status(status: str) -> str:
    return {
        HEALTHY: "#3aa869",   # green
        WATCH:   "#d8a319",   # amber
        ALERT:   "#d35454",   # red
        NEUTRAL: "#a0a0a0",   # grey
    }.get(status, "#a0a0a0")


def _trend_svg(series: list[dict], width: int = 360, height: int = 120) -> str:
    """Mini line chart of historical values."""
    pts = [(s["fy"], s["value"]) for s in series if s["value"] is not None]
    if len(pts) < 1:
        return f'<svg width="{width}" height="{height}"></svg>'
    if len(pts) == 1:
        # Single dot
        x = width // 2
        y = height // 2
        return (f'<svg width="{width}" height="{height}" '
                f'xmlns="http://www.w3.org/2000/svg">'
                f'<circle cx="{x}" cy="{y}" r="4" fill="#1f3864"/>'
                f'<text x="{x}" y="{y - 10}" font-size="11" fill="#1f3864" '
                f'text-anchor="middle">{_fmt(pts[0][1], "")}</text>'
                f'<text x="{x}" y="{height - 6}" font-size="9" fill="#666" '
                f'text-anchor="middle">{html.escape(pts[0][0])}</text></svg>')

    vals = [p[1] for p in pts]
    vmin, vmax = min(vals), max(vals)
    if vmax == vmin:
        vmax = vmin + 1
    pad_x, pad_y = 30, 18
    inner_w = width - 2 * pad_x
    inner_h = height - 2 * pad_y

    def x_at(i): return pad_x + (i / max(1, len(pts) - 1)) * inner_w
    def y_at(v): return pad_y + (1 - (v - vmin) / (vmax - vmin)) * inner_h

    path = " ".join(
        ("M" if i == 0 else "L") + f"{x_at(i):.1f},{y_at(v):.1f}"
        for i, (_, v) in enumerate(pts)
    )
    dots = "".join(
        f'<circle cx="{x_at(i):.1f}" cy="{y_at(v):.1f}" r="3" fill="#1f3864"/>'
        for i, (_, v) in enumerate(pts)
    )
    labels_x = "".join(
        f'<text x="{x_at(i):.1f}" y="{height - 4}" font-size="9" fill="#666" '
        f'text-anchor="middle">{html.escape(fy)}</text>'
        for i, (fy, _) in enumerate(pts)
    )
    # Last value label
    last_x, last_v = x_at(len(pts) - 1), pts[-1][1]
    last_y = y_at(last_v)
    val_label = (
        f'<text x="{last_x:.1f}" y="{last_y - 6:.1f}" font-size="10" '
        f'fill="#1f3864" font-weight="600" text-anchor="end">{_fmt(last_v, "")}</text>'
    )
    return (
        f'<svg width="{width}" height="{height}" xmlns="http://www.w3.org/2000/svg">'
        f'<path d="{path}" stroke="#1f3864" stroke-width="1.8" fill="none"/>'
        f'{dots}{labels_x}{val_label}</svg>'
    )


def _zone_bar_svg(metric: dict, width: int = 360, height: int = 56) -> str:
    """Horizontal benchmark zone bar with current value markers."""
    bm = metric["benchmark"]
    if not bm:
        return f'<svg width="{width}" height="{height}"></svg>'

    # Determine the numeric axis domain
    candidates = []
    for v in [bm.healthy_min, bm.healthy_max, bm.watch_min, bm.watch_max,
              metric["current"], metric["prior"]]:
        if v is not None:
            try:
                candidates.append(float(v))
            except (TypeError, ValueError):
                pass
    if not candidates:
        return f'<svg width="{width}" height="{height}"></svg>'
    lo = min(candidates)
    hi = max(candidates)
    if hi == lo:
        hi = lo + 1
    # Pad the domain a bit so markers aren't at the very edge
    span = hi - lo
    lo -= span * 0.1
    hi += span * 0.1

    pad_x = 12
    bar_y = 22
    bar_h = 14
    inner_w = width - 2 * pad_x

    def x_at(v): return pad_x + ((v - lo) / (hi - lo)) * inner_w

    bands = []
    if bm.direction == "higher":
        # alert: lo..watch_min ; watch: watch_min..healthy_min ; healthy: healthy_min..hi
        if bm.watch_min is not None:
            bands.append(("alert", lo, bm.watch_min))
            bands.append(("watch", bm.watch_min, bm.healthy_min if bm.healthy_min is not None else hi))
        if bm.healthy_min is not None:
            bands.append(("healthy", bm.healthy_min, hi))
    elif bm.direction == "lower":
        if bm.healthy_max is not None:
            bands.append(("healthy", lo, bm.healthy_max))
            bands.append(("watch", bm.healthy_max, bm.watch_max if bm.watch_max is not None else hi))
        if bm.watch_max is not None:
            bands.append(("alert", bm.watch_max, hi))
    elif bm.direction == "band":
        if bm.watch_min is not None and bm.healthy_min is not None:
            bands.append(("alert",   lo, bm.watch_min))
            bands.append(("watch",   bm.watch_min, bm.healthy_min))
            bands.append(("healthy", bm.healthy_min, bm.healthy_max if bm.healthy_max is not None else hi))
        if bm.healthy_max is not None and bm.watch_max is not None:
            bands.append(("watch", bm.healthy_max, bm.watch_max))
            bands.append(("alert", bm.watch_max, hi))

    band_colors = {"healthy": "#c5e8c8", "watch": "#fde2a0", "alert": "#f5c6c0"}
    band_svg = "".join(
        f'<rect x="{x_at(s):.1f}" y="{bar_y}" width="{(x_at(e) - x_at(s)):.1f}" '
        f'height="{bar_h}" fill="{band_colors[name]}" />'
        for name, s, e in bands if s is not None and e is not None and e > s
    )

    # Current + prior markers
    markers = ""
    cur = metric["current"]
    prior = metric["prior"]
    if cur is not None:
        cx = x_at(cur)
        markers += (
            f'<circle cx="{cx:.1f}" cy="{bar_y + bar_h/2}" r="5" fill="#1f3864" '
            f'stroke="white" stroke-width="2"/>'
            f'<text x="{cx:.1f}" y="{bar_y - 4}" font-size="10" '
            f'fill="#1f3864" font-weight="700" text-anchor="middle">'
            f'{html.escape(_fmt(cur, metric["unit"]))}</text>'
        )
    if prior is not None and prior != cur:
        px = x_at(prior)
        markers += (
            f'<circle cx="{px:.1f}" cy="{bar_y + bar_h/2}" r="3.5" '
            f'fill="white" stroke="#888" stroke-width="1.5"/>'
        )

    # Axis labels
    axis = (
        f'<text x="{pad_x}" y="{height - 4}" font-size="9" fill="#888">'
        f'{html.escape(_fmt(lo, metric["unit"]))}</text>'
        f'<text x="{width - pad_x}" y="{height - 4}" font-size="9" fill="#888" '
        f'text-anchor="end">{html.escape(_fmt(hi, metric["unit"]))}</text>'
    )

    return (f'<svg width="{width}" height="{height}" xmlns="http://www.w3.org/2000/svg">'
            f'{band_svg}{markers}{axis}</svg>')


def _gauge_svg(score: int, size: int = 96) -> str:
    """Circular gauge for the 0-100 health score. Pure SVG, no JS animation
    on the print path so it renders identically to the screen view."""
    if score is None:
        return ""
    s = max(0, min(100, int(score)))
    r = size * 0.42
    cx = cy = size / 2
    c = 2 * 3.14159 * r
    offset = c * (1 - s / 100)
    color = "#3aa869" if s >= 70 else ("#d8a319" if s >= 45 else "#d35454")
    return (
        f'<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}" '
        f'xmlns="http://www.w3.org/2000/svg">'
        f'<circle cx="{cx}" cy="{cy}" r="{r}" stroke="#EEF2FF" stroke-width="8" fill="none"/>'
        f'<circle cx="{cx}" cy="{cy}" r="{r}" stroke="{color}" stroke-width="8" '
        f'fill="none" stroke-linecap="round" '
        f'stroke-dasharray="{c:.1f}" stroke-dashoffset="{offset:.1f}" '
        f'transform="rotate(-90 {cx} {cy})"/>'
        f'</svg>'
    )


def _mindmap_svg(pillars: list, by_name: dict, width: int = 920, height: int = 460) -> str:
    """Mind-map view of the 29-metric framework: 5 pillar hubs in a ring, with
    member metrics radiating out. Status of each metric colors the node.
    Cross-pillar relationships drawn as faint indigo curves (e.g. CCC = DSO + DIO - DPO).
    """
    cx, cy = width / 2, height / 2
    import math
    pillar_radius = 150     # distance from center to pillar hub
    metric_radius = 90      # distance from pillar hub to its metrics
    n_pillars = len(pillars)

    # Position each pillar hub on a circle around the center
    pillar_pos = {}
    for i, (label, key, names) in enumerate(pillars):
        angle = -math.pi / 2 + (2 * math.pi * i / n_pillars)
        pillar_pos[label] = (cx + pillar_radius * math.cos(angle),
                             cy + pillar_radius * math.sin(angle), angle)

    # Status colors
    s_color = {"healthy": "#3aa869", "watch": "#d8a319",
               "alert": "#d35454", "neutral": "#a0a0a0"}

    # Place metric nodes radially around their pillar hub
    metric_pos = {}
    spokes = []
    for label, key, names in pillars:
        hx, hy, hang = pillar_pos[label]
        n = len(names)
        # Spread metrics over a ~140° arc facing outward from the center
        spread = math.radians(140)
        start = hang - spread / 2
        for j, mname in enumerate(names):
            t = j / max(1, n - 1) if n > 1 else 0.5
            ang = start + spread * t
            mx = hx + metric_radius * math.cos(ang)
            my = hy + metric_radius * math.sin(ang)
            metric_pos[mname] = (mx, my)
            spokes.append((hx, hy, mx, my))

    # Pre-defined cross-pillar relationships (the "wow" — show how metrics interrelate)
    relationships = [
        ("Cash Conversion Cycle",       "Days of Receivable Outstanding"),
        ("Cash Conversion Cycle",       "Days of Inventory Outstanding"),
        ("Cash Conversion Cycle",       "Days of Payable Outstanding"),
        ("Working Capital Cycle",       "Days of Receivable Outstanding"),
        ("Working Capital Cycle",       "Days of Inventory Outstanding"),
        ("Total Leverage",              "Operating Leverage"),
        ("Total Leverage",              "Financial Leverage"),
        ("Return on Equity",            "Net Profit Margin"),
        ("Return on Equity",            "Asset Turnover"),
        ("Return on Total Assets",      "Asset Turnover"),
        ("Profit Growth",               "Revenue Growth"),
        ("Margin of Safety",            "Break Even Sales"),
        ("Net Margin Expansion",        "Gross Margin Expansion"),
    ]

    # Spokes from pillar hub -> metric (faint cream)
    spoke_svg = "".join(
        f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" '
        f'stroke="#e0d8c8" stroke-width="1"/>'
        for x1, y1, x2, y2 in spokes
    )

    # Cross-relationship curves (indigo, low alpha, dashed)
    rel_svg = ""
    for a, b in relationships:
        if a not in metric_pos or b not in metric_pos:
            continue
        ax, ay = metric_pos[a]; bx, by = metric_pos[b]
        # Quadratic curve via the chart center for organic feel
        rel_svg += (f'<path d="M{ax:.1f},{ay:.1f} Q{cx},{cy} {bx:.1f},{by:.1f}" '
                    f'stroke="#6366F1" stroke-width="1" stroke-dasharray="3 3" '
                    f'opacity="0.35" fill="none"/>')

    # Pillar hubs
    pillar_svg = ""
    for i, (label, key, names) in enumerate(pillars):
        x, y, _ = pillar_pos[label]
        # Status counts to color the hub ring
        counts = {"healthy": 0, "watch": 0, "alert": 0, "neutral": 0}
        for n in names:
            m = by_name.get(n)
            if m: counts[m["status"]] = counts.get(m["status"], 0) + 1
        ring_color = ("#d35454" if counts["alert"] >= 2
                      else ("#d8a319" if (counts["alert"] + counts["watch"]) >= 2
                            else "#3aa869"))
        # Two-line label split at first space
        pieces = label.split(" ", 1)
        line1 = pieces[0]; line2 = pieces[1] if len(pieces) > 1 else ""
        pillar_svg += (
            f'<g class="pillar-node">'
            f'<circle cx="{x:.1f}" cy="{y:.1f}" r="26" fill="white" stroke="{ring_color}" stroke-width="2.5"/>'
            f'<circle cx="{x:.1f}" cy="{y:.1f}" r="22" fill="#EEF2FF" opacity="0.7"/>'
            f'<text x="{x:.1f}" y="{y - 1:.1f}" text-anchor="middle" font-size="10" '
            f'font-weight="700" fill="#14213d">{html.escape(line1)}</text>'
            f'<text x="{x:.1f}" y="{y + 11:.1f}" text-anchor="middle" font-size="9" '
            f'fill="#6b6660">{html.escape(line2)}</text>'
            f'</g>'
        )

    # Metric nodes
    metric_svg = ""
    for label, key, names in pillars:
        for n in names:
            if n not in metric_pos: continue
            mx, my = metric_pos[n]
            m = by_name.get(n)
            color = s_color.get((m or {}).get("status", "neutral"), "#a0a0a0")
            # Position label outward from center
            dx, dy = mx - cx, my - cy
            d = (dx * dx + dy * dy) ** 0.5 or 1
            tx = mx + (dx / d) * 12
            ty = my + (dy / d) * 12
            anchor = "start" if dx > 4 else ("end" if dx < -4 else "middle")
            short = n if len(n) <= 22 else n[:20] + "…"
            metric_svg += (
                f'<g class="metric-node">'
                f'<title>{html.escape(n)} — {(m or {}).get("status","")}</title>'
                f'<circle cx="{mx:.1f}" cy="{my:.1f}" r="5" fill="{color}" '
                f'stroke="white" stroke-width="1.5"/>'
                f'<text x="{tx:.1f}" y="{ty:.1f}" font-size="9.5" fill="#1f1f1f" '
                f'text-anchor="{anchor}" alignment-baseline="middle">{html.escape(short)}</text>'
                f'</g>'
            )

    # Center node — overall framework label
    center_svg = (
        f'<circle cx="{cx}" cy="{cy}" r="36" fill="white" stroke="#4F46E5" stroke-width="2"/>'
        f'<circle cx="{cx}" cy="{cy}" r="32" fill="#F5F3FF"/>'
        f'<text x="{cx}" y="{cy - 4}" text-anchor="middle" font-size="11" font-weight="700" '
        f'fill="#4F46E5">29-METRIC</text>'
        f'<text x="{cx}" y="{cy + 9}" text-anchor="middle" font-size="11" font-weight="700" '
        f'fill="#4F46E5">FRAMEWORK</text>'
    )

    return (f'<svg width="100%" viewBox="0 0 {width} {height}" '
            f'xmlns="http://www.w3.org/2000/svg" style="display:block;">'
            f'{rel_svg}{spoke_svg}{center_svg}{pillar_svg}{metric_svg}'
            f'</svg>')


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# HTML rendering
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

CSS = """
:root {
  --brand:        #4F46E5;   /* AskMyCFO indigo */
  --brand-bright: #6366F1;
  --brand-soft:   #EEF2FF;
  --brand-pale:   #F5F3FF;
  --ink:          #14213d;   /* navy headings (kept) */
  --ink-soft:     #1f3864;
  --paper:        #faf7f2;   /* cream base (kept) */
  --paper-warm:   #f8f4ec;
  --rule:         #e0d8c8;
  --muted:        #6b6660;
  --green:        #3aa869;
  --green-soft:   #c5e8c8;
  --amber:        #d8a319;
  --amber-soft:   #fde2a0;
  --red:          #d35454;
  --red-soft:     #f5c6c0;
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body { font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
       background: var(--paper); color: #1f1f1f; margin: 0; padding: 0; line-height: 1.5; }
.container { max-width: 980px; margin: 0 auto; padding: 32px 24px 80px; }
h1 { font-family: Georgia, "Times New Roman", serif; font-size: 36px; margin: 0 0 6px; color: var(--ink); letter-spacing: -0.4px; }
h2 { font-family: Georgia, serif; font-size: 22px; margin: 28px 0 12px; color: var(--ink); border-bottom: 1px solid var(--rule); padding-bottom: 6px; }
h3 { font-size: 15px; margin: 18px 0 6px; color: var(--ink); font-weight: 600; }
.sub { color: var(--muted); font-style: italic; margin-bottom: 22px; }
.tag { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 10.5px;
       font-weight: 700; letter-spacing: 0.5px; text-transform: uppercase; }
.tag.healthy { background: var(--green-soft); color: #1d6a30; }
.tag.watch   { background: var(--amber-soft); color: #7a5300; }
.tag.alert   { background: var(--red-soft); color: #9a2222; }
.tag.neutral { background: #e6e0d6; color: #555; }

/* Reveal-on-scroll: cards fade up as they enter the viewport */
.reveal { opacity: 0; transform: translateY(14px); transition: opacity 0.5s ease, transform 0.5s ease; }
.reveal.in { opacity: 1; transform: translateY(0); }

.tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
         gap: 12px; margin-bottom: 24px; }
.tile { background: white; border: 1px solid var(--rule); border-radius: 10px;
        padding: 14px 16px; transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease; }
.tile:hover { transform: translateY(-2px); box-shadow: 0 8px 24px -8px rgba(79, 70, 229, 0.18);
              border-color: rgba(79, 70, 229, 0.35); }
.tile .label { font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.6px;
               color: var(--muted); }
.tile .value { font-size: 24px; font-weight: 700; color: var(--ink); margin-top: 4px; }
.tile .delta { font-size: 11px; color: var(--muted); margin-top: 2px; }

/* Hero score-card: indigo gradient + circular gauge */
.score-card { display: flex; gap: 22px; align-items: center;
              background: linear-gradient(135deg, var(--brand-soft), white);
              border: 1px solid rgba(79, 70, 229, 0.18);
              border-radius: 14px; padding: 22px 26px; margin: 14px 0 26px;
              box-shadow: 0 4px 24px -10px rgba(79, 70, 229, 0.18); }
.score-gauge { position: relative; width: 96px; height: 96px; flex-shrink: 0; }
.score-gauge .num { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center;
                    font-family: Georgia, serif; font-size: 28px; font-weight: 700; color: var(--ink); }
.score-gauge .num small { font-size: 11px; color: var(--muted); margin-left: 2px; font-weight: 500; }
.score-text { color: var(--ink); }
.score-text strong { font-size: 15px; }
.score-text .hint { font-size: 12.5px; color: var(--muted); margin-top: 2px; display: block; }

.pillar-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
               gap: 10px; margin: 14px 0 28px; }
.pillar { background: white; border: 1px solid var(--rule); border-radius: 10px; padding: 14px;
          transition: transform 0.2s ease, box-shadow 0.2s ease; }
.pillar:hover { transform: translateY(-2px); box-shadow: 0 8px 18px -10px rgba(20, 33, 61, 0.2); }
.pillar .name { font-size: 12.5px; font-weight: 600; color: var(--ink); }
.pillar .summary { font-size: 11px; color: var(--muted); margin-top: 4px; }
.pillar .bars { display: flex; gap: 2px; margin-top: 8px; height: 6px; border-radius: 3px; overflow: hidden; }
.pillar .bars .seg { flex: 1; }

table.master { width: 100%; border-collapse: collapse; margin: 12px 0 28px; font-size: 12.5px;
               background: white; border-radius: 10px; overflow: hidden;
               box-shadow: 0 1px 0 var(--rule); }
table.master th { text-align: left; font-size: 10px; text-transform: uppercase;
                  letter-spacing: 0.5px; color: var(--muted); padding: 10px;
                  border-bottom: 2px solid var(--rule); }
table.master td { padding: 9px 10px; border-bottom: 1px solid #eee5d4; }
table.master tr:hover td { background: var(--brand-pale); }
table.master td.metric { font-weight: 500; }
table.master td.value { font-variant-numeric: tabular-nums; text-align: right; font-weight: 600; }
table.master tr.section td { background: linear-gradient(90deg, var(--brand-pale), transparent);
                              font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.6px;
                              color: var(--brand); padding: 7px 10px; font-weight: 700; }
table.master tr.section:hover td { background: linear-gradient(90deg, var(--brand-pale), transparent); }

.metric-card { background: white; border: 1px solid var(--rule); border-radius: 10px;
               padding: 18px 20px; margin: 10px 0 14px;
               transition: transform 0.2s ease, box-shadow 0.25s ease, border-color 0.2s ease; }
.metric-card:hover { transform: translateY(-2px); border-color: rgba(79, 70, 229, 0.3);
                     box-shadow: 0 10px 32px -16px rgba(79, 70, 229, 0.22); }
.metric-head { display: flex; justify-content: space-between; align-items: center;
               gap: 10px; border-bottom: 1px solid var(--rule); padding-bottom: 9px; margin-bottom: 10px; }
.metric-head .title { font-size: 16px; font-weight: 600; color: var(--ink); flex: 1; }
.metric-head .meta { display: flex; gap: 10px; align-items: center; }
.metric-head .value { font-weight: 700; color: var(--ink); }

/* Explain button — hooks into Web Speech API */
.explain-btn { display: inline-flex; align-items: center; gap: 5px;
               padding: 4px 10px; font-size: 11px; font-weight: 600;
               color: var(--brand); background: var(--brand-soft);
               border: 1px solid rgba(79, 70, 229, 0.2);
               border-radius: 14px; cursor: pointer;
               transition: background 0.15s ease, transform 0.15s ease; }
.explain-btn:hover { background: white; transform: scale(1.04); }
.explain-btn.playing { background: var(--brand); color: white; border-color: var(--brand); }
.explain-btn .ico { width: 11px; height: 11px; display: inline-block; }

.metric-charts { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; align-items: center;
                 margin: 8px 0 14px; }
.metric-charts .chart-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.6px;
                              color: var(--muted); text-align: center; margin-bottom: 6px; font-weight: 600; }
.guidance { font-size: 12.5px; color: var(--muted); font-style: italic; padding: 9px 12px;
            background: var(--brand-pale); border-left: 3px solid var(--brand); border-radius: 4px;
            margin: 6px 0 12px; }
.commentary { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; font-size: 12.5px; }
.commentary .block { background: var(--paper-warm); padding: 11px 13px; border-radius: 6px;
                     border-top: 2px solid transparent; transition: border-color 0.2s ease; }
.commentary .block:hover { border-top-color: var(--brand); }
.commentary .block .heading { font-size: 10px; text-transform: uppercase; letter-spacing: 0.5px;
                              color: var(--brand); font-weight: 700; margin-bottom: 4px; }
.commentary .block.empty { color: #aaa; font-style: italic; }

.insights { margin: 12px 0 28px; background: white; border: 1px solid var(--rule);
            border-radius: 10px; padding: 18px 22px; }
.insights ol { margin: 0; padding-left: 22px; }
.insights li { margin: 8px 0; padding-left: 4px; }
.insights li::marker { color: var(--brand); font-weight: 700; }

/* Mind Map — interactive metric-relationship graph */
.mindmap-wrap { background: white; border: 1px solid var(--rule); border-radius: 10px;
                padding: 18px; margin: 8px 0 24px; }
.mindmap-wrap .legend { display: flex; gap: 14px; font-size: 11px; color: var(--muted);
                        margin-top: 6px; flex-wrap: wrap; }
.mindmap-wrap .legend span { display: inline-flex; align-items: center; gap: 5px; }
.mindmap-wrap .legend .swatch { width: 10px; height: 10px; border-radius: 50%; display: inline-block; }
.mindmap-wrap svg .pillar-node circle { transition: r 0.2s ease, stroke-width 0.2s ease; }
.mindmap-wrap svg .pillar-node:hover circle { r: 30; stroke-width: 3; }
.mindmap-wrap svg .metric-node circle { transition: r 0.2s ease; }
.mindmap-wrap svg .metric-node:hover circle { r: 7; }

.footer { font-size: 10.5px; color: #888; text-align: center; margin-top: 60px;
          padding-top: 16px; border-top: 1px solid var(--rule); }

@media print {
  body { background: white; }
  .reveal { opacity: 1 !important; transform: none !important; }
  .metric-card, .mindmap-wrap, .insights, .pillar { break-inside: avoid; }
  .explain-btn { display: none; }
}
"""


def _render(data: dict) -> str:
    company = data["company_name"]
    fy = data["fy_label"]
    score = data["health_score"]
    metrics = data["metrics"]
    pillars = data["pillars"]
    insights = data["critical_insights"]

    by_name = {m["name"]: m for m in metrics}

    # Header tiles — pick the most-quoted metrics
    tile_names = ["Revenue Growth", "Net Profit Margin", "Return on Equity",
                  "Current Ratio", "Cash Conversion Cycle", "Interest Coverage"]
    tiles_html = []
    for name in tile_names:
        m = by_name.get(name)
        if not m:
            continue
        delta = ""
        if m["current"] is not None and m["prior"] is not None:
            d = m["current"] - m["prior"]
            arrow = "▲" if d > 0 else ("▼" if d < 0 else "≈")
            delta = f"{arrow} {_fmt(abs(d), m['unit'])} vs prior"
        tiles_html.append(
            f'<div class="tile"><div class="label">{html.escape(name)}</div>'
            f'<div class="value">{html.escape(_fmt(m["current"], m["unit"]))}</div>'
            f'<div class="delta">{html.escape(delta)}</div></div>'
        )

    # Pillar summaries — count statuses per pillar
    pillar_html = []
    for label, key, names in pillars:
        members = [by_name[n] for n in names if n in by_name]
        counts = {HEALTHY: 0, WATCH: 0, ALERT: 0, NEUTRAL: 0}
        for m in members:
            counts[m["status"]] = counts.get(m["status"], 0) + 1
        bars = "".join(
            f'<div class="seg" style="background:{_color_for_status(s)};'
            f'flex:{counts[s] or 0.001};"></div>'
            for s in [HEALTHY, WATCH, ALERT]
        )
        verdict = (HEALTHY if counts[ALERT] == 0 and counts[WATCH] <= 1
                   else (ALERT if counts[ALERT] >= 2 else WATCH))
        pillar_html.append(
            f'<div class="pillar"><div class="name">{html.escape(label)} '
            f'<span class="tag {verdict}">{verdict.upper()}</span></div>'
            f'<div class="bars">{bars}</div>'
            f'<div class="summary">{counts[HEALTHY]} healthy · {counts[WATCH]} watch · '
            f'{counts[ALERT]} alert</div></div>'
        )

    # Master table — every metric in its pillar
    master_rows = []
    for label, key, names in pillars:
        master_rows.append(
            f'<tr class="section"><td colspan="4">{html.escape(label)}</td></tr>'
        )
        for name in names:
            m = by_name.get(name)
            if not m:
                continue
            cur = _fmt(m["current"], m["unit"])
            prior = _fmt(m["prior"], m["unit"])
            master_rows.append(
                f'<tr><td class="metric">{html.escape(m["name"])}</td>'
                f'<td class="value">{html.escape(cur)}</td>'
                f'<td class="value" style="color:#888">{html.escape(prior)}</td>'
                f'<td><span class="tag {m["status"]}">{m["status"].upper()}</span></td></tr>'
            )

    # Per-metric detail cards (grouped by pillar)
    detail_html = []
    for label, key, names in pillars:
        detail_html.append(f'<h2>{html.escape(label)}</h2>')
        for name in names:
            m = by_name.get(name)
            if not m:
                continue
            cur = _fmt(m["current"], m["unit"])
            obs = m["obs"] or ""
            recom = m["recom"] or ""
            future = m["future"] or ""
            blocks = []
            for heading, text in [("Observation", obs),
                                  ("Recommendation", recom),
                                  ("Future plan", future)]:
                if text.strip():
                    blocks.append(
                        f'<div class="block"><div class="heading">{heading}</div>'
                        f'{html.escape(text)}</div>'
                    )
                else:
                    blocks.append(
                        f'<div class="block empty"><div class="heading">{heading}</div>'
                        f'(not generated yet)</div>'
                    )

            # Build the speakable script: Observation + Recommendation joined.
            # Stored on a data-attr so the JS Explain handler picks it up
            # without needing per-card script tags. Strip newlines and quotes
            # for safe attribute embedding.
            speak_lines = []
            speak_lines.append(f"{m['name']}.")
            cur_spoken = _fmt(m["current"], m["unit"]) if m["current"] is not None else "no value"
            speak_lines.append(f"Current value: {cur_spoken}.")
            if m.get("status") and m["status"] != "neutral":
                speak_lines.append(f"Status: {m['status']}.")
            if obs.strip():
                speak_lines.append(obs)
            if recom.strip():
                speak_lines.append(f"Recommendation: {recom}")
            speak_text = " ".join(speak_lines).replace("\n", " ").replace('"', "'")[:1400]

            detail_html.append(
                f'<div class="metric-card reveal">'
                f'<div class="metric-head">'
                f'<div class="title">{html.escape(m["name"])}</div>'
                f'<div class="meta">'
                f'<button class="explain-btn" data-speak="{html.escape(speak_text)}" '
                f'onclick="explainMetric(this)">'
                f'<svg class="ico" viewBox="0 0 24 24" fill="currentColor">'
                f'<path d="M3 9v6h4l5 5V4L7 9H3zm13.5 3a4.5 4.5 0 0 0-2.5-4v8a4.5 4.5 0 0 0 2.5-4z"/></svg>'
                f'Explain</button>'
                f'<span class="tag {m["status"]}">{m["status"].upper()}</span> '
                f'<span class="value">{html.escape(cur)}</span>'
                f'</div></div>'
                f'<div class="guidance">{html.escape(m["guidance"])}</div>'
                f'<div class="metric-charts">'
                f'<div><div class="chart-label">Trend</div>{_trend_svg(m["series"])}</div>'
                f'<div><div class="chart-label">vs benchmark</div>'
                f'{_zone_bar_svg(m)}</div>'
                f'</div>'
                f'<div class="commentary">{"".join(blocks)}</div>'
                f'</div>'
            )

    insights_html = ""
    if insights:
        items = "".join(f"<li>{html.escape(i)}</li>" for i in insights)
        insights_html = (
            f'<div class="insights"><h2>Critical insights</h2>'
            f'<ol>{items}</ol></div>'
        )

    score_html = ""
    if score is not None:
        score_html = (
            f'<div class="score-card reveal">'
            f'<div class="score-gauge">{_gauge_svg(score)}'
            f'<div class="num">{score}<small>/100</small></div></div>'
            f'<div class="score-text"><strong>Overall financial health</strong>'
            f'<span class="hint">Generated from the 29-metric snapshot. Benchmark-aware.</span></div></div>'
        )

    # Banner that explains when observations are missing or matched loosely.
    # This is the difference between "(not generated yet)" appearing 29 times
    # for no apparent reason vs. an actionable hint.
    obs_meta = data.get("obs_meta") or {}
    quality = obs_meta.get("match_quality", "none")
    obs_url = ""
    if obs_meta.get("obs_record_id"):
        obs_url = f"https://airtable.com/{data.get('company_id','')[:0]}".rstrip()  # placeholder
    banner_html = ""
    if quality == "none":
        banner_html = (
            '<div style="background:#fde2a0;color:#7a5300;padding:12px 16px;'
            'border-radius:6px;margin:14px 0 8px;font-size:12.5px;">'
            '<strong>Observations not yet generated for this company.</strong> '
            'Click <code>⚙ Run observations</code> in the Diagnostic Report '
            'panel on the dashboard, then reload this page. The pipeline '
            'will read the bs&amp;pl rows for this company, call GPT-4o-mini '
            'for each of the 29 metrics, and write the output to the '
            '<code>observation_and_recommendation</code> table on Airtable.'
            '</div>'
        )
    elif quality in ("loose_recent", "loose_legacy"):
        obs_fy = obs_meta.get("obs_fy_iso") or "(no financial year set)"
        banner_html = (
            '<div style="background:#cfe2ff;color:#0a3a82;padding:12px 16px;'
            'border-radius:6px;margin:14px 0 8px;font-size:12.5px;">'
            f'<strong>Note:</strong> No observation row matched this exact '
            f'financial year ({html.escape(data.get("fy_label", ""))}). '
            f'Showing observations from the most recent row with content '
            f'(financial year on that row: {html.escape(str(obs_fy))}). '
            f'Re-run observations for this company to refresh against the '
            f'current FY.'
            '</div>'
        )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>{html.escape(company)} — Financial Diagnostic ({html.escape(fy)})</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>{CSS}</style>
</head>
<body>
<div class="container">
  <h1>{html.escape(company)}</h1>
  <p class="sub">Financial Diagnostic Report · {html.escape(fy)} · benchmark-aware</p>
  {banner_html}
  {score_html}
  <h2>Executive snapshot</h2>
  <div class="tiles reveal">{''.join(tiles_html)}</div>
  <h2>Health by pillar</h2>
  <div class="pillar-grid reveal">{''.join(pillar_html)}</div>
  {insights_html}
  <h2>Mind map — how the 29 metrics relate</h2>
  <div class="mindmap-wrap reveal">
    {_mindmap_svg(pillars, by_name)}
    <div class="legend">
      <span><span class="swatch" style="background:#3aa869;"></span>Healthy</span>
      <span><span class="swatch" style="background:#d8a319;"></span>Watch</span>
      <span><span class="swatch" style="background:#d35454;"></span>Alert</span>
      <span><span class="swatch" style="background:#a0a0a0;"></span>Neutral / no value</span>
      <span style="margin-left:auto; font-size:10.5px;"><i style="color:#6366F1;">— — —</i> &nbsp;cross-pillar relationship (e.g. Cash Conversion Cycle = DSO + DIO − DPO)</span>
    </div>
  </div>
  <h2>Master table — every metric, every period</h2>
  <table class="master reveal">
    <thead><tr><th>Metric</th><th style="text-align:right">Current</th>
    <th style="text-align:right">Prior</th><th>Status</th></tr></thead>
    <tbody>{''.join(master_rows)}</tbody>
  </table>
  <h2>Detail by pillar</h2>
  {''.join(detail_html)}
  <div class="footer">
    Generated {html.escape(data["generated_at"])} · AMC_RAG · {html.escape(company)}<br>
    Status zones derived from industry defaults; per-company overrides applied where present.
  </div>
</div>
<script>
// ── Reveal-on-scroll: fade up cards as they enter the viewport ──
(function() {{
  if (!('IntersectionObserver' in window)) {{
    document.querySelectorAll('.reveal').forEach(function (el) {{ el.classList.add('in'); }});
    return;
  }}
  var io = new IntersectionObserver(function (entries) {{
    entries.forEach(function (e) {{
      if (e.isIntersecting) {{ e.target.classList.add('in'); io.unobserve(e.target); }}
    }});
  }}, {{ threshold: 0.08, rootMargin: '0px 0px -40px 0px' }});
  document.querySelectorAll('.reveal').forEach(function (el) {{ io.observe(el); }});
}})();

// ── Explain button: reads observation + recommendation aloud via Web Speech API ──
// No external service. Browser's built-in voices. Click again to stop.
window.__explainCurrent = null;
function explainMetric(btn) {{
  var synth = window.speechSynthesis;
  if (!synth) {{ alert('Your browser does not support speech synthesis.'); return; }}

  // If THIS button is already playing, treat click as stop.
  if (window.__explainCurrent === btn) {{
    synth.cancel();
    btn.classList.remove('playing');
    btn.lastChild.nodeValue = 'Explain';
    window.__explainCurrent = null;
    return;
  }}
  // If another button is playing, stop it first.
  if (window.__explainCurrent) {{
    synth.cancel();
    window.__explainCurrent.classList.remove('playing');
    window.__explainCurrent.lastChild.nodeValue = 'Explain';
  }}

  var text = btn.getAttribute('data-speak') || '';
  if (!text) return;
  var u = new SpeechSynthesisUtterance(text);
  u.rate = 1.0; u.pitch = 1.0; u.volume = 1.0;

  // Prefer an Indian English voice if available, else any English voice.
  function pickVoice() {{
    var voices = synth.getVoices() || [];
    return voices.find(function (v) {{ return /en-IN/i.test(v.lang); }})
        || voices.find(function (v) {{ return /^en/i.test(v.lang); }})
        || voices[0];
  }}
  var v = pickVoice();
  if (v) u.voice = v;

  u.onstart = function () {{
    btn.classList.add('playing');
    btn.lastChild.nodeValue = 'Stop';
    window.__explainCurrent = btn;
  }};
  u.onend = u.onerror = function () {{
    btn.classList.remove('playing');
    btn.lastChild.nodeValue = 'Explain';
    if (window.__explainCurrent === btn) window.__explainCurrent = null;
  }};

  // Voices may load async on Chrome — wait once if needed
  if (!synth.getVoices().length && synth.onvoiceschanged === null) {{
    synth.onvoiceschanged = function () {{ var nv = pickVoice(); if (nv) u.voice = nv; synth.speak(u); }};
  }} else {{
    synth.speak(u);
  }}
}}
</script>
</body>
</html>
"""


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Public API
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def build_report_for_company(company_id: Optional[str] = None,
                              company_name: Optional[str] = None,
                              fy_iso: Optional[str] = None) -> dict[str, Any]:
    """Build an HTML report for a company. Either company_id or company_name
    must be provided."""
    token, _, base_id = _creds()
    if not company_id and company_name:
        rec = _resolve_company(token, base_id, None, company_name)
        if not rec:
            raise RuntimeError(f"Company '{company_name}' not found")
        company_id = rec["id"]
    if not company_id:
        raise ValueError("company_id or company_name required")

    data = _gather_data(company_id, fy_iso=fy_iso)
    html_str = _render(data)
    return {"html": html_str, "data": data}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CLI
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--company-id",   default="", help="Airtable record id (recXXX)")
    ap.add_argument("--company-name", default="", help="Or look up by company_name field")
    ap.add_argument("--fy",           default="", help="FY end_date ISO (YYYY-MM-DD); "
                                                       "defaults to most recent")
    ap.add_argument("--out",          default="report.html",
                    help="Output HTML path (default: report.html)")
    ap.add_argument("--json",         action="store_true",
                    help="Also write the structured data as <out>.json")
    args = ap.parse_args()

    result = build_report_for_company(
        company_id=args.company_id or None,
        company_name=args.company_name or None,
        fy_iso=args.fy or None,
    )
    out_path = Path(args.out)
    out_path.write_text(result["html"], encoding="utf-8")
    print(f"[report] wrote {out_path.resolve()}")
    if args.json:
        # Strip the dataclass benchmark before serialising
        data_clean = json.loads(json.dumps(
            result["data"],
            default=lambda o: getattr(o, "__dict__", str(o)),
        ))
        json_path = out_path.with_suffix(out_path.suffix + ".json")
        json_path.write_text(json.dumps(data_clean, indent=2, default=str),
                             encoding="utf-8")
        print(f"[report] wrote {json_path.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
