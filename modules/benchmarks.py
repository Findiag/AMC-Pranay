"""
benchmarks.py — industry benchmark catalog for the 29 metrics tracked in the
observations pipeline.

Each metric has a healthy/watch/alert band. The classification logic depends
on the metric's `direction`:

  higher = more is better. Healthy ≥ healthy_min; alert < alert_max.
  lower  = less is better. Healthy ≤ healthy_max; alert > alert_min.
  band   = there's a sweet spot; both above and below it are bad.

Industry defaults below come from the Aion Tech / Wendt diagnostic templates.
A per-company override can come from the Airtable `base` table — its
*_range fields take precedence. The override format we accept is:
  - a single number (treated as the healthy threshold)
  - "x-y"   (a band)
  - {"healthy": [...], "watch": [...]} (full structured override)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Benchmark:
    metric:       str
    direction:    str         # "higher" | "lower" | "band"
    healthy_min:  Optional[float] = None
    healthy_max:  Optional[float] = None
    watch_min:    Optional[float] = None
    watch_max:    Optional[float] = None
    unit:         str = ""
    description:  str = ""


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Industry default catalog — sane starting points for Indian SMEs.
# Override per-company via Airtable's `base` table.
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

DEFAULTS: dict[str, Benchmark] = {
    # ---- Growth ----
    "Revenue Growth":            Benchmark("Revenue Growth",       "higher",
        healthy_min=12, watch_min=5,  unit="%",
        description="Industry-typical SME growth target."),
    "Profit Growth":             Benchmark("Profit Growth",        "higher",
        healthy_min=15, watch_min=0,  unit="%",
        description="Profit should grow at least as fast as revenue."),
    "Gross Margin Expansion":    Benchmark("Gross Margin Expansion","higher",
        healthy_min=0.5, watch_min=-1, unit="pp",
        description="Year-over-year change. Positive = improving."),
    "Net Margin Expansion":      Benchmark("Net Margin Expansion", "higher",
        healthy_min=0,   watch_min=-2, unit="pp",
        description="YoY change. Sustained negatives erode value."),

    # ---- Profitability ----
    "Gross Profit Margin":       Benchmark("Gross Profit Margin",  "higher",
        healthy_min=25, watch_min=15, unit="%",
        description="Manufacturing 25-35%, services 40%+, retail 15-25%."),
    "Net Profit Margin":         Benchmark("Net Profit Margin",    "higher",
        healthy_min=10, watch_min=5,  unit="%",
        description="Manufacturing 5-10%, services 10-20%, premium 20%+."),
    "Return on Total Assets":    Benchmark("Return on Total Assets","higher",
        healthy_min=10, watch_min=5,  unit="%",
        description="Below 5% suggests under-utilised assets."),
    "Return on Equity":          Benchmark("Return on Equity",     "higher",
        healthy_min=15, watch_min=8,  unit="%",
        description="Below cost-of-equity (~15-18% in India) destroys value."),

    # ---- Efficiency (days) — lower is better ----
    "Days of Receivable Outstanding": Benchmark("Days of Receivable Outstanding","lower",
        healthy_max=60, watch_max=80, unit="days",
        description="Industry-standard credit terms. >90 is concerning."),
    "Days of Inventory Outstanding":  Benchmark("Days of Inventory Outstanding", "lower",
        healthy_max=45, watch_max=75, unit="days",
        description="Industry-dependent. >90 for high-churn SKUs is bad."),
    "Days of Payable Outstanding":    Benchmark("Days of Payable Outstanding",   "band",
        healthy_min=45, healthy_max=60, watch_min=30, watch_max=75, unit="days",
        description="Sweet spot is industry-standard payment terms."),
    "Working Capital Cycle":          Benchmark("Working Capital Cycle",         "lower",
        healthy_max=75, watch_max=100, unit="days"),
    "Cash Conversion Cycle":          Benchmark("Cash Conversion Cycle",         "lower",
        healthy_max=60, watch_max=100, unit="days",
        description="A 60-day improvement is often worth more than a quarter of profit."),

    # ---- Turnover (×) ----
    "Receivables Turnover": Benchmark("Receivables Turnover","higher",
        healthy_min=4, watch_min=2, unit="x"),
    "Inventory Turnover":   Benchmark("Inventory Turnover",  "higher",
        healthy_min=4, watch_min=2, unit="x"),
    "Payables Turnover":    Benchmark("Payables Turnover",   "lower",
        healthy_max=6, watch_max=8, unit="x",
        description="Lower is better — using supplier credit as free working capital."),
    "Asset Turnover":       Benchmark("Asset Turnover",      "higher",
        healthy_min=1,   watch_min=0.5, unit="x"),
    "Equity Turnover":      Benchmark("Equity Turnover",     "higher",
        healthy_min=2,   watch_min=1,   unit="x"),

    # ---- Liquidity ----
    "Current Ratio":   Benchmark("Current Ratio", "band",
        healthy_min=2, healthy_max=4, watch_min=1.2, watch_max=5, unit="x",
        description="Below 1.2 risks default. Above 4 = idle cash."),
    "Quick Ratio":     Benchmark("Quick Ratio",   "band",
        healthy_min=1, healthy_max=3, watch_min=0.7, watch_max=4, unit="x"),
    "Cash Ratio":      Benchmark("Cash Ratio",    "higher",
        healthy_min=0.5, watch_min=0.2, unit="x",
        description="Below 0.2x means a one-month shock would crater operations."),
    "Cash Burn":       Benchmark("Cash Burn",     "higher",
        healthy_min=12, watch_min=6, unit="months",
        description="Months of runway at current burn rate."),

    # ---- Leverage / Risk ----
    "Operating Leverage": Benchmark("Operating Leverage", "lower",
        healthy_max=2, watch_max=5, unit="x",
        description="High operating leverage = fragile in downturns."),
    "Financial Leverage": Benchmark("Financial Leverage", "lower",
        healthy_max=1.5, watch_max=3, unit="x"),
    "Total Leverage":     Benchmark("Total Leverage",     "lower",
        healthy_max=3,   watch_max=6, unit="x"),
    "Equity Ratio":       Benchmark("Equity Ratio",       "higher",
        healthy_min=0.5, watch_min=0.3, unit="x"),
    "Interest Coverage":  Benchmark("Interest Coverage",  "higher",
        healthy_min=3,   watch_min=1.5, unit="x",
        description="Below 1.5 = distress; above 5 = comfort."),

    # ---- Safety ----
    "Break Even Sales": Benchmark("Break Even Sales", "lower",
        unit="INR",
        description="Should be well below trough revenue. No fixed band."),
    "Margin of Safety": Benchmark("Margin of Safety", "higher",
        healthy_min=30, watch_min=20, unit="%",
        description="Above 30% resilient; below 20% exposed."),
}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Status evaluation
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

HEALTHY = "healthy"
WATCH   = "watch"
ALERT   = "alert"
NEUTRAL = "neutral"   # benchmark not defined or value missing


def status_for(metric_name: str, value: Optional[float],
               override: Optional[Benchmark] = None) -> str:
    """Return one of healthy/watch/alert/neutral for a metric value."""
    if value is None:
        return NEUTRAL
    bm = override or DEFAULTS.get(metric_name)
    if not bm:
        return NEUTRAL
    try:
        v = float(value)
    except (TypeError, ValueError):
        return NEUTRAL

    if bm.direction == "higher":
        if bm.healthy_min is not None and v >= bm.healthy_min: return HEALTHY
        if bm.watch_min   is not None and v >= bm.watch_min:   return WATCH
        return ALERT

    if bm.direction == "lower":
        if bm.healthy_max is not None and v <= bm.healthy_max: return HEALTHY
        if bm.watch_max   is not None and v <= bm.watch_max:   return WATCH
        return ALERT

    if bm.direction == "band":
        if (bm.healthy_min is not None and bm.healthy_max is not None
                and bm.healthy_min <= v <= bm.healthy_max):
            return HEALTHY
        if (bm.watch_min is not None and bm.watch_max is not None
                and bm.watch_min <= v <= bm.watch_max):
            return WATCH
        return ALERT

    return NEUTRAL


def zone_bounds(metric_name: str, override: Optional[Benchmark] = None
                ) -> dict[str, tuple[Optional[float], Optional[float]]]:
    """Return drawing bounds for the benchmark zone bar (used by the report's
    inline SVG charts). Keys: 'alert_lo', 'watch_lo', 'healthy', 'watch_hi',
    'alert_hi'. Each value is (start, end) on the same numeric axis."""
    bm = override or DEFAULTS.get(metric_name)
    if not bm:
        return {}

    if bm.direction == "higher":
        return {
            "alert":   (None, bm.watch_min),
            "watch":   (bm.watch_min, bm.healthy_min),
            "healthy": (bm.healthy_min, None),
        }
    if bm.direction == "lower":
        return {
            "healthy": (None, bm.healthy_max),
            "watch":   (bm.healthy_max, bm.watch_max),
            "alert":   (bm.watch_max, None),
        }
    if bm.direction == "band":
        return {
            "alert_lo":   (None, bm.watch_min),
            "watch_lo":   (bm.watch_min, bm.healthy_min),
            "healthy":    (bm.healthy_min, bm.healthy_max),
            "watch_hi":   (bm.healthy_max, bm.watch_max),
            "alert_hi":   (bm.watch_max, None),
        }
    return {}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Per-company override — read from the Airtable `base` table
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _parse_range(raw) -> Optional[tuple[float, float]]:
    """Best-effort parse of an Airtable range value into (lo, hi) numeric pair."""
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return (float(raw), float(raw))
    s = str(raw).strip().replace("%", "").replace(",", "")
    if not s:
        return None
    # "lo - hi" or "lo to hi"
    import re
    m = re.match(r"^\s*(-?\d+\.?\d*)\s*(?:[-–—]|to)\s*(-?\d+\.?\d*)\s*$", s, re.IGNORECASE)
    if m:
        return (float(m.group(1)), float(m.group(2)))
    # ">= x" / "> x" / "<= x" / "< x"
    m = re.match(r"^\s*(>=?|<=?)\s*(-?\d+\.?\d*)\s*$", s)
    if m:
        op, val = m.group(1), float(m.group(2))
        return (val, val)
    try:
        v = float(s)
        return (v, v)
    except ValueError:
        return None


def override_from_airtable(metric_name: str, base_field_value) -> Optional[Benchmark]:
    """Return a Benchmark built from a single Airtable range field, or None
    if the value can't be parsed. The default healthy/watch boundaries are
    inferred from the parsed range."""
    parsed = _parse_range(base_field_value)
    if not parsed:
        return None
    lo, hi = parsed
    default = DEFAULTS.get(metric_name)
    if not default:
        return None
    if default.direction == "higher":
        return Benchmark(metric_name, "higher",
                         healthy_min=lo, watch_min=lo * 0.5 if lo > 0 else lo,
                         unit=default.unit, description=default.description)
    if default.direction == "lower":
        return Benchmark(metric_name, "lower",
                         healthy_max=hi, watch_max=hi * 1.5,
                         unit=default.unit, description=default.description)
    # band
    return Benchmark(metric_name, "band",
                     healthy_min=lo, healthy_max=hi,
                     watch_min=lo * 0.7, watch_max=hi * 1.3,
                     unit=default.unit, description=default.description)
