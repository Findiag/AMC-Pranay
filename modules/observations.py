"""
observations.py — generate per-metric observations + recommendations + future
plans for a company's bs&pl record(s) and write them to the Airtable
observations table.

Refined from the standalone bspl_to_observations.py:
  - reads credentials from config.json (no env vars required)
  - clean UTF-8 throughout (no mojibake)
  - importable as `run_pipeline(...)` and from /api/observations/run

Usage:
  from observations import run_pipeline
  res = run_pipeline(company_id="recXXXX", dry_run=True)
  print(res["counts"])

CLI:
  python modules/observations.py --company-id recXXXX --dry-run
"""
from __future__ import annotations

import argparse
import io
import json
import os
import sys
import time
from contextlib import redirect_stdout
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import requests

sys.path.insert(0, str(Path(__file__).parent.parent))
from config import load_config


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Tables
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TBL_BSPL    = "tblSRySB4ePxa5xsH"
TBL_OBS     = "tblKOfXEEZln1Ol2n"
TBL_BASE    = "tblJIH3NqKOEqlyaW"
TBL_COMPANY = "tblv7I9Ex10IXVHE3"

AIRTABLE_SLEEP_S = 0.22


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Metric catalog — 29 metrics across growth/profitability/efficiency/
# liquidity/turnover/leverage/safety. Field names match Airtable EXACTLY.
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

@dataclass
class Metric:
    name:         str
    bspl_field:   str
    base_field:   Optional[str]
    obs_field:    str
    recom_field:  str
    future_field: str
    unit:         str
    category:     str
    guidance:     str
    direction:    str = "higher"   # higher | lower | band — used by benchmarks.py


METRICS: list[Metric] = [
    Metric("Revenue Growth", "revenue_growth_number", "revenue_growth_range",
           "Revenue Growth Obs", "Revenue Growth Recom", "Revenue Growth Future Plan",
           "%", "growth", "YoY top-line growth.", "higher"),
    Metric("Profit Growth", "profit_growth_number", "profit_growth_range",
           "Profit Growth Obs", "Profit Growth Recom", "Profit Growth Future Plan",
           "%", "growth", "YoY net profit growth.", "higher"),
    Metric("Gross Margin Expansion", "gross_margin_expenses", "gross_margin_expension_range",
           "Gross Margin Expansion Obs", "Gross Margin Expansion Recom",
           "Gross Margin Expansion Future Plan",
           "pp", "growth", "YoY change in gross margin (pp).", "higher"),
    Metric("Net Margin Expansion", "net_margin_expenses", "net_margin_expension_range",
           "Net Margin Expansion Obs", "Net Margin Expansion Recom",
           "Net Margin Expansion Future Plan",
           "pp", "growth", "YoY change in net margin (pp).", "higher"),
    Metric("Gross Profit Margin", "gross_profit_margin(number)", "gross_profit_margin_range",
           "Gross Profit Margin Obs", "Gross Profit Margin Recom",
           "Gross Profit Margin Future Plan",
           "%", "profitability", "Gross profit / revenue.", "higher"),
    Metric("Net Profit Margin", "net_profit_margin(number)", "net_profit_margin_range",
           "Net Profit Margin Obs", "Net Profit Margin Recom",
           "Net Profit Margin Future Plan",
           "%", "profitability", "Net profit / revenue.", "higher"),
    Metric("Return on Total Assets", "return_on_total_assets(number)",
           "return_on_total_assets_range",
           "Return on Total Assets Obs", "Return on Total Assets Recom",
           "Return on Total Assets Future Plan",
           "%", "returns", "Net profit / total assets.", "higher"),
    Metric("Return on Equity", "return_on_equity(number)", "return_on_equity_range",
           "Return on Equity Obs", "Return on Equity Recom", "Return on Equity Future Plan",
           "%", "returns", "Net profit / shareholder equity.", "higher"),
    Metric("Days of Receivable Outstanding", "days_of_outstanding", "days_of_outstanding_range",
           "Days of Receivable Outstanding Obs", "Days of Receivable Outstanding Recom",
           "Days of Receivable Outstanding Future Plan",
           "days", "efficiency", "Avg days to collect from customers.", "lower"),
    Metric("Days of Inventory Outstanding", "days_of_inventory_outstanding",
           "days_of_inventory_outstanding_range",
           "Days of Inventory Outstanding Obs", "Days of Inventory Outstanding Recom",
           "Days of Inventory Outstanding Future Plan",
           "days", "efficiency", "Avg days inventory is held.", "lower"),
    Metric("Days of Payable Outstanding", "days_of_purchase", "days_of_purchase_range",
           "Days of payable outstanding Obs", "Days of payable outstanding Recom",
           "Days of payable outstanding Future Plan",
           "days", "efficiency", "Avg days to pay suppliers.", "band"),
    Metric("Working Capital Cycle", "working_capital_cycle", "working_capital_cycle_range",
           "Working Capital Cycle Obs", "Working Capital Cycle Recom",
           "Working Capital Cycle Future Plan",
           "days", "efficiency", "DOO + DIO.", "lower"),
    Metric("Cash Conversion Cycle", "cash_conversion_cycle", "cash_conversion_cycle_range",
           "Cash Conversion Cycle Obs", "Cash Conversion Cycle Recom",
           "Cash Conversion Cycle Future Plan",
           "days", "efficiency", "DIO + DOO - DPO.", "lower"),
    Metric("Receivables Turnover", "account_receivable_turnover",
           "account_receivable_turnover_range",
           "Receivables Turnover Obs", "Receivables Turnover Recom",
           "Receivables Turnover Future Plan",
           "x", "turnover", "Revenue / receivables.", "higher"),
    Metric("Inventory Turnover", "inventory_turnover_ratio", "inventory_turnover_ratio_range",
           "Inventory Turnover Obs", "Inventory Turnover Recom",
           "Inventory Turnover Future Plan",
           "x", "turnover", "COGS / inventory.", "higher"),
    Metric("Payables Turnover", "account_payable_turnover_ratio",
           "account_payable_turnover_ratio_range",
           "Payales Turnover Obs", "Payales Turnover Recom", "Payales Turnover Future Plan",
           "x", "turnover", "Purchases / payables.", "lower"),
    Metric("Asset Turnover", "asset_turnover_ratio", "asset_turnover_ratio_range",
           "Asset Turnover Obs", "Asset Turnover Recom", "Asset Turnover Future Plan",
           "x", "turnover", "Revenue / total assets.", "higher"),
    Metric("Equity Turnover", "equity_turnover_ratio", "equity_turnover_ratio_range",
           "Equity Turnover Obs", "Equity Turnover Recom", "Equity Turnover Future Plan",
           "x", "turnover", "Revenue / equity.", "higher"),
    Metric("Current Ratio", "current_ratio", "current_ratio_range",
           "Current Ratio Obs", "Current Ratio Recom", "Current Ratio Future Plan",
           "x", "liquidity", "Current assets / current liabilities.", "band"),
    Metric("Quick Ratio", "quick_ratio", "quick_ratio_range",
           "Quick Ratio Obs", "Quick Ratio Recom", "Quick Ratio Future Plan",
           "x", "liquidity", "(CA - inventory) / CL.", "band"),
    Metric("Cash Ratio", "cash_ratio", "cash_ratio_range",
           "Cash Ratio Obs", "Cash Ratio Recom", "Cash Ratio Future Plan",
           "x", "liquidity", "Cash / CL.", "higher"),
    Metric("Cash Burn", "cash_burn", "cash_burn_range",
           "Cash Burn Obs", "Cash Burn Recom", "Cash Burn Future Plan",
           "months", "liquidity", "Months of runway at current burn rate.", "higher"),
    Metric("Operating Leverage", "operating_leverage", "operating_leverage_range",
           "Operating Leverage Obs", "Operating Leverage Recom",
           "Operating Leverage Future Plan",
           "x", "leverage", "%change-EBIT / %change-Sales.", "lower"),
    Metric("Financial Leverage", "financial_leverage", "financial_leverage_range",
           "Financial Leverage Obs", "Financial Leverage Recom",
           "Financial Leverage Future Plan",
           "x", "leverage", "EBIT / (EBIT - Interest).", "lower"),
    Metric("Total Leverage", "total_leverage", "total_leverage_range",
           "Total Leverage Obs", "Total Leverage Recom", "Total Leverage Future Plan",
           "x", "leverage", "Operating x Financial leverage.", "lower"),
    Metric("Equity Ratio", "equity_ratio", "leverage/debt:equity_ratio_range",
           "Equity Ratio Obs", "Equity Ratio Recom", "Equity Ratio Future Plan",
           "x", "leverage", "Equity / total assets.", "lower"),
    Metric("Interest Coverage", "interest_coverage", "interest_coverage_range",
           "Interest Coverage Obs", "Interest Coverage Recom",
           "Interest Coverage Future Plan",
           "x", "leverage", "EBIT / interest.", "higher"),
    Metric("Break Even Sales", "break_even_point", None,
           "Break Even Sales Obs", "Break Even Sales Recom", "Break Even Sales Future Plan",
           "INR", "safety", "Revenue at which net profit = 0.", "lower"),
    Metric("Margin of Safety", "margin_of_safety", None,
           "Margin of Safety Obs", "Margin of Safety Recom", "Margin of Safety Future Plan",
           "%", "safety", "Cushion above breakeven.", "higher"),
]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Helpers
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _creds() -> tuple[str, str, str]:
    cfg = load_config()
    at = (cfg.get("airtable_api_token") or os.environ.get("AIRTABLE_TOKEN", "")).strip()
    oa = (cfg.get("openai_api_key") or os.environ.get("OPENAI_API_KEY", "")).strip()
    bs = cfg.get("airtable_base_id", "appDDld0aAuplEL2l").strip()
    return at, oa, bs


def is_blank(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


def sanitize_value(v: Any) -> Any:
    """Normalize Airtable formula special values before they reach GPT.

    Airtable returns dicts like {'specialValue': 'Infinity'} or
    {'specialValue': 'NaN'} when a formula divides by zero. Passing those
    raw produces noise like 'inventory turnover is reported as Infinity'.
    Convert them to a clear textual marker so the prompt's
    'Insufficient data' rule fires instead.
    """
    if v is None:
        return None
    if isinstance(v, dict):
        sv = v.get("specialValue")
        if sv in ("Infinity", "-Infinity", "NaN"):
            return "Insufficient data (division-by-zero in source)"
        return None
    if isinstance(v, float):
        import math
        if math.isnan(v) or math.isinf(v):
            return "Insufficient data"
    if isinstance(v, str):
        s = v.strip().lower()
        if s in ("infinity", "-infinity", "nan", "inf", "-inf"):
            return "Insufficient data"
    return v


def at_headers(token: str) -> dict:
    if not token:
        raise RuntimeError("Airtable token missing (set airtable_api_token in config.json).")
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def at_list_all(table_id: str, token: str, base_id: str, params: Optional[dict] = None) -> list:
    out = []
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}"
    offset = None
    while True:
        q = dict(params or {})
        q["pageSize"] = 100
        if offset:
            q["offset"] = offset
        r = requests.get(url, headers=at_headers(token), params=q, timeout=30)
        if r.status_code != 200:
            raise RuntimeError(f"Airtable list failed [{r.status_code}]: {r.text[:300]}")
        data = r.json()
        out.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset:
            break
        time.sleep(AIRTABLE_SLEEP_S)
    return out


def at_find_one(table_id: str, formula: str, token: str, base_id: str) -> Optional[dict]:
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}"
    r = requests.get(url, headers=at_headers(token),
                     params={"filterByFormula": formula, "maxRecords": 1}, timeout=20)
    r.raise_for_status()
    recs = r.json().get("records", [])
    return recs[0] if recs else None


def at_create(table_id: str, fields: dict, token: str, base_id: str) -> dict:
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}"
    r = requests.post(url, headers=at_headers(token),
                      json={"fields": fields, "typecast": True}, timeout=30)
    if r.status_code >= 300:
        raise RuntimeError(f"Airtable create failed [{r.status_code}]: {r.text[:500]}")
    time.sleep(AIRTABLE_SLEEP_S)
    return r.json()


def at_update(table_id: str, record_id: str, fields: dict, token: str, base_id: str) -> dict:
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}/{record_id}"
    r = requests.patch(url, headers=at_headers(token),
                       json={"fields": fields, "typecast": True}, timeout=30)
    if r.status_code >= 300:
        raise RuntimeError(f"Airtable update failed [{r.status_code}]: {r.text[:500]}")
    time.sleep(AIRTABLE_SLEEP_S)
    return r.json()


def _openai_client(api_key: str):
    from openai import OpenAI
    return OpenAI(api_key=api_key)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# GPT calls
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

METRIC_FOCUS: dict[str, str] = {
    "Revenue Growth": (
        "Provide a detailed analysis of the revenue trend. Critically analyse "
        "the revenue growth, taking into account gross-profit-margin variation "
        "and any movement in days-of-receivable-outstanding. If sufficient years "
        "are not available for CAGR, write 'Insufficient data for CAGR'. If "
        "growth is below the industry / healthy range or the trend is "
        "decelerating or volatile, give specific recommendations to improve "
        "revenue performance."
    ),
    "Profit Growth": (
        "Deconstruct YoY net-profit growth into the sales-driven component "
        "(revenue change at prior margin) and the cost-driven component "
        "(margin change at current revenue). State each separately. If profit "
        "is volatile, below the healthy range, or falling, recommend specific "
        "cost or pricing actions."
    ),
    "Gross Margin Expansion": (
        "Evaluate pricing power vs cost pass-through. If gross margin is flat "
        "or contracting YoY, recommend procurement renegotiation, mix shifts, "
        "or pricing actions."
    ),
    "Net Margin Expansion": (
        "Assess sustainability — is the YoY change driven by recurring "
        "operating leverage or one-offs? If flat or declining, recommend "
        "cost-base actions and overhead discipline."
    ),
    "Days of Receivable Outstanding": (
        "If DSO is rising or above the healthy range, recommend collections "
        "tightening (credit-policy review, dunning cadence, factoring)."
    ),
    "Cash Conversion Cycle": (
        "Reconcile DIO + DSO − DPO. Flag whichever component is the main "
        "drag. A 60-day improvement is often worth more than a quarter of profit."
    ),
    "Interest Coverage": (
        "If below 1.5×, this is distress territory. Above 5× is comfort. "
        "Recommend specific actions if below the healthy range."
    ),
}


def gpt_metric_call(client, model: str, company_name: str, fy_label: str,
                    metric: Metric, current_value, prior_value, industry_range) -> dict:
    sys_prompt = (
        "You are a senior CFO of an Indian company. Your analysis must be "
        "strictly data-driven, decision-oriented, and based ONLY on the "
        "numerical values provided in the input. All monetary values are "
        "treated as INR (Lakhs).\n\n"
        "CRITICAL RULES:\n"
        "- Use ONLY the numbers provided in the input. Do NOT assume, infer, "
        "or generate sample values.\n"
        "- NEVER use placeholder sequences such as 120, 130, 140, 150, 160.\n"
        "- If sufficient data is not available for a calculation (e.g. CAGR), "
        "explicitly state 'Insufficient data'.\n"
        "- All calculations must be derived strictly from the values given.\n"
        "- Do NOT hallucinate competitor data; use relative benchmarking only.\n"
        "- Every numeric reference in the output must match the provided dataset.\n\n"
        "Return ONLY a strict JSON object with exactly three keys:\n"
        '  "observation"    - 2-4 sentences of factual interpretation grounded '
        "in the numbers.\n"
        '  "recommendation" - 2-4 sentences of specific actionable advice.\n'
        '  "future_plan"    - 2-3 sentences of forward-looking moves over the '
        "next 1-2 quarters.\n"
        "No markdown, no preamble, no extra keys, no explanation outside the JSON."
    )
    focus = METRIC_FOCUS.get(metric.name, "")
    user_prompt = (
        f"Company: {company_name}\n"
        f"Financial Year: {fy_label}\n"
        f"Metric: {metric.name}  (unit: {metric.unit}; category: {metric.category})\n"
        f"Definition: {metric.guidance}\n\n"
        f"Current-year value: {current_value}\n"
        f"Prior-year value:   {prior_value if prior_value is not None else 'N/A'}\n"
        f"Industry / healthy range: {industry_range if industry_range is not None else 'N/A'}\n\n"
        + (f"Metric-specific focus:\n{focus}\n\n" if focus else "")
        + "Write observation, recommendation, and future plan grounded in the "
          "numbers above."
    )
    resp = client.chat.completions.create(
        model=model, temperature=0.4,
        response_format={"type": "json_object"},
        messages=[{"role": "system", "content": sys_prompt},
                  {"role": "user",   "content": user_prompt}],
    )
    raw = resp.choices[0].message.content or "{}"
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        parsed = {}
    return {
        "observation":    str(parsed.get("observation", "")).strip(),
        "recommendation": str(parsed.get("recommendation", "")).strip(),
        "future_plan":    str(parsed.get("future_plan", "")).strip(),
    }


def gpt_summary_call(client, model: str, company_name: str, fy_label: str,
                     metric_table: list[dict]) -> dict:
    sys_prompt = (
        "You are a senior CFO. Return JSON object with EXACTLY:\n"
        '  "health_score"      - integer 0..100.\n'
        '  "critical_insights" - array of EXACTLY 5 strings (1-2 sentences each), '
        "ranked by urgency.\n"
        "No markdown, no extra keys."
    )
    rows = "\n".join(
        f"- {m['name']:32s} | curr={m['curr']!s:>14s} | prior={m['prior']!s:>14s} "
        f"| range={m['range']!s} | category={m['category']}"
        for m in metric_table
    )
    user_prompt = (
        f"Company: {company_name}\nFinancial Year: {fy_label}\n\n"
        f"29-metric snapshot:\n{rows}\n\n"
        "Score overall health 0-100 and give the 5 most important board-level insights."
    )
    resp = client.chat.completions.create(
        model=model, temperature=0.3,
        response_format={"type": "json_object"},
        messages=[{"role": "system", "content": sys_prompt},
                  {"role": "user",   "content": user_prompt}],
    )
    raw = resp.choices[0].message.content or "{}"
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        parsed = {}
    score = parsed.get("health_score", 0)
    try:
        score = max(0, min(100, int(round(float(score)))))
    except Exception:
        score = 0
    insights = parsed.get("critical_insights") or []
    if not isinstance(insights, list):
        insights = []
    insights = [str(x).strip() for x in insights][:5]
    while len(insights) < 5:
        insights.append("")
    return {"health_score": score, "critical_insights": insights}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Lookups
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def get_company_name(company_records_by_id: dict, company_link: list) -> str:
    if not company_link:
        return "Unknown Company"
    rec = company_records_by_id.get(company_link[0])
    if not rec:
        return company_link[0]
    return rec.get("fields", {}).get("company_name", company_link[0])


def get_financial_year(bspl_fields: dict) -> tuple[Optional[str], str]:
    end_date = bspl_fields.get("end_date")
    if not end_date:
        return None, "Unknown FY"
    try:
        dt = datetime.strptime(end_date, "%Y-%m-%d")
        if dt.month >= 4:
            label = f"FY{str(dt.year)[-2:]}-{str(dt.year + 1)[-2:]}"
        else:
            label = f"FY{str(dt.year - 1)[-2:]}-{str(dt.year)[-2:]}"
        return end_date, label
    except Exception:
        return end_date, end_date


def get_prior_value(bspl_fields: dict, bspl_field_name: str,
                    bspl_records_by_id: dict):
    prev_link = bspl_fields.get("previous_bs&pl") or []
    if not prev_link:
        return None
    prev_rec = bspl_records_by_id.get(prev_link[0])
    if not prev_rec:
        return None
    return prev_rec.get("fields", {}).get(bspl_field_name)


def get_industry_range(base_records_by_company: dict, company_link: list,
                       base_field_name: Optional[str]):
    if not base_field_name or not company_link:
        return None
    base_rec = base_records_by_company.get(company_link[0])
    if not base_rec:
        return None
    return base_rec.get("fields", {}).get(base_field_name)


def find_existing_obs(company_link: list, fy_iso: Optional[str],
                      token: str, base_id: str) -> Optional[dict]:
    if not company_link or not fy_iso:
        return None
    company_id = company_link[0]
    formula = (
        f"AND("
        f"FIND('{company_id}', ARRAYJOIN({{company}})),"
        f"IS_SAME({{financial year}}, '{fy_iso}', 'day')"
        f")"
    )
    return at_find_one(TBL_OBS, formula, token, base_id)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Per-record processing
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _process_record(client, model: str, bspl_rec: dict, bspl_records_by_id: dict,
                    company_records_by_id: dict, base_records_by_company: dict,
                    token: str, base_id: str,
                    dry_run: bool, force: bool, log) -> str:
    f = bspl_rec.get("fields", {})
    company_link = f.get("company") or []
    user_link    = f.get("user") or []
    company_name = get_company_name(company_records_by_id, company_link)
    fy_iso, fy_label = get_financial_year(f)

    log(f"\n{'='*72}")
    log(f"COMPANY : {company_name}")
    log(f"FY      : {fy_label}   (ISO: {fy_iso})")
    log(f"BS&PL id: {bspl_rec['id']}")
    log(f"{'='*72}")

    existing = find_existing_obs(company_link, fy_iso, token, base_id)
    existing_fields = (existing or {}).get("fields", {}) if existing else {}
    if existing:
        log(f"  > existing observation row: {existing['id']}  (fill-blanks mode)")
    else:
        log(f"  > no existing row — will create with all fields")

    payload: dict = {}
    metric_table: list[dict] = []
    gen_count = skip_count = 0

    # ── Build the work list first (no GPT yet) ──
    # This separates "what to do" from "do it" so we can fan out the GPT
    # calls in parallel below. 29 sequential GPT calls take ~145s; firing
    # them in a thread pool of 8 cuts that to roughly 30s per record.
    jobs_to_run = []
    for i, metric in enumerate(METRICS, 1):
        curr  = sanitize_value(f.get(metric.bspl_field))
        prior = sanitize_value(get_prior_value(f, metric.bspl_field, bspl_records_by_id))
        rng   = sanitize_value(get_industry_range(base_records_by_company,
                                                  company_link, metric.base_field))

        metric_table.append({
            "name": metric.name, "curr": curr, "prior": prior,
            "range": rng, "category": metric.category,
        })

        if force:
            obs_blank = rec_blank = fut_blank = True
        else:
            obs_blank = is_blank(existing_fields.get(metric.obs_field))
            rec_blank = is_blank(existing_fields.get(metric.recom_field))
            fut_blank = is_blank(existing_fields.get(metric.future_field))

        if not (obs_blank or rec_blank or fut_blank):
            skip_count += 1
            log(f"  [{i:02d}/29] {metric.name:32s} -- already filled, SKIP")
            continue

        jobs_to_run.append({
            "i": i, "metric": metric, "curr": curr, "prior": prior, "rng": rng,
            "obs_blank": obs_blank, "rec_blank": rec_blank, "fut_blank": fut_blank,
        })

    # ── Fan out GPT calls in parallel ──
    if jobs_to_run:
        from concurrent.futures import ThreadPoolExecutor, as_completed
        import time as _time
        _t0 = _time.time()
        log(f"  Firing {len(jobs_to_run)} metric GPT calls in parallel "
            f"(model={model})...")

        def _one(job: dict):
            try:
                res = gpt_metric_call(client, model, company_name, fy_label,
                                      job["metric"], job["curr"], job["prior"], job["rng"])
                return job, res, None
            except Exception as e:
                return job, {"observation": f"Error: {e}",
                             "recommendation": "", "future_plan": ""}, e

        results: list[tuple] = []
        # 8 workers is a sweet spot for OpenAI — enough concurrency to saturate
        # network without tripping the per-org TPM limit on gpt-4o-mini.
        with ThreadPoolExecutor(max_workers=8) as ex:
            futs = [ex.submit(_one, j) for j in jobs_to_run]
            for fut in as_completed(futs):
                results.append(fut.result())
        log(f"  All {len(jobs_to_run)} GPT calls done in {(_time.time() - _t0):.1f}s")

        # Sort back to original metric order so the log is readable
        results.sort(key=lambda r: r[0]["i"])
        for job, result, err in results:
            i = job["i"]; metric = job["metric"]
            obs_blank = job["obs_blank"]; rec_blank = job["rec_blank"]; fut_blank = job["fut_blank"]

            if obs_blank: payload[metric.obs_field]    = result["observation"]
            if rec_blank: payload[metric.recom_field]  = result["recommendation"]
            if fut_blank: payload[metric.future_field] = result["future_plan"]

            gen_count += 1
            which = "".join(["O" if obs_blank else "-",
                             "R" if rec_blank else "-",
                             "F" if fut_blank else "-"])
            if err:
                log(f"  [{i:02d}/29] {metric.name:32s} OPENAI ERROR: {err}")
                continue
            log(f"  [{i:02d}/29] {metric.name:32s} fill[{which}] "
                f"curr={str(job['curr'])[:10]:<12s} prior={str(job['prior'])[:10]:<12s} "
                f"range={str(job['rng'])[:14]:<14s}")
            if obs_blank and result["observation"]:
                log(f"           OBS  : {result['observation'].replace(chr(10), ' ')[:160]}")
            if rec_blank and result["recommendation"]:
                log(f"           RECOM: {result['recommendation'].replace(chr(10), ' ')[:160]}")
            if fut_blank and result["future_plan"]:
                log(f"           FUTURE: {result['future_plan'].replace(chr(10), ' ')[:160]}")

    log(f"\n  metrics summary: generated={gen_count}  skipped={skip_count}  / 29")

    if force:
        score_blank = True
        ci_blanks = [1, 2, 3, 4, 5]
    else:
        score_blank = is_blank(existing_fields.get("health_score"))
        ci_blanks = [n for n in range(1, 6) if is_blank(existing_fields.get(f"ci_{n}"))]

    if score_blank or ci_blanks:
        try:
            summary = gpt_summary_call(client, model, company_name, fy_label, metric_table)
        except Exception as e:
            log(f"  [summary] OPENAI ERROR: {e}")
            summary = {"health_score": 0, "critical_insights": ["", "", "", "", ""]}
        if score_blank:
            payload["health_score"] = summary["health_score"] / 100.0
            log(f"  [summary] health_score blank -> filling {summary['health_score']}/100")
        for n in ci_blanks:
            payload[f"ci_{n}"] = summary["critical_insights"][n - 1]
            log(f"  [summary] ci_{n} blank -> filling")
    else:
        log(f"  [summary] health_score + all ci_ already filled, SKIP")

    if not payload:
        log(f"\n  *** NOTHING TO WRITE — all fields already filled. ***")
        return "skipped"

    log(f"\n  TOTAL FIELDS TO WRITE: {len(payload)}")

    fields_to_send = dict(payload)
    if not existing:
        if company_link: fields_to_send["company"] = company_link
        if user_link:    fields_to_send["User"] = user_link
        if fy_iso:       fields_to_send["financial year"] = fy_iso
    else:
        # Backfill 'financial year' on legacy rows that don't have it set —
        # without this the report's FY-strict lookup keeps missing them.
        if fy_iso and is_blank(existing_fields.get("financial year")):
            fields_to_send["financial year"] = fy_iso
            log(f"  [legacy backfill] setting 'financial year' = {fy_iso} "
                f"on existing row {existing['id']}")

    if dry_run:
        log(f"  [dry-run] {len(fields_to_send)} fields would be written; no Airtable call.")
        return "dry-run"

    if existing:
        at_update(TBL_OBS, existing["id"], fields_to_send, token, base_id)
        rec_id = existing["id"]
        log(f"  >>> UPDATED Airtable observation row {rec_id} "
            f"({len(fields_to_send)} fields written)")
        log(f"      view at: https://airtable.com/{base_id}/{TBL_OBS}/{rec_id}")
        log(f"      ✓ ACK: {company_name} / {fy_label} / "
            f"{gen_count} metrics generated, {skip_count} pre-filled, "
            f"{len(fields_to_send)} fields patched")
        return "updated"
    else:
        rec = at_create(TBL_OBS, fields_to_send, token, base_id)
        rec_id = rec.get("id")
        log(f"  >>> CREATED Airtable observation row {rec_id} "
            f"({len(fields_to_send)} fields written)")
        log(f"      view at: https://airtable.com/{base_id}/{TBL_OBS}/{rec_id}")
        log(f"      ✓ ACK: {company_name} / {fy_label} / "
            f"{gen_count} metrics generated, {len(fields_to_send)} fields written")
        return "created"


def _filter_bspl(bspl_records: list, user_id: str = "", company_id: str = "") -> list:
    out = []
    for r in bspl_records:
        f = r.get("fields", {})
        company_link = f.get("company") or []
        user_link    = f.get("user") or []
        if company_id and company_id not in company_link:
            continue
        if user_id and user_id not in user_link:
            continue
        out.append(r)
    return out


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Public entry point
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def run_pipeline(user_id: str = "", company_id: str = "",
                 limit: int = 0, dry_run: bool = False, force: bool = False,
                 capture_log: bool = False, model: Optional[str] = None,
                 bspl_ids: Optional[list] = None) -> dict:
    """Run observations pipeline. Returns counts + optional captured log.

    Filters (any combo):
      user_id    — only bs&pl rows linked to this user
      company_id — only bs&pl rows linked to this company (also speeds up the
                   Airtable list call by filtering server-side)
      bspl_ids   — process ONLY these specific bs&pl record ids (most useful
                   from the auto-upload hook: pass the 1-2 just-created
                   record ids and skip the historical backfill).

    Both/all empty = process every record (use carefully).
    """
    buf = io.StringIO() if capture_log else None
    cm = redirect_stdout(buf) if capture_log else _NoOp()

    counts = {"created": 0, "updated": 0, "skipped": 0, "dry-run": 0, "error": 0}

    with cm:
        token, openai_key, base_id = _creds()
        if not token:
            raise RuntimeError("airtable_api_token missing in config.json")
        if not openai_key:
            raise RuntimeError("openai_api_key missing in config.json")

        cfg = load_config()
        model = model or cfg.get("observations_model", "gpt-4o-mini")

        if force:
            print("*** FORCE mode: regenerating everything ***")
        if user_id or company_id or bspl_ids:
            print(f"*** Filters: user_id={user_id or '(any)'}  "
                  f"company_id={company_id or '(any)'}  "
                  f"bspl_ids={bspl_ids or '(any)'} ***")

        print("Loading Airtable data ...")
        # Speed win: when bspl_ids is provided we don't need ALL 248 bs&pl
        # rows. Fetch just those records by id. When company_id is provided
        # we filter server-side so we get ~5-10 rows instead of 248.
        if bspl_ids:
            bspl_records = []
            for rid in bspl_ids:
                rec = at_find_one(TBL_BSPL,
                                  f"RECORD_ID() = '{rid}'", token, base_id)
                if rec:
                    bspl_records.append(rec)
        elif company_id:
            params = {"filterByFormula":
                      f"FIND('{company_id}', ARRAYJOIN({{company}}))"}
            bspl_records = at_list_all(TBL_BSPL, token, base_id, params=params)
        else:
            bspl_records = at_list_all(TBL_BSPL, token, base_id)

        company_records = at_list_all(TBL_COMPANY, token, base_id)
        base_records    = at_list_all(TBL_BASE,    token, base_id)

        bspl_by_id    = {r["id"]: r for r in bspl_records}
        company_by_id = {r["id"]: r for r in company_records}
        base_by_company: dict = {}
        for r in base_records:
            for cid in r.get("fields", {}).get("company", []):
                base_by_company[cid] = r

        # If we filtered server-side by bspl_ids, get_prior_value() needs the
        # FULL bs&pl-by-id map to look up linked previous_bs&pl rows that
        # weren't in the bspl_ids list. Lazy-load only when needed.
        if bspl_ids:
            need_prev = any((r.get("fields", {}).get("previous_bs&pl") or [])
                            for r in bspl_records)
            if need_prev:
                prev_ids = set()
                for r in bspl_records:
                    for pid in (r.get("fields", {}).get("previous_bs&pl") or []):
                        if pid not in bspl_by_id:
                            prev_ids.add(pid)
                for pid in prev_ids:
                    prev = at_find_one(TBL_BSPL,
                                       f"RECORD_ID() = '{pid}'", token, base_id)
                    if prev:
                        bspl_by_id[prev["id"]] = prev

        print(f"  bs&pl: {len(bspl_records)} | company: {len(company_records)} | "
              f"base: {len(base_records)}")

        # If bspl_ids was used we already fetched only those — no extra
        # filtering needed. Otherwise apply the user/company filter.
        if not bspl_ids:
            bspl_records = _filter_bspl(bspl_records, user_id, company_id)
            print(f"  after user/company filter: {len(bspl_records)} bs&pl record(s)")

        if limit and limit > 0:
            bspl_records = bspl_records[:limit]
            print(f"  limit applied -> processing {len(bspl_records)} record(s)")

        client = _openai_client(openai_key)

        for rec in bspl_records:
            try:
                status = _process_record(client, model, rec, bspl_by_id, company_by_id,
                                         base_by_company, token, base_id,
                                         dry_run=dry_run, force=force, log=print)
                counts[status] = counts.get(status, 0) + 1
            except Exception as e:
                counts["error"] += 1
                print(f"  !!! FAILED on record {rec['id']}: {e}")

        print(f"\n{'='*72}\nDONE")
        for k, v in counts.items():
            print(f"  {k:8s}: {v}")
        print(f"{'='*72}")

    return {"counts": counts, "log": (buf.getvalue() if buf else "")}


class _NoOp:
    def __enter__(self): return self
    def __exit__(self, *a): return False


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CLI
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--user-id",    default="")
    ap.add_argument("--company-id", default="")
    ap.add_argument("--limit",      type=int, default=0)
    ap.add_argument("--dry-run",    action="store_true")
    ap.add_argument("--force",      action="store_true",
                    help="Regenerate everything (ignore existing filled fields)")
    ap.add_argument("--model",      default=None,
                    help="OpenAI model override (default from config or gpt-4o-mini)")
    args = ap.parse_args()
    run_pipeline(
        user_id=args.user_id.strip(),
        company_id=args.company_id.strip(),
        limit=args.limit, dry_run=args.dry_run, force=args.force,
        model=args.model,
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
