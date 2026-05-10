"""
llm_mapper_v2.py — Accuracy-first LLM mapper for BS/PL → template.

Implements the upgraded mapping pipeline (10-point engineering plan):

  (2) Strict structured outputs + section-conditional field enums
  (3) Hierarchical mapping: section header context preserved, field choice
      space restricted per section
  (4) Programmatic post-checks (balance, coverage, section-total cross-check)
      with TARGETED re-prompt that names the specific failing constraint
  (5) Optional self-consistency (N runs, majority vote)  — flag-gated
  (6) Optional critic pass (second LLM call hunts for errors) — flag-gated
  (9) Negative examples baked into the prompt to prevent boundary errors

The output shape matches the legacy `map_bs_hybrid` / `map_pl_hybrid`
contract so this is a drop-in replacement behind the `mapper_mode` config flag.
"""
from __future__ import annotations

import json
import os
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Reuse the existing section detector + cell mapping — they encode hard-won
# knowledge from the 71-company audit.
from matcher import (
    BS_FIELD_TO_CELL,
    PL_FIELD_TO_CELL,
    _detect_sections,
)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Config / knobs (read from config.json on every call so edits take effect)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

DEFAULT_MODEL = "gpt-4o-2024-08-06"     # supports strict structured outputs
DEFAULT_CRITIC_MODEL = "gpt-4o-mini"
BALANCE_TOL = 1.0                        # |Assets - L+E| < 1.0 (Lakhs)
MAX_RETRY_ROUNDS = 2                     # constraint-driven re-prompts


def _load_cfg() -> Dict[str, Any]:
    try:
        from config import load_config
        return load_config()
    except Exception:
        return {}


def _get_client():
    cfg = _load_cfg()
    api_key = (cfg.get("openai_api_key") or os.environ.get("OPENAI_API_KEY", "")).strip()
    if not api_key:
        return None, cfg
    try:
        from openai import OpenAI
        return OpenAI(api_key=api_key), cfg
    except ImportError:
        return None, cfg


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Section → allowed cells (hierarchical / section-conditional enums)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BS_SECTION_CELLS: Dict[str, List[str]] = {
    "networth":         ["F5", "F6", "F7", "F8", "F9"],
    "current_liab":     ["F14", "F15", "F16", "F17", "F18", "F19"],
    "noncurrent_liab":  ["F24", "F25", "F26", "F27", "F28"],
    "current_asset":    ["F35", "F36", "F37", "F38", "F39", "F40", "F41", "F42", "F43"],
    "noncurrent_asset": ["F48", "F49", "F50", "F51", "F52", "F53", "F54", "F55"],
}

ALL_BS_CELLS: List[str] = sorted({c for cells in BS_SECTION_CELLS.values() for c in cells})

PL_TARGET_KEYS: List[str] = [
    "F66", "F67", "F69", "F70", "F71", "F73",
    "other_income", "net_profit", "exceptional_items",
]

# Cell → human-readable description (used inside the prompt so the model
# doesn't have to invent semantics from the cell ID alone).
BS_CELL_DESC: Dict[str, str] = {
    "F5":  "Share Capital (paid-up + authorised)",
    "F6":  "Retained Earnings (only if separately disclosed; else 0)",
    "F7":  "General Reserves & Surplus (Old GAAP 'Reserves and surplus' goes here; capital reserve, securities premium, statutory reserve, revaluation reserve)",
    "F8":  "Other Equity (Ind AS 'Other Equity' single-line; share application money pending allotment)",
    "F9":  "Non-Controlling Interest / Minority Interest (consolidated only); else 0",
    "F14": "Trade Payables (SUM all sub-items: MSME dues + non-MSME / sundry creditors / dues to others)",
    "F15": "Provisions (current) — short-term provisions, employee benefit obligations (current)",
    "F16": "Short Term Borrowings — bank overdraft, cash credit, working capital, current maturities of LT debt, bills payable",
    "F17": "Other Current Liabilities — statutory dues, advances from customers, contract liabilities, current tax liabilities, unearned revenue",
    "F18": "Other Financial Liabilities (current) — unpaid dividend, interest accrued, due to others",
    "F19": "Lease Liabilities (current) — 0 if not separately present",
    "F24": "Long Term Borrowings — term loans, debentures, secured/unsecured loans (non-current)",
    "F25": "Provisions (non-current) — long-term provisions, gratuity, leave encashment (non-current)",
    "F26": "Others (non-current liab): SUM of Lease Liabilities (NC) + Deferred Tax Liabilities + Other Non-Current Liabilities + Government grant (NC)",
    "F27": "Other Financial Liabilities (non-current)",
    "F28": "spare (non-current liab) — usually 0",
    "F35": "Bank Balance — 'Other Bank Balances' / earmarked / margin money. If extract has only 'Cash and Bank Balances' combined, put 0 here and full amount in F36.",
    "F36": "Cash & Cash Equivalents — including combined 'Cash and Bank Balances' if not split",
    "F37": "Inventory / Inventories — finished goods, stores, stock-in-trade. 0 for IT/service companies.",
    "F38": "Investments (current) — current investments, short-term investments, mutual funds (current)",
    "F39": "Loans (current) / Short Term Loans & Advances — ICDs, advances to suppliers/employees, security deposits (current)",
    "F40": "Trade Receivables / Accounts Receivable / Sundry Debtors / Contract Assets",
    "F41": "Other Current Assets — current tax assets, advance income tax, TDS receivable, GST input, prepaid expenses, held-for-sale assets",
    "F42": "Other Financial Assets (current)",
    "F43": "spare (current asset) — usually 0",
    "F48": "Fixed Assets — SUM of: PPE + Tangible Assets + Intangible Assets + Right-of-Use Assets + Goodwill + Software + Investment Property. EXCLUDE CWIP (that goes in F51).",
    "F49": "Investments (non-current) — long-term investments, investments in subsidiaries/associates/JVs",
    "F50": "Loans (non-current) / Long-term loans & advances — security deposits (NC)",
    "F51": "Capital Work-in-Progress + Intangible Assets under Development (SUM)",
    "F52": "Other Non-Current Assets — capital advances, non-current tax assets, non-financial NC assets",
    "F53": "Other Financial Assets (non-current)",
    "F54": "Deferred Tax Assets, MAT credit entitlement",
    "F55": "spare (non-current asset) — usually 0",
}

PL_CELL_DESC: Dict[str, str] = {
    "F66": "Revenue from Operations ONLY (exclude Other Income — that's reported separately as 'other_income')",
    "F67": "Cost of Goods Sold. Manufacturing: 'Cost of Materials Consumed' + 'Changes in Inventories of FG/WIP'. Trading: 'Purchase of traded goods' + 'Changes in inventories of stock-in-trade'. Services/IT: 'Operating Expenses' if it's the main cost line. 0 if no clear COGS.",
    "F69": "Employee Benefits Expense — staff cost, salaries & wages, contribution to PF, managerial remuneration",
    "F70": "Finance Costs / Interest expense / borrowing cost",
    "F71": "Depreciation and Amortization Expense (combined). Includes depletion. EXCLUDE impairment.",
    "F73": "Total Tax Expense — SUM of Current Tax + Deferred Tax + Tax of Earlier Years + MAT credit (signed)",
    "other_income": "Other Income — interest income, dividend, rental, gain on sale of investments/assets, profit on sale",
    "net_profit": "Final 'Profit after tax' / 'Profit for the year' / 'Net Profit' figure (after tax, after exceptional, BEFORE OCI)",
    "exceptional_items": "Exceptional / extraordinary items / prior-period adjustments. Positive = income, negative = expense, 0 if absent.",
}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Negative examples (point #9) — common BOUNDARY errors as anti-patterns
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BS_NEGATIVE_EXAMPLES = """COMMON ERRORS — DO NOT MAKE THESE:
- WRONG: "Others" under Non-Current Financial Assets → F41 (Other Current Assets).
  RIGHT: F53 (Other Financial Assets non-current). The section header determines current vs non-current.
- WRONG: "Capital Work-in-Progress" → F48 (Fixed Assets).
  RIGHT: F51 (CWIP). F48 is for completed PPE only.
- WRONG: "Deferred Tax Liability" → F25 (Provisions NC).
  RIGHT: F26 (Others NC liab). DTL is not a provision.
- WRONG: "Current Tax Liabilities" → F25 or F26 (provisions).
  RIGHT: F17 (Other Current Liabilities). Current tax liab is a current item.
- WRONG: "Lease Liabilities" under non-current section → F19 (Lease Liab Current).
  RIGHT: F26 (Others NC liab) when under non-current section.
- WRONG: "Trade Payables — total outstanding dues of MSME" stored in a different cell from "Trade Payables — total outstanding dues of others".
  RIGHT: BOTH go to F14. F14 is the SUM of all Trade Payables sub-rows.
- WRONG: "Other Bank Balances" → F36 (Cash & Equivalents).
  RIGHT: F35 (Bank Balance) — earmarked / margin money goes here, not in cash equivalents.
- WRONG: "Cash and Bank Balances" (combined) → split half-half between F35 and F36.
  RIGHT: Put the FULL amount in F36 when not separately disclosed.
- WRONG: "Investment in Subsidiaries" under Current Assets → F38 (current investments).
  RIGHT: Read the section. Investments in subsidiaries/associates/JVs are non-current → F49.
- WRONG: A row with "Total Current Liabilities" → mapped to a cell.
  RIGHT: section_total or _skip — totals are NEVER assigned to F-cells (would double-count).
- WRONG: A section-header row like "Non-Current Assets" with no value → mapped to a cell.
  RIGHT: _skip. Section headers are organizational, not data.
"""

PL_NEGATIVE_EXAMPLES = """COMMON ERRORS — DO NOT MAKE THESE:
- WRONG: "Other Income" → F66 (Revenue).
  RIGHT: other_income field. F66 is ONLY 'Revenue from Operations'.
- WRONG: "Total Income" (Revenue + Other Income) → F66.
  RIGHT: _skip. Use the underlying Revenue from Operations row only.
- WRONG: "Profit before tax" → net_profit.
  RIGHT: _skip. net_profit is the AFTER-tax figure.
- WRONG: "Other comprehensive income" → F66 or other_income.
  RIGHT: _skip. OCI is below net profit and not part of P&L mapping here.
- WRONG: "Impairment of assets" → F71 (Depreciation).
  RIGHT: _skip (or other_expense, computed as residual via F72).
- WRONG: "Earnings per share / Basic / Diluted" → mapped to a cell.
  RIGHT: _skip. EPS is a derived metric, not an expense.
- WRONG: Tax components ("Current tax", "Deferred tax", "Tax of earlier years") split across cells.
  RIGHT: ALL go to F73. F73 is the SUM of every tax line in the P&L.
- WRONG: "Cost of materials consumed" + "Changes in inventories" both put in F67 separately when extract also has a "Cost of revenue" total.
  RIGHT: Use the components, not the total. Don't double-count.
"""


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Item formatting (section-aware — point #3)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

_SECTION_LABEL = {
    "networth":         "EQUITY / NETWORTH",
    "current_liab":     "CURRENT LIABILITIES",
    "noncurrent_liab":  "NON-CURRENT LIABILITIES",
    "current_asset":    "CURRENT ASSETS",
    "noncurrent_asset": "NON-CURRENT ASSETS",
}


def _format_bs_items_grouped(items: List[Dict]) -> str:
    """Format BS items grouped by section, preserving the structural context
    the LLM needs to disambiguate "Others"-style labels."""
    enriched = _detect_sections(items)
    by_section: Dict[Optional[str], List[Dict]] = defaultdict(list)
    for it in enriched:
        if it["cur"] == 0 and it["pri"] == 0:
            continue
        by_section[it.get("_section")].append(it)

    lines: List[str] = []
    # Emit known sections in canonical BS order
    for sec in ["networth", "current_liab", "noncurrent_liab",
                "current_asset", "noncurrent_asset"]:
        if sec not in by_section:
            continue
        lines.append(f"\n=== SECTION: {_SECTION_LABEL[sec]} (section_code={sec}) ===")
        for it in by_section[sec]:
            lines.append(f"  ROW [{it['label'][:120]}] CY={it['cur']} PY={it['pri']}")
    # Items where section detection failed go into UNKNOWN block — model must infer
    if None in by_section:
        lines.append("\n=== SECTION: UNKNOWN (model must infer section_code from label) ===")
        for it in by_section[None]:
            lines.append(f"  ROW [{it['label'][:120]}] CY={it['cur']} PY={it['pri']}")
    return "\n".join(lines)


def _format_pl_items(items: List[Dict]) -> str:
    lines = []
    for it in items:
        if it["cur"] == 0 and it["pri"] == 0:
            continue
        lines.append(f"  ROW [{it['label'][:120]}] CY={it['cur']} PY={it['pri']}")
    return "\n".join(lines)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Strict JSON Schemas (point #2)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _bs_schema() -> Dict:
    section_codes = list(BS_SECTION_CELLS.keys()) + ["section_total", "_skip"]
    target_cells = ALL_BS_CELLS + ["_skip"]
    return {
        "name": "bs_mapping",
        "strict": True,
        "schema": {
            "type": "object",
            "additionalProperties": False,
            "required": ["mappings", "balance_check_note"],
            "properties": {
                "balance_check_note": {
                    "type": "string",
                    "description": "Brief note on whether Total Assets ≈ Total L+E based on your mappings (one sentence)."
                },
                "mappings": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["source_label", "section_code", "target_cell",
                                     "current_year_value", "prior_year_value", "rationale"],
                        "properties": {
                            "source_label": {
                                "type": "string",
                                "description": "Verbatim label from the input ROW."
                            },
                            "section_code": {
                                "type": "string",
                                "enum": section_codes,
                                "description": "Section the row belongs to. Use 'section_total' for total/subtotal rows; '_skip' for headers/junk."
                            },
                            "target_cell": {
                                "type": "string",
                                "enum": target_cells,
                                "description": "F-cell to assign this row's value to. Must be a cell within target_cell's section. Use '_skip' for total/header rows."
                            },
                            "current_year_value": {"type": "number"},
                            "prior_year_value":   {"type": "number"},
                            "rationale": {
                                "type": "string",
                                "description": "One short sentence explaining the choice — the section_code and a key phrase from the label."
                            }
                        }
                    }
                }
            }
        }
    }


def _pl_schema() -> Dict:
    target_cells = PL_TARGET_KEYS + ["_skip", "other_expense"]
    return {
        "name": "pl_mapping",
        "strict": True,
        "schema": {
            "type": "object",
            "additionalProperties": False,
            "required": ["mappings"],
            "properties": {
                "mappings": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["source_label", "target_cell",
                                     "current_year_value", "prior_year_value", "rationale"],
                        "properties": {
                            "source_label": {"type": "string"},
                            "target_cell":  {"type": "string", "enum": target_cells},
                            "current_year_value": {"type": "number"},
                            "prior_year_value":   {"type": "number"},
                            "rationale": {"type": "string"}
                        }
                    }
                }
            }
        }
    }


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Prompt builders (cacheable static prefix; dynamic items at the end)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _bs_template_block() -> str:
    lines = ["BS TEMPLATE FIELDS (target_cell candidates):"]
    for sec in ["networth", "current_liab", "noncurrent_liab", "current_asset", "noncurrent_asset"]:
        lines.append(f"\n{_SECTION_LABEL[sec]}  (section_code={sec})")
        for cell in BS_SECTION_CELLS[sec]:
            lines.append(f"  {cell}: {BS_CELL_DESC[cell]}")
    return "\n".join(lines)


def _pl_template_block() -> str:
    lines = ["PL TEMPLATE FIELDS (target_cell candidates):"]
    for cell, desc in PL_CELL_DESC.items():
        lines.append(f"  {cell}: {desc}")
    return "\n".join(lines)


def _bs_system_prompt() -> str:
    return (
        "You are an expert Indian financial analyst (Ind AS + Old GAAP) mapping "
        "balance-sheet line items into a fixed template. You are precise, "
        "section-aware, and you NEVER double-count or drop values.\n\n"
        "RULES:\n"
        "1. Every input ROW with non-zero values must appear in your `mappings`. "
        "If a row should not contribute to any F-cell (totals, section headers, junk), "
        "set target_cell='_skip' and section_code='_skip' or 'section_total'.\n"
        "2. The target_cell MUST be a cell that belongs to the row's section_code "
        "(e.g., a row in `noncurrent_asset` cannot map to a current_asset cell).\n"
        "3. Multi-row aggregations: if the source has 'Trade Payables — MSME dues' "
        "and 'Trade Payables — others' as separate rows, BOTH map to F14. The "
        "template will sum them automatically.\n"
        "4. Use the section_code from the section header above each row when present. "
        "When a row appears under SECTION: UNKNOWN, infer the section from the label.\n"
        "5. After mapping, mentally check: SUM(F5:F9) + SUM(F14:F19) + SUM(F24:F28) "
        "should ≈ SUM(F35:F43) + SUM(F48:F55). Note your check in `balance_check_note`.\n"
        "6. NEVER invent fields or cells outside the schema enums.\n\n"
        + BS_NEGATIVE_EXAMPLES + "\n"
        + _bs_template_block()
    )


def _pl_system_prompt() -> str:
    return (
        "You are an expert Indian financial analyst mapping P&L line items into a "
        "fixed template.\n\n"
        "RULES:\n"
        "1. Every input ROW with non-zero values must appear in your `mappings`.\n"
        "2. Headers, totals (e.g., 'Total Income', 'Profit before tax'), and OCI items "
        "must use target_cell='_skip'.\n"
        "3. F72 will be COMPUTED by the caller as a residual — DO NOT assign anything "
        "to F72. Other expenses that aren't COGS/Employee/Interest/Depreciation/Tax "
        "should use target_cell='other_expense' (the caller folds this into F72).\n"
        "4. Tax components (current tax, deferred tax, tax of earlier years) ALL go to F73 — "
        "the template sums them.\n"
        "5. NEVER invent cells outside the schema enums.\n\n"
        + PL_NEGATIVE_EXAMPLES + "\n"
        + _pl_template_block()
    )


def _bs_user_prompt(items_block: str, retry_diagnosis: str = "") -> str:
    parts = [
        "INPUT — extracted balance sheet rows, grouped by detected section:\n",
        items_block,
        "\n\nReturn JSON matching the bs_mapping schema. Cover EVERY non-zero input row "
        "exactly once. Use _skip for totals/headers.",
    ]
    if retry_diagnosis:
        parts.append("\n\n=== PREVIOUS ATTEMPT FAILED A CONSTRAINT ===\n")
        parts.append(retry_diagnosis)
        parts.append("\nFix the specific assignments named above. Keep the rest unchanged "
                     "unless the diagnosis says otherwise.")
    return "".join(parts)


def _pl_user_prompt(items_block: str, retry_diagnosis: str = "") -> str:
    parts = [
        "INPUT — extracted P&L rows:\n",
        items_block,
        "\n\nReturn JSON matching the pl_mapping schema. Cover EVERY non-zero input row "
        "exactly once. Use _skip for totals/headers/OCI.",
    ]
    if retry_diagnosis:
        parts.append("\n\n=== PREVIOUS ATTEMPT FAILED A CONSTRAINT ===\n")
        parts.append(retry_diagnosis)
        parts.append("\nFix the specific assignments named above.")
    return "".join(parts)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# OpenAI call with strict structured outputs (point #2)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _call_structured(client, model: str, system: str, user: str,
                     schema: Dict, temperature: float = 0.0,
                     max_tokens: int = 4000) -> Optional[Dict]:
    """One call with strict JSON-schema response_format. Returns parsed dict or None."""
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            response_format={"type": "json_schema", "json_schema": schema},
            temperature=temperature,
            max_tokens=max_tokens,
        )
        content = resp.choices[0].message.content
        if not content:
            return None
        return json.loads(content)
    except Exception as e:
        print(f"    [llm_mapper_v2] structured call failed: {type(e).__name__}: {e}")
        return None


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Aggregation: per-row mappings → per-cell sums
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _aggregate_bs(parsed: Dict) -> Dict[str, Dict[str, float]]:
    cy = {c: 0.0 for c in ALL_BS_CELLS}
    py = {c: 0.0 for c in ALL_BS_CELLS}
    for m in parsed.get("mappings", []):
        cell = m.get("target_cell", "_skip")
        if cell == "_skip" or cell not in cy:
            continue
        cy[cell] += float(m.get("current_year_value", 0) or 0)
        py[cell] += float(m.get("prior_year_value", 0) or 0)
    return {"current_year": {k: round(v, 2) for k, v in cy.items()},
            "prior_year":   {k: round(v, 2) for k, v in py.items()}}


def _aggregate_pl(parsed: Dict) -> Dict[str, Dict[str, float]]:
    cy = {k: 0.0 for k in PL_TARGET_KEYS}
    py = {k: 0.0 for k in PL_TARGET_KEYS}
    other_cy, other_py = 0.0, 0.0
    for m in parsed.get("mappings", []):
        cell = m.get("target_cell", "_skip")
        cv = float(m.get("current_year_value", 0) or 0)
        pv = float(m.get("prior_year_value", 0) or 0)
        if cell == "_skip":
            continue
        if cell == "other_expense":
            other_cy += cv
            other_py += pv
        elif cell in cy:
            cy[cell] += cv
            py[cell] += pv
    cy["other_expense_total"] = other_cy
    py["other_expense_total"] = other_py
    return {"current_year": {k: round(v, 2) for k, v in cy.items()},
            "prior_year":   {k: round(v, 2) for k, v in py.items()}}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Constraint checks (point #4)
#
# Each check returns (ok: bool, diagnosis: str). Diagnosis is the precise
# message we feed back to the model on retry — it names the failing rows
# and the arithmetic gap.
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _bs_balance_check(agg: Dict[str, Dict[str, float]], year: str,
                      tol: float) -> Tuple[bool, str]:
    m = agg[year]
    nw  = sum(m.get(c, 0) for c in BS_SECTION_CELLS["networth"])
    cl  = sum(m.get(c, 0) for c in BS_SECTION_CELLS["current_liab"])
    ncl = sum(m.get(c, 0) for c in BS_SECTION_CELLS["noncurrent_liab"])
    ca  = sum(m.get(c, 0) for c in BS_SECTION_CELLS["current_asset"])
    nca = sum(m.get(c, 0) for c in BS_SECTION_CELLS["noncurrent_asset"])
    assets = round(ca + nca, 2)
    liab   = round(nw + cl + ncl, 2)
    diff   = round(assets - liab, 2)
    if abs(diff) < tol:
        return True, ""
    sign = "Assets > Liab+Equity" if diff > 0 else "Liab+Equity > Assets"
    diag = (
        f"Balance check FAILED for {year}: Total Assets={assets:.2f}, "
        f"Total Equity+Liab={liab:.2f}, gap={diff:+.2f} ({sign}).\n"
        f"  - Networth (F5:F9)        = {nw:.2f}\n"
        f"  - Current Liab (F14:F19)  = {cl:.2f}\n"
        f"  - NC Liab (F24:F28)       = {ncl:.2f}\n"
        f"  - Current Assets (F35:F43)= {ca:.2f}\n"
        f"  - NC Assets (F48:F55)     = {nca:.2f}\n"
        f"Look for rows that may have been put in the wrong section "
        f"(e.g., a non-current item placed in a current cell or vice versa). "
        f"A gap of ~{abs(diff):.2f} often means a single misclassified row of that magnitude."
    )
    return False, diag


def _coverage_check(items: List[Dict], parsed: Dict, side: str) -> Tuple[bool, str]:
    """Every input row with non-zero values must appear in mappings."""
    src_labels = [it["label"].strip() for it in items
                  if (it["cur"] != 0 or it["pri"] != 0)]
    mapped_labels = [m.get("source_label", "").strip()
                     for m in parsed.get("mappings", [])]
    mapped_count = Counter(mapped_labels)
    missing = []
    for lbl in src_labels:
        if mapped_count.get(lbl, 0) == 0:
            # tolerate small differences (whitespace / casing)
            found = any(lbl.lower() == ml.lower() for ml in mapped_labels)
            if not found:
                missing.append(lbl[:80])
    if not missing:
        return True, ""
    diag = (
        f"Coverage check FAILED ({side}): {len(missing)} source rows are missing "
        f"from your `mappings` array. Add an entry for each (use target_cell='_skip' "
        f"if it's a total/header). Missing rows:\n"
        + "\n".join(f"  - {m}" for m in missing[:15])
    )
    return False, diag


def _section_consistency_check(parsed: Dict) -> Tuple[bool, str]:
    """target_cell must belong to its declared section_code (BS only)."""
    bad = []
    for m in parsed.get("mappings", []):
        sec = m.get("section_code", "")
        cell = m.get("target_cell", "")
        if cell == "_skip" or sec in ("_skip", "section_total"):
            continue
        allowed = BS_SECTION_CELLS.get(sec, [])
        if cell not in allowed:
            bad.append(f"  - row '{m.get('source_label','')[:60]}' "
                       f"section={sec} but target_cell={cell} "
                       f"(allowed for that section: {allowed[:6]}...)")
    if not bad:
        return True, ""
    diag = "Section/cell consistency FAILED:\n" + "\n".join(bad[:10])
    return False, diag


def _section_total_cross_check(items: List[Dict],
                               agg: Dict[str, Dict[str, float]],
                               year: str, tol: float) -> Tuple[bool, str]:
    """If the extract contains a 'Total Current Liabilities' (etc.) row, our
    aggregated sum should match it. This catches missed/double-counted rows."""
    label_for_section = {
        r"^total\s+(non.?\s*current\s+asset|fixed\s+asset|long.?\s*term\s+asset)":
            ("noncurrent_asset", "Non-current Assets"),
        r"^total\s+current\s+asset":   ("current_asset", "Current Assets"),
        r"^total\s+(non.?\s*current\s+liabilit|long.?\s*term\s+liabilit)":
            ("noncurrent_liab", "Non-current Liabilities"),
        r"^total\s+current\s+liabilit": ("current_liab", "Current Liabilities"),
        r"^total\s+(equity|networth|shareholder)": ("networth", "Equity / Networth"),
    }
    issues = []
    val_key = "cur" if year == "current_year" else "pri"
    m_year = agg[year]
    for it in items:
        lbl = it["label"].lower().strip()
        for pat, (sec, human) in label_for_section.items():
            if re.match(pat, lbl) and (it["cur"] != 0 or it["pri"] != 0):
                expected = float(it[val_key])
                got = sum(m_year.get(c, 0) for c in BS_SECTION_CELLS[sec])
                if abs(expected - got) > tol:
                    issues.append(
                        f"  - {human} ({year}): extract reports total={expected:.2f} "
                        f"but your mappings sum to {got:.2f} (gap {expected-got:+.2f}). "
                        f"This usually means a row in this section was missed or sent elsewhere."
                    )
                break
    if not issues:
        return True, ""
    return False, "Section-total cross-check FAILED:\n" + "\n".join(issues[:8])


def _run_bs_constraints(items: List[Dict], parsed: Dict,
                        agg: Dict[str, Dict[str, float]],
                        tol: float) -> List[str]:
    """Return list of failure diagnoses (empty list = all passed)."""
    diags = []
    ok, d = _section_consistency_check(parsed)
    if not ok:
        diags.append(d)
    ok, d = _coverage_check(items, parsed, "BS")
    if not ok:
        diags.append(d)
    for year in ("current_year", "prior_year"):
        ok, d = _bs_balance_check(agg, year, tol)
        if not ok:
            diags.append(d)
        ok, d = _section_total_cross_check(items, agg, year, tol)
        if not ok:
            diags.append(d)
    return diags


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Public API — drop-in replacements for map_bs_hybrid / map_pl_hybrid
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def map_bs_v2(items: List[Dict]) -> Optional[Dict]:
    """Map balance-sheet items via strict-structured-output LLM with constraint
    checks and targeted re-prompts. Returns mapping shape compatible with the
    existing pipeline, or None if the LLM is unavailable."""
    client, cfg = _get_client()
    if client is None:
        print("    [llm_mapper_v2] No OpenAI key configured; cannot run.")
        return None

    model = cfg.get("llm_mapper_model", DEFAULT_MODEL)
    tol   = float(cfg.get("balance_tolerance", BALANCE_TOL))
    enable_critic = bool(cfg.get("enable_critic", False))

    items_block = _format_bs_items_grouped(items)
    if not items_block.strip():
        print("    [llm_mapper_v2] No non-zero BS rows; skipping.")
        return None

    system = _bs_system_prompt()
    schema = _bs_schema()

    parsed = None
    diagnoses_history: List[str] = []
    for round_idx in range(MAX_RETRY_ROUNDS + 1):
        retry_diag = "\n\n".join(diagnoses_history[-1:]) if diagnoses_history else ""
        user = _bs_user_prompt(items_block, retry_diag)
        print(f"    [llm_mapper_v2] BS round {round_idx+1}: model={model}")
        parsed = _call_structured(client, model, system, user, schema)
        if not parsed:
            break

        agg = _aggregate_bs(parsed)
        diags = _run_bs_constraints(items, parsed, agg, tol)
        if not diags:
            print(f"    [llm_mapper_v2] BS round {round_idx+1}: ALL CONSTRAINTS PASS")
            break
        if round_idx >= MAX_RETRY_ROUNDS:
            print(f"    [llm_mapper_v2] BS round {round_idx+1}: still failing — keeping best effort")
            break
        # Combine all diagnoses into one re-prompt
        combined = "\n\n".join(diags)
        diagnoses_history.append(combined)
        print(f"    [llm_mapper_v2] BS round {round_idx+1}: {len(diags)} constraint(s) failed; retrying")

    if not parsed:
        return None

    # Optional critic pass (point #6) — only after constraints pass, to catch
    # silent errors that don't violate arithmetic but look semantically wrong
    if enable_critic:
        crit = _critic_pass(client, "BS", items_block, parsed,
                            cfg.get("critic_model", DEFAULT_CRITIC_MODEL))
        if crit:
            parsed = _apply_critic_fixes(parsed, crit)

    agg = _aggregate_bs(parsed)
    warnings = _collect_bs_warnings(items, parsed, agg, tol)

    mapping = {
        "current_year": agg["current_year"],
        "prior_year":   agg["prior_year"],
        "notes":        f"llm_mapper_v2 ({model}); rounds={round_idx+1}",
        "_warnings":    warnings,
        "_per_row":     parsed.get("mappings", []),  # for debugging / eval
    }
    _print_bs_summary(parsed, agg)
    return mapping


def map_pl_v2(items: List[Dict]) -> Optional[Dict]:
    """LLM-driven P&L mapping. F72 is computed as a residual by the caller."""
    client, cfg = _get_client()
    if client is None:
        print("    [llm_mapper_v2] No OpenAI key configured; cannot run.")
        return None

    model = cfg.get("llm_mapper_model", DEFAULT_MODEL)
    enable_critic = bool(cfg.get("enable_critic", False))

    items_block = _format_pl_items(items)
    if not items_block.strip():
        return None

    system = _pl_system_prompt()
    schema = _pl_schema()

    parsed = None
    diagnoses_history: List[str] = []
    for round_idx in range(MAX_RETRY_ROUNDS + 1):
        retry_diag = "\n\n".join(diagnoses_history[-1:]) if diagnoses_history else ""
        user = _pl_user_prompt(items_block, retry_diag)
        print(f"    [llm_mapper_v2] PL round {round_idx+1}: model={model}")
        parsed = _call_structured(client, model, system, user, schema)
        if not parsed:
            break

        diags = []
        ok, d = _coverage_check(items, parsed, "PL")
        if not ok:
            diags.append(d)
        if not diags:
            print(f"    [llm_mapper_v2] PL round {round_idx+1}: coverage OK")
            break
        if round_idx >= MAX_RETRY_ROUNDS:
            break
        diagnoses_history.append("\n\n".join(diags))
        print(f"    [llm_mapper_v2] PL round {round_idx+1}: coverage failed; retrying")

    if not parsed:
        return None

    if enable_critic:
        crit = _critic_pass(client, "PL", items_block, parsed,
                            cfg.get("critic_model", DEFAULT_CRITIC_MODEL))
        if crit:
            parsed = _apply_critic_fixes(parsed, crit)

    agg = _aggregate_pl(parsed)
    cy = agg["current_year"]
    py = agg["prior_year"]

    mapping = {
        "current_year": {
            "F66": cy["F66"], "F67": cy["F67"], "F69": cy["F69"],
            "F70": cy["F70"], "F71": cy["F71"], "F73": cy["F73"],
            "other_income": cy["other_income"],
            "net_profit":   cy["net_profit"],
            "exceptional":  cy["exceptional_items"],
            "exceptional_items": cy["exceptional_items"],
        },
        "prior_year": {
            "F66": py["F66"], "F67": py["F67"], "F69": py["F69"],
            "F70": py["F70"], "F71": py["F71"], "F73": py["F73"],
            "other_income": py["other_income"],
            "net_profit":   py["net_profit"],
            "exceptional":  py["exceptional_items"],
            "exceptional_items": py["exceptional_items"],
        },
        "notes":     f"llm_mapper_v2 ({model}); rounds={round_idx+1}",
        "_warnings": [],
        "_per_row":  parsed.get("mappings", []),
    }
    _print_pl_summary(parsed)
    return mapping


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Critic pass (point #6) — second LLM call hunts for likely mistakes
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

CRITIC_SCHEMA = {
    "name": "critic_review",
    "strict": True,
    "schema": {
        "type": "object",
        "additionalProperties": False,
        "required": ["likely_errors"],
        "properties": {
            "likely_errors": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "required": ["source_label", "current_target_cell",
                                 "suggested_target_cell", "reason"],
                    "properties": {
                        "source_label":         {"type": "string"},
                        "current_target_cell":  {"type": "string"},
                        "suggested_target_cell":{"type": "string"},
                        "reason":               {"type": "string"}
                    }
                }
            }
        }
    }
}


def _critic_pass(client, side: str, items_block: str, parsed: Dict,
                 critic_model: str) -> Optional[Dict]:
    sys_msg = (
        f"You are an expert reviewer of {side} mappings. You did NOT create the "
        "mapping below — your job is to find errors. Be skeptical. Only flag "
        "assignments that look CLEARLY wrong (section mismatch, double-counted, "
        "obvious miscategorization). Do not flag judgment calls."
    )
    user_msg = (
        f"INPUT ROWS:\n{items_block}\n\n"
        f"PROPOSED MAPPING:\n{json.dumps(parsed.get('mappings', []), indent=2)[:8000]}\n\n"
        "List up to 8 likely errors. Empty list if everything looks fine."
    )
    try:
        resp = client.chat.completions.create(
            model=critic_model,
            messages=[
                {"role": "system", "content": sys_msg},
                {"role": "user", "content": user_msg},
            ],
            response_format={"type": "json_schema", "json_schema": CRITIC_SCHEMA},
            temperature=0.0,
            max_tokens=2000,
        )
        return json.loads(resp.choices[0].message.content or "{}")
    except Exception as e:
        print(f"    [llm_mapper_v2] critic pass failed: {e}")
        return None


def _apply_critic_fixes(parsed: Dict, critic: Dict) -> Dict:
    fixes = critic.get("likely_errors", [])
    if not fixes:
        return parsed
    by_label = {m.get("source_label", "").strip(): m
                for m in parsed.get("mappings", [])}
    applied = 0
    for fx in fixes:
        m = by_label.get(fx.get("source_label", "").strip())
        if not m:
            continue
        new_cell = fx.get("suggested_target_cell")
        if not new_cell or new_cell == m.get("target_cell"):
            continue
        # Only apply if the new cell is in our schema
        if new_cell in ALL_BS_CELLS or new_cell in PL_TARGET_KEYS \
                or new_cell in ("_skip", "other_expense"):
            print(f"    [llm_mapper_v2 CRITIC] {m.get('source_label','')[:50]}: "
                  f"{m.get('target_cell')} → {new_cell} ({fx.get('reason','')[:80]})")
            m["target_cell"] = new_cell
            applied += 1
    if applied:
        print(f"    [llm_mapper_v2 CRITIC] applied {applied} fix(es)")
    return parsed


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Reporting helpers
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _print_bs_summary(parsed: Dict, agg: Dict):
    rows = parsed.get("mappings", [])
    skipped = sum(1 for r in rows if r.get("target_cell") == "_skip")
    mapped  = len(rows) - skipped
    print(f"    [llm_mapper_v2] BS: mapped {mapped} / skipped {skipped}")
    for r in rows:
        cell = r.get("target_cell", "_skip")
        if cell == "_skip":
            continue
        cy = r.get("current_year_value", 0)
        print(f"      • {cell:>4s}  {r.get('source_label','')[:55]:55s}  CY={cy}")


def _print_pl_summary(parsed: Dict):
    rows = parsed.get("mappings", [])
    skipped = sum(1 for r in rows if r.get("target_cell") == "_skip")
    mapped  = len(rows) - skipped
    print(f"    [llm_mapper_v2] PL: mapped {mapped} / skipped {skipped}")
    for r in rows:
        cell = r.get("target_cell", "_skip")
        if cell == "_skip":
            continue
        cy = r.get("current_year_value", 0)
        print(f"      • {cell:>14s}  {r.get('source_label','')[:50]:50s}  CY={cy}")


def _collect_bs_warnings(items, parsed, agg, tol) -> List[str]:
    warns: List[str] = []
    for year in ("current_year", "prior_year"):
        ok, d = _bs_balance_check(agg, year, tol)
        if not ok:
            # Compress to single-line warning for the WARNINGS sheet
            first_line = d.split("\n", 1)[0]
            warns.append(f"v2: {first_line}")
        ok, d = _section_total_cross_check(items, agg, year, tol)
        if not ok:
            warns.append(f"v2: {d.split(chr(10),1)[0]}")
    return warns
