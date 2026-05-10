"""
claude_mapper.py — Claude-based BS/PL → template mapper.

Drop-in replacement for `llm_mapper_v2` that uses the Anthropic Claude API
instead of OpenAI. Same input shape (list of {label, cur, pri} dicts) and
same output shape ({"current_year": {...}, "prior_year": {...}, "_warnings": [...]})
so the rest of the pipeline (validation, F72 residual, Excel writer,
Airtable upload) is untouched.

Approach:
  - Strict structured output via Claude's tool-use mechanism. We define a
    `submit_bs_mapping` / `submit_pl_mapping` tool with a JSON-schema
    input_schema and force_tool_choice. Claude must respond with a tool_use
    block whose `input` matches the schema — this gives us the same
    correctness guarantees as OpenAI's response_format=json_schema.
  - Same negative examples + section-conditional cell enums as v2 so accuracy
    on Indian Ind-AS / Old GAAP statements is preserved.
  - Programmatic post-checks (balance, coverage, section consistency) with
    a targeted retry round when constraints fail.

Config (config.json):
  anthropic_api_key      — sk-ant-... key
  claude_model           — defaults to "claude-sonnet-4-6"
  balance_tolerance      — |Assets − L+E| tolerance, default 1.0
"""
from __future__ import annotations

import json
import os
import re
from collections import Counter, defaultdict
from typing import Any, Dict, List, Optional, Tuple

# Reuse the existing section detector and cell maps so we keep the hard-won
# routing knowledge from the prior 71-company audit.
from matcher import _detect_sections


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Defaults
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

DEFAULT_MODEL = "claude-sonnet-4-6"
BALANCE_TOL   = 1.0
# Text-only retry budget. After exhausting these, if a PDF is supplied we
# escalate to ONE PDF-grounded remediation pass that lets Claude re-read
# the source statement directly. This catches the failure mode where the
# extractor missed a row (Finolex equity gap of ₹471 Cr), which no amount
# of text-only retrying can fix because the data isn't in the items list.
#
# 2026-05 tuning: dropped from 2 to 1. With the new "blocking vs warning"
# split (only balance failure / coverage / section-consistency block;
# section-total cross-check is informational), retries fire much less
# often, so a single retry round is enough. Saves ~30-60s per job in the
# benign-but-noisy case (Shivam-style: BS balanced but cross-check off
# by rounding).
MAX_RETRY_ROUNDS = 1

# Section-total cross-check tolerance, expressed as a fraction of the
# section's own size. Anything inside this band is treated as benign
# rounding (published total vs sum-of-parts disagreement) and surfaces
# as a WARNING rather than a retry trigger.
SECTION_TOTAL_REL_TOL = 0.005   # 0.5% of section magnitude


def _load_cfg() -> Dict[str, Any]:
    try:
        from config import load_config
        return load_config()
    except Exception:
        return {}


def _get_client():
    cfg = _load_cfg()
    api_key = (cfg.get("anthropic_api_key")
               or os.environ.get("ANTHROPIC_API_KEY", "")).strip()
    if not api_key:
        return None, cfg
    try:
        from anthropic import Anthropic
        return Anthropic(api_key=api_key), cfg
    except ImportError:
        print("    [claude_mapper] anthropic SDK not installed. "
              "Run: pip install anthropic")
        return None, cfg


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Section → cell catalogue (must stay identical to llm_mapper_v2)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BS_SECTION_CELLS: Dict[str, List[str]] = {
    # No "spare" / catch-all cells — every source row must land in a
    # named bucket. F9 / F28 / F43 / F55 were dropped per user feedback;
    # NCI now folds into F8 (Other Equity) for consolidated reports.
    "networth":         ["F5", "F6", "F7", "F8"],
    "current_liab":     ["F14", "F15", "F16", "F17", "F18", "F19"],
    "noncurrent_liab":  ["F24", "F25", "F26", "F27"],
    "current_asset":    ["F35", "F36", "F37", "F38", "F39", "F40", "F41", "F42"],
    "noncurrent_asset": ["F48", "F49", "F50", "F51", "F52", "F53", "F54"],
}

ALL_BS_CELLS: List[str] = sorted({c for cells in BS_SECTION_CELLS.values() for c in cells})

PL_TARGET_KEYS: List[str] = [
    "F66", "F67", "F69", "F70", "F71", "F73",
    "other_income", "net_profit", "exceptional_items",
]

BS_CELL_DESC: Dict[str, str] = {
    "F5":  "Equity Share Capital (paid-up)",
    "F6":  "Retained Earnings (only if separately disclosed)",
    "F7":  "General Reserves & Surplus (capital reserve, securities premium, statutory, revaluation)",
    "F8":  "Other Equity (Ind AS line; folds reserves/RE if not separate; share warrants; NCI in consolidated)",
    "F14": "Trade Payables (SUM all sub-items: MSME + non-MSME)",
    "F15": "Provisions current (short-term, employee benefits current)",
    "F16": "Short Term Borrowings (incl current maturities of LT debt, bank OD, bills payable)",
    "F17": "Other Current Liab (statutory dues, advances, contract liab, current tax liab, unearned revenue; DTL only if source classifies current)",
    "F18": "Other Financial Liab CURRENT + Lease Liab CURRENT (Ind AS 116 — never F19)",
    "F19": "DEPRECATED — leave 0; lease folds into F18",
    "F24": "Long Term Borrowings (term loans, debentures, secured/unsecured)",
    "F25": "Provisions NC (gratuity, leave encashment)",
    "F26": "Others NC liab = DTL + Other NCL + Govt grant NC. (Lease NC → F27 not here)",
    "F27": "Other Financial Liab NC + Lease Liab NC (Ind AS 116; never F26)",
    "F35": "Bank Balance (Other Bank, earmarked, margin). Combined Cash+Bank → F36 entire.",
    "F36": "Cash & Cash Equivalents (incl combined Cash+Bank if not split)",
    "F37": "Inventory (finished goods, stores, stock-in-trade; 0 for IT/services)",
    "F38": "Investments current + held-for-sale investments",
    "F39": "Loans current (ICDs, advances, security deposits current)",
    "F40": "Trade Receivables / Sundry Debtors / Contract Assets",
    "F41": "Other Current Assets (Current Tax Assets, advance tax, TDS, GST input, prepaid, held-for-sale)",
    "F42": "Other Financial Assets current (incl current-portion lease receivable)",
    "F48": "Fixed Assets = PPE + Tangible + Intangible + ROU + Goodwill + Software + Investment Property. EXCLUDE CWIP (→F51).",
    "F49": "Investments NC (long-term, Investment Property, Inv in Subs/Assoc/JVs when in NC Financial Assets)",
    "F50": "Loans NC (security deposits NC, long-term advances)",
    "F51": "CWIP ONLY (Intangible AUD → F52)",
    "F52": "Other NCA (capital advances, Intangible AUD, Biological Assets, Inv in Subs if source places here). NEVER tax assets — those → F54.",
    "F53": "Other Financial Assets NC",
    "F54": "ALL NC tax assets ONE bucket: DTA + MAT credit + Non-current Tax Assets / Income Tax Assets (net). Never split with F52.",
}

PL_CELL_DESC: Dict[str, str] = {
    "F66": "Revenue from Operations ONLY (Other Income separate)",
    "F67": "COGS: Manufacturing→Materials+Changes in Inv FG/WIP; Trading→Purchase stock-in-trade+Changes; Services→main cost line; else 0",
    "F69": "Employee Benefits (staff cost, salaries, PF, managerial rem)",
    "F70": "Finance Costs / Interest expense / borrowing cost",
    "F71": "Depreciation + Amortization (combined; incl depletion; EXCLUDE impairment)",
    "F73": "Tax Expense SUM: Current + Deferred + Earlier Years + MAT credit (signed)",
    "other_income": "Other Income (interest, dividend, rental, gain on sale)",
    "net_profit": "Final Profit after tax (after exceptional, BEFORE OCI)",
    "exceptional_items": "Exceptional / extraordinary / prior-period adj. Signed: +income, -expense.",
}


BS_NEGATIVE_EXAMPLES = """COMMON ERRORS — DO NOT MAKE:
- "Others" under NC Financial Assets → F53 (NC), NOT F41 (current). Section header drives.
- CWIP → F51, NOT F48. F48 is completed PPE only.
- DTL: source-section drives — F26 if NC (standard), F17 if source classifies current. NEVER F25.
- Current Tax Liab → F17, NEVER F25/F26.
- Inv in Subs/Assoc/JV → F49 if in NC Financial Assets section; F52 only if source explicitly places there.
- Intangible AUD → F52 (Other NCA), NOT F51.
- Non-current Tax Assets / Income Tax Assets (net) → F54 ALWAYS (single tax bucket). NEVER F52.
- Lease Liab Current → F18, NEVER F19. Lease Liab NC → F27, NEVER F26. (Ind AS 116)
- Trade Payables MSME + non-MSME → BOTH to F14 (template sums).
- Other Bank Balances → F35 (earmarked/margin), NOT F36 (cash eq).
- Combined "Cash and Bank Balances" → entire amount to F36, do NOT split.
- Money Received Against Share Warrants → F8 (Other Equity).
- "Total <section>" / "Non-Current Assets" header rows → _skip (would double-count).
"""

PL_NEGATIVE_EXAMPLES = """COMMON ERRORS — DO NOT MAKE:
- "Other Income" → other_income, NEVER F66 (which is Rev from Ops only).
- "Total Income" (Rev+OI) → _skip; use underlying Rev row.
- "Profit before tax" → _skip (net_profit is AFTER tax).
- OCI / "Other comprehensive income" → _skip (below NP).
- "Impairment of assets" → other_expense, NEVER F71.
- EPS Basic/Diluted → _skip.
- Tax components (Current + Deferred + Earlier + MAT) ALL → F73 (one bucket).
"""


_SECTION_LABEL = {
    "networth":         "EQUITY / NETWORTH",
    "current_liab":     "CURRENT LIABILITIES",
    "noncurrent_liab":  "NON-CURRENT LIABILITIES",
    "current_asset":    "CURRENT ASSETS",
    "noncurrent_asset": "NON-CURRENT ASSETS",
}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Item formatting (section-aware)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _format_bs_items_grouped(items: List[Dict]) -> str:
    enriched = _detect_sections(items)
    by_section: Dict[Optional[str], List[Dict]] = defaultdict(list)
    for it in enriched:
        if it["cur"] == 0 and it["pri"] == 0:
            continue
        by_section[it.get("_section")].append(it)

    lines: List[str] = []
    for sec in ["networth", "current_liab", "noncurrent_liab",
                "current_asset", "noncurrent_asset"]:
        if sec not in by_section:
            continue
        lines.append(f"\n=== SECTION: {_SECTION_LABEL[sec]} (section_code={sec}) ===")
        for it in by_section[sec]:
            lines.append(f"  ROW [{it['label'][:120]}] CY={it['cur']} PY={it['pri']}")
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
# Tool input_schemas (Claude's structured-output contract)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _bs_tool() -> Dict:
    section_codes = list(BS_SECTION_CELLS.keys()) + ["section_total", "_skip"]
    target_cells = ALL_BS_CELLS + ["_skip"]
    return {
        "name": "submit_bs_mapping",
        "description": "Submit the balance-sheet mapping. Each input row must "
                       "appear exactly once. Use _skip for headers/totals.",
        "input_schema": {
            "type": "object",
            "required": ["mappings", "balance_check_note"],
            "properties": {
                "balance_check_note": {
                    "type": "string",
                    "description": "One sentence on whether Total Assets ≈ Total Equity+Liab based on your mappings."
                },
                "mappings": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["source_label", "section_code", "target_cell",
                                     "current_year_value", "prior_year_value", "rationale"],
                        "properties": {
                            "source_label":        {"type": "string"},
                            "section_code":        {"type": "string", "enum": section_codes},
                            "target_cell":         {"type": "string", "enum": target_cells},
                            "current_year_value":  {"type": "number"},
                            "prior_year_value":    {"type": "number"},
                            "rationale":           {"type": "string"}
                        }
                    }
                }
            }
        }
    }


def _pl_tool() -> Dict:
    target_cells = PL_TARGET_KEYS + ["_skip", "other_expense"]
    return {
        "name": "submit_pl_mapping",
        "description": "Submit the P&L mapping. Each non-zero input row must appear exactly once.",
        "input_schema": {
            "type": "object",
            "required": ["mappings"],
            "properties": {
                "mappings": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["source_label", "target_cell",
                                     "current_year_value", "prior_year_value", "rationale"],
                        "properties": {
                            "source_label":        {"type": "string"},
                            "target_cell":         {"type": "string", "enum": target_cells},
                            "current_year_value":  {"type": "number"},
                            "prior_year_value":    {"type": "number"},
                            "rationale":           {"type": "string"}
                        }
                    }
                }
            }
        }
    }


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Prompt builders
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _bs_template_block() -> str:
    lines = ["BS TEMPLATE FIELDS (target_cell candidates):"]
    for sec in ["networth", "current_liab", "noncurrent_liab",
                "current_asset", "noncurrent_asset"]:
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
        "You are an expert Indian CA mapping BS line items into fixed template "
        "cells. Precise, section-aware, never double-count or drop values.\n\n"
        "PRINCIPLES:\n"
        "1. Equation MUST hold: SUM(F35:F42)+SUM(F48:F54) = SUM(F5:F8)+SUM(F14:F18)+SUM(F24:F27). "
        "No spare cells — every non-zero row lands in a named bucket.\n"
        "2. Source-section wins over Schedule III pedantry. The audited "
        "report's classification (current vs NC) is final. Read the section "
        "header above each row.\n"
        "3. Signs: parens = negative; pass through. Accumulated losses → F8 negative.\n"
        "4. Multi-row aggregations: e.g. Trade Payables MSME + non-MSME both → F14 (template sums).\n"
        "5. Contra items (\"Less: provision for doubtful debts\"): use the NET parent, _skip the contra.\n"
        "6. Current maturities of LT debt → F16 (current), NOT F24.\n"
        "7. Goodwill on consolidation → F48. NCI → F8 (no separate NCI cell).\n"
        "8. Pass values to 2dp; do NOT round/scale.\n\n"
        "OUTPUT RULES:\n"
        "- Every non-zero input row must appear in `mappings`. Use target_cell='_skip' "
        "(section_code='_skip' or 'section_total') for totals/headers/junk.\n"
        "- target_cell MUST belong to the row's section_code.\n"
        "- Self-check arithmetic before submitting; state check in `balance_check_note` "
        "(both years' Assets, L+E, gap).\n"
        "- If gap > ₹1 lakh, FIX before submitting. Common: contra double-counted, "
        "LT-debt current maturity in F24 not F16, DTL in F25 not F26.\n"
        "- NEVER invent fields outside the schema enums. Always call submit_bs_mapping.\n\n"
        + BS_NEGATIVE_EXAMPLES + "\n"
        + _bs_template_block()
    )


def _pl_system_prompt() -> str:
    return (
        "You are an expert Indian CA mapping P&L line items. Never conflate "
        "Revenue with Other Income, double-count tax, or include OCI in NP.\n\n"
        "PRINCIPLES:\n"
        "1. Identity: NP = Rev + OI − COGS − Emp − Fin − Dep − OE ± Exc − Tax.\n"
        "2. F66 = Revenue from Operations ONLY. Other Operating Revenue stays inside F66; "
        "Other Income → other_income field.\n"
        "3. COGS by entity: Manufacturing→Materials+Purchase+Changes in Inv FG/WIP; "
        "Trading→Purchase+Changes; Services→0 (operating cost falls to other_expense).\n"
        "4. Signs: parens = negative. Changes in Inv: pass through published sign.\n"
        "5. F70 = expense only; interest INCOME → other_income.\n"
        "6. F71 = Dep + Amort + depletion. EXCLUDE impairment (→other_expense).\n"
        "7. F73 SUMS all tax lines: Current + Deferred (signed) + Earlier Years + MAT.\n"
        "8. Exceptional → exceptional_items (signed). Do NOT also bucket in F67/F69/F70/F71/other_expense.\n"
        "9. OCI / Remeasurement / 'Items reclassified' → _skip.\n"
        "10. net_profit = published 'Profit for the year' (AFTER tax, BEFORE OCI). NOT PBT, NOT Total Comp Inc.\n\n"
        "OUTPUT RULES:\n"
        "- Every non-zero input row in `mappings`. Headers/sub-totals/OCI → _skip.\n"
        "- Tax components ALL → F73 with correct signs.\n"
        "- F72 is NOT a target — use other_expense / other_income; caller computes F72 = OE − OI.\n"
        "- Self-check NP from identity (1) for both years, ±₹1 lakh of published.\n"
        "- NEVER invent cells outside schema enums. Always call submit_pl_mapping.\n\n"
        + PL_NEGATIVE_EXAMPLES + "\n"
        + _pl_template_block()
    )


def _format_current_mapping(parsed: Dict) -> str:
    """Render the current per-row mapping as a compact text table for the
    remediation prompt, so Claude can see what its previous attempt did."""
    lines = ["CURRENT (BROKEN) MAPPING — fix these assignments:"]
    for m in parsed.get("mappings", []):
        cell = m.get("target_cell", "_skip")
        sec  = m.get("section_code", "?")
        cy   = m.get("current_year_value", 0)
        py   = m.get("prior_year_value", 0)
        lbl  = (m.get("source_label") or "")[:90]
        lines.append(f"  {cell:>6s} (sec={sec:18s}) CY={cy:>14}  PY={py:>14}  | {lbl}")
    return "\n".join(lines)


def _bs_remediation_user_prompt(items_block: str, parsed: Dict,
                                  diagnoses: str) -> str:
    """Prompt for the PDF-grounded remediation pass. Tells Claude to use
    the PDF as ground truth, shows it the current mapping, names the
    constraints that are failing, and asks for a corrected mapping."""
    return (
        "You have ALREADY mapped this balance sheet once but the result "
        "FAILS hard accounting constraints — see below. The published BS "
        "in the attached PDF is audited and tallies; therefore the gap "
        "is in YOUR mapping, not the source. Read the PDF directly, find "
        "what was misclassified or what the rule-based extractor missed, "
        "and submit a corrected mapping.\n\n"
        "RECONCILIATION DISCIPLINE — what a finance professional would do:\n"
        "1. Read the published 'Total Equity', 'Total Current Liabilities', "
        "'Total Non-Current Liabilities', 'Total Current Assets', "
        "'Total Non-Current Assets', 'TOTAL' rows directly from the PDF. "
        "These are HARD anchors.\n"
        "2. For each section whose mapped sum disagrees with the published "
        "total, identify which item is missing or misclassified. Common "
        "culprits:\n"
        "   - Money Received Against Share Warrants (goes in F8 Other Equity)\n"
        "   - Capital Reserve / Securities Premium / General Reserve "
        "(usually F7 if separately disclosed; else folded into F8)\n"
        "   - Lease Liabilities split current/non-current (F18 vs F27 per Ind AS 116; never F19/F26)\n"
        "   - Current maturities of LT debt → F16 (current borrowings), "
        "NOT F24\n"
        "   - DTL → F26 (Others NC liab) when source classifies it NC "
        "(standard); F17 if source places it in current\n"
        "   - Investment in Subs/Assoc → F49 if source has it under "
        "Non-Current Financial Assets; F52 if source parks it in Other NCA\n"
        "   - Current Tax Assets → F41 (current section) or F52 "
        "(non-current section) — follow the source\n"
        "   - Government Grant (deferred income) — current portion F17, "
        "non-current portion F26\n"
        "3. The accounting equation MUST hold within ₹1 lakh: "
        "Assets = Equity + Liabilities. Mentally verify before submitting.\n"
        "4. Cover EVERY non-zero source row exactly once. If the extracted "
        "items list is missing a row that the PDF clearly shows, ADD a "
        "new mapping entry for it (use the published value from the PDF).\n\n"
        + items_block
        + "\n\n"
        + _format_current_mapping(parsed)
        + "\n\n=== CONSTRAINTS THAT FAILED ===\n"
        + diagnoses
        + "\n\nReturn a CORRECTED mapping via the submit_bs_mapping tool. "
        "Use the PDF to add missing rows / fix misclassifications. State the "
        "balance check explicitly in `balance_check_note`."
    )


def _pl_remediation_user_prompt(items_block: str, parsed: Dict,
                                  diagnoses: str) -> str:
    return (
        "You have ALREADY mapped this P&L once but the result fails the "
        "P&L identity check (Net Profit ≡ Revenue + Other Income − COGS − "
        "Employee − Finance − Depreciation − Other Expense ± Exceptional "
        "− Tax). The published P&L in the attached PDF is audited; the "
        "gap is in YOUR mapping. Re-read the PDF, find the missing line "
        "or the wrong sign / wrong bucket, and submit a corrected mapping.\n\n"
        "FOCUS POINTS:\n"
        "- All tax components (current tax, deferred tax, MAT, tax of "
        "earlier years) must SUM into F73 with correct signs (deferred "
        "tax credit reduces F73).\n"
        "- Exceptional items go to exceptional_items (signed); never "
        "duplicated in any expense bucket.\n"
        "- OCI items (Remeasurement, Items reclassified) must be _skip — "
        "they are below net profit.\n"
        "- net_profit must be the published 'Profit for the Year' / "
        "'Profit after Tax' figure, NOT 'Profit before tax'.\n"
        "- Changes in Inventory of FG/WIP — pass through the published "
        "sign; do not flip.\n\n"
        + items_block
        + "\n\n"
        + _format_current_mapping(parsed)
        + "\n\n=== CONSTRAINTS THAT FAILED ===\n"
        + diagnoses
        + "\n\nReturn a CORRECTED mapping via the submit_pl_mapping tool."
    )


def _bs_user_prompt(items_block: str, retry_diagnosis: str = "") -> str:
    parts = [
        "INPUT — extracted balance sheet rows, grouped by detected section:\n",
        items_block,
        "\n\nReturn your answer by calling the submit_bs_mapping tool. Cover EVERY "
        "non-zero input row exactly once. Use _skip for totals/headers.",
    ]
    if retry_diagnosis:
        parts.append("\n\n=== PREVIOUS ATTEMPT FAILED A CONSTRAINT ===\n")
        parts.append(retry_diagnosis)
        parts.append("\nFix the named assignments. Keep the rest unchanged.")
    return "".join(parts)


def _pl_user_prompt(items_block: str, retry_diagnosis: str = "") -> str:
    parts = [
        "INPUT — extracted P&L rows:\n",
        items_block,
        "\n\nReturn your answer by calling the submit_pl_mapping tool. Cover EVERY "
        "non-zero input row exactly once. Use _skip for totals/headers/OCI.",
    ]
    if retry_diagnosis:
        parts.append("\n\n=== PREVIOUS ATTEMPT FAILED A CONSTRAINT ===\n")
        parts.append(retry_diagnosis)
        parts.append("\nFix the named assignments.")
    return "".join(parts)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Anthropic API call (one round)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _call_claude_tool_with_pdf(client, model: str, system_prompt: str,
                                 user_prompt: str, tool: Dict,
                                 pdf_bytes: bytes,
                                 max_tokens: int = 12000) -> Optional[Dict]:
    """Variant of `_call_claude_tool` that includes the source PDF as a
    document input block. Used for the PDF-grounded remediation pass when
    text-only retries can't satisfy constraints — typically because the
    extractor missed a row and Claude needs to re-read the source to find
    the missing classification.

    The PDF is sent with ephemeral cache control so back-to-back retries
    on the same document only pay for the parse once."""
    import base64 as _b64
    pdf_b64 = _b64.standard_b64encode(pdf_bytes).decode("ascii")
    try:
        resp = client.messages.create(
            model=model,
            max_tokens=max_tokens,
            system=system_prompt,
            tools=[tool],
            tool_choice={"type": "tool", "name": tool["name"]},
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": pdf_b64,
                        },
                        "cache_control": {"type": "ephemeral"},
                    },
                    {"type": "text", "text": user_prompt},
                ],
            }],
        )
    except Exception as e:
        print(f"    [claude_mapper] PDF remediation API call failed: "
              f"{type(e).__name__}: {e}")
        return None
    for block in resp.content:
        if getattr(block, "type", None) == "tool_use" \
                and getattr(block, "name", None) == tool["name"]:
            return block.input  # type: ignore[attr-defined]
    return None


def _call_claude_tool(client, model: str, system_prompt: str,
                      user_prompt: str, tool: Dict,
                      max_tokens: int = 8000) -> Optional[Dict]:
    """Force Claude to call the given tool. Return the tool's input dict."""
    try:
        resp = client.messages.create(
            model=model,
            max_tokens=max_tokens,
            system=[
                # Cache the static prefix — it's ~5K tokens of rules, negative
                # examples, and template descriptions that don't change between
                # calls. cache_control halves the per-call cost on the second
                # variant (consolidated) and on retries.
                {"type": "text",
                 "text": system_prompt,
                 "cache_control": {"type": "ephemeral"}},
            ],
            tools=[tool],
            tool_choice={"type": "tool", "name": tool["name"]},
            messages=[{"role": "user", "content": user_prompt}],
        )
        for block in resp.content:
            # The SDK exposes blocks as objects (block.type) or dicts depending
            # on version — handle both rather than gambling on one.
            btype = getattr(block, "type", None) or (block.get("type") if isinstance(block, dict) else None)
            if btype == "tool_use":
                inp = getattr(block, "input", None)
                if inp is None and isinstance(block, dict):
                    inp = block.get("input")
                return inp
        print("    [claude_mapper] no tool_use block in response")
        return None
    except Exception as e:
        print(f"    [claude_mapper] call failed: {type(e).__name__}: {e}")
        return None


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Aggregation: per-row → per-cell sums
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
# Constraint checks
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
        f"  - Networth (F5:F8)        = {nw:.2f}\n"
        f"  - Current Liab (F14:F19)  = {cl:.2f}\n"
        f"  - NC Liab (F24:F27)       = {ncl:.2f}\n"
        f"  - Current Assets (F35:F42)= {ca:.2f}\n"
        f"  - NC Assets (F48:F54)     = {nca:.2f}\n"
        f"Look for rows that may have been put in the wrong section. A gap of "
        f"~{abs(diff):.2f} often means a single misclassified row of that magnitude."
    )
    return False, diag


def _coverage_check(items: List[Dict], parsed: Dict, side: str) -> Tuple[bool, str]:
    src_labels = [it["label"].strip() for it in items
                  if (it["cur"] != 0 or it["pri"] != 0)]
    mapped_labels = [m.get("source_label", "").strip()
                     for m in parsed.get("mappings", [])]
    missing = []
    for lbl in src_labels:
        found = any(lbl.lower() == ml.lower() for ml in mapped_labels)
        if not found:
            missing.append(lbl[:80])
    if not missing:
        return True, ""
    diag = (
        f"Coverage check FAILED ({side}): {len(missing)} source rows are missing "
        f"from your `mappings`. Add an entry for each (target_cell='_skip' if "
        f"it's a total/header). Missing rows:\n"
        + "\n".join(f"  - {m}" for m in missing[:15])
    )
    return False, diag


def _section_consistency_check(parsed: Dict) -> Tuple[bool, str]:
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
                       f"(allowed: {allowed[:6]}...)")
    if not bad:
        return True, ""
    return False, "Section/cell consistency FAILED:\n" + "\n".join(bad[:10])


def _section_total_cross_check(items: List[Dict],
                               agg: Dict[str, Dict[str, float]],
                               year: str, tol: float) -> Tuple[bool, str]:
    """Compares published 'Total <section>' rows against the mapper's
    sum-of-parts. Tolerance is the LARGER of:
      - the absolute `tol` (₹1 lakh by default)
      - SECTION_TOTAL_REL_TOL × the section magnitude (0.5%)
    so benign rounding in the published source doesn't trigger a retry.
    """
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
                effective_tol = max(tol,
                                     abs(expected) * SECTION_TOTAL_REL_TOL)
                if abs(expected - got) > effective_tol:
                    issues.append(
                        f"  - {human} ({year}): extract reports total={expected:.2f} "
                        f"but your mappings sum to {got:.2f} (gap {expected-got:+.2f})."
                    )
                break
    if not issues:
        return True, ""
    return False, "Section-total cross-check FAILED:\n" + "\n".join(issues[:8])


def _run_bs_constraints(items: List[Dict], parsed: Dict,
                        agg: Dict[str, Dict[str, float]],
                        tol: float) -> Tuple[List[str], List[str]]:
    """Returns (blocking, informational) diagnostic lists.

    BLOCKING — re-prompt the model:
      - section_consistency (cell in wrong section)
      - coverage (input row not represented)
      - bs_balance (Assets ≠ L+E by > tol) — the master equation

    INFORMATIONAL — surface as warnings only, do NOT retry:
      - section_total_cross_check — published totals vs sum-of-parts.
        Often differs by rounding when the source's published total was
        calculated independently of the line-item rounding. Retrying
        rarely improves these and burns ~30s per round.
    """
    blocking: List[str] = []
    informational: List[str] = []
    ok, d = _section_consistency_check(parsed)
    if not ok:
        blocking.append(d)
    ok, d = _coverage_check(items, parsed, "BS")
    if not ok:
        blocking.append(d)
    for year in ("current_year", "prior_year"):
        ok, d = _bs_balance_check(agg, year, tol)
        if not ok:
            blocking.append(d)
        ok, d = _section_total_cross_check(items, agg, year, tol)
        if not ok:
            informational.append(d)
    return blocking, informational


def _pl_identity_check(parsed: Dict, tol: float) -> List[Tuple[str, str]]:
    """Verify the P&L identity holds:
      Net Profit ≈ Revenue + Other Income - COGS - Employee - Finance
                   - Depreciation - Other Expense ± Exceptional - Tax

    F72 (other expense) is computed as the residual by the caller, so we
    use the model's other_expense entries instead. If the published
    net_profit field disagrees with our reconstruction beyond tolerance,
    something is mis-categorised (a tax line missed, exceptional
    double-counted, OCI included, etc.).

    Returns a list of (year_label, diagnosis) tuples. Empty when both
    years pass.
    """
    diags: List[Tuple[str, str]] = []
    by_year: Dict[str, Dict[str, float]] = {
        "current_year": defaultdict(float),
        "prior_year":   defaultdict(float),
    }
    for m in parsed.get("mappings", []):
        cell = m.get("target_cell", "_skip")
        if cell == "_skip":
            continue
        cy = float(m.get("current_year_value", 0) or 0)
        py = float(m.get("prior_year_value", 0) or 0)
        by_year["current_year"][cell] += cy
        by_year["prior_year"][cell]   += py

    for year in ("current_year", "prior_year"):
        m = by_year[year]
        revenue       = m.get("F66", 0)
        cogs          = m.get("F67", 0)
        employee      = m.get("F69", 0)
        finance       = m.get("F70", 0)
        depreciation  = m.get("F71", 0)
        tax           = m.get("F73", 0)
        other_income  = m.get("other_income", 0)
        other_expense = m.get("other_expense", 0)
        exceptional   = m.get("exceptional_items", 0)
        published_np  = m.get("net_profit", 0)

        if published_np == 0 and revenue == 0:
            continue  # nothing to check

        reconstructed = (revenue + other_income
                         - cogs - employee - finance - depreciation
                         - other_expense
                         + exceptional
                         - tax)
        gap = round(published_np - reconstructed, 2)
        if abs(gap) > tol:
            yl = year.replace("_year", "")
            diag = (
                f"P&L identity check FAILED for {yl}: "
                f"published Net Profit={published_np:.2f}, "
                f"reconstructed={reconstructed:.2f}, gap={gap:+.2f}.\n"
                f"  Revenue={revenue:.2f}  +OtherInc={other_income:.2f}\n"
                f"  -COGS={cogs:.2f}  -Emp={employee:.2f}  "
                f"-Finance={finance:.2f}  -Dep={depreciation:.2f}\n"
                f"  -OtherExp={other_expense:.2f}  "
                f"±Exceptional={exceptional:.2f}  -Tax={tax:.2f}\n"
                f"Likely cause: a tax line missed (should sum into F73), "
                f"exceptional double-counted in another bucket, OCI "
                f"included in net_profit, or signs flipped on Changes "
                f"in Inventory."
            )
            diags.append((yl, diag))
    return diags


def _collect_bs_warnings(items, parsed, agg, tol) -> List[str]:
    """Compress each multi-line diagnosis into a single readable warning row.
    The full diagnosis is multi-line — keep the header AND the bulleted
    issues so the WARNINGS sheet is actually actionable."""
    warns: List[str] = []
    for year in ("current_year", "prior_year"):
        yl = year.replace("_year", "")
        ok, d = _bs_balance_check(agg, year, tol)
        if not ok:
            # Keep the first informative sentence (the gap + side).
            head = d.split("\n", 1)[0]
            warns.append(f"claude [{yl}]: {head}")
        ok, d = _section_total_cross_check(items, agg, year, tol)
        if not ok:
            # The cross-check produces "<header>:\n  - issue 1\n  - issue 2".
            # Inline the bullet points so the WARNINGS sheet shows the
            # specific section + gap, not a useless lone "FAILED:".
            lines = [ln.strip() for ln in d.splitlines() if ln.strip()]
            detail = " | ".join(lines[1:]) if len(lines) > 1 else lines[0]
            warns.append(f"claude [{yl}]: {detail}")
    return warns


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Public API — drop-in replacements for map_bs_v2 / map_pl_v2
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def map_bs_claude(items: List[Dict],
                  pdf_path: Optional[str] = None) -> Optional[Dict]:
    client, cfg = _get_client()
    if client is None:
        print("    [claude_mapper] No Anthropic key configured.")
        return None

    model = cfg.get("claude_model", DEFAULT_MODEL)
    tol   = float(cfg.get("balance_tolerance", BALANCE_TOL))

    items_block = _format_bs_items_grouped(items)
    if not items_block.strip():
        print("    [claude_mapper] No non-zero BS rows; skipping.")
        return None

    system = _bs_system_prompt()
    tool   = _bs_tool()

    parsed: Optional[Dict] = None
    diagnoses_history: List[str] = []
    rounds_used = 0
    last_diags: List[str] = []
    last_blocking: List[str] = []
    last_info: List[str] = []
    for round_idx in range(MAX_RETRY_ROUNDS + 1):
        rounds_used = round_idx + 1
        retry_diag = "\n\n".join(diagnoses_history[-1:]) if diagnoses_history else ""
        user = _bs_user_prompt(items_block, retry_diag)
        print(f"    [claude_mapper] BS round {rounds_used}: model={model}")
        parsed = _call_claude_tool(client, model, system, user, tool)
        if not parsed:
            break

        agg = _aggregate_bs(parsed)
        blocking, info = _run_bs_constraints(items, parsed, agg, tol)
        last_blocking, last_info = blocking, info
        if not blocking:
            # Master equation passed. Section-total nits (info) become
            # warnings; we DO NOT retry on them — saves 1-2 rounds × ~30s
            # per job in the very common "BS balances but rounding
            # differs" case (Shivam, Finolex, etc.).
            if info:
                print(f"    [claude_mapper] BS round {rounds_used}: "
                      f"balance + coverage OK; {len(info)} info-only "
                      f"warning(s) (rounding) — not retrying")
            else:
                print(f"    [claude_mapper] BS round {rounds_used}: "
                      f"ALL CONSTRAINTS PASS")
            break
        if round_idx >= MAX_RETRY_ROUNDS:
            print(f"    [claude_mapper] BS round {rounds_used}: still "
                  f"failing — keeping best effort")
            break
        diagnoses_history.append("\n\n".join(blocking))
        print(f"    [claude_mapper] BS round {rounds_used}: "
              f"{len(blocking)} BLOCKING constraint(s) failed; retrying")

    # ── PDF-grounded remediation pass ──
    # ONLY fires when we still have BLOCKING constraint failures (balance
    # broken / coverage gap / section-consistency). Informational rounding
    # nits never trigger a (~₹0.025) PDF call.
    if parsed and last_blocking and pdf_path and os.path.exists(pdf_path):
        try:
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
            print(f"    [claude_mapper] BS remediation: attaching PDF "
                  f"({len(pdf_bytes)/1024:.0f} KB) for grounded fix")
            remed_prompt = _bs_remediation_user_prompt(
                items_block, parsed, "\n\n".join(last_blocking),
            )
            remediated = _call_claude_tool_with_pdf(
                client, model, system, remed_prompt, tool, pdf_bytes,
            )
            if remediated:
                rem_agg = _aggregate_bs(remediated)
                rem_blocking, rem_info = _run_bs_constraints(items, remediated, rem_agg, tol)
                if len(rem_blocking) < len(last_blocking):
                    print(f"    [claude_mapper] BS remediation: improved "
                          f"({len(last_blocking)} → {len(rem_blocking)} "
                          f"blocking) — using remediated mapping")
                    parsed = remediated
                    last_blocking, last_info = rem_blocking, rem_info
                    rounds_used += 1
                else:
                    print(f"    [claude_mapper] BS remediation: no "
                          f"improvement ({len(rem_blocking)} still "
                          f"blocking) — keeping previous mapping")
        except Exception as e:
            print(f"    [claude_mapper] BS remediation error: "
                  f"{type(e).__name__}: {e}")

    if not parsed:
        return None

    agg = _aggregate_bs(parsed)
    warnings = _collect_bs_warnings(items, parsed, agg, tol)

    notes = f"claude_mapper ({model}); rounds={rounds_used}"
    if pdf_path and os.path.exists(pdf_path):
        notes += "; pdf_remediation=available"
    mapping = {
        "current_year": agg["current_year"],
        "prior_year":   agg["prior_year"],
        "notes":        notes,
        "_warnings":    warnings,
        "_per_row":     parsed.get("mappings", []),
    }
    _print_bs_summary(parsed, agg)
    return mapping


def map_pl_claude(items: List[Dict],
                  pdf_path: Optional[str] = None) -> Optional[Dict]:
    client, cfg = _get_client()
    if client is None:
        print("    [claude_mapper] No Anthropic key configured.")
        return None

    model = cfg.get("claude_model", DEFAULT_MODEL)
    tol   = float(cfg.get("balance_tolerance", BALANCE_TOL))

    items_block = _format_pl_items(items)
    if not items_block.strip():
        return None

    system = _pl_system_prompt()
    tool   = _pl_tool()

    parsed: Optional[Dict] = None
    diagnoses_history: List[str] = []
    rounds_used = 0
    pl_warnings: List[str] = []
    last_diags: List[str] = []
    for round_idx in range(MAX_RETRY_ROUNDS + 1):
        rounds_used = round_idx + 1
        retry_diag = "\n\n".join(diagnoses_history[-1:]) if diagnoses_history else ""
        user = _pl_user_prompt(items_block, retry_diag)
        print(f"    [claude_mapper] PL round {rounds_used}: model={model}")
        parsed = _call_claude_tool(client, model, system, user, tool)
        if not parsed:
            break

        diags = []
        ok, d = _coverage_check(items, parsed, "PL")
        if not ok:
            diags.append(d)
        identity_issues = _pl_identity_check(parsed, tol)
        for yl, msg in identity_issues:
            diags.append(msg)
        last_diags = diags
        if not diags:
            print(f"    [claude_mapper] PL round {rounds_used}: coverage + identity OK")
            break
        if round_idx >= MAX_RETRY_ROUNDS:
            for yl, msg in identity_issues:
                pl_warnings.append(f"claude [{yl}]: {msg.split(chr(10), 1)[0]}")
            break
        diagnoses_history.append("\n\n".join(diags))
        print(f"    [claude_mapper] PL round {rounds_used}: "
              f"{len(diags)} constraint(s) failed; retrying")

    # ── PDF-grounded remediation pass for P&L ──
    if parsed and last_diags and pdf_path and os.path.exists(pdf_path):
        try:
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
            print(f"    [claude_mapper] PL remediation: attaching PDF "
                  f"({len(pdf_bytes)/1024:.0f} KB) for grounded fix")
            remed_prompt = _pl_remediation_user_prompt(
                items_block, parsed, "\n\n".join(last_diags),
            )
            remediated = _call_claude_tool_with_pdf(
                client, model, system, remed_prompt, tool, pdf_bytes,
            )
            if remediated:
                rem_diags = []
                ok2, d2 = _coverage_check(items, remediated, "PL")
                if not ok2:
                    rem_diags.append(d2)
                rem_identity = _pl_identity_check(remediated, tol)
                for yl, msg in rem_identity:
                    rem_diags.append(msg)
                if len(rem_diags) < len(last_diags):
                    print(f"    [claude_mapper] PL remediation: improved "
                          f"({len(last_diags)} → {len(rem_diags)} failing) "
                          f"— using remediated mapping")
                    parsed = remediated
                    last_diags = rem_diags
                    pl_warnings = []  # reset since we have a better answer
                    for yl, msg in rem_identity:
                        pl_warnings.append(f"claude [{yl}]: "
                                            f"{msg.split(chr(10), 1)[0]}")
                    rounds_used += 1
                else:
                    print(f"    [claude_mapper] PL remediation: no "
                          f"improvement — keeping previous mapping")
        except Exception as e:
            print(f"    [claude_mapper] PL remediation error: "
                  f"{type(e).__name__}: {e}")

    if not parsed:
        return None

    agg = _aggregate_pl(parsed)
    cy = agg["current_year"]
    py = agg["prior_year"]

    mapping = {
        "current_year": {
            "F66": cy["F66"], "F67": cy["F67"], "F69": cy["F69"],
            "F70": cy["F70"], "F71": cy["F71"], "F73": cy["F73"],
            "other_income":      cy["other_income"],
            # Surface the summed `other_expense` rows so downstream
            # bs_pl_mapper.compute_f72_residual can compute F72 LITERALLY
            # as (other_expense - other_income) instead of as a residual
            # that silently absorbs Exceptional Items.
            "other_expense":     cy.get("other_expense_total", 0),
            "net_profit":        cy["net_profit"],
            "exceptional":       cy["exceptional_items"],
            "exceptional_items": cy["exceptional_items"],
        },
        "prior_year": {
            "F66": py["F66"], "F67": py["F67"], "F69": py["F69"],
            "F70": py["F70"], "F71": py["F71"], "F73": py["F73"],
            "other_income":      py["other_income"],
            "other_expense":     py.get("other_expense_total", 0),
            "net_profit":        py["net_profit"],
            "exceptional":       py["exceptional_items"],
            "exceptional_items": py["exceptional_items"],
        },
        "notes":     f"claude_mapper ({model}); rounds={rounds_used}",
        "_warnings": pl_warnings,
        "_per_row":  parsed.get("mappings", []),
    }
    _print_pl_summary(parsed)
    return mapping


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Reporting helpers
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _print_bs_summary(parsed: Dict, agg: Dict):
    rows = parsed.get("mappings", [])
    skipped = sum(1 for r in rows if r.get("target_cell") == "_skip")
    mapped  = len(rows) - skipped
    print(f"    [claude_mapper] BS: mapped {mapped} / skipped {skipped}")
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
    print(f"    [claude_mapper] PL: mapped {mapped} / skipped {skipped}")
    for r in rows:
        cell = r.get("target_cell", "_skip")
        if cell == "_skip":
            continue
        cy = r.get("current_year_value", 0)
        print(f"      • {cell:>14s}  {r.get('source_label','')[:50]:50s}  CY={cy}")
