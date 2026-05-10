"""
Smart Financial Matcher v3 — Production Grade
================================================
100% offline. Rules + TF-IDF. Zero API calls.
Built from 71-company audit: covers 181 BS + 213 PL edge-case labels.
"""

import os, re
import numpy as np
import pandas as pd
from pathlib import Path
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Noise filter — PDF table extraction occasionally treats page headers
# (company name in ALL-CAPS, "Annual Report", page numbers like "Page 3
# of 8", section labels) as data rows. They land here with a stray small
# value (1.0, the page number, etc.) and would otherwise pollute the
# WARNINGS sheet with bogus "BS unmatched" entries on every run.
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

_NOISE_KEYWORDS = (
    "limited", " ltd", "private", "annual report", "page no",
    "registered office", "corporate identity", "cin", "regd. office",
    "in lakh", "in lacs", "in lakhs", "(rs.", "(₹",
)

_PAGE_NUMBER_RE = re.compile(r"^\s*page\s+\d+(\s+of\s+\d+)?\s*$", re.IGNORECASE)


def _looks_like_pdf_noise(label: str, cur: float, pri: float) -> bool:
    """True when this row is almost certainly a PDF artifact, not a real
    line item. Used to suppress the WARNINGS sheet entry for it.

    Heuristics:
      - Label is mostly uppercase letters AND value is small (<10)
        → page header pseudo-row (e.g. "KALYANI FORGE LIMITED" with CY=1.0).
      - Label matches a known noise keyword (CIN, registered office, etc.).
      - Label is "Page N" or "Page N of M".
      - Label is just digits/whitespace (lone page numbers).
    """
    lbl = (label or "").strip()
    if not lbl:
        return True
    low = lbl.lower()
    for kw in _NOISE_KEYWORDS:
        if kw in low:
            return True
    if _PAGE_NUMBER_RE.match(lbl):
        return True
    if re.fullmatch(r"[\d\s.,/-]+", lbl):
        return True
    letters = [c for c in lbl if c.isalpha()]
    if letters and sum(1 for c in letters if c.isupper()) / len(letters) > 0.8:
        # ALL-CAPS row with a tiny, suspicious value
        if max(abs(cur or 0), abs(pri or 0)) < 10:
            return True
    return False

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CELL MAPPINGS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

PL_FIELD_TO_CELL = {
    "Revenue": "F66", "Cost of goods sold": "F67",
    "Employee benefits expense": "F69", "Interest": "F70",
    "Depreciation": "F71", "Taxes": "F73",
    "Net Profit": "net_profit", "other_income": "other_income",
    "exceptional": "exceptional",
    "Other Expenses less Other Income": "other_expenses",
    "Gross Profit": "_skip", "Net Margin": "_skip",
}

BS_FIELD_TO_CELL = {
    "Share Capital": "F5", "Retained Earnings": "F6",
    "Reserves and Surplus": "F7", "Other Equity": "F8",
    "Non Controlling Interest": "F9",
    "Trade Payables": "F14", "Provisions Current": "F15",
    "Short Term Borrowings": "F16", "Other Current Liabilities": "F17",
    "Other Financial Liabilities Current": "F18", "Lease Liabilities Current": "F19",
    "Long Term Borrowings": "F24", "Provisions Non Current": "F25",
    "Other Non Current Liabilities": "F26", "Other Financial Liabilities Non Current": "F27",
    "PPE Fixed Assets": "F48", "Investments Non Current": "F49",
    "Loans Non Current": "F50", "CWIP": "F51",
    "Other Non Current Assets": "F52", "Other Financial Assets Non Current": "F53",
    "Deferred Tax Assets": "F54",
    "Bank Balance": "F35", "Cash and Cash Equivalents": "F36",
    "Inventories": "F37", "Current Investments": "F38",
    "Loans Current": "F39", "Trade Receivables": "F40",
    "Other Current Assets": "F41", "Other Financial Assets Current": "F42",
}

BS_SECTION_FALLBACK = {
    "networth": "F8", "current_liab": "F17", "noncurrent_liab": "F26",
    "current_asset": "F41", "noncurrent_asset": "F52",
}

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# BS SECTION DETECTION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BS_SECTION_MARKERS = {
    "networth": [r"shareholder.?s?\s+fund", r"^equity\s*$", r"networth",
                 r"equity\s+and\s+liabilit", r"^equity\s*$"],
    "current_liab": [r"current\s+liabilit"],
    "noncurrent_liab": [r"non.?\s*current\s+liabilit", r"long.?\s*term\s+liabilit"],
    "current_asset": [r"current\s+asset"],
    "noncurrent_asset": [r"non.?\s*current\s+asset", r"long.?\s*term\s+asset", r"fixed\s+asset"],
}

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# BS SKIP PATTERNS (section headers, totals, junk)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BS_SKIP = [
    # Section totals — specific suffixes so MSME "total outstanding dues" passes through
    r"^total\s+(non.?\s*current|current|equity|asset|liabilit|financial|networth|shareholder)",
    r"^total\s+of\s+(equity|non.?\s*current|current|asset|liabilit)",
    r"^total\s+non.?\s*financial\s+(asset|liabilit)",
    r"^total\s+equit",
    r"^total\s*$", r"^sub\s*total",
    # Section headers
    r"^equity\s+and\s+liabilit", r"^assets?\s*$",
    r"^liabilities\s*$",
    r"^equity\s*$", r"^equity\s+attributable",
    r"^shareholder", r"^non.?\s*current\s+(assets?|liabilit)\s*$",
    r"^current\s+(assets?|liabilit)\s*$",
    r"^financial\s+(assets?|liabilit)\s*$",
    r"^non.?\s*financial\s+(assets?|liabilit)\s*$",
    r"^total\s+financial\s+(assets?|liabilit)",
    r"^financial\s+liabilities\s*$", r"^financial\s+assets\s*$",
    r"^attributable\s+to\s+equity\s+holders",
    # Table headers / junk
    r"^particulars", r"^note\s*(no)?$", r"^notes?\s+on\s+financial",
    r"^(i|ii|iii|iv|v|vi)\s*$",
    r"balance\s+sheet\s+(as\s+at|march|dated)",
    r"standalone\s+balance\s+sheet", r"consolidated\s+balance\s+sheet",
    # Company address / registration
    r"regd\s+office", r"registered\s+office", r"cin\s*:",
    r"legislative\s+assembl",
    # Bare single words
    r"^net\s*$", r"^nil\s*$",
    # Revenue/PL items leaking into BS
    r"^revenue\s+from\s+operations",
    r"^total\s+income\s*$",
    r"^other\s+income\s*$",
]

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# BS RULES — 130+ patterns from 71-company audit
# (pattern, field, allowed_sections_or_None)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BS_RULES = [
    # ── EQUITY ──
    (r"share\s+capital", "Share Capital", None),
    (r"equity\s+share\s+capital", "Share Capital", None),
    (r"paid.?up\s+.*capital", "Share Capital", None),
    (r"authorised\s+.*capital", "Share Capital", None),
    (r"retained\s+earnings?", "Retained Earnings", None),
    (r"surplus\s+in\s+(profit|statement)", "Retained Earnings", None),
    (r"reserves?\s+and\s+surplus", "Reserves and Surplus", None),
    (r"general\s+reserve", "Reserves and Surplus", None),
    (r"securities\s+premium", "Reserves and Surplus", None),
    (r"capital\s+reserve", "Reserves and Surplus", None),
    (r"capital\s+redemption", "Reserves and Surplus", None),
    (r"revaluation\s+reserve", "Reserves and Surplus", None),
    (r"statutory\s+reserve", "Reserves and Surplus", None),
    (r"special\s+reserve", "Reserves and Surplus", None),
    (r"other\s+equity", "Other Equity", None),
    (r"share\s+application\s+money", "Other Equity", None),
    (r"money\s+received\s+against\s+share", "Other Equity", None),
    (r"non.?\s*controlling\s+interest", "Non Controlling Interest", None),
    (r"non.?\s*controlling\s+stake", "Non Controlling Interest", None),
    (r"minority\s+interest", "Non Controlling Interest", None),

    # ── TRADE PAYABLES (MSME) ──
    (r"trade\s+payable", "Trade Payables", None),
    (r"accounts?\s+payable", "Trade Payables", None),
    (r"sundry\s+creditor", "Trade Payables", None),
    (r"outstanding\s+dues?\s+(of|to)\s+(micro|creditor|other|msme)", "Trade Payables", None),
    (r"dues?\s+(of|to)\s+(micro|creditor)", "Trade Payables", None),
    (r"^a\s+.*dues", "Trade Payables", None),
    (r"^b\s+.*dues", "Trade Payables", None),
    (r"due\s+to\s+(micro|creditor|other)", "Trade Payables", None),
    (r"(msme|micro\s+(and|&)\s+small)", "Trade Payables", None),
    (r"creditors?\s+(other\s+than|for\s+(goods|supplies))", "Trade Payables", None),

    # ── LOANS & ADVANCES (BEFORE borrowings!) ──
    (r"loan[s]?\s*(and|&)?\s*(advance|icd)", "Loans Non Current", ["noncurrent_asset"]),
    (r"loan[s]?\s*(and|&)?\s*(advance|icd)", "Loans Current", ["current_asset"]),
    (r"^loans?\s*$", "Loans Non Current", ["noncurrent_asset"]),
    (r"^loans?\s*$", "Loans Current", ["current_asset"]),
    (r"^loans?\s+[a-z]\s*$", "Loans Non Current", ["noncurrent_asset"]),
    (r"^loans?\s+[a-z]\s*$", "Loans Current", ["current_asset"]),
    (r"^loan\s*$", "Loans Non Current", ["noncurrent_asset"]),
    (r"^loan\s*$", "Loans Current", ["current_asset"]),
    (r"long.?\s*term\s+loan", "Loans Non Current", None),
    (r"short.?\s*term\s+loan", "Loans Current", None),
    (r"non.?\s*current\s+loan", "Loans Non Current", None),
    (r"current\s+loan", "Loans Current", None),
    (r"advance\s+to\s+(supplier|employee|staff)", "Loans Current", None),
    (r"inter.?\s*corporate\s+deposit", "Loans Current", ["current_asset"]),
    (r"icd", "Loans Current", ["current_asset"]),
    (r"security\s+deposit", "Loans Non Current", ["noncurrent_asset"]),
    (r"security\s+deposit", "Loans Current", ["current_asset"]),

    # ── BORROWINGS ──
    (r"short.?\s*term\s+borrowing", "Short Term Borrowings", None),
    (r"bank\s+overdraft", "Short Term Borrowings", None),
    (r"cash\s+credit", "Short Term Borrowings", None),
    (r"working\s+capital\s+(loan|facility)", "Short Term Borrowings", None),
    (r"current\s+maturities?\s+of\s+long", "Short Term Borrowings", None),
    (r"bills?\s+payable", "Short Term Borrowings", None),
    (r"^borrowings?\s*$", "Short Term Borrowings", ["current_liab"]),
    (r"^borrowings?\s+[a-z]\s*$", "Short Term Borrowings", ["current_liab"]),
    (r"long.?\s*term\s+borrowing", "Long Term Borrowings", None),
    (r"non.?\s*current\s+borrowing", "Long Term Borrowings", None),
    (r"secured\s+loan", "Long Term Borrowings", None),
    (r"unsecured\s+loan", "Long Term Borrowings", None),
    (r"^term\s+loan", "Long Term Borrowings", None),
    (r"debenture", "Long Term Borrowings", None),
    (r"^borrowings?\s*$", "Long Term Borrowings", ["noncurrent_liab"]),
    (r"^borrowings?\s+[a-z]\s*$", "Long Term Borrowings", ["noncurrent_liab"]),

    # ── PROVISIONS (context-aware) ──
    (r"short.?\s*term\s+provision", "Provisions Current", None),
    (r"^provisions?\s*$", "Provisions Current", ["current_liab"]),
    (r"^provisions?\s+[a-z]\s*$", "Provisions Current", ["current_liab"]),
    (r"long.?\s*term\s+provision", "Provisions Non Current", None),
    (r"non.?\s*current\s+provision", "Provisions Non Current", None),
    (r"^provisions?\s*$", "Provisions Non Current", ["noncurrent_liab"]),
    (r"^provisions?\s+[a-z]\s*$", "Provisions Non Current", ["noncurrent_liab"]),
    # Employee benefit obligation → Provisions (NOT Loans!)
    (r"employee.?\s*benefit\s+obligation", "Provisions Non Current", ["noncurrent_liab"]),
    (r"employee.?\s*benefit\s+obligation", "Provisions Current", ["current_liab"]),
    (r"net\s+employee\s+defined\s+benefit", "Provisions Non Current", ["noncurrent_liab"]),
    (r"net\s+employee\s+defined\s+benefit", "Provisions Current", ["current_liab"]),

    # ── LEASE ──
    (r"lease\s+liabilit", "Other Non Current Liabilities", ["noncurrent_liab"]),
    (r"lease\s+liabilit", "Lease Liabilities Current", ["current_liab"]),
    (r"current\s+lease", "Lease Liabilities Current", None),

    # ── CURRENT TAX LIABILITIES → Other CL (NOT NCL!) ──
    (r"current\s+tax\s+liabilit", "Other Current Liabilities", ["current_liab"]),
    (r"current\s+tax\s+liabilit", "Other Current Liabilities", None),
    (r"current\s+tax\s+liabilti", "Other Current Liabilities", None),  # typo: liabilties

    # ── OTHER LIABILITIES ──
    (r"other\s+current\s+liabilit", "Other Current Liabilities", None),
    (r"other\s+current\s+liablit", "Other Current Liabilities", None),  # typo: liablities
    (r"^other\s+liabilit", "Other Current Liabilities", ["current_liab"]),
    (r"^other\s+liabilit", "Other Non Current Liabilities", ["noncurrent_liab"]),
    (r"statutory\s+dues?\s+payable", "Other Current Liabilities", None),
    (r"advance\s+from\s+customer", "Other Current Liabilities", None),
    (r"advances?\s+received", "Other Current Liabilities", None),
    (r"income\s+received\s+in\s+advance", "Other Current Liabilities", None),
    (r"unearned\s+revenue", "Other Current Liabilities", None),
    (r"contract\s+liabilit", "Other Current Liabilities", None),
    (r"revenue\s+received\s+in\s+advance", "Other Current Liabilities", None),
    (r"government\s+grant", "Other Non Current Liabilities", ["noncurrent_liab"]),
    (r"government\s+grant", "Other Current Liabilities", ["current_liab"]),
    (r"other\s+non.?\s*current\s+liabilit", "Other Non Current Liabilities", None),
    (r"other\s+long.?\s*term\s+liabilit", "Other Non Current Liabilities", None),
    (r"deferred\s+tax\s+liabilit", "Other Non Current Liabilities", None),
    (r"deff?erd?\s+tax\s+liabilit", "Other Non Current Liabilities", None),  # typo variants
    (r"^others\s*$", "Other Current Liabilities", ["current_liab"]),
    (r"^others\s*$", "Other Non Current Liabilities", ["noncurrent_liab"]),
    (r"other\s+non.?\s*financial\s+liabilit", "Other Current Liabilities", ["current_liab"]),
    (r"other\s+non.?\s*financial\s+liabilit", "Other Non Current Liabilities", ["noncurrent_liab"]),

    # ── FINANCIAL LIABILITIES ──
    (r"other\s+(current\s+)?financial\s+liabilit", "Other Financial Liabilities Current", ["current_liab"]),
    (r"other\s+financial\s+liabilit", "Other Financial Liabilities Non Current", ["noncurrent_liab"]),
    (r"unpaid\s+dividend", "Other Financial Liabilities Current", None),
    (r"interest\s+accrued\s+(but\s+)?not\s+due", "Other Financial Liabilities Current", None),
    (r"due\s+to\s+others", "Other Current Liabilities", ["current_liab"]),
    (r"liabilit.*associated.*held\s+for\s+sale", "Other Current Liabilities", None),
    (r"liabilit.*classified.*held", "Other Current Liabilities", None),

    # ── PPE / FIXED ASSETS ──
    (r"property.?\s*plant", "PPE Fixed Assets", None),
    (r"^tangible\s+assets?", "PPE Fixed Assets", None),
    (r"intangible\s+assets?\s*$", "PPE Fixed Assets", None),
    (r"other\s+intangible", "PPE Fixed Assets", None),
    (r"goodwill", "PPE Fixed Assets", None),
    (r"right.?\s*of.?\s*use", "PPE Fixed Assets", None),
    (r"right\s+to\s+use", "PPE Fixed Assets", None),
    (r"rou\s+asset", "PPE Fixed Assets", None),
    (r"fixed\s+asset", "PPE Fixed Assets", None),
    (r"plant\s+and\s+machinery", "PPE Fixed Assets", None),
    (r"land\s+and\s+building", "PPE Fixed Assets", None),
    (r"furniture", "PPE Fixed Assets", ["noncurrent_asset"]),
    (r"vehicle", "PPE Fixed Assets", ["noncurrent_asset"]),
    (r"investment\s+propert", "PPE Fixed Assets", None),
    (r"biological\s+asset", "PPE Fixed Assets", None),

    # ── CWIP ──
    (r"capital\s+work.?\s*in.?\s*progress", "CWIP", None),
    (r"cwip", "CWIP", None),
    (r"intangible\s+assets?\s+under\s*dev", "CWIP", None),
    (r"assets?\s+under\s+construction", "CWIP", None),
    (r"intangible\s+assets?\s+underdevelopment", "CWIP", None),

    # ── INVESTMENTS ──
    (r"non.?\s*current\s+investment", "Investments Non Current", None),
    (r"long.?\s*term\s+investment", "Investments Non Current", None),
    (r"investment\s+in\s+(subsidiar|associate|joint)", "Investments Non Current", None),
    (r"investment\s+accounted\s+for\s+using", "Investments Non Current", None),
    (r"investments?\s+in\s+(subsidiar|associate|joint)", "Investments Non Current", None),
    (r"^investments?\s*$", "Investments Non Current", ["noncurrent_asset"]),
    (r"other\s+investments?", "Investments Non Current", ["noncurrent_asset"]),
    (r"current\s+investment", "Current Investments", None),
    (r"short.?\s*term\s+investment", "Current Investments", None),
    (r"^investments?\s*$", "Current Investments", ["current_asset"]),
    (r"other\s+investments?", "Current Investments", ["current_asset"]),
    (r"mutual\s+fund", "Current Investments", ["current_asset"]),
    # Default: bare "investments" without section → NC (more common)
    (r"^investments?\s*$", "Investments Non Current", [None]),
    (r"other\s+investments?", "Investments Non Current", [None]),

    # ── CASH & BANK ──
    (r"other\s+bank\s+balance", "Bank Balance", None),
    (r"bank\s+balance\s+other\s+than", "Bank Balance", None),
    (r"earmarked\s+balance", "Bank Balance", None),
    (r"margin\s+money", "Bank Balance", None),
    (r"cash\s+and\s+(cash\s+)?equivalen", "Cash and Cash Equivalents", None),
    (r"cash\s+(and\s+)?bank\s+balance", "Cash and Cash Equivalents", None),
    (r"cash\s+(on|in)\s+hand", "Cash and Cash Equivalents", None),
    (r"cash\s+cash\s+equivalen", "Cash and Cash Equivalents", None),
    (r"^equivalents?\s*$", "Cash and Cash Equivalents", ["current_asset"]),

    # ── INVENTORY ──
    (r"inventor", "Inventories", None),
    (r"stock\s+in\s+trade", "Inventories", None),
    (r"finished\s+goods?", "Inventories", None),
    (r"stores?\s+and\s+spare", "Inventories", None),

    # ── RECEIVABLES ──
    (r"trade\s+receivable", "Trade Receivables", None),
    (r"accounts?\s+receivable", "Trade Receivables", None),
    (r"sundry\s+debtor", "Trade Receivables", None),
    (r"contract\s+assets?", "Trade Receivables", ["current_asset"]),

    # ── OTHER CURRENT ASSETS ──
    (r"other\s+current\s+asset", "Other Current Assets", None),
    (r"current\s+tax\s+asset", "Other Current Assets", None),
    (r"advance\s+(income\s+)?tax", "Other Current Assets", None),
    (r"tds\s+receivable", "Other Current Assets", None),
    (r"gst\s+input", "Other Current Assets", None),
    (r"income\s+tax\s+asset", "Other Current Assets", ["current_asset"]),
    (r"prepaid\s+expense", "Other Current Assets", ["current_asset"]),
    (r"assets?\s+classified\s+as\s+held", "Other Current Assets", None),
    (r"held\s+for\s+sale", "Other Current Assets", None),
    (r"^others\s*$", "Other Current Assets", ["current_asset"]),
    (r"other\s+non.?\s*financial\s+asset", "Other Current Assets", ["current_asset"]),

    # ── OTHER NON-CURRENT ASSETS ──
    (r"other\s+non.?\s*current\s+asset", "Other Non Current Assets", None),
    (r"capital\s+advance", "Other Non Current Assets", None),
    (r"income\s+tax\s+asset", "Other Non Current Assets", ["noncurrent_asset"]),
    (r"non.?\s*current\s+tax\s+asset", "Other Non Current Assets", None),
    (r"^others\s*$", "Other Non Current Assets", ["noncurrent_asset"]),
    (r"other\s+non.?\s*financial\s+asset", "Other Non Current Assets", ["noncurrent_asset"]),

    # ── DEFERRED TAX ASSETS ──
    (r"deferred\s+tax\s+asset", "Deferred Tax Assets", None),
    (r"mat\s+credit", "Deferred Tax Assets", None),

    # ── FINANCIAL ASSETS ──
    (r"other\s+financial\s+asset", "Other Financial Assets Non Current", ["noncurrent_asset"]),
    (r"other\s+financial\s+asset", "Other Financial Assets Current", ["current_asset"]),
    (r"other.?\s*current\s+financial\s+asset", "Other Financial Assets Current", None),
    (r"others?\s+financial\s+asset", "Other Financial Assets Current", ["current_asset"]),
    (r"others?\s+financial\s+asset", "Other Financial Assets Non Current", ["noncurrent_asset"]),
    # Default: bare "other financial assets" without section → NC
    (r"other\s+financial\s+asset", "Other Financial Assets Non Current", [None]),

    # ── FINANCIAL LIABILITIES (bare, without section) ──
    (r"other\s+financial\s+liabilit", "Other Financial Liabilities Current", [None]),
    (r"^lease\s+liabilit", "Other Non Current Liabilities", [None]),
]

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PL RULES — comprehensive from 71-company audit
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

PL_SKIP = [
    r"^total\s", r"^sub\s*total",
    r"^earnings?\s+per\s+(equity\s+)?share", r"^earning\s+per\s+equity",
    r"^basic\s*$", r"^diluted\s*$", r"^\d+\s*basic", r"^\d+\s*diluted",
    r"^basic\s+(and\s+diluted|in\s|earnings)", r"^diluted\s+(in\s|earnings)",
    r"^a\s+basic\s*$", r"^b\s+diluted\s*$",
    r"^basic\s+in\s", r"^diluted\s+in\s",
    r"^profit\s+before\s+(tax|exceptional)",
    r"^profit\s*(loss)?\s*before\s+(tax|exceptional)",
    r"^expenses?\s*$", r"^income\s*$", r"^revenue\s*$",
    r"^tax\s+expense", r"^continuing\s+operations?",
    r"^profit\s+attributable\s+to\s+non", r"^non\s+controlling\s+interest",
    r"^par\s+value", r"^face\s+value",
    # OCI items
    r"remeasurement", r"re\s+measurement",
    r"income\s+tax\s+(effect|relating\s+to\s+items)",
    r"other\s+comprehensive\s+income",
    r"total\s+comprehensive\s+income",
    r"total\s+other\s+comprehensive",
    r"^exceptional\s+item.*gain.*loss",
    r"^items?\s+that\s+will",
    r"^adjustment", r"^comprehensive\s+income",
    # Profit attribution subtotals
    r"owners?\s+of\s+(the\s+)?(company|parent|holding)",
    r"equity\s+holders?\s+of\s+(the\s+)?parent",
    r"^a\s+owners?\s+of", r"^b\s+non.?\s*controlling",
    # Impairment (separate from D&A)
    r"^impairment",
    r"^operations\s*$",
]

PL_NET_PROFIT = [
    r"profit\s+after\s+tax",
    r"net\s+profit\s+after",
    r"net\s+profit\s+for\s+the",
    r"profit\s*(loss)?\s+for\s+the\s+(year|period)",
    r"^net\s+profit\s*$",
    r"^net\s+income",
    r"profit\s+attributable\s+to\s+owners",
    r"profit\s+loss\s+for\s+the",
]

PL_RULES = [
    # Revenue
    (r"revenue\s+from\s+operations?", "Revenue"),
    (r"sale\s+of\s+(products?|services?|goods)", "Revenue"),
    (r"income\s+from\s+operations?", "Revenue"),
    (r"net\s+sales", "Revenue"),
    (r"^turnover$", "Revenue"),
    (r"^sales\s*$", "Revenue"),

    # Other Income
    (r"other\s+income", "other_income"),
    (r"interest\s+income$", "other_income"),
    (r"dividend\s+income", "other_income"),
    (r"rental\s+income", "other_income"),
    (r"gain\s+on\s+(sale|disposal)", "other_income"),
    (r"profit\s+on\s+sale\s+of", "other_income"),

    # COGS
    (r"cost\s+of\s+(materials?\s+consumed|goods\s+sold|revenue|raw)", "Cost of goods sold"),
    (r"purchases?\s+of\s+(stock|traded|raw|goods)", "Cost of goods sold"),
    (r"changes?\s+in\s+inventor", "Cost of goods sold"),
    (r"cost\s+of\s+sales", "Cost of goods sold"),
    (r"consumption\s+of\s+raw", "Cost of goods sold"),
    (r"excise\s+duty\s+on\s+sale", "Cost of goods sold"),
    (r"cost\s+of\s+construction", "Cost of goods sold"),
    (r"sub.?\s*contract", "Cost of goods sold"),
    (r"direct\s+(cost|expense)", "Cost of goods sold"),
    (r"administrative\s*(and|&)\s*other\s+operating", "Cost of goods sold"),
    (r"operating\s+expenses?\s*$", "Cost of goods sold"),

    # Employee
    (r"employee\s+benefit", "Employee benefits expense"),
    (r"staff\s+(cost|expense)", "Employee benefits expense"),
    (r"salaries?\s+(and\s+)?wages?", "Employee benefits expense"),
    (r"managerial\s+remuneration", "Employee benefits expense"),
    (r"payment\s+to\s+employees?", "Employee benefits expense"),
    (r"contribution\s+to\s+provident", "Employee benefits expense"),
    (r"personnel\s+(cost|expense)", "Employee benefits expense"),

    # Finance
    (r"finance\s+cost", "Interest"),
    (r"interest\s+expense", "Interest"),
    (r"interest\s+(on|and)\s+(borrowing|loan|term|deposit|debenture)", "Interest"),
    (r"borrowing\s+cost", "Interest"),
    (r"bank\s+charges?\s+(and\s+)?interest", "Interest"),

    # Depreciation (both z and s spellings)
    (r"depreciation\s+(and|&)\s+amort", "Depreciation"),
    (r"depreciation\s+expense", "Depreciation"),
    (r"depreciation\s*$", "Depreciation"),
    (r"^amorti[sz]ation", "Depreciation"),
    (r"amorti[sz]ation\s+expense", "Depreciation"),
    (r"depletion", "Depreciation"),

    # Tax (comprehensive)
    (r"current\s+tax", "Taxes"),
    (r"deferred\s+tax", "Taxes"),
    (r"income\s+tax\s+expense", "Taxes"),
    (r"tax\s+(of|for)\s+earlier", "Taxes"),
    (r"tax\s+relating\s+to\s+earlier", "Taxes"),
    (r"mat\s+credit\s+entitlement", "Taxes"),
    (r"short.*excess.*provision.*tax", "Taxes"),
    (r"provision\s+for\s+tax\s+relating", "Taxes"),
    (r"excess\s+provision\s+for\s+tax", "Taxes"),

    # Exceptional (start-only to avoid "profit before exceptional")
    (r"^exceptional\s", "exceptional"),
    (r"^extraordinary\s", "exceptional"),
    (r"prior\s+period\s+(item|adjustment)", "exceptional"),
]

PL_OTHER_EXPENSE = [
    r"other\s+expense", r"advertising", r"rent\s+expense", r"repair",
    r"insurance", r"travelling", r"legal\s+and\s+professional",
    r"audit\s+fee", r"communication", r"printing", r"electricity",
    r"rates\s+and\s+taxes", r"miscellaneous\s+expense",
    r"commission\s+expense", r"bad\s+debts?", r"loss\s+on\s+(sale|disposal)",
    r"donation", r"corporate\s+social", r"freight", r"warranty",
    r"royalt", r"software\s+expense", r"outsourcing",
    r"power\s+and\s+fuel", r"postage", r"telephone",
    r"director.?\s*sitting", r"listing\s+fee",
]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TEXT CLEANING
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _clean(label):
    s = str(label).strip()
    s = re.sub(r'^[\-–—•·]\s*', '', s)                        # dash/bullet prefix
    s = re.sub(r'\b([A-Z])\s([a-z])', r'\1\2', s)             # PDF split-letter
    s = s.lower()
    s = re.sub(r'\([^)]{15,}\)', '', s)                        # long parenthetical
    s = re.sub(r'[()]', '', s)                                 # short parens
    s = re.sub(r'^[ivxlc]+[\s\.\)]+', '', s)                  # roman numeral prefix
    s = re.sub(r'^[a-d]\s+', '', s)                            # letter prefix: "a ", "b "
    s = re.sub(r'[\d,]+\.?\d*', '', s)                         # numbers
    s = re.sub(r'[^\w\s]', ' ', s)                             # non-word chars
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _detect_sections(items):
    """Detect BS sections. Section headers set the context, totals end it.
    Items without an explicit section inherit from the previous item."""
    section = None
    out = []

    # v2 patch: pre-compiled classifiers for section inference when the raw
    # extract has no explicit "Non-Current Assets" / "Current Assets" header
    # (as in shree_pushkar, SIKKO). Maps keyword → section guess.
    NCA_ITEM_SIGS = (
        'property, plant', 'property plant', 'capital work', 'intangible',
        'right of use', 'goodwill', 'investment in subsidiar',
        'investment in associate', 'deferred tax asset',
    )
    CA_ITEM_SIGS = ('inventor', 'trade receivabl', 'cash and cash equiv',
                    'bank balanc')
    NW_ITEM_SIGS = ('equity share capital', 'share capital')

    for item in items:
        label = _clean(item["label"])
        orig_lower = str(item["label"]).lower().strip()

        # Check for section headers
        for sec, patterns in BS_SECTION_MARKERS.items():
            for pat in patterns:
                if re.search(pat, label):
                    section = sec

        # Total rows end the current section (next items may be in new section)
        if re.match(r'^total\s+(non.?\s*current\s+asset|current\s+asset)', label):
            section = None  # will be set by next header
        elif re.match(r'^total\s+(non.?\s*current\s+liabilit|current\s+liabilit)', label):
            section = None
        elif re.match(r'^total\s+equity\s*$', label):
            section = None

        # If still None, try to infer from label itself
        if section is None:
            if any(k in orig_lower for k in ['non-current asset', 'non current asset', 'non- current asset']):
                section = 'noncurrent_asset'
            elif any(k in orig_lower for k in ['current asset']):
                section = 'current_asset'
            elif any(k in orig_lower for k in ['non-current liabilit', 'non current liabilit']):
                section = 'noncurrent_liab'
            elif any(k in orig_lower for k in ['current liabilit']):
                section = 'current_liab'

        # v2 patch: item-signature inference. When section is still None and
        # the item label clearly identifies a section by its nature, set section
        # before emitting. Handles raws that omit the "Non-Current Assets" header.
        if section is None:
            if any(sig in orig_lower for sig in NCA_ITEM_SIGS):
                section = 'noncurrent_asset'
            elif any(sig in orig_lower for sig in CA_ITEM_SIGS):
                section = 'current_asset'
            elif any(sig in orig_lower for sig in NW_ITEM_SIGS):
                section = 'networth'

        out.append({**item, "_section": section})
    return out


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# MATCHER CLASS
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class SmartMatcher:
    def __init__(self, pl_csv=None, bs_csv=None):
        self.pl_tfidf = self.pl_matrix = self.pl_fields = None
        self.bs_tfidf = self.bs_matrix = self.bs_fields = None
        if pl_csv and os.path.exists(pl_csv):
            self._build("pl", pl_csv)
        if bs_csv and os.path.exists(bs_csv):
            self._build("bs", bs_csv)

    def _build(self, which, csv_path):
        df = pd.read_csv(csv_path)
        texts = [_clean(kw) for kw in df["Keyword Variation"]]
        fields = df["Field"].tolist()
        tfidf = TfidfVectorizer(analyzer="word", ngram_range=(1, 3),
                                max_features=25000, sublinear_tf=True)
        matrix = tfidf.fit_transform(texts)
        if which == "pl":
            self.pl_tfidf, self.pl_matrix, self.pl_fields = tfidf, matrix, fields
        else:
            self.bs_tfidf, self.bs_matrix, self.bs_fields = tfidf, matrix, fields

    def _tfidf(self, label, which):
        tfidf = self.pl_tfidf if which == "pl" else self.bs_tfidf
        matrix = self.pl_matrix if which == "pl" else self.bs_matrix
        fields = self.pl_fields if which == "pl" else self.bs_fields
        if tfidf is None:
            return "_unknown", 0.0
        clean = _clean(label)
        if not clean:
            return "_skip", 1.0
        vec = tfidf.transform([clean])
        scores = cosine_similarity(vec, matrix).flatten()
        idx = int(np.argmax(scores))
        return fields[idx], float(scores[idx])

    # ── PL ──

    def match_pl(self, label):
        clean = _clean(label)
        if not clean:
            return {"field": "_skip", "cell": "_skip", "confidence": 1.0, "method": "empty"}
        for pat in PL_NET_PROFIT:
            if re.search(pat, clean):
                return {"field": "Net Profit", "cell": "net_profit", "confidence": 0.95, "method": "rule"}
        for pat in PL_SKIP:
            if re.search(pat, clean):
                return {"field": "_skip", "cell": "_skip", "confidence": 1.0, "method": "rule"}
        for pat, field in PL_RULES:
            if re.search(pat, clean):
                cell = PL_FIELD_TO_CELL.get(field, field)
                return {"field": field, "cell": cell, "confidence": 0.95, "method": "rule"}
        for pat in PL_OTHER_EXPENSE:
            if re.search(pat, clean):
                return {"field": "Other Expenses less Other Income", "cell": "other_expenses",
                        "confidence": 0.85, "method": "rule"}
        field, score = self._tfidf(label, "pl")
        cell = PL_FIELD_TO_CELL.get(field, "other_expenses")
        if score >= 0.45:
            return {"field": field, "cell": cell, "confidence": score, "method": "tfidf"}
        return {"field": "Other Expenses", "cell": "other_expenses", "confidence": score, "method": "fallback"}

    # ── BS ──

    def match_bs(self, label, section=None):
        clean = _clean(label)
        if not clean:
            return {"field": "_skip", "cell": "_skip", "confidence": 1.0, "method": "empty"}
        for pat in BS_SKIP:
            if re.search(pat, clean):
                return {"field": "_skip", "cell": "_skip", "confidence": 1.0, "method": "rule"}
        for pat, field, allowed in BS_RULES:
            if re.search(pat, clean):
                if allowed is None or section in allowed:
                    cell = BS_FIELD_TO_CELL.get(field, "_unknown")
                    return {"field": field, "cell": cell, "confidence": 0.95, "method": "rule"}
        field, score = self._tfidf(label, "bs")
        cell = BS_FIELD_TO_CELL.get(field, "_unknown")
        if score >= 0.40:
            return {"field": field, "cell": cell, "confidence": score, "method": "tfidf"}
        if section and section in BS_SECTION_FALLBACK:
            cell = BS_SECTION_FALLBACK[section]
            return {"field": f"other_{section}", "cell": cell, "confidence": score, "method": "fallback"}
        return {"field": "_unknown", "cell": "_unknown", "confidence": 0.0, "method": "none"}

    # ── Bulk ──

    def classify_pl_items(self, items):
        buckets = {k: {"cur": 0.0, "pri": 0.0, "sources": []}
                   for k in ["F66","F67","F69","F70","F71","F73",
                             "net_profit","other_income","exceptional","other_expenses"]}
        unmatched, all_matches, warnings = [], [], []
        for item in items:
            if item["cur"] == 0 and item["pri"] == 0: continue
            r = self.match_pl(item["label"])
            all_matches.append({**item, "match": r})
            cell = r["cell"]
            if cell == "_skip": continue
            if cell in buckets:
                buckets[cell]["cur"] += item["cur"]
                buckets[cell]["pri"] += item["pri"]
                buckets[cell]["sources"].append(f"{item['label']} ({r['confidence']:.2f})")
            else:
                buckets["other_expenses"]["cur"] += item["cur"]
                buckets["other_expenses"]["pri"] += item["pri"]
                buckets["other_expenses"]["sources"].append(item["label"])
            if r["confidence"] < 0.5 and r["method"] not in ("rule",):
                warnings.append(f"PL low-conf: [{item['label'][:40]}] → {cell} ({r['confidence']:.2f})")
        buckets["unmatched"] = unmatched
        buckets["all_matches"] = all_matches
        buckets["warnings"] = warnings
        return buckets

    def classify_bs_items(self, items):
        enriched = _detect_sections(items)
        buckets = {cell: {"cur": 0.0, "pri": 0.0, "sources": []}
                   for cell in BS_FIELD_TO_CELL.values()}
        unmatched, all_matches, warnings = [], [], []
        for item in enriched:
            if item["cur"] == 0 and item["pri"] == 0: continue
            r = self.match_bs(item["label"], item.get("_section"))
            all_matches.append({**item, "match": r})
            cell = r["cell"]
            if cell == "_skip": continue
            if cell in buckets:
                buckets[cell]["cur"] += item["cur"]
                buckets[cell]["pri"] += item["pri"]
                buckets[cell]["sources"].append(f"{item['label']} ({r['confidence']:.2f})")
            elif cell != "_unknown":
                if cell not in buckets:
                    buckets[cell] = {"cur": 0, "pri": 0, "sources": []}
                buckets[cell]["cur"] += item["cur"]
                buckets[cell]["pri"] += item["pri"]
                buckets[cell]["sources"].append(item["label"])
            else:
                unmatched.append({**item, "match": r})
                # Suppress warnings for clearly-not-a-line-item rows. The PDF
                # extractor occasionally pulls page headers/footers (company
                # name in ALL-CAPS, page numbers, "Annual Report") into the
                # row stream with a stray value of 1.0 / page number — those
                # aren't real unmatched line items, just noise.
                if not _looks_like_pdf_noise(item["label"], item["cur"], item["pri"]):
                    warnings.append(f"BS unmatched: [{item['label'][:40]}] CY={item['cur']}")
            if r["confidence"] < 0.5 and r["method"] not in ("rule",):
                warnings.append(f"BS low-conf: [{item['label'][:40]}] → {cell} ({r['confidence']:.2f})")
        buckets["unmatched"] = unmatched
        buckets["all_matches"] = all_matches
        buckets["warnings"] = warnings
        return buckets


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SINGLETON
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

_matcher = None

def get_matcher(pl_csv=None, bs_csv=None):
    global _matcher
    if _matcher is None:
        base = Path(__file__).parent
        if pl_csv is None:
            for p in [base / "pl_keywords_500_each.csv", base.parent / "pl_keywords_500_each.csv"]:
                if p.exists(): pl_csv = str(p); break
        if bs_csv is None:
            for p in [base / "bs_keywords.csv", base.parent / "bs_keywords.csv"]:
                if p.exists(): bs_csv = str(p); break
        _matcher = SmartMatcher(pl_csv, bs_csv)
    return _matcher
