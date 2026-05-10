"""
airtable_uploader.py — Upload AMC_RAG `*_Report.xlsx` files to Airtable.

What it does:
  1. Auto-detects the company name and financial year from the report filename
     (e.g. "Shahlon_Silk_Industries_FY2425_BS_PL_extracted_Report.xlsx" →
     company "Shahlon Silk Industries", CY 2024-04-01..2025-03-31,
     PY 2023-04-01..2024-03-31).
  2. Parses the vertical Excel template (column D = labels, E = Standalone CY,
     F = Standalone PY) into two dicts of Airtable-field → number.
  3. Looks up the linked user (by email) and company (by name) records.
  4. Uploads Prior Year first, then Current Year linked to it via
     `previous_bs&pl`.

Run as a CLI:
  python modules/airtable_uploader.py path/to/Company_Report.xlsx
  python modules/airtable_uploader.py path/to/Company_Report.xlsx \
                                      --company "Override Name" \
                                      --user-email someone@askmycfo.ai \
                                      --fy FY2425

Or programmatically:
  from airtable_uploader import upload_report
  result = upload_report("path/to/Report.xlsx")

Credentials are read from config.json (see DEFAULTS in config.py). Never
hard-code tokens in this file.

The FIELD_MAP and SECTION_AWARE_FIELDS dicts below are byte-for-byte from the
working uploader — DO NOT "fix" the typos (e.g. `general_resource&surplus`,
`deffered_tax_asset_(NCA)`); they match the actual Airtable field names.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    import pandas as pd
except ImportError as _e:
    pd = None  # type: ignore
    _PANDAS_ERROR = _e

import requests

# Read config from the flask_app root
sys.path.insert(0, str(Path(__file__).parent.parent))
from config import load_config


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# FIELD MAPPINGS — preserved EXACTLY from the working uploader.
# Typos in field names are intentional; they match Airtable.
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

FIELD_MAP: Dict[str, Optional[str]] = {
    # --- Balance Sheet: Equity ---
    "Share Capital":                       "share_capital",
    "Retained Earnings":                   "retained_earnings",
    "General Reserves & Surplus":          "general_resource&surplus",
    "Other Equity":                        "othere_equity",

    # --- Balance Sheet: Current Liabilities ---
    "Accounts Payable":                    "account_payable",
    "Provisions":                          "provisions_(CL)",
    "Short term borrowings":               "short_term _borrowings",  # space before _borrowings (sic)
    "Other Current Liabilities":           "other_current_liabilities",
    "Other Financial Liabilities":         "other_financial_liabilities_(CL)",
    # F19 row no longer exists in the report (Lease Current folds into F18
    # per Ind AS 116). Kept as a no-op alias for any legacy report files.
    "Lease Liabilities (Current)":         None,

    # --- Balance Sheet: Non-Current Liabilities ---
    "Long Term borrowings":                "long_term_borrowings",
    "Provision":                           "provisions_(NCL)",
    # F26 — renamed when Lease NC was moved to F27. Both labels accepted
    # so old report xlsx files (pre-renaming) still upload correctly.
    "Others (DTL + Other NCL)":            "other_non-current_liabilities",
    "Others (Lease NC + DTL + Other NCL)": "other_non-current_liabilities",

    # --- Balance Sheet: Current Assets ---
    "Bank Balance":                        "bank_balance",
    "Cash & Cash Equivalence":             "cash&cash_eqivalance",
    "Inventory":                           "inventory",
    "Accounts Receivable":                 "Accounts Receivable",
    "Other Current Assets":                "other_current_assets",

    # --- Balance Sheet: Non-Current Assets ---
    "Fixed Assets (PPE+Intangible+ROU+Goodwill)": "fixed_assets",
    "Capital Work-in-Progress":            "capital_working_progress",
    "Other Non-Current Assets":            "other_non-current_assets",
    "Deferred Tax Assets":                 "deffered_tax_asset_(NCA)",

    # --- P&L ---
    "Revenue":                             "revenue",
    "Revenue from Operations":             "revenue",
    "Cost of goods sold":                  "cost_of_goods_sold",
    "Employee benefits expense":           "employee_expenses",
    "Interest":                            "interest",
    "Finance costs":                       "interest",
    "Depreciation":                        "depreciation",
    "Depreciation & amortization expense": "depreciation",
    "Other expenses less other income":    "other_expenses_less_other_income",
    "Taxes":                               "taxes",
    # NEW visible rows from the Schedule III re-layout. Their values
    # already roll into F74 "Other expenses less other income" via
    # Excel formula, so don't double-upload them.
    "Other Expenses":                      None,
    "Other Income":                        None,
    "Profit before Exceptional Items & Tax": None,
    "Profit Before Tax":                   None,
}

# (label, section) → Airtable field — for labels that recur across sections
SECTION_AWARE_FIELDS: Dict[Tuple[str, str], str] = {
    ("Investments", "current_assets"):                          "investments_(CA)",
    ("Investments", "non_current_assets"):                      "investments_(NCA)",
    ("Loans", "current_assets"):                                "loans_(CA)",
    ("Loans", "non_current_assets"):                            "loans_(NCA)",
    ("Other Financial Assets", "current_assets"):               "other_financial_assets_(CA)",
    ("Other Financial Assets", "non_current_assets"):           "other_financial_asset_(NCA)",
    ("Other Financial Liabilities", "current_liabilities"):     "other_financial_liabilities_(CL)",
    ("Other Financial Liabilities", "non_current_liabilities"): "other_financial_liabilities_(NCL)",
}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Filename parsing
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# Pieces tacked on by the AMC_RAG pipeline (Stage 1 → 2 → 3) that we want to
# strip off before treating the rest as the company name.
_FILENAME_NOISE_PATTERNS = [
    r"_BS_PL_extracted_Report$",
    r"_extracted_Report$",
    r"_Report$",
    r"_extracted$",
    r"_BS_PL$",
]


def extract_company_name(report_path: str | Path) -> str:
    """Best-effort company-name detection from the report filename.

    Examples:
      Shahlon_Silk_Industries_FY2425_BS_PL_extracted_Report.xlsx
        → "Shahlon Silk Industries"
      Macobs Technologies FY2425_BS_PL_extracted_Report.xlsx
        → "Macobs Technologies"
    """
    stem = Path(report_path).stem
    for pat in _FILENAME_NOISE_PATTERNS:
        stem = re.sub(pat, "", stem)
    # Strip a trailing FY token like "_FY2425" or " FY2425"
    stem = re.sub(r"[\s_]+FY\d{4}\s*$", "", stem)
    # Underscores → spaces, collapse multiples
    stem = re.sub(r"[_]+", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()
    return stem


def extract_fy_dates(report_path: str | Path,
                     fy_override: Optional[str] = None) -> Dict[str, str]:
    """Detect the financial year from the filename and return CY/PY date ranges.

    `fy_override` lets the caller force e.g. "FY2425". Falls back to "today's"
    Indian-FY assumption when no FY token is present in the filename.

    Returns:
      {
        "start_date_current": "YYYY-MM-DD",
        "end_date_current":   "YYYY-MM-DD",
        "start_date_prior":   "YYYY-MM-DD",
        "end_date_prior":     "YYYY-MM-DD",
        "interval_current":   365,
        "interval_prior":     365 or 366,
        "fy_token":           "FY2425",
      }
    """
    src = fy_override or Path(report_path).stem
    cy_start_date_explicit = cy_end_date_explicit = None

    # 1. ISO end-date override (e.g. "2025-03-31" or "2025-12-31"). Used by
    #    the upload form's date picker so users with non-March year-ends can
    #    pick precisely.
    iso_match = re.match(r"^\s*(\d{4})-(\d{2})-(\d{2})\s*$", src)
    if iso_match and fy_override:
        from datetime import date as _date
        ey, em, ed = int(iso_match.group(1)), int(iso_match.group(2)), int(iso_match.group(3))
        cy_end_date_explicit = _date(ey, em, ed)
        # Period start = same calendar day, one year earlier. Picking
        # 2027-01-01 means "12-month period ending 2027-01-01" -> start
        # 2026-01-01 (NOT 2026-01-02 as the previous +1d bug).
        cy_start_date_explicit = _date(ey - 1, em, ed)
        start_yy, end_yy = (cy_start_date_explicit.year - 2000) % 100, (ey - 2000) % 100
    else:
        # Case-insensitive: filenames like "shivam_autotech_fy2223" /
        # "Foo_Fy2526" used to silently miss this regex and fall through
        # to "today's FY" — yielding garbage column labels in the
        # per-company workbook.
        m = re.search(r"FY\s*(\d{2})(\d{2})", src, re.IGNORECASE)
        if m:
            start_yy, end_yy = int(m.group(1)), int(m.group(2))
        else:
            # Fall back to current Indian financial year based on today's date
            import datetime as _dt
            today = _dt.date.today()
            if today.month >= 4:
                start_yy = today.year % 100
                end_yy   = (today.year + 1) % 100
            else:
                start_yy = (today.year - 1) % 100
                end_yy   = today.year % 100

    cy_start_year = 2000 + start_yy
    cy_end_year   = 2000 + end_yy
    py_start_year = cy_start_year - 1
    py_end_year   = cy_end_year - 1

    def _interval(start_year: int, end_year: int) -> int:
        # 366 if the period crosses a leap-year Feb-29
        leap_start = start_year if start_year % 4 == 0 and (start_year % 100 != 0 or start_year % 400 == 0) else None
        leap_end   = end_year   if end_year   % 4 == 0 and (end_year   % 100 != 0 or end_year   % 400 == 0) else None
        return 366 if leap_end == end_year else 365

    if cy_start_date_explicit and cy_end_date_explicit:
        py_start = cy_start_date_explicit.replace(year=cy_start_date_explicit.year - 1)
        py_end   = cy_end_date_explicit.replace(year=cy_end_date_explicit.year - 1)
        return {
            "start_date_current": cy_start_date_explicit.isoformat(),
            "end_date_current":   cy_end_date_explicit.isoformat(),
            "start_date_prior":   py_start.isoformat(),
            "end_date_prior":     py_end.isoformat(),
            "interval_current":   (cy_end_date_explicit - cy_start_date_explicit).days,
            "interval_prior":     (py_end - py_start).days,
            "fy_token":           f"FY{start_yy:02d}{end_yy:02d}",
        }

    return {
        "start_date_current": f"{cy_start_year}-04-01",
        "end_date_current":   f"{cy_end_year}-03-31",
        "start_date_prior":   f"{py_start_year}-04-01",
        "end_date_prior":     f"{py_end_year}-03-31",
        "interval_current":   _interval(cy_start_year, cy_end_year),
        "interval_prior":     _interval(py_start_year, py_end_year),
        "fy_token":           f"FY{start_yy:02d}{end_yy:02d}",
    }


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Excel parser — preserves the section-aware label routing
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def parse_report(file_path: str | Path) -> Tuple[Dict[str, float], Dict[str, float], List[str]]:
    """Parse a `*_Report.xlsx` produced by bs_pl_mapper.process_file.

    Returns (current_year_fields, prior_year_fields, warnings).
    Only writable (numeric) Airtable fields are returned; formula fields are
    skipped because Airtable computes them automatically.
    """
    if pd is None:
        raise RuntimeError(f"pandas not installed: {_PANDAS_ERROR}. "
                           f"Run `pip install pandas openpyxl`.")

    df = pd.read_excel(file_path)

    # The template puts data in column D onwards: D = "Particulars",
    # E = Standalone CY, F = Standalone PY (then optional G/H for Consolidated).
    data = df.iloc[:, 3:6].copy()
    data.columns = ["Particulars", "Current_Year", "Prior_Year"]
    data = data.iloc[1:].reset_index(drop=True)  # drop the header row

    current_year: Dict[str, float] = {}
    prior_year:   Dict[str, float] = {}
    warnings:     List[str] = []
    section: Optional[str] = None

    for _, row in data.iterrows():
        label = str(row["Particulars"]).strip() if pd.notna(row["Particulars"]) else ""
        label_lower = label.lower().strip()

        # Section header detection — sets context for context-aware mapping
        if label_lower in ("current liabilities",):
            section = "current_liabilities"; continue
        if label_lower in ("non - current liabilities", "non-current liabilities"):
            section = "non_current_liabilities"; continue
        if label_lower in ("current assets",):
            section = "current_assets"; continue
        if label_lower in ("non - current assets", "non-current assets"):
            section = "non_current_assets"; continue
        if "p&l" in label_lower:
            section = "pl"; continue
        if "networth" in label_lower and "total" not in label_lower:
            section = "equity"; continue

        if not label or label == "nan":
            continue

        # Skip section totals, calculated rows, irrelevant headers
        if label_lower.startswith("total") or "difference" in label_lower:
            continue
        if label_lower in ("gross profit", "net profit", "balance sheet particulars"):
            continue
        # Exceptional Items shows in the Excel report for accounting clarity
        # (it's needed to make NP balance), but there's no Airtable field
        # for it — fold into NP rather than uploading separately.
        if label_lower.startswith("exceptional"):
            continue

        # Context-aware mapping wins over plain mapping
        airtable_field = SECTION_AWARE_FIELDS.get((label, section))
        if not airtable_field:
            airtable_field = FIELD_MAP.get(label)

        if airtable_field is None:
            if label in FIELD_MAP and FIELD_MAP[label] is None:
                continue  # explicitly mapped to None = intentionally skipped
            if "spare" in label_lower or "nci" in label_lower:
                continue
            warnings.append(f"No Airtable mapping for '{label}' (section={section})")
            continue

        # Parse numeric values
        for store, key in ((current_year, "Current_Year"), (prior_year, "Prior_Year")):
            v = row[key]
            if pd.notna(v):
                try:
                    store[airtable_field] = float(v)
                except (ValueError, TypeError):
                    warnings.append(f"Non-numeric for '{label}' [{key}]: {v!r}")

    return current_year, prior_year, warnings


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Airtable helpers
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _headers(token: str) -> Dict[str, str]:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def find_record_id(base_id: str, table_id: str, field_name: str, value: str,
                   token: str) -> Optional[str]:
    """Look up an Airtable record by a field value. Returns the record id
    or None when not found."""
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}"
    safe = (value or "").replace('"', '\\"')
    params = {
        "filterByFormula": f'{{{field_name}}} = "{safe}"',
        "maxRecords": 1,
    }
    resp = requests.get(url, headers=_headers(token), params=params, timeout=30)
    if resp.status_code != 200:
        return None
    records = resp.json().get("records", [])
    return records[0]["id"] if records else None


def find_company_id_fuzzy(base_id: str, table_id: str, value: str, token: str,
                          log=print) -> Optional[Tuple[str, str]]:
    """Find a company record by name with fuzzy fallback.

    Strategy:
      1. Exact match on `company_name` field.
      2. Case-insensitive contains match (e.g. 'Akiko' -> 'Akiko Global Services Limited').
      3. First-word match (handles 'Shahlon Silk Industries' vs 'Shahlon Silk Industries Ltd').
    Returns (record_id, matched_name) or None.
    """
    if not value:
        return None
    # 1. Exact
    rec_id = find_record_id(base_id, table_id, "company_name", value, token)
    if rec_id:
        return rec_id, value

    # 2. & 3. Pull all companies once and search locally — Airtable's
    # SEARCH() formula doesn't expose a full-text option, and one paged
    # fetch is cheaper than crafting a complex formula.
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}"
    out = []
    offset = None
    while True:
        params = {"pageSize": 100, "fields[]": "company_name"}
        if offset:
            params["offset"] = offset
        r = requests.get(url, headers=_headers(token), params=params, timeout=30)
        if r.status_code != 200:
            break
        data = r.json()
        out.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset:
            break

    needle = value.lower().strip()
    needle_first = needle.split()[0] if needle else ""

    # Case-insensitive contains
    for r in out:
        n = (r.get("fields", {}).get("company_name") or "").strip()
        if not n: continue
        if needle in n.lower() or n.lower() in needle:
            log(f"[airtable] fuzzy company match: '{value}' ~> '{n}' (substring)")
            return r["id"], n

    # First-word match
    if needle_first:
        for r in out:
            n = (r.get("fields", {}).get("company_name") or "").strip()
            if not n: continue
            if n.lower().startswith(needle_first):
                log(f"[airtable] fuzzy company match: '{value}' ~> '{n}' (first-word)")
                return r["id"], n

    return None


def upload_record(base_id: str, table_id: str, fields: Dict, token: str
                  ) -> Tuple[bool, object]:
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}"
    payload = {"records": [{"fields": fields}], "typecast": True}
    resp = requests.post(url, headers=_headers(token), json=payload, timeout=60)
    if resp.status_code == 200:
        return True, resp.json()["records"][0]["id"]
    return False, resp.json()


def update_record(base_id: str, table_id: str, record_id: str, fields: Dict,
                  token: str) -> Tuple[bool, object]:
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}/{record_id}"
    payload = {"fields": fields, "typecast": True}
    resp = requests.patch(url, headers=_headers(token), json=payload, timeout=60)
    if resp.status_code == 200:
        return True, resp.json()["id"]
    return False, resp.json()


def find_bspl_record_id(base_id: str, table_id: str, company_id: str,
                        end_date: str, token: str,
                        company_name: str = "") -> Optional[str]:
    """Find an existing bs&pl row for this company + end_date. Returns the
    Airtable record id or None. Used by upload_report() to upsert instead
    of creating duplicates on every re-upload of the same FY.

    Important Airtable quirk: filterByFormula on a linked-record field
    (`{company}`) does NOT expose the linked record IDs — it expands to the
    linked record's PRIMARY FIELD (the company name). So we have to filter
    by company_name, not company_id, and then VERIFY client-side that the
    company_id actually matches. Without the client-side verify, a search
    for "Ambuja" would also match "Ambuja Cements Foundation" via a naive
    FIND() substring and stomp on the wrong row.

    DATETIME_FORMAT is used instead of IS_SAME because IS_SAME on a Date
    field with a string argument has been observed to silently return false
    for valid same-day pairs in some Airtable bases.
    """
    if not end_date:
        return None
    name = (company_name or "").strip()
    if not name and company_id:
        # Fall back: fetch the company's name via the record id. One extra
        # API call per upload but keeps the function backwards-compatible.
        try:
            cfg = load_config()
            comp_tbl = cfg.get("airtable_company_table_id", "")
            if comp_tbl:
                url = f"https://api.airtable.com/v0/{base_id}/{comp_tbl}/{company_id}"
                r = requests.get(url, headers=_headers(token), timeout=30)
                if r.status_code == 200:
                    name = (r.json().get("fields", {}).get("company_name") or "").strip()
        except Exception:
            pass
    if not name:
        return None
    safe = name.replace("'", "\\'")
    # Separator-wrapped exact match: avoids "Ambuja" matching "Ambuja Cements
    # Foundation". ARRAYJOIN({company}, '|') gives e.g. 'Ambuja Cements';
    # we search for '|Ambuja Cements|' inside '|...|' so partials never hit.
    formula = (f"AND(FIND('|{safe}|', ('|' & ARRAYJOIN({{company}}, '|') & '|')),"
               f"DATETIME_FORMAT({{end_date}}, 'YYYY-MM-DD') = '{end_date}')")
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}"
    resp = requests.get(url, headers=_headers(token),
                        params={"filterByFormula": formula,
                                "maxRecords": 5,
                                "fields[]": ["end_date", "company"]},
                        timeout=30)
    if resp.status_code != 200:
        return None
    recs = resp.json().get("records", [])
    if not recs:
        return None
    # Client-side verify: when a company_id is supplied, only accept a
    # record whose linked-company array contains exactly that id. This is
    # the last line of defence against duplicate-key collisions.
    if company_id:
        for r in recs:
            comp_links = r.get("fields", {}).get("company") or []
            if company_id in comp_links:
                return r["id"]
        return None
    return recs[0]["id"]


def _link_next_year(base_id: str, table_id: str, company_id: str,
                    this_end_date_iso: str, this_record_id: str,
                    token: str, company_name: str, log) -> Optional[str]:
    """Forward-chain backfill: if a bs&pl row for the year AFTER this one
    already exists for this company AND its previous_bs&pl link is empty,
    PATCH that next-year row to point at `this_record_id`.

    Closes the chain when reports are uploaded out of order — e.g. user
    uploads FY24-25 first (creating an unlinked FY23-24 row), then later
    uploads FY23-24. Without this, the FY24-25 row's previous_bs&pl would
    only get set if the user re-uploads FY24-25.

    Never overwrites an existing previous_bs&pl link (defensive against
    manual curation in Airtable). Returns the next-year record id when
    found (linked or not), None otherwise.
    """
    if not this_end_date_iso or not this_record_id:
        return None
    try:
        from datetime import date as _date
        end = _date.fromisoformat(this_end_date_iso)
        try:
            next_end = end.replace(year=end.year + 1)
        except ValueError:
            # Feb 29 -> Feb 28
            next_end = end.replace(year=end.year + 1, day=28)
    except Exception as e:
        log(f"[airtable] forward-link date math failed (non-fatal): {e}")
        return None

    next_id = find_bspl_record_id(base_id, table_id, company_id,
                                   next_end.isoformat(), token,
                                   company_name=company_name)
    if not next_id:
        return None

    # Read the next-year row to see whether previous_bs&pl is already set.
    try:
        url = f"https://api.airtable.com/v0/{base_id}/{table_id}/{next_id}"
        r = requests.get(url, headers=_headers(token),
                         params={"fields[]": "previous_bs&pl"}, timeout=30)
        if r.status_code != 200:
            log(f"[airtable] forward-link read failed for {next_id} "
                f"[{r.status_code}]: {r.text[:200]}")
            return next_id
        existing = r.json().get("fields", {}).get("previous_bs&pl") or []
    except Exception as e:
        log(f"[airtable] forward-link read raised (non-fatal): {e}")
        return next_id

    if existing:
        # Already linked — leave it alone (could be manually curated).
        if this_record_id not in existing:
            log(f"[airtable] next-year row {next_id} (end_date="
                f"{next_end.isoformat()}) already linked to {existing} — "
                f"NOT overwriting with {this_record_id}")
        return next_id

    ok, _ = update_record(base_id, table_id, next_id,
                          {"previous_bs&pl": [this_record_id]}, token)
    if ok:
        log(f"[airtable] forward-linked {next_id} (end_date="
            f"{next_end.isoformat()}) -> previous_bs&pl=[{this_record_id}]")
    else:
        log(f"[airtable] forward-link PATCH failed for {next_id}")
    return next_id


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Public API
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def upload_report(report_path: str | Path,
                  company_override: Optional[str] = None,
                  user_email_override: Optional[str] = None,
                  fy_override: Optional[str] = None,
                  cfg_override: Optional[Dict] = None,
                  dry_run: bool = False,
                  log: Optional[callable] = None) -> Dict:
    """Upload one `*_Report.xlsx` to Airtable. Returns a result dict.

    `cfg_override` lets callers (e.g. tests) skip config.json. `dry_run=True`
    does everything except the actual POSTs and returns the payloads.
    """
    log = log or print
    cfg = cfg_override or load_config()

    token = (cfg.get("airtable_api_token") or "").strip()
    base_id = cfg.get("airtable_base_id", "").strip()
    bs_pl_table = cfg.get("airtable_bs_pl_table_id", "").strip()
    user_table = cfg.get("airtable_user_table_id", "").strip()
    company_table = cfg.get("airtable_company_table_id", "").strip()
    default_email = (user_email_override
                     or cfg.get("airtable_user_email", "").strip())

    if not (token and base_id and bs_pl_table):
        return {"ok": False, "error": "Airtable credentials missing in config.json "
                                       "(airtable_api_token / _base_id / _bs_pl_table_id)."}

    company_name = company_override or extract_company_name(report_path)
    fy = extract_fy_dates(report_path, fy_override=fy_override)

    log(f"[airtable] file={Path(report_path).name}")
    log(f"[airtable] company='{company_name}' user='{default_email}' fy={fy['fy_token']}")
    log(f"[airtable] CY={fy['start_date_current']}..{fy['end_date_current']} "
        f"PY={fy['start_date_prior']}..{fy['end_date_prior']}")

    cy_data, py_data, warns = parse_report(report_path)
    for w in warns:
        log(f"[airtable] WARN: {w}")
    log(f"[airtable] parsed CY: {len(cy_data)} fields, PY: {len(py_data)} fields")

    # Lookup linked records
    user_id = company_id = None
    matched_name = company_name  # canonical primary-field name; used by upsert
    if user_table and default_email:
        user_id = find_record_id(base_id, user_table, "email", default_email, token)
        log(f"[airtable] user lookup '{default_email}' -> {user_id or 'NOT FOUND'}")
    if company_table and company_name:
        match = find_company_id_fuzzy(company_table_id := company_table,
                                       base_id=base_id, value=company_name,
                                       token=token, log=log) if False else None
        # call the helper with positional args (above ternary is awkward)
        match = find_company_id_fuzzy(base_id, company_table, company_name, token, log=log)
        if match:
            company_id, matched_name = match
            log(f"[airtable] company lookup '{company_name}' -> {company_id} "
                f"({'exact' if matched_name == company_name else f'fuzzy: {matched_name}'})")
        else:
            # Not found even fuzzy — auto-create so observations + reports
            # have something to attach to. Better than silently dropping the
            # link and leaving the bs&pl row orphaned.
            log(f"[airtable] company '{company_name}' NOT FOUND — auto-creating new row")
            try:
                new_url = f"https://api.airtable.com/v0/{base_id}/{company_table}"
                resp = requests.post(new_url, headers=_headers(token),
                                     json={"fields": {"company_name": company_name},
                                           "typecast": True}, timeout=30)
                if resp.status_code < 300:
                    company_id = resp.json().get("id")
                    log(f"[airtable] auto-created company '{company_name}' -> {company_id}")
                else:
                    log(f"[airtable] company auto-create FAILED [{resp.status_code}]: "
                        f"{resp.text[:200]}")
            except Exception as e:
                log(f"[airtable] company auto-create exception: {e}")

    def _build_record(year_data: Dict[str, float], year: str) -> Dict:
        rec = dict(year_data)
        if year == "current":
            rec["interval(days)"] = fy["interval_current"]
            rec["start_date"] = fy["start_date_current"]
            rec["end_date"]   = fy["end_date_current"]
        else:
            rec["interval(days)"] = fy["interval_prior"]
            rec["start_date"] = fy["start_date_prior"]
            rec["end_date"]   = fy["end_date_prior"]
        if user_id:
            rec["user"] = [user_id]
        if company_id:
            rec["company"] = [company_id]
        return rec

    prior_record = _build_record(py_data, "prior")
    current_record = _build_record(cy_data, "current")

    # ── Look up the year-BEFORE-PY (PYPY) record so the new PY row can
    # link to it via previous_bs&pl. Without this the chain breaks for the
    # oldest year in any upload — e.g. uploading an FY24-25 report leaves
    # the new FY23-24 (PY) row with no link to the existing FY22-23 row,
    # even when that row already exists in the base from an earlier
    # upload. The CY upsert below already does this for FY24-25 -> FY23-24;
    # this block extends the same idempotent behaviour to PY.
    pypy_id: Optional[str] = None
    try:
        from datetime import date as _date
        py_end = _date.fromisoformat(fy["end_date_prior"])
        try:
            pypy_end = py_end.replace(year=py_end.year - 1)
        except ValueError:
            # Feb 29 -> Feb 28 in non-leap year
            pypy_end = py_end.replace(year=py_end.year - 1, day=28)
        pypy_id = find_bspl_record_id(base_id, bs_pl_table, company_id,
                                       pypy_end.isoformat(), token,
                                       company_name=matched_name)
        if pypy_id:
            log(f"[airtable] PY-of-PY row found at end_date="
                f"{pypy_end.isoformat()} -> {pypy_id} (will link via "
                f"previous_bs&pl on new PY row)")
            prior_record["previous_bs&pl"] = [pypy_id]
    except Exception as e:
        # Don't fail the whole upload if the lookup hiccups — just skip
        # the link. The user can re-upload the older FY later to repair it.
        log(f"[airtable] PY-of-PY lookup skipped (non-fatal): {e}")

    if dry_run:
        return {"ok": True, "dry_run": True, "company": company_name,
                "company_record_id": company_id, "user_record_id": user_id,
                "fy": fy, "prior": prior_record, "current": current_record,
                "pypy_id": pypy_id, "warnings": warns}

    # ── Upsert (not blind-create) to prevent duplicate rows ──
    # Old behavior: every upload created 2 new bs&pl rows. After 3 uploads
    # of the same Akiko FY24-25 PDF, you'd have 6 rows → 6 observations →
    # endless re-runs. Now we look for an existing row at (company, end_date)
    # and PATCH it instead, so re-uploads are idempotent.
    py_existing = find_bspl_record_id(base_id, bs_pl_table, company_id,
                                       fy["end_date_prior"], token,
                                       company_name=matched_name)
    if py_existing:
        log(f"[airtable] PY row exists for end_date={fy['end_date_prior']} -> "
            f"updating {py_existing} (no new duplicate)")
        ok, result = update_record(base_id, bs_pl_table, py_existing, prior_record, token)
        py_action = "updated"
    else:
        log("[airtable] uploading Prior Year (new row)...")
        ok, result = upload_record(base_id, bs_pl_table, prior_record, token)
        py_action = "created"
    if not ok:
        log(f"[airtable] PY {py_action} FAILED: {json.dumps(result)[:300]}")
        return {"ok": False, "stage": "prior_year", "error": result}
    prior_id = result
    log(f"[airtable] PY {py_action} -> {prior_id}")

    # Forward-chain backfill: PY's next year is the CY of THIS upload, which
    # we're about to upsert anyway. So we only forward-link from CY (below).
    # Note: we deliberately don't forward-link from PY here because that
    # would race with the CY upsert that follows immediately.
    time.sleep(0.3)

    current_record["previous_bs&pl"] = [prior_id]
    cy_existing = find_bspl_record_id(base_id, bs_pl_table, company_id,
                                       fy["end_date_current"], token,
                                       company_name=matched_name)
    if cy_existing:
        log(f"[airtable] CY row exists for end_date={fy['end_date_current']} -> "
            f"updating {cy_existing} (no new duplicate)")
        ok, result = update_record(base_id, bs_pl_table, cy_existing,
                                   current_record, token)
        cy_action = "updated"
    else:
        log("[airtable] uploading Current Year (new row)...")
        ok, result = upload_record(base_id, bs_pl_table, current_record, token)
        cy_action = "created"
    if not ok:
        log(f"[airtable] CY {cy_action} FAILED: {json.dumps(result)[:300]}")
        return {"ok": False, "stage": "current_year", "prior_id": prior_id,
                "error": result}
    current_id = result
    log(f"[airtable] CY {cy_action} -> {current_id}")

    # Forward-chain backfill: if a row for the year AFTER this CY already
    # exists (e.g. user uploaded FY24-25 first then is now uploading
    # FY23-24, so FY24-25 is sitting in Airtable with previous_bs&pl=[]),
    # back-link it now. No-op when there's no next-year row, and never
    # overwrites an existing link.
    next_id = _link_next_year(base_id, bs_pl_table, company_id,
                                fy["end_date_current"], current_id, token,
                                matched_name, log)

    return {"ok": True, "prior_id": prior_id, "current_id": current_id,
            "pypy_id": pypy_id, "next_year_id": next_id,
            "py_action": py_action, "cy_action": cy_action,
            "company": company_name, "company_record_id": company_id,
            "user_record_id": user_id, "fy": fy, "warnings": warns}


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CLI
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def main() -> int:
    ap = argparse.ArgumentParser(description="Upload an AMC_RAG *_Report.xlsx to Airtable")
    ap.add_argument("report", help="Path to a single *_Report.xlsx file, "
                                    "OR a directory containing many.")
    ap.add_argument("--company",    help="Override the auto-detected company name")
    ap.add_argument("--user-email", help="Override the user email from config")
    ap.add_argument("--fy",         help="Override FY token (e.g. FY2425) used for date detection")
    ap.add_argument("--dry-run",    action="store_true",
                    help="Build the payloads and print them; don't POST.")
    args = ap.parse_args()

    target = Path(args.report)
    if target.is_dir():
        files = sorted(set(list(target.glob("*_Report.xlsx"))
                           + list(target.glob("*_Report*.xlsx"))))
    else:
        files = [target]

    if not files:
        print(f"[airtable] No *_Report.xlsx files at {target}", file=sys.stderr)
        return 2

    overall_ok = True
    for f in files:
        print(f"\n=== {f.name} ===")
        result = upload_report(
            f,
            company_override=args.company,
            user_email_override=args.user_email,
            fy_override=args.fy,
            dry_run=args.dry_run,
        )
        if not result.get("ok"):
            overall_ok = False
            print(f"[airtable] FAILED: {result.get('error')}")
        elif result.get("dry_run"):
            print(json.dumps(
                {"company": result["company"], "fy": result["fy"]["fy_token"],
                 "current_field_count": len(result["current"]) - 4,
                 "prior_field_count":   len(result["prior"]) - 4,
                 "warnings": len(result["warnings"])},
                indent=2,
            ))
    return 0 if overall_ok else 1


if __name__ == "__main__":
    sys.exit(main())
