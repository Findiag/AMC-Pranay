"""
extract_diagnostic.py — Pre-flight check for `*_extracted.xlsx` files.

Runs ~20 structural and content checks against an extracted xlsx and tells you,
BEFORE you run the mapper, whether the extract has the information the mapper
needs to produce a correct mapping. Surfaces problems like:

  • PPE row missing (the Shalon Silks bug)
  • "Other Equity" present as a single line with no Note 11 breakdown
    (so F6/F7/F8 split is impossible regardless of mapper quality)
  • Trade Payables MSME / non-MSME rows not both present
  • Section headers (NCA / CA / NCL / CL / Equity) absent → section detection
    will have to guess from item signatures
  • Section totals don't reconcile to the items above them (extraction lost rows)
  • Orphan-value rows (values with no label) or value-in-label rows
    (label and value not split by extractor)
  • Magnitude check (figures look like rupees instead of lakhs)
  • PL: Revenue / COGS / Employee / Finance / Depreciation / Tax / Net Profit
    rows all present

Usage:
  python modules/extract_diagnostic.py path/to/Company_extracted.xlsx
  python modules/extract_diagnostic.py path/to/Company_extracted.xlsx --json

  # programmatic:
  from extract_diagnostic import diagnose
  report = diagnose("Company_extracted.xlsx")
  if report.has_failures(): ...
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

# Reuse the existing parser so diagnostics see exactly the same data shape the
# mapper sees — there is no point in checking a representation the mapper
# wouldn't actually consume.
import importlib.util
import os

_THIS_DIR = Path(__file__).parent
sys.path.insert(0, str(_THIS_DIR))


def _load_parser():
    """Late-import process_source / extract_items / classify_sheets to avoid
    pulling the whole openai/chromadb stack just to read an xlsx."""
    from bs_pl_mapper import process_source, extract_items, classify_sheets
    return process_source, extract_items, classify_sheets


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Data model
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

PASS = "pass"
WARN = "warn"
FAIL = "fail"
INFO = "info"


@dataclass
class CheckResult:
    name: str
    status: str             # pass | warn | fail | info
    message: str
    details: List[str] = field(default_factory=list)
    suggestion: str = ""

    @property
    def emoji(self) -> str:
        return {"pass": "✓", "warn": "⚠", "fail": "✗", "info": "·"}.get(self.status, "?")


@dataclass
class DiagnosticReport:
    file_path: str
    variants: Dict[str, Dict[str, List[Dict]]] = field(default_factory=dict)
    checks: List[CheckResult] = field(default_factory=list)

    def add(self, **kwargs) -> None:
        self.checks.append(CheckResult(**kwargs))

    def by_status(self, status: str) -> List[CheckResult]:
        return [c for c in self.checks if c.status == status]

    def has_failures(self) -> bool:
        return any(c.status == FAIL for c in self.checks)

    def has_warnings(self) -> bool:
        return any(c.status == WARN for c in self.checks)

    def verdict(self) -> str:
        if self.has_failures():
            return "MAPPING WILL LIKELY FAIL — fix extraction issues first"
        if self.has_warnings():
            return "MAPPING POSSIBLE WITH CAVEATS — see warnings"
        return "EXTRACTION LOOKS HEALTHY — safe to run mapper"

    def to_markdown(self) -> str:
        lines = [f"# Diagnostic: `{Path(self.file_path).name}`", ""]
        lines.append(f"**Verdict:** {self.verdict()}")
        lines.append(f"**Checks:** {len(self.by_status(PASS))} pass · "
                     f"{len(self.by_status(WARN))} warn · "
                     f"{len(self.by_status(FAIL))} fail · "
                     f"{len(self.by_status(INFO))} info")
        lines.append("")

        # Variants summary
        if self.variants:
            lines.append("## Variants found")
            for v, sheets in self.variants.items():
                cols = ", ".join(f"{t}={len([i for i in items if i['cur'] != 0 or i['pri'] != 0])} valued rows"
                                 for t, items in sheets.items())
                lines.append(f"- **{v}**: {cols}")
            lines.append("")

        for stat_label, stat_code in [("Failures", FAIL), ("Warnings", WARN),
                                       ("Passed", PASS), ("Info", INFO)]:
            items = self.by_status(stat_code)
            if not items:
                continue
            lines.append(f"## {stat_label}")
            for c in items:
                lines.append(f"### {c.emoji} {c.name}")
                lines.append(c.message)
                if c.details:
                    lines.append("")
                    for d in c.details[:15]:
                        lines.append(f"  - {d}")
                    if len(c.details) > 15:
                        lines.append(f"  - …and {len(c.details) - 15} more")
                if c.suggestion:
                    lines.append(f"\n  **Fix:** {c.suggestion}")
                lines.append("")
        return "\n".join(lines)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "file": self.file_path,
            "verdict": self.verdict(),
            "summary": {
                "pass": len(self.by_status(PASS)),
                "warn": len(self.by_status(WARN)),
                "fail": len(self.by_status(FAIL)),
                "info": len(self.by_status(INFO)),
            },
            "variants": {
                v: {t: len([i for i in items if i['cur'] != 0 or i['pri'] != 0])
                    for t, items in sheets.items()}
                for v, sheets in self.variants.items()
            },
            "checks": [
                {"name": c.name, "status": c.status, "message": c.message,
                 "details": c.details, "suggestion": c.suggestion}
                for c in self.checks
            ],
        }


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Helpers
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).lower().strip())


def _find_rows(items: List[Dict], pattern: str) -> List[Dict]:
    rx = re.compile(pattern, re.IGNORECASE)
    return [it for it in items if rx.search(_norm(it["label"]))]


def _has_row(items: List[Dict], pattern: str, with_value: bool = True) -> bool:
    rows = _find_rows(items, pattern)
    if not rows:
        return False
    if with_value:
        return any(r["cur"] != 0 or r["pri"] != 0 for r in rows)
    return True


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Individual checks
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _check_sheets_present(report: DiagnosticReport) -> None:
    if not report.variants:
        report.add(name="Sheets present", status=FAIL,
                   message="No BS or P&L sheets recognised in this file.",
                   suggestion="Check that the xlsx contains sheets with names "
                              "containing 'Balance Sheet' / 'Profit & Loss' / "
                              "'Standalone' / 'Consolidated'.")
        return
    missing = []
    for variant, sheets in report.variants.items():
        if "bs" not in sheets:
            missing.append(f"{variant}: BS missing")
        if "pl" not in sheets:
            missing.append(f"{variant}: P&L missing")
    if missing:
        report.add(name="Sheets present", status=WARN,
                   message="Some variants are missing BS or P&L:",
                   details=missing,
                   suggestion="Confirm whether the source PDF actually has both, "
                              "or whether Stage 1 page detection cut them off.")
    else:
        report.add(name="Sheets present", status=PASS,
                   message=f"All variants have both BS and P&L. "
                           f"Variants: {list(report.variants.keys())}")


def _check_bs_section_headers(report: DiagnosticReport, items: List[Dict],
                               variant: str) -> None:
    # Header regex anchored at start (^) so a row like "Total Non-Current
    # Liabilities" doesn't false-positive as a "non-current liabilities" header
    # via substring match.
    expected = {
        "equity":            r"^(equity\s*$|equity\s+and\s+liab|shareholder|networth)",
        "current_liab":      r"^current\s+liabilit",
        "noncurrent_liab":   r"^non.?\s*current\s+liabilit",
        "current_asset":     r"^current\s+asset",
        "noncurrent_asset":  r"^non.?\s*current\s+asset",
    }
    missing = []
    for sec, pat in expected.items():
        rows = _find_rows(items, pat)
        # Filter out rows that are actually totals (start with "total")
        rows = [r for r in rows if not _norm(r["label"]).startswith("total")]
        if not rows:
            missing.append(sec)
    if not missing:
        report.add(name=f"BS section headers ({variant})", status=PASS,
                   message="All five section headers present.")
    elif len(missing) <= 2:
        report.add(name=f"BS section headers ({variant})", status=WARN,
                   message=f"Section header(s) missing: {', '.join(missing)}",
                   suggestion="Mapper will fall back to item-signature inference. "
                              "Usually OK but introduces a small error rate on "
                              "boundary items like 'Others' or 'Loans'.")
    else:
        report.add(name=f"BS section headers ({variant})", status=FAIL,
                   message=f"Most section headers missing: {', '.join(missing)}",
                   suggestion="Re-run extraction; this PDF's section structure "
                              "wasn't captured.")


def _check_bs_balance(report: DiagnosticReport, items: List[Dict],
                      variant: str, tol_pct: float = 1.0) -> None:
    """If the extract contains 'Total Assets' and 'Total Equity and Liabilities'
    rows, they should match (within tolerance)."""
    total_assets_rows = _find_rows(items, r"^total\s+assets?\s*$")
    total_eql_rows = _find_rows(items, r"^total\s+(equity\s+and\s+liab|liab\s+and\s+equity)")
    if not total_assets_rows or not total_eql_rows:
        report.add(name=f"BS face balances ({variant})", status=INFO,
                   message="No 'Total Assets' / 'Total Equity & Liab' row found "
                           "in extract — cannot verify balance from the face.")
        return
    ta = total_assets_rows[0]["cur"]
    teql = total_eql_rows[0]["cur"]
    if ta == 0 or teql == 0:
        report.add(name=f"BS face balances ({variant})", status=WARN,
                   message=f"Total Assets={ta} or Total Equity+Liab={teql} is "
                           f"zero in extract — extractor probably lost the value.",
                   suggestion="Re-extract; this is a Stage 2 problem.")
        return
    pct_diff = abs(ta - teql) / max(abs(ta), abs(teql)) * 100
    if pct_diff < tol_pct:
        report.add(name=f"BS face balances ({variant})", status=PASS,
                   message=f"Total Assets ({ta:,.2f}) ≈ Total Equity+Liab ({teql:,.2f})")
    else:
        report.add(name=f"BS face balances ({variant})", status=FAIL,
                   message=f"Total Assets ({ta:,.2f}) ≠ Total Equity+Liab "
                           f"({teql:,.2f}); diff {ta-teql:+,.2f} ({pct_diff:.2f}%)",
                   suggestion="The face of the BS in the extract doesn't balance. "
                              "No mapper can fix this — it's an extraction problem.")


def _check_section_totals_reconcile(report: DiagnosticReport, items: List[Dict],
                                     variant: str) -> None:
    """For each 'Total Current Assets', 'Total Non-Current Liabilities', etc.
    sum the items between section header and total row and verify they tie."""
    # Header regex must anchor at start (^) — otherwise "current liabilit"
    # would also match "non-current liabilit" because of substring match,
    # and the section-totals check would span the wrong rows.
    section_pats = [
        ("Current Assets",
         r"^current\s+asset",
         r"^total\s+current\s+asset"),
        ("Non-Current Assets",
         r"^(non.?\s*current\s+asset|fixed\s+asset|long.?\s*term\s+asset)",
         r"^total\s+(non.?\s*current\s+asset|long.?\s*term\s+asset|fixed\s+asset)"),
        ("Current Liabilities",
         r"^current\s+liabilit",
         r"^total\s+current\s+liabilit"),
        ("Non-Current Liabilities",
         r"^(non.?\s*current\s+liabilit|long.?\s*term\s+liab)",
         r"^total\s+(non.?\s*current\s+liabilit|long.?\s*term\s+liab)"),
    ]
    issues = []
    for label, header_pat, total_pat in section_pats:
        header_rx = re.compile(header_pat, re.IGNORECASE)
        total_rx = re.compile(total_pat, re.IGNORECASE)
        # Find first header position and matching total position
        header_idx = next((i for i, it in enumerate(items)
                           if header_rx.search(_norm(it["label"]))
                           and not total_rx.search(_norm(it["label"]))), None)
        total_idx = next((i for i, it in enumerate(items)
                          if i > (header_idx or -1)
                          and total_rx.search(_norm(it["label"]))), None)
        if header_idx is None or total_idx is None:
            continue
        between = items[header_idx + 1:total_idx]
        # Exclude only structural sub-totals (e.g. "Total Non-Current Liab",
        # "Total Equity"), NOT data rows that happen to start with "Total"
        # like "Total outstanding dues of micro and small enterprises".
        subtotal_rx = re.compile(
            r"^total\s+(non.?\s*current|current|equity|networth|shareholder|"
            r"fixed|long.?\s*term|financial|of\s+(equity|asset|liab))",
            re.IGNORECASE,
        )
        valued = [it for it in between
                  if (it["cur"] != 0 or it["pri"] != 0)
                  and not subtotal_rx.match(_norm(it["label"]))]
        item_sum = sum(it["cur"] for it in valued)
        total_val = items[total_idx]["cur"]
        if total_val == 0:
            continue
        gap = total_val - item_sum
        if abs(gap) > max(1.0, abs(total_val) * 0.005):
            issues.append(
                f"{label}: items sum to {item_sum:,.2f} but extract reports "
                f"total {total_val:,.2f} (gap {gap:+,.2f}) — "
                f"{'rows missing from extract' if gap > 0 else 'extra rows or wrong total'}"
            )
    if not issues:
        report.add(name=f"Section totals reconcile ({variant})", status=PASS,
                   message="Every section total matches the sum of its items.")
    else:
        report.add(name=f"Section totals reconcile ({variant})", status=FAIL,
                   message="One or more section totals don't match the items above:",
                   details=issues,
                   suggestion="Extraction lost rows. Re-extract or hand-fix the xlsx "
                              "before running the mapper.")


def _check_other_equity_breakdown(report: DiagnosticReport, items: List[Dict],
                                   variant: str) -> None:
    """Schedule III BS shows 'Other Equity' as a single line. The breakdown
    (Capital Reserve, Securities Premium, Retained Earnings, Revaluation, OCI)
    lives in a Note. Without that breakdown, F6/F7/F8 split is impossible."""
    has_other_equity = _has_row(items, r"other\s+equity|reserves\s+and\s+surplus")
    has_breakdown = any(_has_row(items, p) for p in [
        r"securities\s+premium",
        r"capital\s+reserve",
        r"retained\s+earnings|surplus\s+in\s+(profit|statement)",
        r"general\s+reserve",
        r"revaluation\s+(reserve|surplus)",
    ])
    if not has_other_equity:
        report.add(name=f"Other Equity breakdown ({variant})", status=INFO,
                   message="No Other-Equity / Reserves-and-Surplus row found.")
        return
    if has_breakdown:
        report.add(name=f"Other Equity breakdown ({variant})", status=PASS,
                   message="Other Equity has component rows (premium / reserves / "
                           "retained earnings) — F6/F7/F8 split is possible.")
    else:
        report.add(name=f"Other Equity breakdown ({variant})", status=WARN,
                   message="Only the aggregate 'Other Equity' line is present; "
                           "no breakdown into Capital Reserve / Securities Premium / "
                           "Retained Earnings / OCI.",
                   suggestion="The mapper will be forced to put the entire amount "
                              "into a single cell (F8). If F6/F7 accuracy matters, "
                              "the extractor must capture the Note 11 breakdown.")


def _check_trade_payables(report: DiagnosticReport, items: List[Dict],
                           variant: str) -> None:
    msme = _find_rows(items, r"(msme|micro\s+(and|&)\s+small|"
                              r"outstanding\s+dues?\s+(of|to)\s+(micro|msme))")
    others = _find_rows(items, r"trade\s+payable|sundry\s+creditor|"
                                r"outstanding\s+dues?\s+(of|to)\s+(creditor|other)|"
                                r"creditors?\s+(other\s+than|for)")
    if msme and others:
        report.add(name=f"Trade Payables split ({variant})", status=PASS,
                   message=f"Both MSME and non-MSME trade-payable rows found "
                           f"(MSME: {len(msme)} row(s), Others: {len(others)} row(s)).")
    elif _has_row(items, r"trade\s+payable") or _has_row(items, r"sundry\s+creditor"):
        report.add(name=f"Trade Payables split ({variant})", status=INFO,
                   message="Trade Payables present as a single line "
                           "(no MSME / non-MSME split). Mapping is fine; just "
                           "noting that the multi-row aggregation test is N/A here.")
    else:
        report.add(name=f"Trade Payables split ({variant})", status=WARN,
                   message="No Trade Payables / Sundry Creditors row found at all. "
                           "F14 will be 0 unless the company genuinely has none.",
                   suggestion="If the source PDF has trade payables, re-extract.")


def _check_ppe(report: DiagnosticReport, items: List[Dict],
               variant: str) -> None:
    """The Shalon Silks bug: PPE row dropped during extraction."""
    has_ppe = _has_row(items, r"property.?\s*plant|^tangible\s+asset|"
                              r"plant\s+and\s+machinery|fixed\s+asset")
    if has_ppe:
        report.add(name=f"PPE row present ({variant})", status=PASS,
                   message="Property Plant & Equipment row found.")
    else:
        # CWIP without PPE is a near-definitive sign that PPE was dropped:
        # CWIP only exists as a holding place for in-progress fixed assets, so
        # the company having CWIP but no PPE is structurally implausible.
        has_cwip = _has_row(items, r"capital\s+work.?\s*in.?\s*progress|^cwip|"
                                    r"intangible\s+assets?\s+under\s*dev")
        # Other strong NCA indicators that imply PPE should also exist
        has_other_nca = _has_row(items, r"non.?\s*current\s+investment|"
                                          r"deferred\s+tax\s+asset|"
                                          r"right.?\s*of.?\s*use|"
                                          r"goodwill")
        if has_cwip or has_other_nca:
            report.add(name=f"PPE row present ({variant})", status=FAIL,
                       message="No PPE row found, but other non-current assets are "
                               "present — PPE was probably dropped by the extractor.",
                       suggestion="This is the classic Stage 2 bug (e.g. Shalon Silks). "
                                  "Check the source PDF and re-extract.")
        else:
            report.add(name=f"PPE row present ({variant})", status=INFO,
                       message="No PPE row — could be a service / IT company.")


def _check_orphan_rows(report: DiagnosticReport, items: List[Dict],
                       variant: str) -> None:
    """Rows with no label but values, or labels with values=0 that look like
    wrap continuations — both signal extraction quality issues."""
    orphan_value = [it for it in items
                    if (not it["label"] or it["label"].lower() in ("nan", ""))
                    and (it["cur"] != 0 or it["pri"] != 0)]
    wrap_suffix = ("small", "enterprises", "and", "other", "than",
                   "of", "the", "to", "micro")
    likely_wraps = [it for it in items
                    if it["cur"] == 0 and it["pri"] == 0
                    and any(_norm(it["label"]).endswith(s) for s in wrap_suffix)]
    issues = []
    if orphan_value:
        issues.append(f"{len(orphan_value)} row(s) with values but no label "
                      f"(extractor probably split label/value across rows)")
    if likely_wraps:
        issues.append(f"{len(likely_wraps)} row(s) look like wrapped-label "
                      f"continuations not yet stitched")
    if not issues:
        report.add(name=f"Row hygiene ({variant})", status=PASS,
                   message="No orphan-value or hanging wrap rows.")
    else:
        report.add(name=f"Row hygiene ({variant})", status=WARN,
                   message="Extraction artefacts detected:",
                   details=issues + [
                       *[f"orphan: '{it['label'][:40]}' CY={it['cur']}"
                         for it in orphan_value[:5]],
                       *[f"wrap?: '{it['label'][:60]}'" for it in likely_wraps[:5]],
                   ],
                   suggestion="bs_pl_mapper.extract_items already attempts to stitch "
                              "these. Check whether the existing wrap-suffix list "
                              "needs to be extended for this company's PDF style.")


def _check_value_in_label(report: DiagnosticReport, items: List[Dict],
                           variant: str) -> None:
    """Labels with trailing numbers — the extractor failed to split label/value."""
    bad = []
    for it in items:
        if it["cur"] != 0 or it["pri"] != 0:
            continue
        m = re.search(r"\s+([\d,]+\.?\d+)(\s+([\d,]+\.?\d+))?\s*$", it["label"])
        if m and len(it["label"]) > 15:
            bad.append(f"'{it['label'][:80]}'")
    if not bad:
        report.add(name=f"Value-in-label ({variant})", status=PASS,
                   message="No labels with trailing numbers detected.")
    else:
        report.add(name=f"Value-in-label ({variant})", status=WARN,
                   message=f"{len(bad)} label(s) appear to contain values not "
                           f"separated into columns:",
                   details=bad,
                   suggestion="bs_pl_mapper.extract_items handles common cases of "
                              "this. If it slipped through, the regex needs to be "
                              "extended for this PDF's number formatting.")


def _check_pl_required_rows(report: DiagnosticReport, items: List[Dict],
                             variant: str) -> None:
    required = [
        ("Revenue from Operations",
         r"revenue\s+from\s+operations?|sale\s+of\s+(products?|services?|goods)|"
         r"net\s+sales|^turnover$|^sales\s*$|income\s+from\s+operations?"),
        ("COGS / Materials / Purchases",
         r"cost\s+of\s+(materials?|goods\s+sold|revenue|raw)|"
         r"purchases?\s+of\s+(stock|traded|raw|goods)|"
         r"changes?\s+in\s+inventor|cost\s+of\s+sales|"
         r"operating\s+expenses?\s*$"),
        ("Employee Benefits",
         r"employee\s+benefit|staff\s+(cost|expense)|salaries?\s+(and\s+)?wages?|"
         r"managerial\s+remuneration|payment\s+to\s+employees?|"
         r"contribution\s+to\s+provident|personnel\s+(cost|expense)"),
        ("Finance Cost",
         r"finance\s+cost|interest\s+expense|borrowing\s+cost|"
         r"interest\s+(on|and)\s+(borrowing|loan|term|deposit|debenture)"),
        ("Depreciation",
         r"depreciation|amorti[sz]ation|depletion"),
        ("Tax",
         r"current\s+tax|deferred\s+tax|income\s+tax\s+expense|tax\s+expense|"
         r"tax\s+(of|for)\s+earlier|provision\s+for\s+tax"),
        ("Net Profit",
         r"profit\s+after\s+tax|net\s+profit|profit\s*(loss)?\s+for\s+the\s+(year|period)|"
         r"profit\s+attributable\s+to\s+owners|profit\s+loss\s+for\s+the"),
    ]
    missing = []
    for label, pat in required:
        if not _has_row(items, pat):
            missing.append(label)
    if not missing:
        report.add(name=f"P&L required rows ({variant})", status=PASS,
                   message="All required P&L rows present.")
    elif len(missing) <= 2:
        report.add(name=f"P&L required rows ({variant})", status=WARN,
                   message=f"Optional / minor P&L rows missing: {', '.join(missing)}",
                   suggestion="If the company genuinely has no Finance Cost or "
                              "Depreciation (rare), this is fine. Otherwise re-extract.")
    else:
        report.add(name=f"P&L required rows ({variant})", status=FAIL,
                   message=f"Many P&L rows missing: {', '.join(missing)}",
                   suggestion="Stage 2 extraction is incomplete for the P&L sheet. "
                              "Re-extract or fall back to a different P&L page.")


def _check_value_magnitude(report: DiagnosticReport, items: List[Dict],
                            variant: str) -> None:
    """The template assumes values are in Lakhs. If max absolute value is > 1e8,
    the extract is probably in rupees instead — mapping will produce numbers
    that look right but are a million× wrong."""
    vals = [abs(it["cur"]) for it in items if it["cur"] != 0]
    if not vals:
        return
    mx = max(vals)
    if mx > 1e10:
        report.add(name=f"Value magnitude ({variant})", status=WARN,
                   message=f"Largest absolute value is {mx:,.0f} — looks like "
                           f"the extract is in rupees, not lakhs.",
                   suggestion="Either re-extract specifying the unit, or divide "
                              "all values by 1e5 (Lakh) / 1e7 (Crore) before mapping.")
    elif mx < 10 and len(vals) > 5:
        report.add(name=f"Value magnitude ({variant})", status=WARN,
                   message=f"Largest absolute value is only {mx:.2f} — looks like "
                           f"values may be in Crores, not Lakhs.",
                   suggestion="Multiply by 100 (Crore→Lakh) before mapping.")
    else:
        report.add(name=f"Value magnitude ({variant})", status=PASS,
                   message=f"Value magnitudes look like Lakhs (max abs = {mx:,.2f}).")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Public entry point
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def diagnose(extract_path: str) -> DiagnosticReport:
    """Run all checks against a `*_extracted.xlsx` file and return a structured
    report. Never raises on bad input — wraps parser exceptions as FAIL checks."""
    report = DiagnosticReport(file_path=str(extract_path))

    if not Path(extract_path).exists():
        report.add(name="File exists", status=FAIL,
                   message=f"File not found: {extract_path}")
        return report

    try:
        process_source, _, _ = _load_parser()
        report.variants = process_source(extract_path)
    except Exception as e:
        report.add(name="Parse extract", status=FAIL,
                   message=f"Failed to parse extract: {type(e).__name__}: {e}",
                   suggestion="The xlsx may be corrupt or in an unexpected format.")
        return report

    _check_sheets_present(report)

    for variant, sheets in report.variants.items():
        bs_items = sheets.get("bs", [])
        pl_items = sheets.get("pl", [])

        if bs_items:
            _check_bs_section_headers(report, bs_items, variant)
            _check_bs_balance(report, bs_items, variant)
            _check_section_totals_reconcile(report, bs_items, variant)
            _check_other_equity_breakdown(report, bs_items, variant)
            _check_trade_payables(report, bs_items, variant)
            _check_ppe(report, bs_items, variant)
            _check_orphan_rows(report, bs_items, variant)
            _check_value_in_label(report, bs_items, variant)
            _check_value_magnitude(report, bs_items, variant)

        if pl_items:
            _check_pl_required_rows(report, pl_items, variant)
            _check_orphan_rows(report, pl_items, f"{variant}-PL")
            _check_value_in_label(report, pl_items, f"{variant}-PL")

    return report


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CLI
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def main() -> int:
    ap = argparse.ArgumentParser(description="Diagnostic for *_extracted.xlsx files")
    ap.add_argument("path", help="Path to a single *_extracted.xlsx file, "
                                  "OR a directory containing many.")
    ap.add_argument("--json", action="store_true",
                    help="Emit JSON instead of markdown.")
    ap.add_argument("--quiet", action="store_true",
                    help="Print only the verdict line per file (for batch use).")
    args = ap.parse_args()

    target = Path(args.path)
    if target.is_dir():
        files = sorted(list(target.glob("*_extracted.xlsx"))
                       + list(target.glob("*__extracted.xlsx")))
        if not files:
            print(f"[diagnostic] No *_extracted.xlsx files in {target}")
            return 1
    else:
        files = [target]

    overall_fail = False
    out: List[Dict[str, Any]] = []
    for f in files:
        rep = diagnose(str(f))
        if rep.has_failures():
            overall_fail = True
        if args.json:
            out.append(rep.to_dict())
        elif args.quiet:
            print(f"{rep.emoji_for_verdict()} {f.name}: {rep.verdict()} "
                  f"({len(rep.by_status(FAIL))}F / {len(rep.by_status(WARN))}W)")
        else:
            print(rep.to_markdown())
            print()

    if args.json:
        print(json.dumps(out if len(out) > 1 else out[0], indent=2))

    return 1 if overall_fail else 0


# Convenience for --quiet mode (avoid attaching method on dataclass directly)
def _emoji_for_verdict(self):
    if self.has_failures():
        return "✗"
    if self.has_warnings():
        return "⚠"
    return "✓"
DiagnosticReport.emoji_for_verdict = _emoji_for_verdict


if __name__ == "__main__":
    sys.exit(main())
