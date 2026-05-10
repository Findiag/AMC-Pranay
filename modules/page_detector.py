"""
BS & P&L Page/Section Detector v2
===================================
Finds Balance Sheet and Profit & Loss sections from ANY Indian annual report.

Detection strategy (two-pass):
  PASS 1 — Title-line anchoring: Find lines like "Balance Sheet as at March 31, 2025"
           or "Statement of Profit and Loss" that are SHORT standalone headers,
           then verify data markers exist in the next 80 lines.
  PASS 2 — Data-pattern fallback: Slide a window across the text, score by
           marker density + number density. Catches reports with missing headers.
"""

import re
import os
import zipfile
import glob
from dataclasses import dataclass, field
from typing import Optional

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False


# ─── Structural markers ─────────────────────────────────────────────

BS_MARKERS = [
    "total assets", "total equity", "equity and liabilities",
    "non-current assets", "non current assets",
    "current assets", "share capital",
    "total non-current", "total non current",
    "total current assets", "total current liabilities",
    "trade receivables", "property, plant and equipment",
    "property plant and equipment", "borrowings", "inventories",
    "other equity", "trade payables", "investments", "deferred tax",
    "reserves and surplus", "shareholders' funds",
    "total equity and liabilities", "total equity & liabilities",
    # NBFC format
    "financial assets", "non financial assets", "non-financial assets",
    "total financial assets", "total non financial assets",
    "financial liabilities", "non-financial liabilities",
    "total financial liabilities",
    "liabilities and equity",
]

PL_MARKERS = [
    "revenue from operations", "total income", "total expenses",
    "total expense", "profit before tax", "profit for the year",
    "profit for the period", "tax expense", "total tax expense",
    "employee benefit", "employee benefits expense",
    "depreciation", "finance cost", "finance costs",
    "other income", "other expenses", "earnings per share",
    "earning per share", "comprehensive income",
    "cost of materials consumed", "cost of goods sold",
    "operating expense", "purchase of services",
    "profit before exceptional",
]

NEGATIVE_MARKERS = [
    "notes to financial statements", "notes forming part of",
    "significant accounting policies", "accounting policy",
    "contingent liabilities", "related party", "segment reporting",
    "auditor's report", "auditors' report",
    "director's report", "directors' report",
    "management discussion", "corporate governance",
    "secretarial audit", "risk management",
]

CASHFLOW_MARKERS = [
    "cash flow from operating", "cash flows from operating",
    "cash flow from investing", "cash flows from investing",
    "cash flow from financing", "cash flows from financing",
    "net increase in cash", "net decrease in cash",
]

# Title patterns for section headers
BS_TITLE_PATTERNS = [
    r"(?:standalone|consolidated)?\s*balance\s+sheet\s+(?:as\s+(?:at|on))",
    r"balance\s+sheet\s+as\s+(?:at|on)\s+(?:\d|march|april|sep|december)",
    r"(?:standalone|consolidated)\s+balance\s+sheet",
    r"balance\s+sheet$",  # line that is JUST "Balance Sheet"
]

PL_TITLE_PATTERNS = [
    r"(?:standalone|consolidated)?\s*statement\s+of\s+profit\s+and\s+loss",
    r"(?:standalone|consolidated)?\s*profit\s+and\s+loss\s+(?:account|statement)",
    r"statement\s+of\s+profit\s+or\s+loss",
    r"statement\s+of\s+profit\s+and\s+loss$",
]


@dataclass
class DetectedSection:
    section_type: str           # 'balance_sheet' or 'profit_and_loss'
    variant: str                # 'standalone' or 'consolidated'
    source_type: str            # 'pdf_page', 'text_section', 'image_archive'
    page_number: Optional[int]  # start page (1-indexed)
    page_end: Optional[int] = None  # end page for multi-page tables
    start_line: Optional[int] = None
    end_line: Optional[int] = None
    score: float = 0.0
    text_preview: str = ""
    markers_found: list = field(default_factory=list)


def _count_numbers(text: str) -> int:
    return len(re.findall(r'[\d,]+\.\d{1,2}', text))


def _detect_variant(text: str) -> str:
    if "consolidated" in text[:600].lower():
        return "consolidated"
    return "standalone"


def _is_cashflow(text: str) -> bool:
    lower = text.lower()
    return sum(1 for m in CASHFLOW_MARKERS if m in lower) >= 2


def _score_section(text: str, section_type: str) -> tuple[float, list]:
    lower = text.lower()
    score = 0.0
    matched = []

    markers = BS_MARKERS if section_type == "balance_sheet" else PL_MARKERS
    for m in markers:
        if m in lower:
            score += 5
            matched.append(m)

    # Number density
    num_count = _count_numbers(text)
    density = num_count / max(len(text) / 1000, 0.1)
    if density > 15:
        score += 20
    elif density > 8:
        score += 12
    elif density > 3:
        score += 5

    # Penalties
    for neg in NEGATIVE_MARKERS:
        if neg in lower:
            score -= 8
    if _is_cashflow(text):
        score -= 25

    # Bonuses for column headers
    if re.search(r'\bparticulars\b', lower):
        score += 5
    if re.search(r'as\s+at\s+\d', lower) or re.search(r'as\s+on\s+\d', lower):
        score += 5
    if re.search(r'\bnotes?\b\s', lower) and re.search(r'as\s+at|as\s+on|for\s+the\s+year', lower):
        score += 3

    return score, matched


def _find_end(lines, start, section_type):
    total = len(lines)
    if section_type == "balance_sheet":
        end_kw = ["total equity and liabilities", "total equity & liabilities",
                  "total liabilities and equity", "total liabilities"]
        stop_kw = ["statement of profit and loss", "profit and loss account",
                    "cash flow", "statement of cash flow", "changes in equity"]
    else:
        end_kw = ["earnings per share", "earning per share",
                   "comprehensive income for the year", "comprehensive income for the period",
                   "total comprehensive income"]
        stop_kw = ["balance sheet", "cash flow", "statement of cash flow",
                    "changes in equity"]

    for i in range(start + 8, min(start + 120, total)):
        lower = lines[i].lower().strip()
        for em in end_kw:
            if em in lower:
                return min(i + 5, total)
        if i > start + 15:
            for sm in stop_kw:
                if sm in lower and len(lower) < 120:
                    return i
    return min(start + 80, total)


def detect_file_type(filepath: str) -> str:
    with open(filepath, "rb") as f:
        header = f.read(10)
    if header[:5] == b"%PDF-":
        return "pdf"
    elif header[:2] == b"PK":
        return "zip"
    return "text"


# ─── TEXT file detector ──────────────────────────────────────────────

def detect_from_text(filepath: str) -> list[DetectedSection]:
    with open(filepath, "r", encoding="utf-8", errors="replace") as f:
        lines = f.readlines()

    results = []
    total = len(lines)
    used_ranges = []

    def _overlaps(s, e):
        for (rs, re_, _) in used_ranges:
            if s < re_ and e > rs:
                return True
        return False

    # ── PASS 1: Title-line anchored ──
    for i, line in enumerate(lines):
        lower = line.lower().strip()
        if len(lower) < 10 or len(lower) > 200:
            continue

        found_type = None
        for p in BS_TITLE_PATTERNS:
            if re.search(p, lower):
                found_type = "balance_sheet"
                break
        if not found_type:
            for p in PL_TITLE_PATTERNS:
                if re.search(p, lower):
                    found_type = "profit_and_loss"
                    break
        if not found_type:
            continue

        # Verify data exists nearby (next 15 lines)
        lookahead = "\n".join(lines[i:min(i + 15, total)]).lower()
        if found_type == "balance_sheet":
            if not any(kw in lookahead for kw in ["assets", "equity", "non-current", "current",
                                                   "share capital", "liabilities", "borrowings",
                                                   "payables", "financial assets"]):
                continue
        else:
            if not any(kw in lookahead for kw in ["income", "revenue", "expense", "continuing", "operations"]):
                continue

        # Find section boundaries
        # Start: title line, or earlier if "Particulars" or company header is above
        # Some reports (especially NBFCs) have the title mid-section, so search wider
        start = i
        for j in range(max(0, i - 50), i):
            jlower = lines[j].lower().strip()
            if ("particulars" in jlower and ("note" in jlower or "as at" in jlower or "for the year" in jlower)):
                start = j
                break
            # Also catch the header block: company name + "(Figure in Lakhs)" pattern
            if re.search(r'(?:amount|figure|₹|rs\.?|inr)\s*(?:in\s*)?(?:lakhs?|crore|million)', jlower):
                # Look ahead from here for the actual data start
                for k in range(j, min(j + 8, i + 1)):
                    if "particulars" in lines[k].lower() or "assets" in lines[k].lower() or "equity" in lines[k].lower():
                        start = max(j - 3, 0)
                        break
                if start != i:
                    break

        end = _find_end(lines, start, found_type)

        if _overlaps(start, end):
            continue

        block = "".join(lines[start:end])
        if _is_cashflow(block):
            continue

        score, matched = _score_section(block, found_type)
        if score >= 20 and len(matched) >= 3:
            variant = _detect_variant("".join(lines[max(0, start - 10):start + 5]) + block[:200])
            used_ranges.append((start, end, found_type))
            results.append(DetectedSection(
                section_type=found_type,
                variant=variant,
                source_type="text_section",
                page_number=None,
                start_line=start + 1,
                end_line=end,
                score=score,
                text_preview=block[:300],
                markers_found=matched,
            ))

    # ── PASS 2: Data-pattern fallback ──
    # Only for types not yet found
    for section_type in ["balance_sheet", "profit_and_loss"]:
        existing = [r for r in results if r.section_type == section_type]
        if len(existing) >= 2:
            continue

        existing_variants = {r.variant for r in existing}
        WINDOW = 80
        best_candidates = []

        for i in range(0, total - 20, 25):
            if _overlaps(i, i + WINDOW):
                continue
            block = "".join(lines[i:min(i + WINDOW, total)])
            score, matched = _score_section(block, section_type)

            if score >= 45 and len(matched) >= 5 and not _is_cashflow(block):
                end = _find_end(lines, i, section_type)
                block = "".join(lines[i:end])
                score, matched = _score_section(block, section_type)
                variant = _detect_variant(block)

                if variant not in existing_variants and score >= 40:
                    best_candidates.append((score, i, end, matched, variant, block))

        # Take the best candidate(s)
        best_candidates.sort(reverse=True)
        for score, start, end, matched, variant, block in best_candidates[:2]:
            if not _overlaps(start, end):
                used_ranges.append((start, end, section_type))
                results.append(DetectedSection(
                    section_type=section_type,
                    variant=variant,
                    source_type="text_section",
                    page_number=None,
                    start_line=start + 1,
                    end_line=end,
                    score=score,
                    text_preview=block[:300],
                    markers_found=matched,
                ))

    # Deduplicate: best per (type, variant)
    seen = {}
    for r in results:
        key = (r.section_type, r.variant)
        if key not in seen or r.score > seen[key].score:
            seen[key] = r
    return sorted(seen.values(), key=lambda x: x.score, reverse=True)


# ─── PDF detector (multi-page aware) ─────────────────────────────────

# End markers — table is COMPLETE when we find one of these
BS_END_MARKERS = [
    "total equity and liabilities", "total equity & liabilities",
    "total liabilities and equity", "total liabilities",
]
PL_END_MARKERS = [
    "earnings per share", "earning per share",
    "earnings per equity share",
    "basic and diluted",
]

# Title patterns for page-level detection
BS_PAGE_TITLES = [
    r"balance\s+sheet",
]
PL_PAGE_TITLES = [
    r"statement\s+of\s+profit\s+and\s+loss",
    r"profit\s+and\s+loss\s+(?:account|statement)",
    r"statement\s+of\s+profit\s+or\s+loss",
]

# Markers that mean we've hit a DIFFERENT section (stop scanning)
STOP_TITLES = [
    r"cash\s+flow\s+statement",
    r"statement\s+of\s+cash\s+flow",
    r"statement\s+of\s+changes\s+in\s+equity",
    r"notes\s+to\s+(?:the\s+)?(?:standalone|consolidated)?\s*financial\s+statements",
    r"notes\s+forming\s+part",
    r"significant\s+accounting\s+polic",
]


def _page_has_title(text: str, patterns: list) -> bool:
    """Check if page has a section title."""
    # Only check the first 400 chars (title should be near top)
    top = text[:400].lower()
    return any(re.search(p, top) for p in patterns)


def _page_is_new_section(text: str) -> bool:
    """Check if this page starts a different section (not a continuation)."""
    top = text[:400].lower()
    return any(re.search(p, top) for p in STOP_TITLES)


def _page_has_end_marker(text: str, section_type: str) -> bool:
    """Check if this page contains the end marker for the section."""
    lower = text.lower()
    markers = BS_END_MARKERS if section_type == "balance_sheet" else PL_END_MARKERS
    return any(m in lower for m in markers)


def _is_continuation_page(text: str, section_type: str) -> bool:
    """
    Check if a page is likely a continuation of a financial table.
    Continuation pages have: numbers, financial terms, but NO new section title.
    """
    lower = text.lower()

    # Must not be a new section
    if _page_is_new_section(text):
        return False

    # Must not be a different financial statement title
    if section_type == "balance_sheet":
        if _page_has_title(text, PL_PAGE_TITLES):
            return False
    else:
        if _page_has_title(text, BS_PAGE_TITLES):
            return False

    # Must have some financial numbers
    num_count = _count_numbers(text)
    if num_count < 5:
        return False

    # Should have at least some relevant markers
    markers = BS_MARKERS if section_type == "balance_sheet" else PL_MARKERS
    marker_hits = sum(1 for m in markers if m in lower)
    if marker_hits < 2:
        return False

    return True


def detect_from_pdf(filepath: str) -> list[DetectedSection]:
    """
    Detect BS and P&L from real PDFs with multi-page table support.

    For each detected section:
    1. Find the title page
    2. Check if the table is complete (has end marker)
    3. If not, scan forward until we find the end marker or a new section
    4. Store full page range (page_number to page_end)
    """
    if not HAS_PDFPLUMBER:
        raise ImportError("pdfplumber required")

    # Extract text from all pages, parallelized. pdfplumber per-page text
    # extraction is the main page_detector bottleneck — a 200-page annual
    # report used to take ~60s sequential; ThreadPoolExecutor cuts this to
    # ~15-20s on 4 workers (CPU-bound on most PDFs but pdfplumber releases
    # the GIL for the underlying C extraction work).
    from concurrent.futures import ThreadPoolExecutor
    with pdfplumber.open(filepath) as pdf:
        n = len(pdf.pages)

        def _extract(i: int) -> str:
            try:
                return pdf.pages[i].extract_text() or ""
            except Exception:
                return ""

        # min(8, n) — diminishing returns past ~4-8 workers; pdfplumber's
        # internal locks can throttle higher concurrency.
        with ThreadPoolExecutor(max_workers=min(8, max(1, n))) as ex:
            page_texts = list(ex.map(_extract, range(n)))

    total_pages = len(page_texts)
    results = []
    used_pages = set()  # avoid double-counting

    # PASS 1: Find title pages
    candidates = []
    for idx, text in enumerate(page_texts):
        if len(text.strip()) < 80:
            continue

        page_num = idx + 1

        # Check for BS title
        if _page_has_title(text, BS_PAGE_TITLES) and not _is_cashflow(text):
            score, matched = _score_section(text, "balance_sheet")
            if score >= 25 and len(matched) >= 3:
                variant = _detect_variant(text)
                candidates.append((score, idx, "balance_sheet", variant, matched, text))

        # Check for P&L title
        if _page_has_title(text, PL_PAGE_TITLES) and not _is_cashflow(text):
            score, matched = _score_section(text, "profit_and_loss")
            if score >= 25 and len(matched) >= 3:
                variant = _detect_variant(text)
                candidates.append((score, idx, "profit_and_loss", variant, matched, text))

    # Sort by score descending
    candidates.sort(reverse=True)

    # PASS 2: For each candidate, find the full page range
    for score, start_idx, section_type, variant, matched, text in candidates:
        if start_idx in used_pages:
            continue

        # Check if table is complete on this page
        end_idx = start_idx

        if _page_has_end_marker(page_texts[start_idx], section_type):
            # Complete on single page
            end_idx = start_idx
        else:
            # Scan forward for continuation pages
            for next_idx in range(start_idx + 1, min(start_idx + 5, total_pages)):
                next_text = page_texts[next_idx]

                # Stop if we hit a completely different section
                if _page_is_new_section(next_text):
                    break

                # Stop if this page is a title page for a DIFFERENT type
                if section_type == "balance_sheet" and _page_has_title(next_text, PL_PAGE_TITLES):
                    break
                if section_type == "profit_and_loss" and _page_has_title(next_text, BS_PAGE_TITLES):
                    break

                # This looks like a continuation
                if _is_continuation_page(next_text, section_type) or _page_has_end_marker(next_text, section_type):
                    end_idx = next_idx
                    # Accumulate markers from continuation pages
                    _, extra_matched = _score_section(next_text, section_type)
                    for m in extra_matched:
                        if m not in matched:
                            matched.append(m)

                    if _page_has_end_marker(next_text, section_type):
                        break
                else:
                    break

        # Mark pages as used
        for p in range(start_idx, end_idx + 1):
            used_pages.add(p)

        # Re-score with combined text from all pages
        combined_text = "\n".join(page_texts[start_idx:end_idx + 1])
        final_score, final_matched = _score_section(combined_text, section_type)

        # ── VERIFICATION: confirm this is actual financial data, not notes ──
        if section_type == "balance_sheet":
            lower_combined = combined_text.lower()
            has_total_assets = "total assets" in lower_combined or "total asset" in lower_combined
            has_total_le = any(m in lower_combined for m in BS_END_MARKERS)
            num_count = _count_numbers(combined_text)
            # Notes sections have many keywords but few tabular numbers
            if num_count < 10 and not has_total_assets and not has_total_le:
                # This is likely a notes section, not actual BS
                for p in range(start_idx, end_idx + 1):
                    used_pages.discard(p)
                continue  # skip this candidate

        page_count = end_idx - start_idx + 1
        results.append(DetectedSection(
            section_type=section_type,
            variant=variant,
            source_type="pdf_page",
            page_number=start_idx + 1,
            page_end=end_idx + 1,
            score=final_score,
            text_preview=text[:300],
            markers_found=final_matched,
        ))

    # Deduplicate: keep best per (type, variant)
    seen = {}
    for r in results:
        key = (r.section_type, r.variant)
        if key not in seen or r.score > seen[key].score:
            seen[key] = r
    return sorted(seen.values(), key=lambda x: x.score, reverse=True)


# ─── ZIP detector ────────────────────────────────────────────────────

def detect_from_zip(filepath: str) -> list[DetectedSection]:
    with zipfile.ZipFile(filepath) as z:
        images = [n for n in z.namelist() if n.lower().endswith(('.jpg', '.jpeg', '.png'))]
    return [DetectedSection(
        section_type="unknown", variant="unknown",
        source_type="image_archive", page_number=None,
        start_line=None, end_line=None, score=0,
        text_preview=f"ZIP with {len(images)} page images — needs Vision API.",
        markers_found=[],
    )]


# ─── Main entry ──────────────────────────────────────────────────────

def detect_pages(filepath: str) -> list[DetectedSection]:
    file_type = detect_file_type(filepath)
    filename = os.path.basename(filepath)

    print(f"\n{'='*75}")
    print(f"  {filename}  ({file_type})")
    print(f"{'='*75}")

    if file_type == "pdf":
        results = detect_from_pdf(filepath)
    elif file_type == "zip":
        results = detect_from_zip(filepath)
    else:
        results = detect_from_text(filepath)

    results.sort(key=lambda x: x.score, reverse=True)

    if not results:
        print("  ❌ No BS or P&L sections detected")
    for r in results:
        icon = {"balance_sheet": "📊", "profit_and_loss": "📈"}.get(r.section_type, "🖼️")
        if r.page_number and r.page_end and r.page_end > r.page_number:
            loc = f"Pages {r.page_number}–{r.page_end}"
        elif r.page_number:
            loc = f"Page {r.page_number}"
        elif r.start_line:
            loc = f"Lines {r.start_line}–{r.end_line}"
        else:
            loc = "N/A"
        label = r.section_type.replace("_", " ").upper()
        preview = next((l.strip() for l in r.text_preview.split('\n') if l.strip()), "")

        print(f"\n  {icon} {label} ({r.variant})")
        print(f"     Location : {loc}")
        print(f"     Score    : {r.score:.0f}  |  Markers: {len(r.markers_found)}")
        print(f"     Matched  : {', '.join(r.markers_found[:8])}")
        print(f"     Preview  : {preview[:110]}")

    return results


# ─── Extract detected sections into one combined PDF ────────────────

def extract_pages(filepath: str, output_dir: str = None) -> str:
    """
    Detect BS & P&L, then extract exact table boundaries into
    one single PDF with headings between sections.

    Returns path to created PDF.
    """
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as rl_canvas

    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(filepath))
    os.makedirs(output_dir, exist_ok=True)

    file_type = detect_file_type(filepath)
    stem = os.path.splitext(os.path.basename(filepath))[0]

    sections = detect_pages(filepath)
    sections = [s for s in sections if s.section_type != "unknown"]

    if not sections:
        print("\n  ❌ Nothing to extract")
        return ""

    out_path = os.path.join(output_dir, f"{stem}_BS_PL.pdf")

    if file_type == "text":
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            all_lines = f.readlines()

        c = rl_canvas.Canvas(out_path, pagesize=A4)
        w, h = A4
        margin = 50
        lh = 12

        for i, s in enumerate(sections):
            if not s.start_line:
                continue

            # New page for each section (except first)
            if i > 0:
                c.showPage()

            y = h - margin

            # ── Section heading ──
            label = f"{s.variant.upper()} {s.section_type.replace('_', ' ').upper()}"
            c.setFont("Helvetica-Bold", 14)
            c.drawString(margin, y, label)
            y -= 10
            c.setFont("Helvetica", 9)
            c.setFillColorRGB(0.4, 0.4, 0.4)
            c.drawString(margin, y, f"Lines {s.start_line}–{s.end_line}  |  Score: {s.score:.0f}  |  Markers: {len(s.markers_found)}")
            c.setFillColorRGB(0, 0, 0)
            y -= 8

            # ── Divider line ──
            c.setStrokeColorRGB(0.7, 0.7, 0.7)
            c.line(margin, y, w - margin, y)
            y -= 18

            # ── Exact table content (start_line to end_line, no buffer) ──
            c.setFont("Courier", 8.5)
            start = s.start_line - 1
            end = s.end_line

            for line in all_lines[start:end]:
                if y < margin:
                    c.showPage()
                    c.setFont("Courier", 8.5)
                    y = h - margin
                c.drawString(margin, y, line.rstrip()[:130])
                y -= lh

        c.save()

    elif file_type == "pdf":
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(filepath)
        writer = PdfWriter()
        total = len(reader.pages)
        added_pages = set()

        for s in sections:
            if not s.page_number:
                continue
            start_p = s.page_number - 1
            end_p = (s.page_end - 1) if s.page_end else start_p

            for i in range(start_p, min(end_p + 1, total)):
                if i not in added_pages:
                    writer.add_page(reader.pages[i])
                    added_pages.add(i)

        with open(out_path, "wb") as f:
            writer.write(f)

    print(f"\n  ✅ Created: {os.path.basename(out_path)}")
    print(f"     Sections: {len(sections)}")
    for s in sections:
        label = f"{s.variant} {s.section_type.replace('_', ' ')}"
        if s.start_line:
            loc = f"lines {s.start_line}–{s.end_line}"
        elif s.page_end and s.page_end > s.page_number:
            loc = f"pages {s.page_number}–{s.page_end}"
        elif s.page_number:
            loc = f"page {s.page_number}"
        else:
            loc = "?"
        print(f"     • {label} ({loc})")

    return out_path


# ─── CLI ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage:")
        print("  python page_detector.py report.pdf              # detect only")
        print("  python page_detector.py report.pdf --extract    # detect + extract into single PDF")
        sys.exit(1)

    filepath = sys.argv[1]

    if "--extract" in sys.argv:
        out_dir = None
        if "--output" in sys.argv:
            idx = sys.argv.index("--output")
            if idx + 1 < len(sys.argv):
                out_dir = sys.argv[idx + 1]
        extract_pages(filepath, out_dir)
    else:
        detect_pages(filepath)