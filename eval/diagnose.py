#!/usr/bin/env python3
"""
eval/diagnose.py — batch diagnostic runner.

Walks `eval/ground_truth/` (or a custom dir) and runs the extract_diagnostic
checks against every `*__extracted.xlsx` file. Emits a single consolidated
markdown report plus per-file detail.

Usage:
  python eval/diagnose.py                          # walk eval/ground_truth/
  python eval/diagnose.py path/to/dir              # walk a different dir
  python eval/diagnose.py path/to/file.xlsx        # single file
  python eval/diagnose.py --json                   # machine-readable output
  python eval/diagnose.py --strict                 # exit 1 if any file has failures
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "modules"))

from extract_diagnostic import diagnose, DiagnosticReport, FAIL, WARN, PASS, INFO  # noqa


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("path", nargs="?",
                    default=str(Path(__file__).parent / "ground_truth"),
                    help="File or directory (default: eval/ground_truth/)")
    ap.add_argument("--json", action="store_true",
                    help="Emit JSON to stdout (suitable for piping).")
    ap.add_argument("--strict", action="store_true",
                    help="Exit code 1 if any file has FAIL checks.")
    ap.add_argument("--reports-dir",
                    default=str(Path(__file__).parent / "reports"),
                    help="Directory to write the consolidated markdown report.")
    args = ap.parse_args()

    target = Path(args.path)
    if target.is_dir():
        files = sorted(set(list(target.glob("*_extracted.xlsx"))
                           + list(target.glob("*__extracted.xlsx"))))
    elif target.is_file():
        files = [target]
    else:
        print(f"[diagnose] {target} not found", file=sys.stderr)
        return 2

    if not files:
        print(f"[diagnose] No *_extracted.xlsx files in {target}", file=sys.stderr)
        print(f"[diagnose] Drop pairs into {target} as:", file=sys.stderr)
        print(f"           Company__extracted.xlsx + Company__truth.json", file=sys.stderr)
        return 2

    reports = [diagnose(str(f)) for f in files]

    if args.json:
        print(json.dumps([r.to_dict() for r in reports], indent=2))
    else:
        # Consolidated header
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        n_fail = sum(1 for r in reports if r.has_failures())
        n_warn = sum(1 for r in reports if r.has_warnings() and not r.has_failures())
        n_pass = len(reports) - n_fail - n_warn

        lines = [
            f"# Extraction diagnostic — {ts}",
            "",
            f"- **Files scanned**: {len(reports)}",
            f"- **Healthy**: {n_pass} ✓",
            f"- **With warnings only**: {n_warn} ⚠",
            f"- **With failures**: {n_fail} ✗",
            "",
            "## Summary table",
            "",
            "| File | Verdict | Pass | Warn | Fail |",
            "|---|---|---|---|---|",
        ]
        for r in reports:
            v = "✗ FAIL" if r.has_failures() else ("⚠ WARN" if r.has_warnings() else "✓ OK")
            lines.append(
                f"| `{Path(r.file_path).name}` | {v} | "
                f"{len(r.by_status(PASS))} | {len(r.by_status(WARN))} | "
                f"{len(r.by_status(FAIL))} |"
            )
        lines.append("")
        lines.append("---")
        lines.append("")

        for r in reports:
            lines.append(r.to_markdown())
            lines.append("")
            lines.append("---")
            lines.append("")

        out = "\n".join(lines)

        # Write consolidated report file
        rep_dir = Path(args.reports_dir)
        rep_dir.mkdir(parents=True, exist_ok=True)
        out_path = rep_dir / f"{ts}_diagnostic.md"
        out_path.write_text(out, encoding="utf-8")

        # Print summary table to stdout, full report to file
        print("\n".join(lines[:lines.index("## Summary table") + 4 + len(reports)]))
        print()
        print(f"[diagnose] Full per-file detail written to: {out_path}")

    return 1 if args.strict and any(r.has_failures() for r in reports) else 0


if __name__ == "__main__":
    sys.exit(main())
