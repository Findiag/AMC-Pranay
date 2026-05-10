#!/usr/bin/env python3
"""
eval/run.py — Mapping accuracy harness.

How it works:
  1. Walks `eval/ground_truth/` for paired files:
       <company>__extracted.xlsx   (input — what the table extractor produced)
       <company>__truth.json       (expected mapping — hand-curated)
  2. Runs the configured mapper (v2_llm | hybrid | rules_only) on each input.
  3. Compares produced cell values vs. ground truth, computes per-field
     precision/recall and overall balance-pass rate.
  4. Emits a markdown report under `eval/reports/<timestamp>.md` and prints
     a one-line summary the CI / human can compare against `baseline.json`.

Ground-truth JSON shape:
{
  "company": "Colgate",
  "variant": "standalone",
  "current_year": {"F5": 1234.5, "F14": 567.8, ...},
  "prior_year":   {"F5": 1100.0, ...}
}

Run:
  python eval/run.py                  # default mode (from config.json)
  python eval/run.py --mode hybrid    # force hybrid (offline) for baseline
  python eval/run.py --mode v2_llm    # force v2 LLM mapper
  python eval/run.py --update-baseline  # write current score to baseline.json
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "modules"))


def _load_truth(truth_path: Path) -> dict:
    return json.loads(truth_path.read_text(encoding="utf-8"))


def _run_mapper_on_extract(extract_path: Path, mode: str) -> dict:
    """Returns {"variant_name": {"current_year": {...}, "prior_year": {...}}}.
    The harness bypasses the full process_file path so it doesn't write
    Excel reports for every eval run — it calls the same mapper functions
    directly on the parsed source items."""
    from bs_pl_mapper import process_source

    variants = process_source(str(extract_path))
    out = {}
    for variant, sheets in variants.items():
        per_year = {"current_year": {}, "prior_year": {}}
        if "bs" in sheets:
            bs = None
            if mode == "claude":
                try:
                    from claude_mapper import map_bs_claude
                    bs = map_bs_claude(sheets["bs"])
                except Exception as e:
                    print(f"  [eval] claude BS failed ({e}); falling back to hybrid")
            elif mode == "v2_llm":
                from llm_mapper_v2 import map_bs_v2
                bs = map_bs_v2(sheets["bs"])
            if bs is None:
                from bs_pl_mapper import map_bs_hybrid
                bs = map_bs_hybrid(sheets["bs"])
            if bs:
                for y in ("current_year", "prior_year"):
                    per_year[y].update(bs.get(y, {}))
        if "pl" in sheets:
            pl = None
            if mode == "claude":
                try:
                    from claude_mapper import map_pl_claude
                    pl = map_pl_claude(sheets["pl"])
                except Exception as e:
                    print(f"  [eval] claude PL failed ({e}); falling back to hybrid")
            elif mode == "v2_llm":
                from llm_mapper_v2 import map_pl_v2
                pl = map_pl_v2(sheets["pl"])
            if pl is None:
                from bs_pl_mapper import map_pl_hybrid
                pl = map_pl_hybrid(sheets["pl"])
            if pl:
                for y in ("current_year", "prior_year"):
                    pyd = pl.get(y, {})
                    for k in ("F66", "F67", "F69", "F70", "F71", "F73"):
                        if k in pyd:
                            per_year[y][k] = pyd[k]
                    # Mirror process_file: F72 = literal OE − OI, F75 =
                    # exceptional. Without this the eval harness sees both
                    # as 0 and any truth value for them counts as a miss.
                    other_exp = float(pyd.get("other_expense", 0) or 0)
                    other_inc = float(pyd.get("other_income", 0) or 0)
                    per_year[y]["F72"] = round(other_exp - other_inc, 2)
                    exc = pyd.get("exceptional_items",
                                   pyd.get("exceptional", 0))
                    per_year[y]["F75"] = round(float(exc or 0), 2)
        out[variant] = per_year
    return out


def _score_one(produced: dict, truth: dict, tol: float = 1.0) -> dict:
    """Per-field exact-match + total balance check.
    Returns a dict with field_correct, field_total, balance_ok, gap."""
    variant = truth.get("variant", "standalone")
    p = produced.get(variant, {})
    field_correct = 0
    field_total = 0
    field_misses = []
    for year in ("current_year", "prior_year"):
        truth_y = truth.get(year, {}) or {}
        prod_y = p.get(year, {}) or {}
        for cell, expected in truth_y.items():
            got = float(prod_y.get(cell, 0) or 0)
            field_total += 1
            if abs(got - float(expected)) <= tol:
                field_correct += 1
            else:
                field_misses.append({
                    "year": year, "cell": cell,
                    "expected": float(expected), "got": got,
                })
    # Balance check on produced (CY only)
    cy = p.get("current_year", {}) or {}
    nw  = sum(cy.get(c, 0) for c in ["F5", "F6", "F7", "F8", "F9"])
    cl  = sum(cy.get(c, 0) for c in ["F14", "F15", "F16", "F17", "F18", "F19"])
    ncl = sum(cy.get(c, 0) for c in ["F24", "F25", "F26", "F27", "F28"])
    ca  = sum(cy.get(c, 0) for c in ["F35", "F36", "F37", "F38", "F39",
                                       "F40", "F41", "F42", "F43"])
    nca = sum(cy.get(c, 0) for c in ["F48", "F49", "F50", "F51", "F52",
                                       "F53", "F54", "F55"])
    gap = round((ca + nca) - (nw + cl + ncl), 2)
    return {
        "field_correct": field_correct,
        "field_total":   field_total,
        "balance_ok":    abs(gap) <= tol,
        "gap":           gap,
        "misses":        field_misses,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", default=None,
                    help="mapper mode: v2_llm | hybrid | rules_only "
                         "(default: read from config.json)")
    ap.add_argument("--ground-truth-dir", default=str(Path(__file__).parent / "ground_truth"))
    ap.add_argument("--reports-dir",      default=str(Path(__file__).parent / "reports"))
    ap.add_argument("--update-baseline",  action="store_true",
                    help="Write the current overall score into baseline.json")
    ap.add_argument("--tol", type=float, default=1.0)
    ap.add_argument("--diagnose", action="store_true",
                    help="Run extract_diagnostic on each input before mapping; "
                         "skip files with FAIL checks (their issues mean a low "
                         "score would not actually reflect mapper quality).")
    args = ap.parse_args()

    gt_dir   = Path(args.ground_truth_dir)
    rep_dir  = Path(args.reports_dir)
    rep_dir.mkdir(parents=True, exist_ok=True)
    gt_dir.mkdir(parents=True, exist_ok=True)

    if args.mode:
        # Override config for this run
        from config import save_config
        save_config({"mapper_mode": args.mode})
    from config import load_config
    mode = load_config().get("mapper_mode", "hybrid")

    pairs = []
    for truth_file in sorted(gt_dir.glob("*__truth.json")):
        company = truth_file.name.replace("__truth.json", "")
        extract = gt_dir / f"{company}__extracted.xlsx"
        if extract.exists():
            pairs.append((company, extract, truth_file))

    if not pairs:
        print(f"[eval] No (extract, truth) pairs found in {gt_dir}")
        print("       Drop pairs as: <company>__extracted.xlsx + <company>__truth.json")
        sys.exit(1)

    print(f"[eval] mode={mode}  pairs={len(pairs)}")
    rows = []
    total_correct = 0
    total_fields  = 0
    total_balance_ok = 0
    skipped_due_to_extract = []
    for company, extract, truth_file in pairs:
        if args.diagnose:
            from extract_diagnostic import diagnose
            diag = diagnose(str(extract))
            if diag.has_failures():
                fails = "; ".join(c.name for c in diag.by_status("fail"))
                print(f"[eval]   {company}: SKIPPED (extract has failures: {fails})")
                skipped_due_to_extract.append({"company": company, "fails": fails})
                rows.append({"company": company,
                             "error": f"Skipped — extract failed diagnostic: {fails}"})
                continue
            elif diag.has_warnings():
                warns = "; ".join(c.name for c in diag.by_status("warn"))
                print(f"[eval]   {company}: diagnostic warnings → {warns}")
        print(f"[eval]   running {company}...")
        try:
            produced = _run_mapper_on_extract(extract, mode)
        except Exception as e:
            print(f"[eval]   {company}: ERROR {type(e).__name__}: {e}")
            rows.append({"company": company, "error": str(e)})
            continue
        truth = _load_truth(truth_file)
        score = _score_one(produced, truth, tol=args.tol)
        rows.append({"company": company, **score})
        total_correct += score["field_correct"]
        total_fields  += score["field_total"]
        total_balance_ok += int(score["balance_ok"])

    # Aggregate
    field_acc = (total_correct / total_fields) if total_fields else 0
    balance_pct = (total_balance_ok / len(pairs)) if pairs else 0
    summary = {
        "mode":             mode,
        "pairs":            len(pairs),
        "field_correct":    total_correct,
        "field_total":      total_fields,
        "field_accuracy":   round(field_acc, 4),
        "balance_pass":     total_balance_ok,
        "balance_pct":      round(balance_pct, 4),
        "timestamp":        dt.datetime.now().isoformat(timespec="seconds"),
    }

    # Markdown report
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    md = [f"# Mapper eval — {ts}",
          f"- **Mode**: `{mode}`",
          f"- **Pairs**: {len(pairs)}",
          f"- **Field accuracy**: {total_correct}/{total_fields} ({field_acc*100:.2f}%)",
          f"- **Balance pass**: {total_balance_ok}/{len(pairs)} ({balance_pct*100:.2f}%)",
          "", "| Company | Fields | Acc | Balance | Gap |",
          "|---|---|---|---|---|"]
    for r in rows:
        if "error" in r:
            md.append(f"| {r['company']} | — | — | ERROR | {r['error'][:40]} |")
        else:
            acc = (r["field_correct"] / r["field_total"]) if r["field_total"] else 0
            md.append(f"| {r['company']} | {r['field_correct']}/{r['field_total']} "
                      f"| {acc*100:.1f}% | {'✓' if r['balance_ok'] else '✗'} "
                      f"| {r['gap']:+.2f} |")
    # Per-company miss detail
    md.append("\n## Misses\n")
    for r in rows:
        if r.get("misses"):
            md.append(f"### {r['company']}")
            for m in r["misses"][:20]:
                md.append(f"  - {m['year']} {m['cell']}: expected {m['expected']:.2f}, got {m['got']:.2f}")
    (rep_dir / f"{ts}_{mode}.md").write_text("\n".join(md), encoding="utf-8")

    print(json.dumps(summary, indent=2))

    if args.update_baseline:
        (Path(__file__).parent / "baseline.json").write_text(
            json.dumps(summary, indent=2), encoding="utf-8")
        print("[eval] baseline.json updated.")
    else:
        # Compare to baseline if it exists
        bl = Path(__file__).parent / "baseline.json"
        if bl.exists():
            base = json.loads(bl.read_text(encoding="utf-8"))
            delta_acc = round(field_acc - base.get("field_accuracy", 0), 4)
            delta_bal = round(balance_pct - base.get("balance_pct", 0), 4)
            sign = lambda x: "+" if x >= 0 else ""
            print(f"[eval] vs baseline ({base.get('mode')}): "
                  f"acc {sign(delta_acc)}{delta_acc*100:.2f}pp | "
                  f"balance {sign(delta_bal)}{delta_bal*100:.2f}pp")


if __name__ == "__main__":
    main()
