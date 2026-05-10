# Mapper eval harness

The accuracy harness lives here. **Nothing else in `AMC_RAG/flask_app` should
be modified to claim an accuracy improvement until the harness number moves.**

## How to populate ground truth

For every report you want to track, drop two files into `eval/ground_truth/`:

```
ground_truth/
├── Colgate__extracted.xlsx     # the *_extracted.xlsx that Stage 2 produces
└── Colgate__truth.json         # hand-curated correct mapping
```

The `__truth.json` schema:

```json
{
  "company": "Colgate",
  "variant": "standalone",
  "current_year": {
    "F5":  100.0,
    "F14": 230.5,
    "F40": 1500.0,
    "F66": 8000.0
  },
  "prior_year": {
    "F5":  100.0,
    "F14": 210.0
  }
}
```

You don't need to fill in every cell — only the ones you've verified. The
harness scores each cell that's present in `truth.json`. Cells the truth file
omits don't count for or against the model.

**Recommended starter set:** 5–10 reports across different formats (Schedule III
Ind AS, Old GAAP, manufacturing, services, IT). Don't over-curate at first;
start with the obvious cells (Share Capital, PPE, Trade Receivables, Trade
Payables, Inventories, Revenue) and grow over time.

## Running

```bash
# 1. Lock current pipeline as the baseline (run once)
python eval/run.py --mode hybrid --update-baseline

# 2. Run the new mapper and compare
python eval/run.py --mode v2_llm
# → prints "vs baseline (hybrid): acc +X.XXpp | balance +X.XXpp"

# 3. After every prompt / threshold change:
python eval/run.py --mode v2_llm
# → must move acc forward, never backward
```

## Reading the report

`eval/reports/<timestamp>_<mode>.md` shows:

- **Field accuracy** — per-cell exact-match rate within `--tol` (default 1.0
  Lakhs). This is the primary number.
- **Balance pass** — fraction of reports where Total Assets ≈ Total L+E. A
  cheap secondary signal.
- **Misses table** — every cell that disagreed with truth, so you can see
  *which* fields the model is confusing (this is the actionable signal).

## Pre-flight diagnostic on extracts

Before you spend time wondering why the mapper got something wrong, run the
diagnostic to find out whether the extract itself is mappable:

```bash
# Single file
python modules/extract_diagnostic.py path/to/Company_extracted.xlsx

# Every file in eval/ground_truth/ at once (writes a consolidated report)
python eval/diagnose.py

# Pipe to other tools
python eval/diagnose.py --json | jq '.[] | select(.summary.fail > 0) | .file'

# Skip-on-fail integration with the eval harness
python eval/run.py --mode v2_llm --diagnose
# → companies whose extract has FAIL checks are skipped (their score wouldn't
#   actually reflect mapper quality, only extraction quality)
```

The diagnostic catches issues like:

- **PPE row missing** (the Shalon Silks bug — drops 60–80% of NCA)
- **Other Equity present as one line, no Note 11 breakdown** → F6/F7/F8 split
  impossible regardless of mapper
- **Section totals don't reconcile** to items above them → rows lost in extraction
- **Section headers (NCA / CA / NCL / CL / Equity) absent** → mapper has to
  guess from item signatures, introduces small error rate
- **Trade Payables MSME / non-MSME not both present** → multi-row aggregation
  test won't fire
- **Orphan-value rows** (values with no label) or **value-in-label rows**
  (extractor failed to split label / value)
- **Magnitude check** — values look like rupees / crores instead of lakhs
- **P&L missing required rows** (Revenue / COGS / Employee / Finance / Dep / Tax / NP)

Verdict is one of:

- ✓ **EXTRACTION LOOKS HEALTHY** — safe to run mapper
- ⚠ **MAPPING POSSIBLE WITH CAVEATS** — see warnings (mapper will produce
  output but some cells may be approximate)
- ✗ **MAPPING WILL LIKELY FAIL** — fix extraction issues first; no prompt
  change can recover from missing data

## What "acceptable accuracy" looks like

| Field accuracy | Verdict |
|---|---|
| < 70% | Mapper is failing. Don't deploy. |
| 70–85% | Production-viable with reviewer in the loop. |
| 85–95% | Strong. Reviewer catches edge cases. |
| > 95% | Reviewer-optional for the long tail. |

These thresholds are rough — what matters is **the trend** vs. the baseline.
