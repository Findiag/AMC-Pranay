# Ask My CFO — Annual Report → Standardized Financial Template

Flask service that converts Indian annual-report PDFs (Ind AS / Schedule III / Old GAAP) into a single normalized Balance Sheet + P&L template. Built around Anthropic Claude as the primary mapper, with a rule-based extractor and a PDF-grounded remediation pass for failure recovery.

```
PDF → page_detector → extract_tables (rule-based)
                           ↓
                    sanity check (9 heuristics)
                     ├── broken? → Claude PDF native extraction
                     └── ok     → keep
                           ↓
                    bs_pl_mapper (Claude tool-use, 3 retries + PDF-grounded
                                   remediation pass on failure)
                           ↓
                    sanity check (BS imbalance, zero-density, YoY anomaly)
                     ├── broken? → Claude PDF re-extract + re-map
                     └── ok     → done
                           ↓
                    Airtable upsert (idempotent, chain-linking)
                     + per-company multi-year workbook
```

## Quick start (local dev)

```bash
git clone <your-fork-url> ask-my-cfo
cd ask-my-cfo

python -m venv .venv
source .venv/bin/activate              # Windows: .venv\Scripts\activate

pip install -r requirements.txt

# Configure secrets — pick ONE:
#   (A) edit config.json directly  (gitignored, local only)
#   (B) export env vars (production-safe)
cp config.example.json config.json
# Open config.json, paste your ANTHROPIC_API_KEY (required).
# Airtable keys are optional; needed only if you want auto-upload.

python app.py
# → http://localhost:5000
```

Open http://localhost:5000 and drag in any annual-report PDF.

## Deploying to Railway

See **[DEPLOY_RAILWAY.md](DEPLOY_RAILWAY.md)** for the step-by-step. TL;DR:

1. Push this repo to GitHub
2. Railway → New Project → Deploy from GitHub
3. Set env vars (see `.env.example`) — at minimum `ANTHROPIC_API_KEY` and `API_AUTH_TOKEN`
4. Mount a volume at `/data` for persistence
5. Railway auto-detects `Procfile` and runs gunicorn

## Connecting a Bubble frontend

See **[BUBBLE_INTEGRATION.md](BUBBLE_INTEGRATION.md)** for the API workflow patterns. Summary:

- All `/api/*` calls require `Authorization: Bearer <API_AUTH_TOKEN>`
- Upload PDF: `POST /api/upload` (multipart) → returns `job_id`
- Stream progress: `GET /api/stream/<job_id>` (Server-Sent Events)
- Download report: `GET /api/download/<job_id>/<filename>`

## How it works

### Stages

| Stage | Module | Output |
|---|---|---|
| 1. Page detection | `modules/page_detector.py` | `<name>_BS_PL.pdf` (BS+PL pages only) |
| 2. Table extraction (rule-based) | `modules/extract_tables.py` | `<name>_extracted.xlsx` |
| 2.5. Pre-check + Claude PDF fallback | `modules/claude_pdf_extractor.py` | replaces extracted xlsx if rule-based broke |
| 3. BS/PL mapping | `modules/claude_mapper.py` (Claude tool-use) | mapped cell values |
| 3.5. Post-check + remediation | sanity check + Claude PDF re-map if BS imbalance | balanced final mapping |
| 4. Report generation | `modules/bs_pl_mapper.py` | `<name>_extracted_Report.xlsx` |
| 5. Per-company workbook | `modules/company_excel_writer.py` | `companies/<Company>_FinancialReport.xlsx` |
| 6. Airtable upload (optional) | `modules/airtable_uploader.py` | upserts CY+PY rows, chain-links via `previous_bs&pl` |
| 7. Observations (optional) | `modules/observations.py` | per-metric AI commentary on Airtable rows |

### Key design choices

- **Standalone-only by default** — `STANDALONE_ONLY=true` (config flag). Halves LLM cost; flip to false for both Standalone + Consolidated columns.
- **Unit normalization** — every PDF is scanned for "Rs in Crore" / "INR Lakhs" / etc. and values are scaled to **Lakhs (canonical)** before mapping. Per-company workbooks always speak Lakhs.
- **Idempotent Airtable upsert** — re-uploading the same year UPDATES the existing row; never creates duplicates. Chain links (`previous_bs&pl`) self-heal in any upload order.
- **Self-improving** — `POST /api/corrections` adds high-confidence examples to ChromaDB; future mappings retrieve these as context.
- **Eval-driven** — every prompt change is regression-tested via `python eval/run.py --mode claude` against `eval/baseline.json`.

## API reference (production)

All routes require `Authorization: Bearer <API_AUTH_TOKEN>` when the token is configured. `/api/health` is exempt.

### Public

| Method | Path | Purpose |
|---|---|---|
| `GET` | `/api/health` | Liveness probe (unauth'd) |
| `GET` | `/health` | Detailed status (matcher loaded, keys configured, mapper mode) |

### Job orchestration

| Method | Path | Purpose |
|---|---|---|
| `POST` | `/api/upload` | Upload one or more PDFs. Form fields: `files`, `api_key` (optional override), `skip_stage1`, `fy_override`, `gen_observations`. Returns `{job_id, files, ...}` |
| `GET` | `/api/stream/<job_id>` | Server-Sent Events stream of job progress + log lines |
| `GET` | `/api/status/<job_id>` | Snapshot status + output filenames |
| `GET` | `/api/download/<job_id>/<filename>` | Download a single output file |
| `GET` | `/api/download_all/<job_id>` | Download all outputs as a zip |

### Configuration

| Method | Path | Purpose |
|---|---|---|
| `GET` | `/api/config` | Current config (secrets redacted) |
| `POST` | `/api/config` | Update config keys |

### RAG / corrections

| Method | Path | Purpose |
|---|---|---|
| `GET` | `/api/rag/stats` | Vector DB stats |
| `POST` | `/api/corrections` | Submit a reviewer-validated label→field mapping |

### Airtable + observations

| Method | Path | Purpose |
|---|---|---|
| `POST` | `/api/airtable/upload/<job_id>` | Push a job's Report.xlsx to Airtable |
| `GET` | `/api/companies` | List all companies in the Airtable base |
| `POST` | `/api/observations/run` | Trigger the bs&pl observations pipeline |
| `GET` | `/api/report/<company_id>` | HTML diagnostic report for a company |
| `GET` | `/api/report/by-name/<company_name>` | Same, looked up by name |

## Configuration

### Resolution order (last wins)

1. Hard-coded defaults in `config.py`
2. `config.json` (gitignored — local dev only)
3. **Environment variables** (production)

### Required

| Env var | What |
|---|---|
| `ANTHROPIC_API_KEY` | Required. Claude API key. Get one at https://console.anthropic.com |
| `API_AUTH_TOKEN` | Required for production. Long random string. Bubble must send `Authorization: Bearer <this>`. |

### Optional

See `.env.example` for the full list with descriptions.

## Eval / regression testing

Drop a `<company>__extracted.xlsx` + `<company>__truth.json` pair into `eval/ground_truth/`, then:

```bash
# Lock current accuracy as the baseline
python eval/run.py --mode claude --update-baseline

# After any prompt or mapper change
python eval/run.py --mode claude
# → "vs baseline (claude): acc +X.XXpp | balance +X.XXpp"
```

A negative `acc` means regression — investigate before merging. See `eval/README.md` for the truth-JSON schema.

## Project layout

```
ask-my-cfo/
├── README.md                      ← you are here
├── DEPLOY_RAILWAY.md              ← Railway deploy guide
├── BUBBLE_INTEGRATION.md          ← Bubble frontend integration
├── requirements.txt
├── runtime.txt                    ← Python version pin
├── Procfile                       ← gunicorn start command
├── railway.toml                   ← Railway service config
├── .gitignore
├── .env.example                   ← env-var template
├── config.py                      ← config loader (env > json > defaults)
├── config.example.json            ← non-secret config template
├── app.py                         ← Flask entry point
├── index.html                     ← upload + log-stream UI
├── landing.html                   ← marketing landing
├── Input Templates.xlsx           ← canonical row layout (read-only reference)
├── modules/
│   ├── page_detector.py           ← stage 1
│   ├── extract_tables.py          ← stage 2 (rule-based extractor)
│   ├── claude_pdf_extractor.py    ← stage 2.5 + post-check + YoY anomaly
│   ├── claude_mapper.py           ← stage 3 (primary Claude mapper)
│   ├── llm_mapper_v2.py           ← legacy GPT-4 mapper (mapper_mode=v2_llm)
│   ├── matcher.py                 ← rule-based + TF-IDF fallback
│   ├── rag_matcher.py             ← ChromaDB retrieval
│   ├── bs_pl_mapper.py            ← stage 3 orchestrator + report writer
│   ├── company_excel_writer.py    ← per-company multi-year workbook
│   ├── airtable_uploader.py       ← idempotent upsert + chain backfill
│   ├── observations.py            ← per-metric AI commentary
│   └── ...
└── eval/
    ├── README.md
    ├── run.py                     ← accuracy harness
    ├── ground_truth/              ← <company>__truth.json fixtures
    └── baseline.json              ← current locked accuracy
```

## License

Proprietary — internal use only.
