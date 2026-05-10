#!/usr/bin/env python3
"""
ASK MY CFO — M1 Automation
===========================
Properly chains: page_detector → extract_tables → bs_pl_mapper
using their top-level functions exactly as they work from CLI.

Flow:
  1. extract_pages(pdf)        → *_BS_PL.pdf      (only BS/PL pages)
  2. extract_tables(bs_pl_pdf) → *_extracted.xlsx  (structured rows)
  3. process_file(xlsx, ...)   → *_Report.xlsx     (mapped template)
"""

import os
import sys
import json
import uuid
import io
import time
import zipfile
import threading
import traceback
import contextlib
from pathlib import Path
from datetime import datetime
from queue import Queue, Empty

# Add flask_app root to path so we can import config.py
sys.path.insert(0, str(Path(__file__).parent))
from config import (
    load_config, save_config,
    UPLOAD_DIR, OUTPUT_DIR,  # DATA_DIR-rooted in production via env
)

# Add modules to path
MODULES_DIR = os.path.join(os.path.dirname(__file__), "modules")
sys.path.insert(0, MODULES_DIR)

from flask import (
    Flask, render_template, request, jsonify,
    send_file, Response, stream_with_context
)
from werkzeug.utils import secure_filename

app = Flask(__name__, template_folder=".")
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500MB

BASE_DIR = Path(__file__).parent

# Create dirs at import time (not just __main__) so `flask run`, gunicorn,
# or a PyCharm run-config that imports app.py all work without "path does
# not exist" 500s on first upload.
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CORS + Bearer-token auth
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CORS lets the Bubble frontend (different origin) call /api/* from the
# browser. The allowed origins come from the CORS_ORIGINS env var
# (comma-separated). Default '*' is fine for local dev / when the API is
# reverse-proxied behind the same domain as the UI.
try:
    from flask_cors import CORS
    _cors_origins_raw = (load_config().get("cors_origins") or "*").strip()
    _origins = [o.strip() for o in _cors_origins_raw.split(",") if o.strip()]
    if _origins == ["*"]:
        _origins = "*"
    CORS(app, resources={r"/api/*": {"origins": _origins}}, supports_credentials=False)
except ImportError:
    # flask-cors not installed — local dev usually fine without it.
    print("[startup] flask-cors not installed; CORS disabled. "
          "Install for cross-origin support: pip install flask-cors")


# Bearer-token auth on /api/* routes. When API_AUTH_TOKEN is set
# (env or config.json), every /api/* request must include
#   Authorization: Bearer <token>
# Reads from `request.headers["Authorization"]` or the `?token=` query
# param (the latter is for SSE streams which can't easily set custom headers
# from EventSource in the browser). Routes are otherwise unchanged.
def _api_auth_required(view_func):
    """Decorator-like wrapper used inside before_request below."""
    return view_func  # placeholder; the actual check happens in the hook


@app.before_request
def _check_api_auth():
    # Public routes pass through.
    p = request.path or ""
    if not p.startswith("/api/"):
        return None
    if request.method == "OPTIONS":  # CORS preflight
        return None
    if p in ("/api/health",):  # explicit allowlist of unauth'd endpoints
        return None

    expected = (load_config().get("api_auth_token") or "").strip()
    if not expected:
        return None  # auth disabled (no token configured)

    # Header takes precedence; fall back to ?token= for SSE clients.
    auth = (request.headers.get("Authorization") or "").strip()
    presented = ""
    if auth.lower().startswith("bearer "):
        presented = auth[7:].strip()
    if not presented:
        presented = (request.args.get("token") or "").strip()
    if presented != expected:
        return jsonify({"error": "unauthorized — missing or wrong "
                                  "Bearer token"}), 401
    return None


# ── JSON error handlers (prevent HTML error pages) ──
@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "File too large. Max 500MB."}), 413

@app.errorhandler(500)
def server_error(e):
    # Print to stderr (visible in terminal) AND include the last 20 lines
    # of the traceback in the JSON response (visible in DevTools) — silent
    # 500s are very hard to debug otherwise.
    traceback.print_exc()
    return jsonify({
        "error": f"Server error: {e}",
        "traceback": traceback.format_exc().splitlines()[-20:],
    }), 500

@app.errorhandler(404)
def not_found(e):
    return jsonify({"error": "Not found"}), 404


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# JOB TRACKER
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

jobs = {}


class PipelineJob:
    def __init__(self, job_id, pdf_paths, api_key, skip_stage1, skip_stage3,
                 fy_override=None, gen_observations=None):
        self.job_id = job_id
        self.pdf_paths = pdf_paths
        self.api_key = api_key
        self.skip_stage1 = skip_stage1
        self.skip_stage3 = skip_stage3
        self.fy_override = fy_override   # None => auto-detect from filename
        # None = use config.auto_observations; True/False = force per-job
        self.gen_observations = gen_observations
        self.status = "queued"
        self.stage = 0
        self.stage_label = ""
        self.progress = 0
        self.logs = []
        self.output_files = []
        self.error = None
        self.queue = Queue()

    def log(self, msg, level="info"):
        entry = {"time": datetime.now().strftime("%H:%M:%S"), "level": level, "message": msg}
        self.logs.append(entry)
        self._push({"type": "log", "data": entry, "stage": self.stage,
                     "stage_label": self.stage_label, "progress": self.progress,
                     "status": self.status})

    def set_stage(self, stage, label, progress=None):
        self.stage = stage
        self.stage_label = label
        if progress is not None:
            self.progress = progress
        self._push({"type": "stage", "stage": stage, "stage_label": label,
                     "progress": self.progress, "status": self.status})

    def finish(self):
        self.status = "completed"
        self.progress = 100
        self._finished_at = time.time()
        self._push({"type": "complete", "status": "completed", "progress": 100,
                     "files": [os.path.basename(f) for f in self.output_files]})

    def fail(self, error):
        self.status = "failed"
        self.error = str(error)
        self._finished_at = time.time()
        self._push({"type": "error", "error": str(error), "status": "failed"})

    def _push(self, data):
        self.queue.put(json.dumps(data))


# ── Job GC: drop completed/failed jobs from memory after TTL ────────────────
# Without this the `jobs` dict grows unbounded — each PipelineJob keeps log
# lines + queue + output paths in memory forever.
_JOB_TTL_SEC = 60 * 60  # 1 hour after completion


def _gc_jobs_loop():
    while True:
        time.sleep(300)  # every 5 min
        try:
            now = time.time()
            stale = [jid for jid, j in list(jobs.items())
                     if j.status in ("completed", "failed")
                     and getattr(j, "_finished_at", now) and (now - j._finished_at) > _JOB_TTL_SEC]
            for jid in stale:
                jobs.pop(jid, None)
        except Exception:
            pass  # never crash the GC thread


threading.Thread(target=_gc_jobs_loop, daemon=True).start()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# STDOUT CAPTURE — route module prints → job log
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

class LogCapture(io.TextIOBase):
    def __init__(self, job, level="info"):
        self.job = job
        self.level = level
        self.buffer = ""

    def write(self, s):
        if not s:
            return 0
        self.buffer += s
        while "\n" in self.buffer:
            line, self.buffer = self.buffer.split("\n", 1)
            line = line.rstrip()
            if line:
                lvl = self.level
                if any(c in line for c in ["✓", "✅", "succeeded", "balanced"]):
                    lvl = "success"
                elif any(c in line for c in ["✗", "ERROR", "FAILED", "failed!"]):
                    lvl = "error"
                elif any(c in line for c in ["⚠", "WARNING", "retrying", "FLAGGED"]):
                    lvl = "warning"
                elif "━" in line or "──" in line or "==" in line:
                    lvl = "header"
                self.job.log(line, lvl)
        return len(s)

    def flush(self):
        if self.buffer.strip():
            self.job.log(self.buffer.strip(), self.level)
            self.buffer = ""


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# PIPELINE — uses EXACT same top-level functions as CLI
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def run_pipeline(job: PipelineJob):
    out_dir = OUTPUT_DIR / job.job_id
    out_dir.mkdir(parents=True, exist_ok=True)

    # Per-thread stdout/stderr capture via contextlib — does NOT mutate
    # global sys.stdout, so concurrent uploads don't bleed logs into each
    # other's SSE streams.
    cap_out = LogCapture(job, "info")
    cap_err = LogCapture(job, "error")

    try:
        job.status = "running"
        total = len(job.pdf_paths)

        for fi, pdf_path in enumerate(job.pdf_paths):
            stem = Path(pdf_path).stem
            base_pct = int(fi / total * 100)
            span = int(100 / total)

            job.log(f"━━━ File {fi+1}/{total}: {Path(pdf_path).name} ━━━", "header")

            # ═══════════════════════════════════════
            # STAGE 1: extract_pages() → clean PDF
            # ═══════════════════════════════════════
            working_pdf = pdf_path

            if job.skip_stage1:
                job.set_stage(1, "Skipped — using PDF as-is", base_pct + 2)
                job.log("Stage 1 skipped (PDF assumed to already contain only BS/PL)")
            else:
                job.set_stage(1, f"Detecting pages — {stem}", base_pct + 2)
                job.log("Running page_detector.extract_pages()...")

                try:
                    from page_detector import extract_pages

                    with contextlib.redirect_stdout(cap_out), \
                         contextlib.redirect_stderr(cap_err):
                        bs_pl_pdf = extract_pages(pdf_path, str(out_dir))
                    cap_out.flush()
                    cap_err.flush()

                    if bs_pl_pdf and os.path.exists(bs_pl_pdf):
                        working_pdf = bs_pl_pdf
                        job.log(f"✓ Created {os.path.basename(bs_pl_pdf)}", "success")
                        job.output_files.append(bs_pl_pdf)
                    else:
                        job.log("⚠ No sections found — falling back to full PDF", "warning")

                except Exception as e:
                    job.log(f"⚠ Stage 1 error: {e}", "warning")
                    job.log("Falling back to full PDF for extraction", "warning")

            # ═══════════════════════════════════════
            # STAGE 2: extract_tables() → Excel
            # ═══════════════════════════════════════
            job.set_stage(2, f"Extracting tables — {stem}", base_pct + int(span * 0.25))
            job.log(f"Running extract_tables.extract_tables({os.path.basename(working_pdf)})...")

            extracted_xlsx = None
            try:
                from extract_tables import extract_tables as run_extract

                with contextlib.redirect_stdout(cap_out), \
                     contextlib.redirect_stderr(cap_err):
                    run_extract(working_pdf, str(out_dir))
                cap_out.flush()
                cap_err.flush()

                # Find output — extract_tables saves as {stem}_extracted.xlsx
                expected = out_dir / f"{Path(working_pdf).stem}_extracted.xlsx"
                if expected.exists():
                    extracted_xlsx = str(expected)
                else:
                    # Fallback: find any new _extracted.xlsx
                    for f in sorted(out_dir.glob("*_extracted.xlsx"), key=os.path.getmtime, reverse=True):
                        extracted_xlsx = str(f)
                        break

                if extracted_xlsx:
                    job.log(f"✓ Created {os.path.basename(extracted_xlsx)}", "success")
                    job.output_files.append(extracted_xlsx)
                else:
                    job.log("✗ No extracted Excel was generated!", "error")
                    continue

            except Exception as e:
                job.log(f"✗ Stage 2 failed: {e}", "error")
                job.log(traceback.format_exc(), "error")
                continue

            # ═══════════════════════════════════════
            # STAGE 2.5: sanity-check extracted xlsx; on failure, hand the
            # PDF to Claude for native-PDF extraction. The downstream mapper
            # consumes the same xlsx shape either way.
            # ═══════════════════════════════════════
            claude_used = False  # tracks whether we've already burned a
                                  # Claude PDF call, so Stage 3.5 doesn't loop
            try:
                from claude_pdf_extractor import (
                    is_extraction_broken, extract_with_claude,
                )
                broken, reason = is_extraction_broken(extracted_xlsx)
            except Exception as e:
                broken, reason = False, f"sanity check skipped: {e}"
                job.log(f"⚠ Sanity check error (non-fatal): {e}", "warning")

            if broken:
                job.log(f"⚠ Extraction looks broken: {reason}", "warning")
                cfg = load_config()
                anth_key = (job.api_key
                            if (job.api_key or "").startswith("sk-ant-")
                            else (cfg.get("anthropic_api_key") or "").strip())
                if anth_key:
                    job.log("Falling back to Claude native PDF extraction...",
                             "warning")
                    try:
                        with contextlib.redirect_stdout(cap_out), \
                             contextlib.redirect_stderr(cap_err):
                            claude_xlsx = extract_with_claude(
                                working_pdf, str(out_dir),
                                api_key=anth_key,
                                model=cfg.get("claude_model"),
                            )
                        cap_out.flush()
                        cap_err.flush()
                        if claude_xlsx and os.path.exists(claude_xlsx):
                            extracted_xlsx = claude_xlsx
                            claude_used = True
                            if claude_xlsx not in job.output_files:
                                job.output_files.append(claude_xlsx)
                            job.log(f"✓ Claude extraction succeeded: "
                                     f"{os.path.basename(claude_xlsx)}",
                                     "success")
                        else:
                            job.log("⚠ Claude fallback returned no output — "
                                     "continuing with rule-based xlsx",
                                     "warning")
                    except Exception as e:
                        job.log(f"⚠ Claude fallback raised: {e}", "warning")
                        job.log(traceback.format_exc(), "warning")
                else:
                    job.log("⚠ No anthropic_api_key configured — cannot run "
                             "Claude fallback. Mapper will run on the broken "
                             "extract.", "warning")

            # ═══════════════════════════════════════
            # STAGE 3: process_file() → Report
            # ═══════════════════════════════════════
            job.set_stage(3, f"Mapping to template — {stem}", base_pct + int(span * 0.50))
            job.log(f"Running bs_pl_mapper.process_file({os.path.basename(extracted_xlsx)})...")

            report_path = None
            try:
                from bs_pl_mapper import process_file

                with contextlib.redirect_stdout(cap_out), \
                     contextlib.redirect_stderr(cap_err):
                    report_path = process_file(
                        input_file=extracted_xlsx,
                        output_dir=str(out_dir),
                        fy_override=job.fy_override,
                        source_pdf_path=working_pdf,
                    )
                cap_out.flush()
                cap_err.flush()

                if report_path and os.path.exists(report_path):
                    job.log(f"✓ Created {os.path.basename(report_path)}", "success")
                    job.output_files.append(report_path)
                else:
                    job.log("⚠ process_file returned no output", "warning")

            except Exception as e:
                job.log(f"✗ Stage 3 failed: {e}", "error")
                job.log(traceback.format_exc(), "error")

            # ═══════════════════════════════════════
            # STAGE 3.5: post-mapper sanity check — runs for EVERY job.
            # Catches:
            #   - mostly-zero report (rare API edge / schema mismatch)
            #   - BS imbalance > 1% of Total Assets (Finolex-style: a row
            #     missed by the extractor that the mapper couldn't recover)
            #   - all-variants-empty report
            # When broken AND we haven't already burned a Claude PDF call
            # in Stage 2.5, do it now and re-map once. If Claude was
            # already used and the result is STILL broken, surface a
            # strong warning rather than looping (re-extracting from the
            # same PDF would just yield the same data).
            # ═══════════════════════════════════════
            if report_path and os.path.exists(report_path):
                try:
                    from claude_pdf_extractor import (
                        is_report_broken, extract_with_claude,
                        check_yoy_anomalies, append_warnings_to_report,
                    )
                    rep_broken, rep_reason = is_report_broken(report_path)
                except Exception as e:
                    rep_broken, rep_reason = False, f"post-check skipped: {e}"

                if rep_broken and claude_used:
                    # Already tried Claude PDF in Stage 2.5; re-running it
                    # would reuse the same source bytes and produce the
                    # same result. Surface the gap loudly and move on so
                    # the user (or downstream Airtable WARNINGS) sees it.
                    job.log(f"⚠ Report still broken AFTER Claude fallback: "
                             f"{rep_reason}", "error")
                    job.log("⚠ Manual review required — see WARNINGS sheet "
                             "in the Report xlsx.", "warning")
                elif rep_broken:
                    job.log(f"⚠ Mapper output looks broken: {rep_reason}",
                             "warning")
                    cfg = load_config()
                    anth_key = (job.api_key
                                if (job.api_key or "").startswith("sk-ant-")
                                else (cfg.get("anthropic_api_key") or "").strip())
                    if anth_key:
                        job.log("Retrying via Claude native PDF extraction "
                                 "+ remap...", "warning")
                        try:
                            with contextlib.redirect_stdout(cap_out), \
                                 contextlib.redirect_stderr(cap_err):
                                claude_xlsx = extract_with_claude(
                                    working_pdf, str(out_dir),
                                    api_key=anth_key,
                                    model=cfg.get("claude_model"),
                                )
                            cap_out.flush()
                            cap_err.flush()

                            if claude_xlsx and os.path.exists(claude_xlsx):
                                extracted_xlsx = claude_xlsx
                                claude_used = True
                                if claude_xlsx not in job.output_files:
                                    job.output_files.append(claude_xlsx)
                                with contextlib.redirect_stdout(cap_out), \
                                     contextlib.redirect_stderr(cap_err):
                                    new_report = process_file(
                                        input_file=extracted_xlsx,
                                        output_dir=str(out_dir),
                                        fy_override=job.fy_override,
                                        source_pdf_path=working_pdf,
                                    )
                                cap_out.flush()
                                cap_err.flush()
                                if new_report and os.path.exists(new_report):
                                    report_path = new_report
                                    if new_report not in job.output_files:
                                        job.output_files.append(new_report)
                                    job.log(f"✓ Recovered via Claude: "
                                             f"{os.path.basename(new_report)}",
                                             "success")
                                    # Re-check the recovered report; if it's
                                    # still broken, log it but don't loop.
                                    rep_broken2, rep_reason2 = is_report_broken(report_path)
                                    if rep_broken2:
                                        job.log(f"⚠ Recovered report still "
                                                 f"broken: {rep_reason2}",
                                                 "error")
                                else:
                                    job.log("⚠ Mapper retry produced no "
                                             "output", "warning")
                            else:
                                job.log("⚠ Claude fallback returned no xlsx",
                                         "warning")
                        except Exception as e:
                            job.log(f"⚠ Recovery raised: {e}", "warning")
                            job.log(traceback.format_exc(), "warning")
                    else:
                        job.log("⚠ No anthropic_api_key — cannot recover. "
                                 "Report flagged as broken in WARNINGS sheet.",
                                 "warning")

                # ── YoY anomaly check (always runs on the final report) ──
                # Catches decimal-point bugs, wrong-unit uploads, mapper
                # cell-swap mistakes — anything that produces a
                # statistically implausible year-over-year swing on a
                # known-stable line. Findings are appended to WARNINGS;
                # they don't block the upload, just flag for review.
                try:
                    if report_path and os.path.exists(report_path):
                        anomalies = check_yoy_anomalies(report_path)
                        if anomalies:
                            crit = sum(1 for s, _ in anomalies if s == "critical")
                            warn = len(anomalies) - crit
                            job.log(f"⚠ YoY anomaly check: {crit} critical, "
                                     f"{warn} warn — see WARNINGS sheet",
                                     "warning" if not crit else "error")
                            for sev, msg in anomalies[:5]:
                                tag = "‼" if sev == "critical" else "⚠"
                                job.log(f"   {tag} {msg}",
                                         "error" if sev == "critical" else "warning")
                            append_warnings_to_report(
                                report_path,
                                [("YoY", sev.upper(), msg)
                                 for sev, msg in anomalies],
                            )
                except Exception as e:
                    job.log(f"⚠ YoY anomaly check failed (non-fatal): {e}",
                             "warning")

        # ═══════════════════════════════════════
        # POST-PIPELINE: auto-upload + auto-observations
        # ═══════════════════════════════════════
        # Once a *_Report.xlsx is produced, push it to Airtable and trigger
        # observations for that company so the diagnostic report is ready
        # to view immediately. Both gated by config (default: ON for obs,
        # OFF for upload to avoid surprise writes).
        cfg = load_config()
        do_upload = bool(cfg.get("airtable_auto_upload", False))
        # Per-job observations toggle wins over the config default. None
        # means "no per-job preference" — fall back to config.
        if job.gen_observations is None:
            do_obs = bool(cfg.get("auto_observations", True))
        else:
            do_obs = bool(job.gen_observations)
            job.log(f"ℹ Observations toggle (per-job) = {do_obs}", "info")
        # company_id -> set of bs&pl record ids that were just uploaded.
        # Auto-obs runs only on these specific records — not on the
        # company's full history — to avoid 7×30s loops on every re-upload.
        company_to_bspl_ids: dict[str, set[str]] = {}

        if do_upload:
            reports = [f for f in job.output_files if f.endswith("_Report.xlsx")]
            if reports:
                job.log(f"━━━ Auto-upload to Airtable: {len(reports)} report(s) ━━━", "header")
                try:
                    from airtable_uploader import upload_report
                    for rp in reports:
                        try:
                            res = upload_report(rp,
                                                fy_override=job.fy_override,
                                                log=lambda m: job.log(m, "info"))
                            if res.get("ok"):
                                py_act = res.get("py_action", "?")
                                cy_act = res.get("cy_action", "?")
                                job.log(f"✓ Airtable {os.path.basename(rp)} -> "
                                        f"company={res.get('company')}, "
                                        f"PY {py_act}={res.get('prior_id')}, "
                                        f"CY {cy_act}={res.get('current_id')}",
                                        "success")
                                cid = res.get("company_record_id")
                                if cid:
                                    s = company_to_bspl_ids.setdefault(cid, set())
                                    if res.get("current_id"): s.add(res["current_id"])
                                    if res.get("prior_id"):   s.add(res["prior_id"])
                            else:
                                job.log(f"⚠ Airtable upload failed for {os.path.basename(rp)}: "
                                        f"{res.get('error')}", "warning")
                        except Exception as e:
                            job.log(f"⚠ Airtable upload exception for "
                                    f"{os.path.basename(rp)}: {e}", "warning")
                except ImportError as e:
                    job.log(f"⚠ airtable_uploader unavailable: {e}", "warning")

        if do_obs and company_to_bspl_ids:
            total_rows = sum(len(s) for s in company_to_bspl_ids.values())
            job.log(f"━━━ Auto-generating observations for "
                    f"{total_rows} just-uploaded bs&pl row(s) "
                    f"({len(company_to_bspl_ids)} company) ━━━", "header")
            try:
                from observations import run_pipeline as run_obs
                # Stream the per-metric prints from observations.py straight
                # into the dashboard SSE — no buffering — so the user sees
                # OBS / RECOM / FUTURE / ACK lines as they're generated.
                # bspl_ids filter scopes processing to ONLY the new rows so
                # historical observations aren't regenerated on every upload.
                for cid, ids in company_to_bspl_ids.items():
                    try:
                        with contextlib.redirect_stdout(cap_out), \
                             contextlib.redirect_stderr(cap_err):
                            out = run_obs(company_id=cid,
                                          bspl_ids=list(ids),
                                          dry_run=False, force=False,
                                          capture_log=False)
                        cap_out.flush()
                        cap_err.flush()
                        c = out.get("counts", {})
                        job.log(f"✓ Observations for company {cid}: "
                                f"created={c.get('created',0)} "
                                f"updated={c.get('updated',0)} "
                                f"skipped={c.get('skipped',0)} "
                                f"error={c.get('error',0)}", "success")
                    except Exception as e:
                        job.log(f"⚠ Observations failed for {cid}: {e}", "warning")
            except ImportError as e:
                job.log(f"⚠ observations module unavailable: {e}", "warning")
        elif do_obs and not do_upload:
            job.log("ℹ Auto-observations skipped — enable airtable_auto_upload "
                    "in config.json to trigger them automatically after each run.",
                    "info")

        # All files done
        if job.output_files:
            job.log(f"━━━ Done — {len(job.output_files)} file(s) generated ━━━", "header")
            job.finish()
        else:
            job.fail("No output files were generated.")

    except Exception as e:
        job.log(f"Fatal: {e}", "error")
        job.log(traceback.format_exc(), "error")
        job.fail(str(e))


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ROUTES
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/landing")
def landing():
    """Marketing landing page — links into /dashboard."""
    return render_template("landing.html")


# Unauth'd lightweight liveness probe — Railway / load balancers hit this.
# Stays separate from /health (which requires loading the matcher and is
# more expensive); /api/health is intentionally cheap.
@app.route("/api/health")
def api_health():
    return jsonify({"status": "ok"}), 200


@app.route("/health")
def health():
    try:
        from matcher import get_matcher
        m = get_matcher()
        has_matcher = m.pl_tfidf is not None
    except Exception:
        has_matcher = False
    cfg = load_config()
    anthropic_set = bool((cfg.get("anthropic_api_key") or "").strip())
    openai_set    = bool((cfg.get("openai_api_key") or "").strip())
    airtable_set  = bool((cfg.get("airtable_api_token") or "").strip()
                         and (cfg.get("airtable_base_id") or "").strip())
    auto_upload   = bool(cfg.get("airtable_auto_upload", False))
    auto_obs      = bool(cfg.get("auto_observations", True))
    # Mode summary for the dashboard badge:
    #   "ai_full"     LLM + airtable both wired (full AI + storage)
    #   "ai_only"     LLM only — mapping uses Claude/GPT, no Airtable storage
    #   "offline"     no LLM key — pipeline runs offline (rules+TF-IDF only)
    llm_set = anthropic_set or openai_set
    if llm_set and airtable_set:
        mode = "ai_full"
    elif llm_set:
        mode = "ai_only"
    else:
        mode = "offline"
    return jsonify({
        "status":                   "ok",
        "matcher":                  has_matcher,
        "anthropic_key_configured": anthropic_set,
        "api_key_configured":       openai_set,           # legacy field name
        "openai_key_configured":    openai_set,
        "airtable_configured":      airtable_set,
        "auto_upload":              auto_upload,
        "auto_observations":        auto_obs,
        "mapper_mode":              cfg.get("mapper_mode", "claude"),
        "claude_model":             cfg.get("claude_model", "claude-sonnet-4-6"),
        "mode":                     mode,
    })


@app.route("/api/config", methods=["GET", "POST"])
def api_config():
    """Read or update persistent config (config.json)."""
    if request.method == "GET":
        cfg = load_config()
        # Redact API keys in GET responses — return whether each is set, not the value
        return jsonify({
            "anthropic_api_key_set":     bool(cfg.get("anthropic_api_key")),
            "anthropic_api_key_preview": (cfg.get("anthropic_api_key", "")[:11] + "…")
                                          if cfg.get("anthropic_api_key") else "",
            "claude_model":              cfg.get("claude_model", "claude-sonnet-4-6"),
            "openai_api_key_set":        bool(cfg.get("openai_api_key")),
            "openai_api_key_preview":    (cfg.get("openai_api_key", "")[:7] + "…")
                                          if cfg.get("openai_api_key") else "",
            "mapper_mode":               cfg.get("mapper_mode", "claude"),
            "use_gpt_fallback":          cfg.get("use_gpt_fallback", True),
            "balance_tolerance":         cfg.get("balance_tolerance", 1.0),
            "standalone_only":           cfg.get("standalone_only", True),
        })
    # POST: update keys
    try:
        data = request.get_json(force=True, silent=True) or {}
    except Exception:
        data = {}
    updates = {}
    if "anthropic_api_key" in data:
        key = str(data["anthropic_api_key"]).strip()
        # Sanity: empty (clear) or starts with 'sk-ant-'
        if key and not key.startswith("sk-ant-"):
            return jsonify({"error": "Anthropic keys start with 'sk-ant-'. Paste the full key."}), 400
        updates["anthropic_api_key"] = key
    if "claude_model" in data:
        updates["claude_model"] = str(data["claude_model"]).strip()
    if "openai_api_key" in data:
        key = str(data["openai_api_key"]).strip()
        if key and not key.startswith("sk-"):
            return jsonify({"error": "OpenAI keys start with 'sk-'. Paste the full key."}), 400
        updates["openai_api_key"] = key
    if "mapper_mode" in data:
        m = str(data["mapper_mode"]).strip()
        if m not in ("claude", "v2_llm", "hybrid", "rules_only"):
            return jsonify({"error": "mapper_mode must be one of "
                                       "claude|v2_llm|hybrid|rules_only"}), 400
        updates["mapper_mode"] = m
    if "use_gpt_fallback" in data:
        updates["use_gpt_fallback"] = bool(data["use_gpt_fallback"])
    if "balance_tolerance" in data:
        try:
            updates["balance_tolerance"] = float(data["balance_tolerance"])
        except (TypeError, ValueError):
            return jsonify({"error": "balance_tolerance must be a number"}), 400
    if "standalone_only" in data:
        updates["standalone_only"] = bool(data["standalone_only"])
    if not updates:
        return jsonify({"error": "no recognized fields to update"}), 400
    save_config(updates)
    return jsonify({"status": "saved", "fields": list(updates.keys())})


@app.route("/api/rag/stats", methods=["GET"])
def api_rag_stats():
    """Return vector DB stats: how many BS/PL entries, which models, threshold."""
    try:
        from rag_matcher import stats as rag_stats
        return jsonify(rag_stats())
    except Exception as e:
        return jsonify({"error": f"RAG not available: {e}"}), 503


@app.route("/api/corrections", methods=["POST"])
def api_corrections():
    """Human reviewer confirms the correct mapping for a label.
    Payload: {"label": "...", "field": "Trade Payables",
              "side": "BS"|"PL", "section": "current_liab" (optional),
              "company": "Colgate" (optional)}
    The correction is added to ChromaDB with confidence=1.0 so it outweighs
    seed data in future retrieval."""
    try:
        from rag_matcher import add_correction
    except Exception as e:
        return jsonify({"error": f"RAG not available: {e}"}), 503

    data = request.get_json(force=True, silent=True) or {}
    label = str(data.get("label", "")).strip()
    field = str(data.get("field", "")).strip()
    side = str(data.get("side", "BS")).upper()
    section = str(data.get("section", "")).strip()
    company = str(data.get("company", "")).strip()

    if not label or not field:
        return jsonify({"error": "label and field are required"}), 400
    if side not in ("BS", "PL"):
        return jsonify({"error": "side must be 'BS' or 'PL'"}), 400

    ok = add_correction(label, field, side=side, section=section,
                        source_company=company)
    if not ok:
        return jsonify({"error": "failed to store correction"}), 500
    return jsonify({"status": "stored", "label": label, "field": field})


@app.route("/api/upload", methods=["POST"])
def upload():
    try:
        if "files" not in request.files:
            return jsonify({"error": "No files uploaded"}), 400

        files = request.files.getlist("files")
        if not files or not any(f.filename for f in files):
            return jsonify({"error": "No files selected"}), 400

        api_key = request.form.get("api_key", "").strip()
        # Fall back to config.json when form field is empty. Prefer the
        # Claude key (primary mapper); only fall back to the OpenAI key when
        # Claude isn't configured.
        if not api_key:
            cfg = load_config()
            api_key = (cfg.get("anthropic_api_key", "").strip()
                       or cfg.get("openai_api_key", "").strip())
        skip_stage1 = request.form.get("skip_stage1") == "true"
        skip_stage3 = False

        # FY override (optional): user picks the financial year on upload.
        # Empty string => auto-detect from filename (existing behavior).
        # Format: "FY2425" or an ISO end-date like "2025-03-31".
        fy_override = (request.form.get("fy_override") or "").strip() or None

        # Per-job observations toggle. Form may send "true"/"false" or be
        # absent (None = use config.auto_observations). Anything other than
        # the literal string "false" is treated as ON, matching the
        # default-on UI checkbox.
        raw_obs = request.form.get("gen_observations")
        if raw_obs is None:
            gen_observations = None
        else:
            gen_observations = raw_obs.lower() != "false"

        job_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + str(uuid.uuid4())[:6]
        job_dir = UPLOAD_DIR / job_id
        job_dir.mkdir(parents=True, exist_ok=True)

        pdf_paths = []
        for f in files:
            if f.filename:
                fname = secure_filename(f.filename)
                fpath = str(job_dir / fname)
                f.save(fpath)
                pdf_paths.append(fpath)

        if not pdf_paths:
            return jsonify({"error": "No valid files uploaded"}), 400

        job = PipelineJob(job_id, pdf_paths, api_key, skip_stage1, skip_stage3,
                          fy_override=fy_override,
                          gen_observations=gen_observations)
        jobs[job_id] = job

        thread = threading.Thread(target=run_pipeline, args=(job,), daemon=True)
        thread.start()

        return jsonify({
            "job_id": job_id,
            "files": [os.path.basename(p) for p in pdf_paths],
            "skip_stage1": skip_stage1,
            "skip_stage3": skip_stage3,
            "fy_override": fy_override,
            "gen_observations": gen_observations,
        })
    except Exception as e:
        return jsonify({"error": f"Upload error: {str(e)}"}), 500


@app.route("/api/stream/<job_id>")
def stream(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404

    def generate():
        while True:
            try:
                msg = job.queue.get(timeout=60)
                yield f"data: {msg}\n\n"
                data = json.loads(msg)
                if data.get("type") in ("complete", "error"):
                    break
            except Empty:
                yield f"data: {json.dumps({'type': 'ping'})}\n\n"
                if job.status in ("completed", "failed"):
                    break

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/api/download/<job_id>/<path:filename>")
def download(job_id, filename):
    fpath = OUTPUT_DIR / job_id / filename
    if fpath.exists():
        return send_file(str(fpath), as_attachment=True)
    return jsonify({"error": "File not found"}), 404


@app.route("/api/companies", methods=["GET"])
def api_companies():
    """List companies from the Airtable company table for the diagnostic-report
    UI. Returns minimal {id, name} pairs sorted by name. Cached in-memory for
    60s to avoid hammering Airtable on every dropdown render."""
    import time as _time
    cache = api_companies._cache  # type: ignore[attr-defined]
    now = _time.time()
    if cache and (now - cache["at"]) < 60:
        return jsonify(cache["data"])
    try:
        from observations import at_list_all, TBL_COMPANY, _creds
    except ImportError as e:
        return jsonify({"error": f"observations module unavailable: {e}"}), 503
    try:
        token, _, base_id = _creds()
        if not token:
            return jsonify({"error": "airtable_api_token missing in config.json",
                            "companies": []}), 400
        recs = at_list_all(TBL_COMPANY, token, base_id)
        out = []
        for r in recs:
            name = (r.get("fields", {}).get("company_name") or "").strip()
            if name:
                out.append({"id": r["id"], "name": name})
        out.sort(key=lambda x: x["name"].lower())
        payload = {"companies": out}
        api_companies._cache = {"at": now, "data": payload}  # type: ignore[attr-defined]
        return jsonify(payload)
    except Exception as e:
        return jsonify({"error": f"{type(e).__name__}: {e}",
                        "companies": []}), 500
api_companies._cache = None  # type: ignore[attr-defined]


@app.route("/api/observations/run", methods=["POST"])
def api_observations_run():
    """Trigger the bs&pl observations pipeline for a company.

    JSON body (all optional):
      {"company_id": "recXXX",   // Airtable id; pipeline processes this company only
       "user_id":    "recYYY",   // alt. filter — bs&pl rows for this user
       "limit":      5,           // cap number of bs&pl rows processed
       "force":      false,       // regenerate everything
       "dry_run":    true,        // build payloads but don't write
       "model":      "gpt-4o-mini"}
    """
    body = request.get_json(silent=True) or {}
    try:
        from observations import run_pipeline
    except ImportError as e:
        return jsonify({"error": f"observations module unavailable: {e}"}), 503
    try:
        result = run_pipeline(
            user_id=body.get("user_id", ""),
            company_id=body.get("company_id", ""),
            limit=int(body.get("limit", 0) or 0),
            force=bool(body.get("force", False)),
            dry_run=bool(body.get("dry_run", False)),
            model=body.get("model"),
            capture_log=True,
        )
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": f"{type(e).__name__}: {e}"}), 500


@app.route("/api/report/<company_id>", methods=["GET"])
def api_report_html(company_id):
    """HTML financial diagnostic report for a company.
    Query params:
      ?fy=2024-03-31     (default: most recent)
      ?as=json           (return structured data instead of HTML)
    """
    fy = request.args.get("fy", "").strip() or None
    fmt = request.args.get("as", "").lower()
    try:
        from report_generator import build_report_for_company
    except ImportError as e:
        return jsonify({"error": f"report_generator unavailable: {e}"}), 503
    try:
        result = build_report_for_company(company_id=company_id, fy_iso=fy)
    except Exception as e:
        return jsonify({"error": f"{type(e).__name__}: {e}"}), 500
    if fmt == "json":
        # Strip dataclass benchmark before serializing
        data = json.loads(json.dumps(result["data"],
                                      default=lambda o: getattr(o, "__dict__", str(o))))
        return jsonify(data)
    return Response(result["html"], mimetype="text/html; charset=utf-8")


@app.route("/api/report/by-name/<path:company_name>", methods=["GET"])
def api_report_by_name(company_name):
    """Same as /api/report/<id> but looks up by company_name field."""
    fy = request.args.get("fy", "").strip() or None
    try:
        from report_generator import build_report_for_company
    except ImportError as e:
        return jsonify({"error": f"report_generator unavailable: {e}"}), 503
    try:
        result = build_report_for_company(company_name=company_name, fy_iso=fy)
    except Exception as e:
        return jsonify({"error": f"{type(e).__name__}: {e}"}), 500
    return Response(result["html"], mimetype="text/html; charset=utf-8")


@app.route("/api/airtable/upload/<job_id>", methods=["POST"])
def api_airtable_upload(job_id):
    """Push the job's *_Report.xlsx file(s) to Airtable.

    Optional JSON body:
      {"company": "Override Name",     // skip filename auto-detection
       "user_email": "...",            // override default
       "fy": "FY2425",                 // override FY auto-detection
       "dry_run": true}                // build payloads without POSTing
    """
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404

    out_dir = OUTPUT_DIR / job_id
    if not out_dir.exists():
        return jsonify({"error": "Job output directory missing"}), 404

    reports = sorted(out_dir.glob("*_Report.xlsx"))
    if not reports:
        return jsonify({"error": "No *_Report.xlsx in this job's output"}), 404

    body = request.get_json(silent=True) or {}
    try:
        from airtable_uploader import upload_report
    except ImportError as e:
        return jsonify({"error": f"airtable_uploader unavailable: {e}"}), 503

    results = []
    for report in reports:
        try:
            r = upload_report(
                str(report),
                company_override=body.get("company"),
                user_email_override=body.get("user_email"),
                fy_override=body.get("fy"),
                dry_run=bool(body.get("dry_run")),
                log=lambda m: job.log(m, "info"),
            )
            r["report"] = report.name
            results.append(r)
        except Exception as e:
            results.append({"report": report.name, "ok": False, "error": str(e)})

    return jsonify({"job_id": job_id, "results": results})


@app.route("/api/download_all/<job_id>")
def download_all(job_id):
    """Download all output files as a single zip."""
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fpath in job.output_files:
            if os.path.exists(fpath):
                zf.write(fpath, os.path.basename(fpath))
    buf.seek(0)

    zip_name = f"ASK_MY_CFO_Output_{job_id}.zip"
    return send_file(buf, as_attachment=True, download_name=zip_name, mimetype="application/zip")


@app.route("/api/status/<job_id>")
def status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "status": job.status, "stage": job.stage, "progress": job.progress,
        "files": [os.path.basename(f) for f in job.output_files],
        "error": job.error,
    })


if __name__ == "__main__":
    UPLOAD_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    # Honor PORT env var so Railway / Heroku / Fly can inject the bound port.
    # Fallback to 5000 for local dev. host=0.0.0.0 is required on Railway.
    port = int(os.environ.get("PORT", "5000"))
    print()
    print("  ╔═══════════════════════════════════════════╗")
    print(f"  ║  ASK MY CFO — listening on port {port:<5}    ║")
    print("  ╚═══════════════════════════════════════════╝")
    print()
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
