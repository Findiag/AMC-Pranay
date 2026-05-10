# Deploy to Railway — step by step

End-to-end deploy of Ask My CFO to [Railway](https://railway.com). After this guide you'll have a publicly-reachable HTTPS URL ready to wire to Bubble.

**Cost estimate:** Hobby plan ($5/mo) + ~$0.25/mo per GB volume + Anthropic API calls (~₹0.025–₹0.50 per report depending on PDF size).

---

## 1. Prep your repo

```bash
# 1. Verify .gitignore is in place — it should EXCLUDE config.json
git status                          # config.json must NOT be in the list
git status --ignored | grep config  # should show "Ignored files: config.json"

# 2. Initial commit
git init
git add .
git commit -m "Initial commit"

# 3. Push to GitHub
git remote add origin https://github.com/<you>/ask-my-cfo.git
git branch -M main
git push -u origin main
```

> **CRITICAL:** if `config.json` shows up in `git status` (not "Ignored"), DO NOT COMMIT. Open `.gitignore` and confirm `config.json` is on its own line. Run `git rm --cached config.json` if it's already tracked.

---

## 2. Create the Railway project

1. Go to https://railway.com/dashboard → **New Project** → **Deploy from GitHub repo**
2. Select your `ask-my-cfo` repo
3. Railway will detect Python via `runtime.txt` + `requirements.txt` and start the first build automatically

The first build takes ~3–5 minutes (chromadb + sentence-transformers are large).

---

## 3. Set environment variables

In your service → **Variables** tab → **Raw Editor**, paste:

```bash
# Required
ANTHROPIC_API_KEY=sk-ant-api03-...
API_AUTH_TOKEN=<generate a long random string — see below>

# Recommended
DATA_DIR=/data
CORS_ORIGINS=https://yourapp.bubbleapps.io
STANDALONE_ONLY=true
MAPPER_MODE=claude

# Airtable (only if you want auto-upload)
AIRTABLE_API_TOKEN=pat...
AIRTABLE_BASE_ID=app...
AIRTABLE_BS_PL_TABLE_ID=tbl...
AIRTABLE_USER_TABLE_ID=tbl...
AIRTABLE_COMPANY_TABLE_ID=tbl...
AIRTABLE_USER_EMAIL=you@yourdomain.com
AIRTABLE_AUTO_UPLOAD=false        # set true to push every job automatically
AUTO_OBSERVATIONS=true
```

Generate a strong `API_AUTH_TOKEN`:

```bash
python -c "import secrets; print(secrets.token_urlsafe(32))"
# → e.g. 'X3hJk2_aF9pQz4vR8nL6tY1wU5sB7iE0mC4dN2gH'
```

Save your token — Bubble will need this exact string.

---

## 4. Add a persistent volume

ChromaDB and the sentence-transformers model cache need to survive deploys. Without this, every deploy re-downloads ~100 MB of model weights and wipes accumulated `/api/corrections` feedback.

1. Service → **Volumes** → **Add Volume**
2. **Mount path:** `/data`
3. **Size:** start at 5 GB (resize later if reports pile up)

The `DATA_DIR=/data` env var (set in step 3) tells the app to use this volume for `uploads/`, `output/`, `chroma_db/`, and `model_cache/`.

---

## 5. Generate a public URL

1. Service → **Settings** → **Networking**
2. **Generate Domain** → Railway gives you `<service>-production.up.railway.app`
3. (Optional) **Custom Domain** → add `api.yourdomain.com` and follow the CNAME instructions

---

## 6. Verify the deploy

```bash
# Liveness — no auth required
curl https://<your-app>.up.railway.app/api/health
# → {"status":"ok"}

# Auth check — should fail without the token
curl https://<your-app>.up.railway.app/api/config
# → {"error":"unauthorized — missing or wrong Bearer token"}

# Auth check — should succeed
curl -H "Authorization: Bearer YOUR_API_AUTH_TOKEN" \
     https://<your-app>.up.railway.app/api/config
# → {"anthropic_api_key_set": true, "claude_model": "claude-sonnet-4-6", ...}

# Full health snapshot
curl -H "Authorization: Bearer YOUR_API_AUTH_TOKEN" \
     https://<your-app>.up.railway.app/health
# → {"status":"ok","anthropic_key_configured":true,"mode":"ai_full",...}
```

If `anthropic_key_configured` is `false`, the env var isn't reaching the app — check Variables tab and re-deploy.

---

## 7. Test an actual upload

```bash
curl -X POST \
  -H "Authorization: Bearer YOUR_API_AUTH_TOKEN" \
  -F "files=@./sample_annual_report.pdf" \
  -F "skip_stage1=false" \
  https://<your-app>.up.railway.app/api/upload
# → {"job_id":"20260510_120000_a1b2c3","files":["sample_annual_report.pdf"],...}

# Stream progress (Ctrl-C to stop)
curl -H "Authorization: Bearer YOUR_API_AUTH_TOKEN" \
     https://<your-app>.up.railway.app/api/stream/20260510_120000_a1b2c3
# → SSE stream of log lines

# Download the final Report once status='completed'
curl -H "Authorization: Bearer YOUR_API_AUTH_TOKEN" \
     -o report.xlsx \
     https://<your-app>.up.railway.app/api/download/20260510_120000_a1b2c3/sample_annual_report_extracted_Report.xlsx
```

---

## 8. Wire to Bubble

See **[BUBBLE_INTEGRATION.md](BUBBLE_INTEGRATION.md)** for the API workflows.

---

## Operational notes

### Scaling

The default `Procfile` uses `gunicorn --workers 2 --threads 4`. That's right for Hobby (1 vCPU). For Pro plan (8 vCPU):

```bash
# Update Procfile or set START_COMMAND env var:
gunicorn --bind 0.0.0.0:$PORT --workers 4 --threads 8 --timeout 600 app:app
```

A single mapping job uses ~1 worker for ~60–90 seconds. Two workers handle ~80 jobs/hour.

### Logs

Service → **Deployments** → click the active deployment → **View Logs**. Streams gunicorn access logs + every `print()` from the pipeline.

### Restarting after env-var changes

Updating env vars in Railway triggers an automatic redeploy. No manual restart needed.

### Volume backups

Railway volumes don't auto-snapshot on Hobby plan. For ChromaDB durability, periodically dump:

```bash
# Inside the Railway shell (or via CLI: `railway run bash`)
tar -czf /tmp/chroma_backup_$(date +%F).tar.gz /data/chroma_db
# Then download via railway-cli or copy to S3
```

### Cost monitoring

- Anthropic spend: https://console.anthropic.com → Usage. Set a budget alert.
- Railway: Settings → Usage. The hobby $5 plan covers ~750 single-worker hours/mo.

### Updating the deployed code

```bash
git push origin main
# Railway auto-deploys on push. Watch the Deployments tab.
```

Roll back by clicking an older deployment → **Redeploy**.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `Application failed to respond` on `/api/health` | gunicorn didn't start | Check Deployment logs for ImportError; usually missing env var or syntax error |
| 401 on every request from Bubble | Wrong / missing `Authorization` header | Confirm Bubble's "Authorization" param matches `API_AUTH_TOKEN` exactly (case-sensitive, no `Bearer ` prefix duplicated) |
| Upload returns 500 | Volume not mounted, can't write to `/data` | Verify volume exists at `/data`; check `DATA_DIR=/data` env var is set |
| `anthropic_key_configured: false` in `/health` | Env var not set or service not redeployed | Set `ANTHROPIC_API_KEY` and redeploy |
| First upload after deploy is slow (~30s extra) | Sentence-transformers model downloading | One-time. Subsequent uploads use the cached model in `/data/model_cache/` |
| ChromaDB lost after deploy | No volume mounted | Add a volume at `/data`; redeploy |
| CORS errors in Bubble browser console | `CORS_ORIGINS` doesn't match Bubble URL | Set `CORS_ORIGINS=https://yourapp.bubbleapps.io` (no trailing slash) |
| 504 timeout on large PDFs | gunicorn timeout < extraction time | Already set to 600s in Procfile; large PDFs (>200 pages) may need 900 |
