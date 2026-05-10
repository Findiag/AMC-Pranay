# Bubble integration guide

Wire your Bubble app to the Ask My CFO API deployed on Railway. Three Bubble workflows cover 95% of use cases:

1. **Upload PDF** → kick off a job and store the `job_id`
2. **Poll status** → check `/api/status/<job_id>` until completed
3. **Download report** → either link to the file or pull it server-side

## Prerequisites

Before starting in Bubble:

- Your Railway URL (e.g. `https://askmycfo-production.up.railway.app`)
- Your `API_AUTH_TOKEN` (the one you set in Railway env vars)
- Bubble paid plan (Personal or higher) — the free plan can't make HTTPS calls to external APIs

---

## 1. Install the API Connector plugin

Bubble dashboard → **Plugins** → **+ Add plugins** → search **"API Connector"** → Install.

---

## 2. Configure the API Connector

Plugins tab → **API Connector** → **+ Add another API**:

| Field | Value |
|---|---|
| API Name | `AskMyCFO` |
| Authentication | **Private key in header** |
| Key name | `Authorization` |
| Key value | `Bearer YOUR_API_AUTH_TOKEN` (replace with your real token) |
| Use as | `Action` (Bubble will let you choose per-call later) |

Now add API calls inside this group.

---

## 3. API call: Upload PDF

| Field | Value |
|---|---|
| Name | `upload_report` |
| Use as | **Action** |
| Method | `POST` |
| URL | `https://YOUR-RAILWAY-URL/api/upload` |
| Body type | **Form-data** |
| Body | `files` = (file param), `skip_stage1` = `false`, `gen_observations` = `true` |

Mark the `files` row's checkbox **"Send file"** so Bubble passes it as multipart.

**Initialize call:** click **Initialize call** with a sample PDF. Bubble will read the response shape:

```json
{
  "job_id": "20260510_120000_a1b2c3",
  "files": ["sample.pdf"],
  "skip_stage1": false,
  "fy_override": null,
  "gen_observations": true
}
```

---

## 4. API call: Get job status

| Field | Value |
|---|---|
| Name | `get_status` |
| Use as | **Data** (Bubble lets you display fields directly) |
| Method | `GET` |
| URL | `https://YOUR-RAILWAY-URL/api/status/[job_id]` |

The `[job_id]` square-bracket syntax tells Bubble it's a parameter you'll pass at runtime.

Initialize with a real job_id from a previous upload. Response shape:

```json
{
  "status": "completed",
  "stage": 3,
  "progress": 100,
  "files": [
    "sample_BS_PL.pdf",
    "sample_extracted.xlsx",
    "sample_extracted_Report.xlsx"
  ],
  "error": null
}
```

`status` will be one of: `queued`, `running`, `completed`, `failed`.

---

## 5. API call: Download report file

| Field | Value |
|---|---|
| Name | `download_file` |
| Use as | **Action** |
| Method | `GET` |
| URL | `https://YOUR-RAILWAY-URL/api/download/[job_id]/[filename]` |
| Return type | **File** |

This returns the xlsx as a file Bubble can attach to a download button or save into the database.

---

## 6. Wiring the Bubble workflow

Typical happy path on a Bubble page:

```
PAGE: Upload Annual Report

  [File Uploader]   ← user drops a PDF
  [Button: Process] ← onclick:

    Step 1: API Action — AskMyCFO/upload_report
            Files = FileUploader's value
            → Set custom state job_id = result of step 1's job_id

    Step 2: Schedule a Custom Event — "check_status"
            After 5s
            Param: job_id = current state job_id

CUSTOM EVENT: check_status (job_id text)

  Step 1: API Data Source — AskMyCFO/get_status, job_id = passed param
          → Set custom state status = result's status
                          progress = result's progress

  Step 2: Conditional
    If status = "completed":
      → Show success message + download link to /api/download/{job_id}/{filename}
    Else if status = "failed":
      → Show error message
    Else:
      → Schedule another check_status in 3s
```

### Realtime progress (optional, advanced)

The `/api/stream/<job_id>` endpoint streams Server-Sent Events. Bubble's API Connector doesn't natively support SSE, but you can fall back to **polling `/api/status` every 2–3 seconds** which is good enough for most UX. SSE only matters if you want to show every log line live.

If you really need SSE, drop a small `<iframe>` element on the page that points at a static HTML file you host on Bubble's static hosting:

```html
<script>
  const es = new EventSource("https://YOUR-RAILWAY-URL/api/stream/JOB_ID?token=YOUR_TOKEN");
  es.onmessage = e => parent.postMessage(JSON.parse(e.data), "*");
</script>
```

Then a Bubble JavaScript-to-Bubble plugin captures the postMessage and updates state. Note the `?token=` query param — SSE clients can't easily set custom headers, so we accept the token as a query parameter as a fallback (only for `/api/stream`).

---

## 7. Showing the download link

After the status workflow detects `completed`:

```
Element: Link
  Destination URL: https://YOUR-RAILWAY-URL/api/download/<job_id>/<filename>
  Display text: "Download Financial Report"
```

Browsers can't send `Authorization` headers via plain anchor tags. Two options:

**Option A (simpler):** disable auth. Set `API_AUTH_TOKEN=""` in Railway and rely on the URL being unguessable (`job_id` is 20 chars, ~98 bits of entropy when combined with the random suffix). NOT recommended for production.

**Option B (recommended):** make the download go through Bubble's backend. Add a Backend Workflow:

```
Backend Workflow: serve_download (job_id text, filename text)

  Step 1: API Action — AskMyCFO/download_file
          job_id = passed param
          filename = passed param
  Step 2: Return data from API
          file = result of step 1
```

Then your link points at `https://yourapp.bubbleapps.io/api/1.1/wf/serve_download/?job_id=X&filename=Y`. Bubble injects the auth header server-side, and the user gets a clean download URL.

---

## 8. Triggering Airtable upload + observations on demand

If you set `AIRTABLE_AUTO_UPLOAD=false` in Railway (recommended — gives users control), then add two more API calls:

**push_to_airtable**
```
Method: POST
URL: https://YOUR-RAILWAY-URL/api/airtable/upload/[job_id]
Body type: JSON
Body: {}
```

**generate_observations**
```
Method: POST
URL: https://YOUR-RAILWAY-URL/api/observations/run
Body type: JSON
Body: {"company_id": "[company_id]", "force": false}
```

Wire each to its own button (e.g. "Push to Airtable", "Generate observations"). Users click only when ready.

---

## 9. Listing companies for a dropdown

```
API Data Source: list_companies
Method: GET
URL: https://YOUR-RAILWAY-URL/api/companies
```

Response:

```json
{
  "companies": [
    {"id": "recABC...", "name": "Ambuja Cements"},
    {"id": "recDEF...", "name": "Finolex Cables"}
  ]
}
```

Bind to a Dropdown element's "Choices source." Bubble auto-handles the array.

---

## 10. Embedding the diagnostic HTML report

Each company has an HTML report at `/api/report/<company_id>` (or by name at `/api/report/by-name/<company_name>`). To embed inside Bubble:

```
Element: HTML
  Content: <iframe src="https://YOUR-RAILWAY-URL/api/report/COMPANY_ID?token=YOUR_TOKEN"
                   style="width:100%; height:800px; border:0;"></iframe>
```

Same `?token=` fallback as SSE — iframes can't set headers.

---

## Common gotchas

| Symptom | Fix |
|---|---|
| Bubble shows "Workflow error: 401" | Token mismatch. Check `Bearer <token>` in API Connector header EXACTLY matches `API_AUTH_TOKEN` env var on Railway. |
| File upload returns 413 | PDF > 500 MB. Increase `app.config["MAX_CONTENT_LENGTH"]` in `app.py` and redeploy. |
| Status polling never reaches "completed" | First-ever job downloads ~100 MB sentence-transformers model — takes 30s extra. Subsequent jobs are fast. |
| CORS error in browser console | `CORS_ORIGINS` env var doesn't match your Bubble URL. Set to `https://yourapp.bubbleapps.io` (no path, no trailing slash). |
| Download link 404 | The `<filename>` portion is the BASENAME (no path). Get it from `status.files[]`. |
| Airtable upload says "company not found" | Either pre-create the company in Airtable, or the auto-create branch will fire (logged as `auto-created company '<name>' -> recXXX`). |

---

## Sample full workflow JSON (Bubble export)

If you want to import a ready-made workflow, paste this into a Bubble plugin's "API Connector" import:

```json
{
  "api_name": "AskMyCFO",
  "auth_type": "private_key_header",
  "header_key": "Authorization",
  "header_value_prefix": "Bearer ",
  "calls": [
    {
      "name": "upload_report",
      "method": "POST",
      "use_as": "action",
      "endpoint": "/api/upload",
      "body_type": "form-data",
      "params": ["files", "skip_stage1", "gen_observations", "fy_override"]
    },
    {
      "name": "get_status",
      "method": "GET",
      "use_as": "data",
      "endpoint": "/api/status/[job_id]"
    },
    {
      "name": "download_file",
      "method": "GET",
      "use_as": "action",
      "endpoint": "/api/download/[job_id]/[filename]",
      "return_type": "file"
    },
    {
      "name": "push_to_airtable",
      "method": "POST",
      "use_as": "action",
      "endpoint": "/api/airtable/upload/[job_id]",
      "body_type": "json"
    },
    {
      "name": "list_companies",
      "method": "GET",
      "use_as": "data",
      "endpoint": "/api/companies"
    }
  ]
}
```

(Bubble's UI is the source of truth — this is just a reference of the shape.)
