# Workflow: Deploy & Use the Dues Report HTTP API (Vercel)

## Objective
Trigger the same dues report generation as the CLI tool (`tools/generate_dues_report.py`), but over HTTP — for use from external systems (a webhook, a scheduled job, a form) rather than a local terminal. Hosted on Vercel as a Python serverless function.

## Architecture
`api/generate_report.py` is a thin HTTP wrapper around `tools/generate_dues_report.run_report()` — the exact same Airtable-querying, xlsx-building, and emailing logic the CLI uses, unchanged. The only real differences from local CLI usage:
- Generated files are written to `/tmp` (Vercel's one writable, non-persistent directory) instead of `reports/` — there is no persistent storage; email is the only delivery mechanism for the cloud version.
- Gmail auth uses env vars (`GMAIL_CLIENT_ID`/`GMAIL_CLIENT_SECRET`/`GMAIL_REFRESH_TOKEN`) instead of local `credentials.json`/`token.json` files, since there's no browser available for interactive consent in a serverless function.

## One-Time Setup

### 1. Google Cloud Console — switch OAuth consent screen to "Internal"
`lowcoding.dev` is a Google Workspace domain, so: **APIs & Services → OAuth consent screen → User Type → Internal**. This removes the 7-day refresh-token expiry that applies to "Testing" status apps, and requires no Google verification review. **Do this before relying on the cloud deployment long-term** — otherwise the deployed function's emailing will silently start failing about a week after the token was issued.

### 2. Deploy to Vercel
```
npm i -g vercel   # if not already installed
vercel login
vercel            # first deploy, links the project
```

### 3. Set environment variables
In the Vercel dashboard (Project → Settings → Environment Variables) or via `vercel env add <NAME> production`:
- `AIRTABLE_API_KEY` — same Personal Access Token as local `.env`
- `AIRTABLE_BASE_ID`, `AIRTABLE_TABLE_ID`, `AIRTABLE_TEAMS_TABLE_ID` — optional, same defaults as `.env.example`
- `API_SHARED_SECRET` — a long random string you generate; callers must send this as the `X-API-Key` header
- `GMAIL_CLIENT_ID`, `GMAIL_CLIENT_SECRET` — from local `credentials.json`
- `GMAIL_REFRESH_TOKEN` — from local `token.json` (obtained during the one-time local `--email` OAuth consent)

Never paste these values into chat or a shared doc — set them directly via the Vercel dashboard or CLI.

### 4. Redeploy after setting env vars
```
vercel --prod
```

## Request / Response Contract

`POST /api/generate_report`

Headers: `X-API-Key: <API_SHARED_SECRET value>`, `Content-Type: application/json`

Body:
```json
{
  "team": "Brooklyn FC (Championship)",
  "month": "May",
  "year": 2026,
  "email": true,
  "to": ["philipp@lowcoding.dev", "someone-else@x.com"],
  "additional_mail": "manager@x.com"
}
```
- `team`: exact Airtable `Team Name`, or `"all"` for every team with matching data that period (same semantics as the CLI's `--team all`).
- `month`: number (1-12) or name.
- `email` (optional, default `false`): if true, emails the generated report(s) as attachments, same as CLI `--email`.
- `to` (optional): overrides the default recipient (`philipp@lowcoding.dev`). Accepts either a single string or an array of email addresses — all listed addresses go in the To field of one email.
- `additional_mail` (optional): one or more email addresses (string or array) CC'd on the same email.
- `dry_run` (optional): query and summarize without generating/emailing anything.

Example:
```
curl -X POST https://<your-deployment>.vercel.app/api/generate_report \
  -H "X-API-Key: <secret>" \
  -H "Content-Type: application/json" \
  -d '{"team": "Brooklyn FC (Championship)", "month": "May", "year": 2026, "email": true}'
```

Response (200):
```json
{
  "status": "ok",
  "is_batch": false,
  "month": "May",
  "year": 2026,
  "results": [{"team": "...", "matched": 20, "merged": 10, "total_dues": "1366.30", "total_hq": "406.06", "output_path": "/tmp/reports/(PU# 791101) Brooklyn FC (Championship) _ May_2026.xlsx"}],
  "emailed": true,
  "email_error": null
}
```
Note: `output_path` reflects the new deliverable-format filename — `(PU# <number>) <Team> _ <Month>_<Year>.xlsx`, where `PU#` comes from the Teams table's `CWA PU Number` field (falls back to `(PU# unknown)` if a team has none on file). Since this filename has no timestamp, re-running the same team/month/year overwrites rather than versioning.

Status codes: `400` bad/missing input, `401` missing or wrong `X-API-Key`, `500` server-side config/schema problem, `502` upstream Airtable failure. `200` with `"status": "no_data"` means the request was valid but nothing matched (not an error — mirrors the CLI's exit-2 case).

## Edge Cases & Known Behaviors
All CLI edge cases still apply (duplicate-player merging, Dues Paid/Player Invoiced filtering, batch skip behavior for teams with no data that period) since the underlying logic is identical — see `workflows/generate_team_dues_report.md`.
- **No persistent storage**: files exist only in `/tmp` for the duration of one request/response. If `email` is false, the generated report is produced and then effectively discarded — there's no way to retrieve it afterward. Always pass `"email": true` for the cloud version unless you're just testing with `"dry_run": true`.
- **A failed email doesn't fail the request**: if report generation succeeds but the Gmail send fails, the response is still `200` with `"emailed": false` and `"email_error"` describing what went wrong — the caller should check that field explicitly rather than assuming `200` means the email was sent.
- **`GMAIL_REFRESH_TOKEN` expiry**: only relevant if the OAuth consent screen setup step (Internal) wasn't done — see the note above.
- **`--team all` runtime**: observed ~60s locally for 66 teams; `vercel.json` sets `maxDuration: 120` for headroom, well under Vercel Hobby's 300s ceiling.

## Learnings
_(empty — update after the first real deployed request, especially anything about cold-start latency, the actual `/tmp` write behavior in production, or whether the Internal consent screen change fully eliminated token expiry.)_
