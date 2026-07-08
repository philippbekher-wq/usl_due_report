#!/usr/bin/env python3
"""Vercel Python Function: GET /api/cron_brooklyn_report

Invoked by Vercel Cron (see the "crons" entry in vercel.json) to send the
Brooklyn FC (Championship) May 2026 dues report to philipp@lowcoding.dev on
a recurring schedule. Team/month/year/recipient are intentionally hardcoded
since this endpoint exists for exactly one recurring report; use
POST /api/generate_report (see workflows/deploy_and_use_cloud_report_api.md)
for anything else.

Vercel signs cron-triggered requests with an `Authorization: Bearer
$CRON_SECRET` header (set CRON_SECRET in the project's env vars) so this
endpoint can reject requests that didn't come from Vercel's scheduler.
"""

import json
import os
import sys
from decimal import Decimal
from http.server import BaseHTTPRequestHandler
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tools"))

import generate_dues_report as gdr  # noqa: E402

TEAM = "Brooklyn FC (Championship)"
MONTH = "May"
YEAR = 2026
RECIPIENT = "philipp@lowcoding.dev"


def _json_safe(value):
    if isinstance(value, Decimal):
        return str(value.quantize(Decimal("0.01")))
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return {k: _json_safe(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_json_safe(v) for v in value]
    return value


class handler(BaseHTTPRequestHandler):
    def _send_json(self, status: int, body: dict) -> None:
        payload = json.dumps(_json_safe(body)).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        expected_secret = os.environ.get("CRON_SECRET")
        provided_auth = self.headers.get("Authorization")
        if not expected_secret or provided_auth != f"Bearer {expected_secret}":
            self._send_json(401, {"error": "Missing or invalid cron authorization."})
            return

        try:
            result = gdr.run_report(
                TEAM, MONTH, YEAR,
                dry_run=False,
                email=True,
                to=RECIPIENT,
                cc=None,
                output_dir="/tmp/reports",
                template=str(Path(__file__).resolve().parent.parent / "report_sample.xlsx"),
            )
        except (gdr.ConfigError, gdr.AuthError) as e:
            self._send_json(500, {"error": f"Server configuration error: {e}"})
            return
        except gdr.SchemaMismatchError as e:
            self._send_json(500, {"error": f"Airtable schema mismatch: {e}"})
            return
        except (gdr.RateLimitError, gdr.AirtableAPIError) as e:
            self._send_json(502, {"error": f"Airtable API error: {e}"})
            return
        except ImportError as e:
            self._send_json(500, {"error": f"Gmail dependencies not installed: {e}"})
            return

        if result["no_data"]:
            self._send_json(200, {
                "status": "no_data", "team": TEAM, "month": MONTH, "year": YEAR,
            })
            return

        self._send_json(200, {
            "status": "ok",
            "results": result["results"],
            "emailed": result["emailed"],
            "email_error": result["email_error"],
        })
