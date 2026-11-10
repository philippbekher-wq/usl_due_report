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

Each run checks in with Sentry Crons (monitor MONITOR_SLUG, auto-created from
MONITOR_CONFIG), which alerts when a run fails, overruns, or never happens.
Any run that doesn't email the report is also reported to Sentry with details.
"""

import json
import os
import sys
import time
from decimal import Decimal
from http.server import BaseHTTPRequestHandler
from pathlib import Path

from sentry_sdk.crons import capture_checkin
from sentry_sdk.crons.consts import MonitorStatus

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tools"))

import generate_dues_report as gdr  # noqa: E402
import monitoring  # noqa: E402

ENDPOINT = "GET /api/cron_brooklyn_report"
TEAM = "Brooklyn FC (Championship)"
MONTH = "May"
YEAR = 2026
RECIPIENT = "philipp@lowcoding.dev"

MONITOR_SLUG = "cron-brooklyn-report"
MONITOR_CONFIG = {
    "schedule": {"type": "crontab", "value": "0 3 * * *"},  # keep in sync with vercel.json "crons"
    "timezone": "UTC",  # Vercel cron schedules are UTC
    # Minutes. Vercel Hobby only guarantees a cron fires somewhere within its
    # scheduled hour, so a tighter margin would raise false "missed" alerts.
    "checkin_margin": 90,
    "max_runtime": 2,  # minutes; the function itself is killed at maxDuration=60s
    "failure_issue_threshold": 1,
    "recovery_threshold": 1,
}


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


def _problem(status: int, body: dict):
    """Why this run didn't deliver the report, or None if it did."""
    if status >= 400:
        return body["error"]
    if body["status"] == "no_data":
        return f"No matching records for {TEAM} {MONTH} {YEAR}; nothing was sent"
    if not body["emailed"]:
        return f"Report generated but not emailed: {body['email_error']}"
    return None


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
            if not expected_secret:
                # Vercel only sends the Authorization header when CRON_SECRET is
                # set, so every scheduled run is being rejected. (With it set, a
                # mismatch is just a stray request and not worth an alert.)
                monitoring.report_problem(
                    ENDPOINT, "CRON_SECRET is not set; scheduled runs are rejected with 401", status=401,
                )
                capture_checkin(MONITOR_SLUG, status=MonitorStatus.ERROR, monitor_config=MONITOR_CONFIG)
                monitoring.flush()
            self._send_json(401, {"error": "Missing or invalid cron authorization."})
            return

        check_in_id = capture_checkin(
            MONITOR_SLUG, status=MonitorStatus.IN_PROGRESS, monitor_config=MONITOR_CONFIG,
        )
        started = time.monotonic()
        try:
            status, body = self._run()
            problem = _problem(status, body)
            if problem:
                monitoring.report_problem(ENDPOINT, problem, status=status, response=_json_safe(body))
        except Exception as e:
            status, body = 500, {"error": f"Unexpected server error ({type(e).__name__})."}
            problem = str(e)
            monitoring.report_exception(ENDPOINT, e, status=status)
        capture_checkin(
            MONITOR_SLUG, check_in_id=check_in_id,
            status=MonitorStatus.ERROR if problem else MonitorStatus.OK,
            duration=time.monotonic() - started,
        )
        monitoring.flush()
        self._send_json(status, body)

    def _run(self) -> tuple[int, dict]:
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
            return 500, {"error": f"Server configuration error: {e}"}
        except gdr.SchemaMismatchError as e:
            return 500, {"error": f"Airtable schema mismatch: {e}"}
        except (gdr.RateLimitError, gdr.AirtableAPIError) as e:
            return 502, {"error": f"Airtable API error: {e}"}
        except ImportError as e:
            return 500, {"error": f"Gmail dependencies not installed: {e}"}

        if result["no_data"]:
            return 200, {"status": "no_data", "team": TEAM, "month": MONTH, "year": YEAR}

        return 200, {
            "status": "ok",
            "results": result["results"],
            "emailed": result["emailed"],
            "email_error": result["email_error"],
        }
