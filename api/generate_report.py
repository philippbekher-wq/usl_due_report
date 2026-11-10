#!/usr/bin/env python3
"""Vercel Python Function: POST /api/generate_report

Wraps tools/generate_dues_report.run_report() over HTTP. See
workflows/deploy_and_use_cloud_report_api.md for the request/response
contract and setup steps.

Every error the caller sees, plus a 200 where no report went out (no data, or
the email failed), is also reported to Sentry so the operator gets notified.
"""

import json
import os
import sys
from decimal import Decimal
from http.server import BaseHTTPRequestHandler
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tools"))

import generate_dues_report as gdr  # noqa: E402
import monitoring  # noqa: E402

ENDPOINT = "POST /api/generate_report"
MAX_DURATION_SECONDS = 120  # keep in sync with this function's maxDuration in vercel.json


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


def _problem(status: int, body: dict, request: dict):
    """What the operator should hear about this response, or None if it's fine."""
    if status >= 400:
        return body["error"]
    if body.get("status") == "no_data":
        return (
            f"No matching records for team={request.get('team')!r}, "
            f"{body['month']} {body['year']}; no report was generated"
        )
    if body.get("email_error") and not request.get("dry_run"):
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

    def do_POST(self):
        content_length = int(self.headers.get("Content-Length", 0))
        raw_body = self.rfile.read(content_length) if content_length else b""
        http_request = monitoring.describe_request(self, raw_body)  # what the caller sent, for Sentry
        with monitoring.timeout_watchdog(ENDPOINT, MAX_DURATION_SECONDS, http_request):
            request = {}  # parsed body, filled in by _process
            try:
                status, body = self._process(raw_body or b"{}", request)
                problem = _problem(status, body, request)
                if problem:
                    monitoring.report_problem(
                        ENDPOINT, problem, status=status, http_request=http_request,
                        contexts={"response_to_caller": _json_safe(body)},
                    )
            except Exception as e:
                status, body = 500, {"error": f"Unexpected server error ({type(e).__name__})."}
                monitoring.report_exception(ENDPOINT, e, status=status, http_request=http_request)
            monitoring.flush()
        self._send_json(status, body)

    def _process(self, raw_body: bytes, request: dict) -> tuple[int, dict]:
        try:
            body = json.loads(raw_body)
        except json.JSONDecodeError:
            return 400, {"error": "Request body must be valid JSON."}
        if not isinstance(body, dict):
            return 400, {"error": "Request body must be a JSON object."}
        request.update(body)

        expected_key = os.environ.get("API_SHARED_SECRET")
        provided_key = self.headers.get("X-API-Key")
        if not expected_key or provided_key != expected_key:
            return 401, {"error": "Missing or invalid X-API-Key header."}

        team = body.get("team")
        month = body.get("month")
        year = body.get("year")
        if not team or not month or not year:
            return 400, {"error": "Request body must include 'team', 'month', and 'year'."}

        try:
            result = gdr.run_report(
                team, str(month), int(year),
                dry_run=bool(body.get("dry_run", False)),
                email=bool(body.get("email", False)),
                to=body.get("to", gdr.DEFAULT_RECIPIENT_EMAIL),
                cc=body.get("additional_mail"),
                output_dir="/tmp/reports",
                template=str(Path(__file__).resolve().parent.parent / "report_sample.xlsx"),
                skip_schema_check=bool(body.get("skip_schema_check", False)),
                month_filter_override=body.get("month_filter_override"),
            )
        except (gdr.ConfigError, gdr.AuthError) as e:
            return 500, {"error": f"Server configuration error: {e}"}
        except ValueError as e:
            return 400, {"error": str(e)}
        except gdr.SchemaMismatchError as e:
            return 500, {"error": f"Airtable schema mismatch: {e}"}
        except (gdr.RateLimitError, gdr.AirtableAPIError) as e:
            return 502, {"error": f"Airtable API error: {e}"}
        except ImportError as e:
            return 500, {"error": f"Gmail dependencies not installed: {e}"}

        if result["no_data"]:
            return 200, {
                "status": "no_data",
                "team": team, "month": result["month_name"], "year": result["year"],
                "message": "No matching records found for this team/period.",
            }

        return 200, {
            "status": "ok",
            "is_batch": result["is_batch"],
            "month": result["month_name"],
            "year": result["year"],
            "results": result["results"],
            "emailed": result["emailed"],
            "email_error": result["email_error"],
        }

    def do_GET(self):
        self._send_json(200, {"status": "ok", "message": "POST a request to generate a report."})
