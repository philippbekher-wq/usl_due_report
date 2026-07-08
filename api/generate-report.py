#!/usr/bin/env python3
"""Vercel Python Function: POST /api/generate-report

Wraps tools/generate_dues_report.run_report() over HTTP. See
workflows/deploy_and_use_cloud_report_api.md for the request/response
contract and setup steps.
"""

import json
import os
import sys
from decimal import Decimal
from http.server import BaseHTTPRequestHandler
from pathlib import Path

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tools"))

import generate_dues_report as gdr  # noqa: E402


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

    def do_POST(self):
        content_length = int(self.headers.get("Content-Length", 0))
        raw_body = self.rfile.read(content_length) if content_length else b"{}"
        try:
            body = json.loads(raw_body)
        except json.JSONDecodeError:
            self._send_json(400, {"error": "Request body must be valid JSON."})
            return

        expected_key = os.environ.get("API_SHARED_SECRET")
        provided_key = self.headers.get("X-API-Key")
        if not expected_key or provided_key != expected_key:
            self._send_json(401, {"error": "Missing or invalid X-API-Key header."})
            return

        team = body.get("team")
        month = body.get("month")
        year = body.get("year")
        if not team or not month or not year:
            self._send_json(400, {"error": "Request body must include 'team', 'month', and 'year'."})
            return

        try:
            result = gdr.run_report(
                team, str(month), int(year),
                dry_run=bool(body.get("dry_run", False)),
                email=bool(body.get("email", False)),
                to=body.get("to", gdr.DEFAULT_RECIPIENT_EMAIL),
                output_dir="/tmp/reports",
                template=str(Path(__file__).resolve().parent.parent / "report_sample.xlsx"),
                skip_schema_check=bool(body.get("skip_schema_check", False)),
                month_filter_override=body.get("month_filter_override"),
            )
        except (gdr.ConfigError, gdr.AuthError) as e:
            self._send_json(500, {"error": f"Server configuration error: {e}"})
            return
        except ValueError as e:
            self._send_json(400, {"error": str(e)})
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
                "status": "no_data",
                "team": team, "month": result["month_name"], "year": result["year"],
                "message": "No matching records found for this team/period.",
            })
            return

        self._send_json(200, {
            "status": "ok",
            "is_batch": result["is_batch"],
            "month": result["month_name"],
            "year": result["year"],
            "results": result["results"],
            "emailed": result["emailed"],
            "email_error": result["email_error"],
        })

    def do_GET(self):
        self._send_json(200, {"status": "ok", "message": "POST a request to generate a report."})
