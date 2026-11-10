#!/usr/bin/env python3
"""Sentry error reporting for the HTTP endpoints in api/.

Callers only see the HTTP response, so this is how the operator finds out a
request went wrong. Sentry turns each report into an issue and emails per the
project's alert rules. The cron endpoint additionally checks in with Sentry
Crons, which alerts when a scheduled run fails, overruns, or never happens.

Every report carries what the caller sent (Sentry's "Request" section: method,
URL, safe headers, body) so the cause can be diagnosed from the issue alone.

Initialized on import from the SENTRY_DSN env var; with it unset (local runs)
every call here is a silent no-op. See the "Error Notifications" section of
workflows/deploy_and_use_cloud_report_api.md.
"""

import json
import os
import threading
from contextlib import contextmanager

import sentry_sdk

FLUSH_TIMEOUT_SECONDS = 3
MAX_RAW_BODY_CHARS = 4000
SAFE_HEADERS = ("User-Agent", "Content-Type", "Content-Length")
SECRET_HEADERS = ("X-API-Key", "Authorization")  # only their presence is reported

sentry_sdk.init(
    dsn=os.environ.get("SENTRY_DSN"),
    environment=os.environ.get("VERCEL_ENV", "local"),
    release=os.environ.get("VERCEL_GIT_COMMIT_SHA") or None,
    send_default_pii=False,
    # Frame locals would carry secrets (API_SHARED_SECRET/CRON_SECRET in the
    # handlers, the Airtable key in run_report's Config). File/line/code of
    # each frame are still reported.
    include_local_variables=False,
)


def describe_request(handler, raw_body: bytes = b"") -> dict:
    """What the caller sent, in Sentry's HTTP request format, minus secrets.
    The body is parsed JSON when possible, otherwise the raw text, so input
    that failed to parse is still visible."""
    headers = {h: handler.headers[h] for h in SAFE_HEADERS if handler.headers.get(h)}
    headers.update({h: "[Filtered]" for h in SECRET_HEADERS if handler.headers.get(h)})
    path, _, query = handler.path.partition("?")
    info = {
        "method": handler.command,
        "url": f"https://{handler.headers.get('Host', '')}{path}",
        "query_string": query,
        "headers": headers,
    }
    if raw_body:
        text = raw_body.decode("utf-8", errors="replace")
        try:
            info["data"] = json.loads(text)
        except json.JSONDecodeError:
            info["data"] = text[:MAX_RAW_BODY_CHARS]
    return info


def _capture(capture, endpoint, status, http_request, contexts) -> None:
    # A fresh scope per report, so nothing leaks between requests on a warm instance.
    with sentry_sdk.new_scope() as scope:
        scope.set_tag("endpoint", endpoint)
        if status is not None:
            scope.set_tag("http_status", str(status))
        for name, value in (contexts or {}).items():
            if value:
                scope.set_context(name, dict(value))
        if http_request:
            def add_request(event, hint):
                event["request"] = http_request
                return event
            scope.add_event_processor(add_request)
        capture()


def report_problem(endpoint: str, message: str, *, status=None, http_request=None, contexts=None) -> None:
    """Report a handled failure (an error response, or a 200 where no report
    went out). Same message => same Sentry issue, so repeats don't re-alert
    unless the alert rule says so. `contexts` is {name: dict} of extra detail."""
    _capture(
        lambda: sentry_sdk.capture_message(f"{endpoint}: {message}", level="error"),
        endpoint, status, http_request, contexts,
    )


def report_exception(endpoint: str, exc: BaseException, *, status=None, http_request=None, contexts=None) -> None:
    """Report an unexpected exception with its full traceback."""
    _capture(lambda: sentry_sdk.capture_exception(exc), endpoint, status, http_request, contexts)


def flush() -> None:
    """Send queued events now. Call before returning: Vercel freezes the
    function once the response is out, which would strand them."""
    sentry_sdk.flush(timeout=FLUSH_TIMEOUT_SECONDS)


@contextmanager
def timeout_watchdog(endpoint: str, max_duration_seconds: int, http_request: dict, margin_seconds: int = 15):
    """Report if the wrapped block is still running `margin_seconds` before
    Vercel's hard kill, which no except block can observe."""
    fire_after = max(1, max_duration_seconds - margin_seconds)

    def fire():
        report_problem(
            endpoint,
            f"still running after {fire_after}s; Vercel kills it at {max_duration_seconds}s",
            http_request=http_request,
        )
        flush()

    timer = threading.Timer(fire_after, fire)
    timer.daemon = True
    timer.start()
    try:
        yield
    finally:
        timer.cancel()
