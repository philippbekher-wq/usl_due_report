#!/usr/bin/env python3
"""Sentry error reporting for the HTTP endpoints in api/.

Callers only see the HTTP response, so this is how the operator finds out a
request went wrong. Sentry turns each report into an issue and emails per the
project's alert rules. The cron endpoint additionally checks in with Sentry
Crons, which alerts when a scheduled run fails, overruns, or never happens.

Initialized on import from the SENTRY_DSN env var; with it unset (local runs)
every call here is a silent no-op. See the "Error Notifications" section of
workflows/deploy_and_use_cloud_report_api.md.
"""

import os
import threading
from contextlib import contextmanager

import sentry_sdk

FLUSH_TIMEOUT_SECONDS = 3

sentry_sdk.init(
    dsn=os.environ.get("SENTRY_DSN"),
    environment=os.environ.get("VERCEL_ENV", "local"),
    release=os.environ.get("VERCEL_GIT_COMMIT_SHA") or None,
    send_default_pii=False,
)


def _scope_kwargs(endpoint, status, request, response) -> dict:
    tags = {"endpoint": endpoint}
    if status is not None:
        tags["http_status"] = str(status)
    contexts = {}
    # Custom keys: "request"/"response" are reserved Sentry context types.
    if request:
        contexts["report_request"] = dict(request)
    if response:
        contexts["report_response"] = dict(response)
    return {"tags": tags, "contexts": contexts}


def report_problem(endpoint: str, message: str, *, status=None, request=None, response=None) -> None:
    """Report a handled failure (an error response, or a 200 where no report
    went out). Same message => same Sentry issue, so repeats don't re-alert
    unless the alert rule says so."""
    sentry_sdk.capture_message(
        f"{endpoint}: {message}", level="error",
        **_scope_kwargs(endpoint, status, request, response),
    )


def report_exception(endpoint: str, exc: BaseException, *, status=None, request=None, response=None) -> None:
    """Report an unexpected exception with its full traceback."""
    sentry_sdk.capture_exception(exc, **_scope_kwargs(endpoint, status, request, response))


def flush() -> None:
    """Send queued events now. Call before returning: Vercel freezes the
    function once the response is out, which would strand them."""
    sentry_sdk.flush(timeout=FLUSH_TIMEOUT_SECONDS)


@contextmanager
def timeout_watchdog(endpoint: str, max_duration_seconds: int, request: dict, margin_seconds: int = 15):
    """Report if the wrapped block is still running `margin_seconds` before
    Vercel's hard kill, which no except block can observe. `request` is read
    when the timer fires, so the caller may keep filling it in afterwards."""
    fire_after = max(1, max_duration_seconds - margin_seconds)

    def fire():
        report_problem(
            endpoint,
            f"still running after {fire_after}s; Vercel kills it at {max_duration_seconds}s",
            request=request,
        )
        flush()

    timer = threading.Timer(fire_after, fire)
    timer.daemon = True
    timer.start()
    try:
        yield
    finally:
        timer.cancel()
