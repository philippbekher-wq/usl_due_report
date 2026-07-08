#!/usr/bin/env python3
"""Generate a monthly per-team USL dues report as an .xlsx file, sourced from Airtable."""

import argparse
import calendar
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import requests
from dotenv import load_dotenv
import openpyxl

DEFAULT_BASE_ID = "appiVO0LP0Xm603VO"
DEFAULT_TABLE_ID = "tblgvGu3ueD1HOm8I"
DEFAULT_TEAMS_TABLE_ID = "tblxtS2ioQsKvjcZ3"
DEFAULT_RECIPIENT_EMAIL = "philipp@lowcoding.dev"
EMAIL_SUBJECT = "Report"

EXPECTED_RECORD_FIELDS = [
    "Unique ID", "Player", "Full Name", "Current Base Compensation", "Dues Paid",
    "Street Address Line 1", "Street Address Line 2", "City", "State", "Zip",
    "Computed Wage", "Computed PC", "Computed MRF", "Total Due to National HQ", "Created",
]
FILTER_FIELDS = ["Team Name", "Payroll Date Month", "Payroll Year", "Dues Paid", "Player Invoiced"]

COLUMN_FIELD_MAP = {
    "A": "Unique ID",
    "B": "Full Name",  # "Player" is a linked-record field (returns record IDs); the sample's B/C columns are always identical, so both source from "Full Name".
    "C": "Full Name",
    "D": "Current Base Compensation",
    "E": "Dues Paid",
    "F": "Street Address Line 1",
    "G": "Street Address Line 2",
    "H": "City",
    "I": "State",
    "J": "Zip",
    "K": "Computed Wage",
    "L": "Computed PC",
    "M": "Computed MRF",
    "N": "Total Due to National HQ",
    "O": "Created",  # Airtable's "Created" (createdTime) field populates the report's "Sending Date" column.
}

CURRENCY_FORMAT = '"$"#,##0.00'
DATE_FORMAT = "mm/dd/yy"
NUMBER_FORMAT_COLUMNS = {
    "D": CURRENCY_FORMAT, "E": CURRENCY_FORMAT,
    "K": CURRENCY_FORMAT, "L": CURRENCY_FORMAT,
    "M": CURRENCY_FORMAT, "N": CURRENCY_FORMAT,
    "O": DATE_FORMAT,
}

MAX_PAGES = 50
RATE_LIMIT_COOLDOWN_SECONDS = 30
MAX_RATE_LIMIT_RETRIES = 3


class ConfigError(Exception):
    pass


class AuthError(Exception):
    pass


class RateLimitError(Exception):
    pass


class AirtableAPIError(Exception):
    pass


class SchemaMismatchError(Exception):
    pass


@dataclass
class Config:
    api_key: str
    base_id: str
    table_id: str
    teams_table_id: str


def load_config() -> Config:
    load_dotenv()
    api_key = os.getenv("AIRTABLE_API_KEY")
    if not api_key:
        raise ConfigError(
            "AIRTABLE_API_KEY is not set. Copy .env.example to .env and fill in "
            "your Airtable Personal Access Token."
        )
    return Config(
        api_key=api_key,
        base_id=os.getenv("AIRTABLE_BASE_ID", DEFAULT_BASE_ID),
        table_id=os.getenv("AIRTABLE_TABLE_ID", DEFAULT_TABLE_ID),
        teams_table_id=os.getenv("AIRTABLE_TEAMS_TABLE_ID", DEFAULT_TEAMS_TABLE_ID),
    )


def normalize_month(month_arg: str) -> tuple[int, str]:
    s = month_arg.strip()
    if s.isdigit():
        num = int(s)
        if not 1 <= num <= 12:
            raise ValueError(f"Month number must be 1-12, got {num}")
        return num, calendar.month_name[num]
    s_lower = s.lower()
    for i in range(1, 13):
        if s_lower in (calendar.month_name[i].lower(), calendar.month_abbr[i].lower()):
            return i, calendar.month_name[i]
    raise ValueError(f"Could not parse month {month_arg!r} — use a number (1-12) or a name (e.g. 'July').")


def _request_with_retry(cfg: Config, url: str, params: dict) -> requests.Response:
    headers = {"Authorization": f"Bearer {cfg.api_key}"}
    attempt = 0
    while True:
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        if resp.status_code == 401:
            raise AuthError(
                f"Airtable rejected the API token (401) — check AIRTABLE_API_KEY in .env "
                f"and confirm it has access to base {cfg.base_id}."
            )
        if resp.status_code == 429:
            attempt += 1
            if attempt > MAX_RATE_LIMIT_RETRIES:
                raise RateLimitError("Airtable rate limit (429) persisted after retries.")
            # Airtable's documented policy is a ~30s cooldown after a 429, not classic backoff.
            time.sleep(RATE_LIMIT_COOLDOWN_SECONDS)
            continue
        if not resp.ok:
            raise AirtableAPIError(f"Airtable API error {resp.status_code}: {resp.text[:500]}")
        return resp


def get_table_schema(cfg: Config) -> dict[str, dict]:
    url = f"https://api.airtable.com/v0/meta/bases/{cfg.base_id}/tables"
    resp = _request_with_retry(cfg, url, {})
    data = resp.json()
    for table in data.get("tables", []):
        if table.get("id") == cfg.table_id:
            schema = {}
            for f in table.get("fields", []):
                ftype = f.get("type")
                if ftype in ("formula", "rollup"):
                    ftype = (f.get("options") or {}).get("result", {}).get("type", ftype)
                schema[f["name"]] = {"type": ftype, "id": f.get("id")}
            return schema
    raise SchemaMismatchError(f"Table {cfg.table_id} not found in base {cfg.base_id} metadata response.")


def validate_expected_fields(schema: dict, expected: list[str]) -> None:
    missing = [f for f in expected if f not in schema]
    if missing:
        available = ", ".join(sorted(schema.keys()))
        raise SchemaMismatchError(
            f"Missing expected field(s) in Airtable table: {', '.join(missing)}.\n"
            f"Available fields: {available}"
        )


def build_filter_formula(team, month_num, month_name, year, field_types, override=None) -> str:
    team_escaped = team.replace('"', '\\"')
    team_clause = f'{{Team Name}} = "{team_escaped}"'

    if override:
        month_clause = override
    else:
        month_type = field_types.get("Payroll Date Month")
        if month_type == "number":
            month_clause = f"{{Payroll Date Month}} = {month_num}"
        elif month_type in ("date", "dateTime"):
            month_clause = f'DATETIME_FORMAT({{Payroll Date Month}}, "M") = "{month_num}"'
        else:
            month_clause = f'{{Payroll Date Month}} = "{month_name}"'

    year_type = field_types.get("Payroll Year")
    if year_type == "number":
        year_clause = f"{{Payroll Year}} = {year}"
    else:
        year_clause = f'{{Payroll Year}} = "{year}"'

    # Records with no Dues Paid value or with Player Invoiced checked are excluded —
    # they represent players not yet processed for this payroll period.
    dues_paid_clause = "NOT({Dues Paid} = BLANK())"
    not_invoiced_clause = "NOT({Player Invoiced})"

    return f"AND({team_clause}, {month_clause}, {year_clause}, {dues_paid_clause}, {not_invoiced_clause})"


def fetch_all_records(
    cfg: Config, formula: str = None, table_id: str = None, fields: list[str] = None, max_pages: int = MAX_PAGES,
) -> list[dict]:
    url = f"https://api.airtable.com/v0/{cfg.base_id}/{table_id or cfg.table_id}"
    records = []
    offset = None
    page_count = 0
    while True:
        params = {"pageSize": 100}
        if formula:
            params["filterByFormula"] = formula
        if fields:
            params["fields[]"] = fields
        if offset:
            params["offset"] = offset
        resp = _request_with_retry(cfg, url, params)
        data = resp.json()
        records.extend(data.get("records", []))
        offset = data.get("offset")
        page_count += 1
        if not offset:
            break
        if page_count >= max_pages:
            raise AirtableAPIError(f"Pagination exceeded {max_pages} pages; aborting as a safety guard.")
    return records


def get_team_metadata(cfg: Config) -> dict[str, dict]:
    """Fetch Name + CWA PU Number for every team in the Teams table. Keyed by
    the exact 'Name' field value; used both to resolve --team all's team list
    and to look up each team's PU# for the output filename."""
    records = fetch_all_records(cfg, table_id=cfg.teams_table_id, fields=["Name", "CWA PU Number"])
    metadata = {}
    for r in records:
        fields = r.get("fields", {})
        name = fields.get("Name")
        if not name or not str(name).strip():
            continue
        metadata[str(name).strip()] = {"pu_number": fields.get("CWA PU Number")}
    if not metadata:
        raise AirtableAPIError(
            f"No team names found in Teams table {cfg.teams_table_id} — check AIRTABLE_TEAMS_TABLE_ID "
            f"and confirm the table has a 'Name' field."
        )
    return metadata


def lookup_pu_number(team_metadata: dict, team_name: str):
    if team_name in team_metadata:
        return team_metadata[team_name]["pu_number"]
    lower_map = {k.lower(): v["pu_number"] for k, v in team_metadata.items()}
    return lower_map.get(team_name.strip().lower())


DUPLICATE_SUM_FIELDS = ["Dues Paid", "Computed Wage", "Computed PC", "Computed MRF", "Total Due to National HQ"]


def merge_duplicate_players(records: list[dict]) -> list[dict]:
    """Airtable sometimes returns multiple matching records for the same player
    (same Full Name) within a team/month/year — e.g. a correction or a split
    payroll entry. Collapse those into one row per player, summing the
    dues/wage fields; other fields (address, base compensation, etc.) are
    taken from the first record encountered for that player."""
    merged: dict[str, dict] = {}
    order: list[str] = []
    for rec in records:
        fields = rec.get("fields", {})
        key = fields.get("Full Name") or rec.get("id")
        if key not in merged:
            merged[key] = {"fields": dict(fields)}
            order.append(key)
        else:
            merged_fields = merged[key]["fields"]
            for f in DUPLICATE_SUM_FIELDS:
                merged_fields[f] = (merged_fields.get(f) or 0) + (fields.get(f) or 0)
    return [merged[key] for key in order]


def compute_team_totals(records: list[dict]) -> tuple[Decimal, Decimal]:
    total_dues = Decimal("0")
    total_hq = Decimal("0")
    for rec in records:
        fields = rec.get("fields", {})
        dues = fields.get("Dues Paid")
        hq = fields.get("Total Due to National HQ")
        if dues is not None:
            total_dues += Decimal(str(dues))
        if hq is not None:
            total_hq += Decimal(str(hq))
    return total_dues, total_hq


FILENAME_ILLEGAL_CHARS = re.compile(r'[\/\\:*?"<>|]')


def build_filename(team_name: str, pu_number, month_name: str, year: int) -> str:
    if pu_number is None or pu_number == "":
        pu_display = "PU# unknown"
    else:
        if isinstance(pu_number, float) and pu_number.is_integer():
            pu_number = int(pu_number)
        pu_display = f"PU# {pu_number}"
    raw = f"({pu_display}) {team_name} _ {month_name}_{year}.xlsx"
    return FILENAME_ILLEGAL_CHARS.sub("_", raw)


def build_report_workbook(template_path, records, total_dues, total_hq):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    row_idx = 2
    for rec in records:
        fields = rec.get("fields", {})
        for col, field_name in COLUMN_FIELD_MAP.items():
            cell = ws[f"{col}{row_idx}"]
            value = fields.get(field_name)
            if col == "O" and value:
                # Airtable date fields come back as ISO 8601 strings.
                try:
                    value = datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
                except (ValueError, AttributeError):
                    pass
            cell.value = value
            if col in NUMBER_FORMAT_COLUMNS:
                cell.number_format = NUMBER_FORMAT_COLUMNS[col]
        row_idx += 1

    last_data_row = row_idx - 1
    ws["P2"] = f"${total_dues:,.2f}"
    ws["Q2"] = f"${total_hq:,.2f}"

    if ws.max_row > last_data_row:
        ws.delete_rows(last_data_row + 1, ws.max_row - last_data_row)

    return wb


def process_team_group(team_name, records, dry_run, template, pu_number, month_name, year, output_dir) -> dict:
    merged_records = merge_duplicate_players(records)
    total_dues, total_hq = compute_team_totals(merged_records)
    result = {
        "team": team_name,
        "matched": len(records),
        "merged": len(merged_records),
        "total_dues": total_dues,
        "total_hq": total_hq,
        "output_path": None,
    }
    if not dry_run:
        wb = build_report_workbook(template, merged_records, total_dues, total_hq)
        output_dir.mkdir(parents=True, exist_ok=True)
        filename = build_filename(team_name, pu_number, month_name, year)
        output_path = output_dir / filename
        wb.save(output_path)
        result["output_path"] = output_path
    return result


def print_team_summary(result: dict, month_name: str, year: int) -> None:
    print(f"Matched {result['matched']} record(s) for {result['team']} — {month_name} {year}")
    duplicate_count = result["matched"] - result["merged"]
    if duplicate_count:
        print(f"Merged {duplicate_count} duplicate record(s) by Full Name into {result['merged']} player row(s).")
    print(f"Team Total Dues Paid: ${result['total_dues']:,.2f}")
    print(f"Team Total Due to National HQ: ${result['total_hq']:,.2f}")
    if result["output_path"] is None:
        print("(dry run — no file written)")
    else:
        print(f"Report written to {result['output_path']}")


def print_batch_line(result: dict) -> None:
    if result["matched"] == 0:
        print(f"{result['team']}: no matching records — skipped")
        return
    out = result["output_path"] if result["output_path"] is not None else "(dry run)"
    print(
        f"{result['team']}: {result['matched']} matched -> {result['merged']} merged | "
        f"Dues ${result['total_dues']:,.2f} | HQ ${result['total_hq']:,.2f} | {out}"
    )


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate a monthly USL team dues report from Airtable as an .xlsx file."
    )
    parser.add_argument(
        "--team", required=True,
        help='Exact Team Name as stored in Airtable (case-sensitive), or "all" to generate one report '
             "per team in the Teams table with matching payroll data for the given month/year.",
    )
    parser.add_argument("--month", required=True, help='Month as a number (1-12) or name (e.g. "July").')
    parser.add_argument("--year", required=True, type=int, help="Four-digit year, e.g. 2026.")
    parser.add_argument("--output-dir", default="reports", help="Directory to write the report into (default: reports/).")
    parser.add_argument("--template", default="report_sample.xlsx", help="Path to the template .xlsx to base formatting on.")
    parser.add_argument("--skip-schema-check", action="store_true", help="Skip the one-time Airtable schema validation call.")
    parser.add_argument("--dry-run", action="store_true", help="Query Airtable and print a summary without writing a file.")
    parser.add_argument(
        "--month-filter-override", default=None,
        help="Raw filterByFormula clause to use for the month condition instead of the auto-derived one.",
    )
    parser.add_argument(
        "--email", action="store_true",
        help=f'Email the generated report(s) via Gmail to --to (default {DEFAULT_RECIPIENT_EMAIL}) '
             f'with subject "{EMAIL_SUBJECT}". Requires credentials.json in the project root '
             f"(see workflow doc for setup). Has no effect with --dry-run (nothing to attach).",
    )
    parser.add_argument(
        "--to", default=DEFAULT_RECIPIENT_EMAIL,
        help="Email recipient(s), comma-separated (only used with --email).",
    )
    parser.add_argument(
        "--cc", default=None,
        help="CC recipient(s), comma-separated (only used with --email).",
    )
    return parser


def _normalize_emails(value) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [e.strip() for e in value.split(",") if e.strip()]
    return [str(e).strip() for e in value if str(e).strip()]


def run_report(
    team, month, year, *,
    dry_run=False, email=False, to=DEFAULT_RECIPIENT_EMAIL, cc=None,
    output_dir="reports", template="report_sample.xlsx",
    skip_schema_check=False, month_filter_override=None,
) -> dict:
    """Core orchestration shared by the CLI and the HTTP handler. Raises
    ConfigError/AuthError/ValueError/SchemaMismatchError/RateLimitError/
    AirtableAPIError/ImportError on hard failures. "No matching records" is
    NOT raised — it's a normal outcome reflected via the returned dict's
    "no_data" flag, since callers (CLI exit code, HTTP status) decide how to
    present it. Prints progress/summary lines exactly as the CLI always has;
    harmless when called from a server context (becomes log output)."""
    cfg = load_config()
    month_num, month_name = normalize_month(month)
    is_batch = team.strip().lower() == "all"
    run_timestamp = datetime.now().strftime("%Y%m%d-%H%M")

    field_types = {}
    if not skip_schema_check:
        schema = get_table_schema(cfg)
        validate_expected_fields(schema, EXPECTED_RECORD_FIELDS + FILTER_FIELDS)
        field_types = {name: info["type"] for name, info in schema.items()}

    team_metadata = get_team_metadata(cfg)
    team_names = sorted(team_metadata.keys()) if is_batch else [team]

    out_dir = (
        Path(output_dir) / f"ALL_{year}-{month_num:02d}_{run_timestamp}"
        if is_batch else Path(output_dir)
    )

    results = []
    for t in team_names:
        formula = build_filter_formula(
            t, month_num, month_name, year, field_types, override=month_filter_override,
        )
        records = fetch_all_records(cfg, formula)
        if not records:
            result = {
                "team": t, "matched": 0, "merged": 0,
                "total_dues": Decimal("0"), "total_hq": Decimal("0"), "output_path": None,
            }
        else:
            pu_number = lookup_pu_number(team_metadata, t)
            result = process_team_group(t, records, dry_run, template, pu_number, month_name, year, out_dir)
        results.append(result)
        if is_batch:
            print_batch_line(result)

    no_data = all(r["matched"] == 0 for r in results)

    if not no_data:
        if not is_batch:
            print_team_summary(results[0], month_name, year)
        else:
            files_written = sum(1 for r in results if r["output_path"] is not None)
            skipped = sum(1 for r in results if r["matched"] == 0)
            dry_note = " (dry run)" if dry_run else ""
            print(
                f"\nBatch complete{dry_note}: {len(results)} team(s) checked, "
                f"{files_written} file(s) written, {skipped} skipped (no data)."
            )

    emailed = False
    email_error = None
    if email and not no_data:
        attachment_paths = [r["output_path"] for r in results if r["output_path"] is not None]
        to_list = _normalize_emails(to)
        cc_list = _normalize_emails(cc)
        if not attachment_paths:
            email_error = "no files were written (dry_run or nothing generated)"
            print(f"Skipping email: {email_error}.", file=sys.stderr)
        elif not to_list:
            email_error = "no recipient email address(es) provided"
            print(f"Skipping email: {email_error}.", file=sys.stderr)
        else:
            import gmail_client
            period = f"{month_name} {year}"
            body = (
                f"Attached: {len(attachment_paths)} team dues report(s) for {period}."
                if is_batch else f"Attached: {results[0]['team']} dues report for {period}."
            )
            try:
                gmail_client.send_email_with_attachments(
                    to_list, EMAIL_SUBJECT, body, attachment_paths, cc=cc_list or None,
                )
                emailed = True
                cc_note = f" (cc: {', '.join(cc_list)})" if cc_list else ""
                print(f"Emailed {len(attachment_paths)} report(s) to {', '.join(to_list)}{cc_note}.")
            except (gmail_client.GmailAuthError, gmail_client.GmailSendError) as e:
                email_error = str(e)
                print(f"Error: {e}", file=sys.stderr)

    return {
        "is_batch": is_batch,
        "month_num": month_num,
        "month_name": month_name,
        "year": year,
        "no_data": no_data,
        "results": results,
        "emailed": emailed,
        "email_error": email_error,
    }


def main():
    args = build_arg_parser().parse_args()

    try:
        result = run_report(
            args.team, args.month, args.year,
            dry_run=args.dry_run, email=args.email, to=args.to, cc=args.cc,
            output_dir=args.output_dir, template=args.template,
            skip_schema_check=args.skip_schema_check,
            month_filter_override=args.month_filter_override,
        )
    except (ConfigError, AuthError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except SchemaMismatchError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(3)
    except (RateLimitError, AirtableAPIError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(4)
    except ImportError as e:
        print(f"Error: Gmail dependencies not installed ({e}). Run: pip install -r requirements.txt", file=sys.stderr)
        sys.exit(5)

    if result["no_data"]:
        if result["is_batch"]:
            print(
                f"No records found for any team, month={result['month_name']} ({result['month_num']}), "
                f"year={result['year']}. Verify the period is correct.",
                file=sys.stderr,
            )
        else:
            print(
                f"No records found for team={args.team!r}, month={result['month_name']} ({result['month_num']}), "
                f"year={result['year']}. Verify the team name matches Airtable's 'Team Name' field exactly "
                f"(case-sensitive) and that data exists for this period.",
                file=sys.stderr,
            )
        sys.exit(2)

    if args.email and result["email_error"]:
        sys.exit(5)


if __name__ == "__main__":
    main()
