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


def fetch_all_records(cfg: Config, formula: str) -> list[dict]:
    url = f"https://api.airtable.com/v0/{cfg.base_id}/{cfg.table_id}"
    records = []
    offset = None
    page_count = 0
    while True:
        params = {"filterByFormula": formula, "pageSize": 100}
        if offset:
            params["offset"] = offset
        resp = _request_with_retry(cfg, url, params)
        data = resp.json()
        records.extend(data.get("records", []))
        offset = data.get("offset")
        page_count += 1
        if not offset:
            break
        if page_count >= MAX_PAGES:
            raise AirtableAPIError(f"Pagination exceeded {MAX_PAGES} pages; aborting as a safety guard.")
    return records


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


def sanitize_filename(team: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", team).strip("_")


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


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate a monthly USL team dues report from Airtable as an .xlsx file."
    )
    parser.add_argument("--team", required=True, help="Exact Team Name as stored in Airtable (case-sensitive).")
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
    return parser


def main():
    args = build_arg_parser().parse_args()

    try:
        cfg = load_config()
    except ConfigError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        month_num, month_name = normalize_month(args.month)
    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        field_types = {}
        if not args.skip_schema_check:
            schema = get_table_schema(cfg)
            validate_expected_fields(schema, EXPECTED_RECORD_FIELDS + FILTER_FIELDS)
            field_types = {name: info["type"] for name, info in schema.items()}

        formula = build_filter_formula(
            args.team, month_num, month_name, args.year, field_types,
            override=args.month_filter_override,
        )
        records = fetch_all_records(cfg, formula)
    except AuthError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except SchemaMismatchError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(3)
    except RateLimitError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(4)
    except AirtableAPIError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(4)

    if not records:
        print(
            f"No records found for team={args.team!r}, month={month_name} ({month_num}), year={args.year}. "
            f"Verify the team name matches Airtable's 'Team Name' field exactly (case-sensitive) "
            f"and that data exists for this period.",
            file=sys.stderr,
        )
        sys.exit(2)

    merged_records = merge_duplicate_players(records)
    total_dues, total_hq = compute_team_totals(merged_records)

    print(f"Matched {len(records)} record(s) for {args.team} — {month_name} {args.year}")
    duplicate_count = len(records) - len(merged_records)
    if duplicate_count:
        print(f"Merged {duplicate_count} duplicate record(s) by Full Name into {len(merged_records)} player row(s).")
    print(f"Team Total Dues Paid: ${total_dues:,.2f}")
    print(f"Team Total Due to National HQ: ${total_hq:,.2f}")

    if args.dry_run:
        print("(dry run — no file written)")
        return

    wb = build_report_workbook(args.template, merged_records, total_dues, total_hq)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M")
    filename = f"{sanitize_filename(args.team)}_{args.year}-{month_num:02d}_{timestamp}.xlsx"
    output_path = output_dir / filename
    wb.save(output_path)
    print(f"Report written to {output_path}")


if __name__ == "__main__":
    main()
