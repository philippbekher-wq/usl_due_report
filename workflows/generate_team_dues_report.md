# Workflow: Generate Team Dues Report

## Objective
Produce a monthly dues report, as an `.xlsx` file, for a single USL team — listing each player's dues/wage figures for that month, sourced live from Airtable. Used by the players association to reconcile what a team owes National HQ.

## Required Inputs
- **Team name** — must match Airtable's `Team Name` field exactly (case-sensitive, plain text match). Or the literal word **"all"** (case-insensitive) to generate one report per team listed in the Teams table (`tblxtS2ioQsKvjcZ3`), for the given month/year, in a single run. No eligibility filtering is applied — every team in that table is attempted; teams without matching payroll data for the period are simply skipped (see Edge Cases).
- **Month** — a number (1–12) or a name (e.g. "July").
- **Year** — four digits, e.g. 2026.

If any of these are missing or ambiguous, ask the user before running the tool.

## Record Matching Criteria
Beyond team/month/year, a record only qualifies for the report if:
- `Dues Paid` is **not blank** — a blank value means the player hasn't been processed for that payroll period yet, so they're excluded entirely (not shown as a $0 row).
- `Player Invoiced` is **unchecked** — checked means the player has already been invoiced separately and shouldn't appear again in this report.

Both conditions are applied server-side in the Airtable query (not just at the totals-computation step), so excluded records never appear as blank rows in the output.

**Duplicate players**: after fetching, records are grouped by `Full Name`. If a player has more than one matching record (e.g. a correction or split payroll entry), they're collapsed into a single output row, summing `Dues Paid`, `Computed Wage`, `Computed PC`, `Computed MRF`, and `Total Due to National HQ` across the duplicates. Other fields (Unique ID, address, base compensation, etc.) are taken from the first record encountered for that player.

## One-Time Setup
1. `pip install -r requirements.txt`
2. Copy `.env.example` to `.env` and fill in `AIRTABLE_API_KEY` with a Personal Access Token that has read access to base `appiVO0LP0Xm603VO`. Never put the token anywhere else.
3. **Only needed for `--email`**: create a Google Cloud OAuth client (Gmail API enabled, OAuth consent screen configured, OAuth client ID of type "Desktop app"), download it, and save it as `credentials.json` in the project root. The first run with `--email` opens a browser for one-time consent, then caches `token.json` (both gitignored) so future runs don't prompt again.

## Tool to Use
`tools/generate_dues_report.py`

```
python tools/generate_dues_report.py --team "Chicago Fire" --month July --year 2026
```

Run with `--dry-run` first when trying a new team/month/year combination for the first time — it queries Airtable and prints the matched record count and computed totals without writing a file, which is a cheap way to sanity-check spelling and data availability before generating the real report.

To generate a report for every team at once:
```
python tools/generate_dues_report.py --team all --month July --year 2026
```
This fetches every team name from the Teams table, then runs the exact same single-team query/report logic once per team (not a grouped/aggregated query) — so batch-mode numbers for any given team always match what a single-team run for that same team would produce.

To also email the generated report(s), add `--email`:
```
python tools/generate_dues_report.py --team "Chicago Fire" --month July --year 2026 --email
python tools/generate_dues_report.py --team all --month July --year 2026 --email
```
Sends via Gmail (OAuth, see One-Time Setup) to `philipp@lowcoding.dev` by default (override with `--to`), subject `"Report"`, with every generated `.xlsx` attached to a single email — for `--team all` that means one email with all of that run's team reports attached, not one email per team. Has no effect combined with `--dry-run` (nothing was written to attach) — the tool prints a warning and skips emailing rather than failing.

## Expected Output
Single-team: `reports/<Team>_<YYYY>-<MM>_<timestamp>.xlsx`, e.g. `reports/Chicago_Fire_2026-07_20260708-1420.xlsx`.

Batch (`--team all`): `reports/ALL_<YYYY>-<MM>_<timestamp>/<Team>_<YYYY>-<MM>_<timestamp>.xlsx` — one file per team with matching data, all sharing the same run timestamp in a dedicated subfolder so a batch is identifiable and movable as one unit.

Layout matches `report_sample.xlsx` exactly: one row per player (columns A–O, pulled directly from Airtable fields), with `Team Total Dues Paid` and `Team Total Due to National HQ` (columns P/Q) computed by the tool and written only into row 2, as currency-formatted text (e.g. `"$1,366.30"`).

**Note on delivery**: this report is written locally rather than uploaded to a cloud service, as an intentional exception to this project's usual "deliverables live in cloud services" convention — the user retrieves the file manually.

## Edge Cases & Known Behaviors
- **Zero matching records**: no file is written; the tool prints a message and exits with code 2. Double-check the team name spelling/case and that Airtable has data for that period.
- **Missing `Unique ID` on a player record**: expected — some real records have this blank. Passed through as a blank cell, not an error.
- **Missing `Dues Paid`**: the record is excluded from the report entirely (see Record Matching Criteria above), not shown as a $0 row.
- **`Player Invoiced` checked**: the record is excluded from the report entirely.
- **Missing `Total Due to National HQ` on an otherwise-matching record**: treated as $0 toward the team total, no warning.
- **Re-running for the same team/month/year**: does not overwrite — each run produces a new timestamped file in `reports/`.
- **Airtable rate limiting (429)**: the tool automatically waits ~30s and retries, up to 3 times.
- **Missing/renamed Airtable field**: the tool validates the live schema before querying and fails loudly (exit code 3), naming exactly which expected field is missing and listing the base's actual field names — it will not silently produce blank columns.
- **Duplicate records for the same player**: merged into one row with summed dues/wage fields (see Record Matching Criteria above). The tool prints how many duplicates were merged.
- **`Payroll Date Month` / `Payroll Year` filter values**: the tool auto-detects the Airtable field type (number vs. text vs. date) and builds the filter accordingly. If a report unexpectedly returns zero records for a team/period known to have data, use `--month-filter-override` to pass a raw `filterByFormula` clause by hand, then update the Learnings section below with what the correct format turned out to be.
- **`--team all` and junk/test teams**: the Teams table contains obvious test entries (e.g. "Testing team 3", "Test Team") and no eligibility filtering is applied. These simply produce zero matching payroll records for any real period and are skipped like any other team with no data that month — not specially excluded.
- **`--team all` and Team Name mismatches**: a team's exact name in the Teams table (`Name` field) must match its `Team Name` value in the payroll table for that team to get a report. A mismatch (typo, casing, extra whitespace) silently produces a skip, not an error.
- **`--team all` with zero data everywhere**: if every team in the Teams table comes back with zero payroll matches for the given month/year (almost certainly a wrong period), no report is written and the tool exits with code 2 — same signal as the single-team zero-match case.
- **A real team literally named "All"**: cannot be targeted individually via `--team all` — that value always triggers batch mode instead. Not currently an issue in the live data, but worth knowing if a team is ever renamed.
- **`--email` without `credentials.json`**: fails clearly (exit code 5) naming the missing file and what to do — report generation itself still succeeds and the local file is still written; only the email step is skipped/failed.
- **`--email` with `--dry-run`**: prints a warning and skips emailing rather than erroring — there's nothing generated to attach.
- **`--email` in batch mode**: all generated reports from that run are attached to a single email (not one email per team). If the total attachment size ever exceeds ~20MB (unlikely at current report sizes — each is a few KB), the tool raises a clear error rather than silently truncating or failing partway through.
- **Gmail token expiry**: `token.json` auto-refreshes via its stored refresh token on each run; if Google ever revokes it (e.g. consent revoked, or the OAuth app is still in "Testing" mode and the token ages out), delete `token.json` and re-run with `--email` to redo the one-time browser consent.

## Learnings
Confirmed against the live base on 2026-07-08:
- `Payroll Date Month` and `Payroll Year` are both Airtable `number` fields, so the filter uses plain numeric equality (`{Payroll Date Month} = 3`) — no override needed in practice.
- The report's "Sending Date" column has no literal matching field in Airtable. It's mapped to the built-in `Created` field (`createdTime` type) instead. If a future report needs to reflect a different date (e.g. actual dues-mailing date), revisit this mapping.
- `Player` is a linked-record field (returns Airtable record IDs, not text) — using it directly crashes the xlsx writer. Both the "Player" and "Full Name" output columns are sourced from the `Full Name` field instead, since the sample data always shows those two columns identical.
- `Street Address Line 1/2`, `City`, `State`, `Zip` are frequently blank in real records — expected, not a bug.
- `Player Invoiced` is a `checkbox` field. Discovered via a real test run (Brooklyn FC (Championship), May 2026) that roughly half of a team's matching records can have a blank `Dues Paid` — these are players not yet processed for that payroll period and must be excluded from the query (see Record Matching Criteria), not included as blank/zero rows.
- Duplicate records per player are real, not hypothetical: the same May 2026 Brooklyn FC (Championship) test returned 20 matching records that merged down to 10 unique players. The merged output (post Dues Paid/Player Invoiced filtering and duplicate-summing) matched `report_sample.xlsx`'s original values exactly, confirming the sample itself was built from this same team/period with duplicates already combined.
- First real `--team all` run (May 2026): 66 teams in the Teams table, 37 produced a report, 29 skipped for no matching payroll data — including all 5 junk/test entries ("Test Team", "Testing Team", "Testing team 2", "Testing team 3", "New Test Team"), confirming no special-case exclusion logic was needed for them. Batch-mode totals for Brooklyn FC (Championship) matched the already-verified single-team run exactly ($1,366.30 / $406.06), confirming per-team queries inside the batch loop behave identically to standalone single-team invocations. Runtime for all 66 sequential per-team queries was ~60 seconds with no rate-limit (429) errors encountered — no additional throttling needed at this scale.
