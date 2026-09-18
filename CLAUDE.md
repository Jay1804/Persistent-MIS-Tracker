# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo does

Automates a recurring MIS report for the AuthBridge internal query portal (`mis.authbridge.com/export_query`). It logs in with Playwright, pulls two separate exports for client Persistent Systems Limited (client code `7084`), joins them, and produces a formatted Excel workbook with a real (COM-created) Excel PivotTable summary. There is both a CLI entry point and a Streamlit UI on top of the same logic.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt
playwright install chromium   # one-time browser install for Playwright

# Run the full pipeline headed (visible browser), writes the report to disk
python app.py

# Run the Streamlit UI (headless browser under the hood)
streamlit run streamlit_app.py
```

There is no test suite, linter, or build step in this repo — it's a small automation script plus a thin UI layer.

## Architecture

**`app.py`** contains all the logic, structured as one linear pipeline (`generate_report()`) built from single-purpose functions:

1. `login(page)` — logs into the MIS portal (credentials from `.env`: `MIS_USERNAME`, `MIS_PASSWORD`).
2. `select_report(page)` / `select_tracker_report(page)` — drive the portal's cascading `Host → Database → Data Time Slab → Query` dropdowns for two different exports:
   - Main data: Host=`Bridge Live`, Data Time Slab=`Client Specific Tracker`, Query=`Persistent Systems Limited Cleint Code-7084` (note: typo in the portal's own query name — must match exactly).
   - Joining-date source: Data Time Slab=`Tracker - Query`, Query=`Client Wise All Cases - Bridge`, with a `Client Code` text field (`7084`).
3. `export_report(page)` — clicks the portal's Export button (`#run_query`) and captures the resulting `.csv.zip` download.
4. `unzip_report(zip_path)` — extracts the CSV and deletes the zip.
5. `build_doj_map(tracker_csv_path)` — the "Client Wise All Cases" export has no column literally named "Date of Joining"; it's hidden in generic `case_flex_field*` columns with inconsistent date formats (`5-May-2025`, `June 23,2025`, `Not Mentioned`, etc.). `parse_doj()` uses `dateutil` to normalize these. The join key is `case_ars_no` (↔ `Vendor ID` in the main export) — **not** employee code, since one employee can have multiple cases.
6. `build_report_workbook(csv_path, doj_map)` — loads the main CSV, renames columns to the business-facing names in `COLUMN_NAMES`, maps in the Date of Joining column, drops rows whose `Overall case status` is in `EXCLUDED_STATUSES`, and writes the `Data Tab` sheet via pandas/openpyxl.
7. `build_pivot_summary(report_path)` — reopens the saved workbook via `win32com` (Excel COM automation — **requires Excel installed on the host**, and only works on Windows). Applies header styling (bold + light grey), auto-fits/caps column widths with wrap text for long text columns, fixes row height (`ROW_HEIGHT`), and builds a **real, refreshable Excel PivotTable** (not a static computed table) on a new `Summary` sheet, counting on the `Vendor ID` field (chosen because it's never blank — counting on a field with blank cells silently undercounts in Excel pivots).
8. `generate_report(headless=...)` — orchestrates all of the above; returns the final report path. Shared by both the CLI (`main()`, headed browser) and the Streamlit app (headless browser).

**`streamlit_app.py`** is a thin UI: a button that calls `generate_report(headless=True)`, then a download button plus a preview of both the pivot summary and the Data Tab, read back from the generated `.xlsx` via pandas.

### Key gotchas for future changes

- **Windows + Excel only.** `build_pivot_summary` uses `win32com.client` to drive a real Excel instance; this pipeline cannot run on Linux/macOS or a machine without Excel installed.
- **File locking.** If the output workbook (`Persistent MIS Tracker Report.xlsx`) is open in Excel, `generate_report()` fails with a `PermissionError` — close it first.
- **Session timeouts.** The portal's login session can expire quickly; if automation fails partway through, re-login is usually the fix.
- **Portal query names are exact strings**, including at least one typo (`"Cleint"` instead of `"Client"`) — copy them verbatim rather than "fixing" the spelling.
- **Pivot count field must be a never-blank column** (currently `Vendor ID`). If you change what's counted, check for blank cells first (Excel's pivot `Count` skips blanks and silently undercounts).
