# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

An ETL pipeline (`app.py`) that pulls the "Client Wise All Cases - Bridge" report for a specific client (Wipro, client code `3017`) from Authbridge's internal MIS Query Browser (`https://mis.authbridge.com/export_query/`), reshapes it into a client-ready workbook with derived columns, and builds a native-Excel pivot "Summary Sheet" on top of it — all via one script, with an optional Streamlit UI (`streamlit_app.py`) as a front end for the same pipeline.

There is no test suite in this repo. Dependencies are pinned in `requirements.txt`. Generated output artifacts (CSV/XLSX exports) are written to the project root; `.playwright-mcp/` is browser-automation debris from an earlier exploratory session, safe to ignore/delete. `Wipro Limited - Dashboard.xlsx` is a **reference-only** file (an earlier hand-built dashboard) used to model the Summary Sheet's pivot layout — the pipeline has no live link/connection to it, it was just read once to copy its structure.

## Running it

```
pip install -r requirements.txt
python app.py
```

or, for the Streamlit front end:

```
python -m streamlit run streamlit_app.py
```

`pywin32` is Windows-only (see `requirements.txt`'s platform marker) and requires Excel to be installed — it drives Excel via COM to build the pivot tables (see `build_summary_pivots` below). Everything else is pure Python and cross-platform.

Credentials are required via the `MIS_USERNAME`/`MIS_PASSWORD` environment variables (no hardcoded fallback — `app.py` raises `KeyError` if either is unset). A local `.env` file (gitignored, not committed) holds these for development; export its values into the shell before running, e.g. `set -a && source .env && set +a` on bash, or a tool like `python-dotenv` if you want it loaded automatically. `streamlit_app.py` doesn't add its own credential UI — it just reuses `app.py`'s resolution as-is.

## Architecture

`app.py` runs three stages sequentially from `__main__` (and `streamlit_app.py` calls the same three functions directly, with progress shown via `st.status`):

1. **`export_csv()`** — logs into the MIS Query Browser and downloads the report **via direct HTTP form POSTs**, not browser automation. It replicates two requests captured from the real UI:
   - `POST login.php` with `username`/`password`/`login=Login` (session-cookie auth via `requests.Session`, no CSRF token).
   - `POST process.php` (multipart) with the report selection encoded in `EXPORT_PAYLOAD` — `hostname`, `database`, `access_time`, `csv_query`, `client1`. These are opaque internal IDs/names from the site's `<select>` options (e.g. `hostname=4` / `database=checkpoint_live` mean "Bridge Live", `csv_query=263` means the "Client Wise All Cases - Bridge" report). There's no API to look these up — to target a different report or client, drive the UI once with browser devtools (or Playwright) open on the Network tab, submit the export form, and copy the resulting `process.php` POST body into `EXPORT_PAYLOAD`.
   - The response is a `.zip` containing the CSV; it's extracted in-memory and written to disk.

2. **`process_csv_to_xlsx(csv_path)`** — reads the exported CSV and:
   - Filters rows to the **current calendar month** by `received_date` (compared against `datetime.date.today()` at run time — this is dynamic, not a fixed date).
   - Drops the columns listed in `FLEX_FIELDS_TO_DROP` (`case_flex_field27`–`30`).
   - Renames the *remaining* columns, in order, to the business-friendly names in `NEW_HEADERS`. This mapping is purely **positional**: `len(NEW_HEADERS)` must exactly equal the source column count minus the dropped flex fields (currently 44 − 4 = 40). A count mismatch raises `RuntimeError`, but if the *source report's* column order ever changes without changing the count, columns will get silently mislabeled — verify against a fresh header row before trusting a schema change. Note `NEW_HEADERS` contains `"Value"` twice (columns M and O) — a known artifact of the positional mapping — so those two source columns can only be addressed unambiguously by position, not name (see `COL_VALUE_ENTITY` below).
   - Calls `add_formula_columns()` (via `openpyxl`) to append five derived columns after the 40 renamed ones — **AO–AS**:
     - `Hours` (AO) — `=IF(AN="","DQC Date time Blank",IF(AC="NA","I Verify Date and time Blank",AN-AC))`, formatted `[h]:mm:ss`. `AN` = `DQC_RELEASED_DATE`, `AC` = `I verify date & time`.
     - `New Hours - Days` (AP) — `=IF(AN="","",INT(AN-AC))`.
     - `New Bucket` (AQ) — buckets `Hours` into `ABOVE 24` / `BELOW 24` / `DQC Pending`, falling back to `Hours`'s own text (e.g. `"I Verify Date and time Blank"`) via `ISNUMBER(AO)` whenever `AO` isn't a plain number — this is what stops it erroring to `#VALUE!` when `Hours` itself returned a status message instead of a duration.
     - `Location` (AR) — **not a formula**; a static value computed in Python from `Account Name` (column R) via `LOCATION_MAP`/`normalize_location()`, since normalizing typos/abbreviations (`IND`, `Inda`, `MY`, blank → defaults to `India`, etc.) isn't cleanly expressible as an Excel lookup formula.
     - `Inflow Time` (AS) — buckets `API_Created_stamp` (AF) into an hour range like `"10 AM to 11 AM"`.
   - Also applies cosmetic formatting to the whole header row (bold white-on-blue `#1F6FC5`, frozen top row, autofilter, sized column widths).
   - Writes the result to a new `.xlsx` in the project root.

3. **`build_summary_pivots(xlsx_path)`** — drives Excel via COM (`pywin32`) to add a `Summary Sheet` tab with 9 native PivotTables, modeled on the equivalent tables in `Wipro Limited - Dashboard.xlsx` (reference only, per above). Renames the data sheet to `Raw data` (idempotent — looks it up by name first so re-running doesn't collide) and rebuilds `Summary Sheet` from scratch each call. `PIVOT_SPECS` defines each pivot's anchor column, title, row field(s) — addressed by **1-based column position**, not name, because of the duplicate `"Value"` header noted above — and an optional single-item filter (`SummaryPivot5` is isolated to `Digital`, `SummaryPivot9` to `GOOGLE`, matching the reference file). Each pivot gets a merged blue title banner, the built-in `PivotStyleMedium2` style, and a border. Attaches to Excel if the workbook is already open (leaves that instance running on exit) or opens a temporary invisible instance otherwise. Wraps the whole COM session in `pythoncom.CoInitialize()`/`CoUninitialize()` — required because the calling thread (e.g. Streamlit's ScriptRunner thread) may not already have COM initialized on it the way a plain script's main thread does; omitting this throws `CoInitialize has not been called`.

Both write steps (`export_csv`'s `_write_with_fallback_name` and `process_csv_to_xlsx`'s output loop) handle the target filename being locked (e.g. open in Excel) by falling back to an auto-incrementing `" (N)"` suffix instead of failing — mirroring how Windows/Chrome name duplicate downloads. `build_summary_pivots`, by contrast, edits the workbook live via COM and has no such fallback — if it opens the file itself it expects to fully own that handle.

## Streamlit front end (`streamlit_app.py`)

A thin UI over the same three functions: shows the last-generated report's timestamp with a download button if one exists, and a "Run report now" button that calls `export_csv()` → `process_csv_to_xlsx()` → `build_summary_pivots()` in sequence with progress streamed via `st.status`, then offers the fresh file for download. No credential fields — relies entirely on `app.py`'s existing env-var/hardcoded-default resolution. Since `build_summary_pivots` needs Excel via COM, this only works run locally on Windows with Excel installed, not on a non-Windows host.
