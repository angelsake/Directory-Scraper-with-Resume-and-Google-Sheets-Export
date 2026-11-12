"""
Directory Scraper with Resume & Google Sheets Export
----------------------------------------------------
Scrapes a paginated, searchable table (several search terms = "categories"),
combines and de-duplicates the results, saves a checkpoint after every page so
an interrupted run can be resumed, and exports to Excel and (optionally)
Google Sheets. Failures are reported, never hidden.

Practice target (a site built for learning scraping):
    https://www.scrapethissite.com/pages/forms/   (hockey teams, with search)

Usage:
    pip install -r requirements.txt
    python scraper.py                                  # all teams, Excel report
    python scraper.py --terms new,san,los              # three search terms
    python scraper.py --resume                         # continue an interrupted run
    python scraper.py --sheet-id YOUR_SHEET_ID         # also write to Google Sheets
"""

import argparse
import json
import os
import re
import time
from datetime import datetime, timezone
from urllib.parse import urlencode

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_URL = "https://www.scrapethissite.com/pages/forms/"
PAGE_PARAM, QUERY_PARAM, PER_PAGE_PARAM = "page_num", "q", "per_page"  # verify on the live site
OUTPUT_FILE = "teams_report.xlsx"
CHECKPOINT_FILE = "checkpoint.json"
MAX_RETRIES = 3          # bounded retries per page
BACKOFF_SECONDS = 1      # wait grows with each retry
DELAY_SECONDS = 0.5      # be polite between pages
META_COLUMNS = ["search_term", "page_number", "source_url", "scraped_at"]


# --------------------------------------------------------------------------
# Fetching and parsing
# --------------------------------------------------------------------------
def make_getter(timeout=15):
    """Return get(url, params) -> html, using one reusable HTTP session."""
    import requests

    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 (portfolio project)"})

    def get(url, params):
        response = session.get(url, params=params, timeout=timeout)
        response.raise_for_status()
        response.encoding = "utf-8"
        return response.text

    return get


def normalize_header(text):
    """'Win %' -> 'win_pct', 'Goals For (GF)' -> 'goals_for_gf', '+ / -' -> 'diff'."""
    text = text.strip().lower().replace("%", " pct ")
    return re.sub(r"[^a-z0-9]+", "_", text).strip("_") or "diff"


def parse_table(html):
    """
    Read the data table using its header names (not fixed CSS classes).
    Returns a list of dicts. An empty list means 'no data rows on this page'.
    Raises ValueError if the table is malformed, so it is reported, not hidden.
    """
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "html.parser")
    table = soup.select_one("table.table") or soup.find("table")
    if table is None:
        return []

    headers = [normalize_header(th.get_text(" ", strip=True)) for th in table.select("tr th")]
    rows = []
    for tr in table.select("tr"):
        cells = [td.get_text(" ", strip=True) for td in tr.find_all("td")]
        if not cells:
            continue
        if not headers:
            raise ValueError("table has data rows but no header row")
        if len(cells) != len(headers):
            raise ValueError(f"row has {len(cells)} cells but the table has {len(headers)} headers")
        rows.append(dict(zip(headers, cells)))
    return rows


def scrape_page(get, base_url, term, page, per_page):
    params = {PAGE_PARAM: page, PER_PAGE_PARAM: per_page}
    if term:
        params[QUERY_PARAM] = term
    html = get(base_url, params)
    return parse_table(html), f"{base_url}?{urlencode(params)}"


def fetch_with_retries(get, base_url, term, page, per_page, retries=MAX_RETRIES):
    """Returns (result, error_message). result is None if every attempt failed."""
    last_error = ""
    for attempt in range(1, retries + 1):
        try:
            return scrape_page(get, base_url, term, page, per_page), ""
        except Exception as error:  # network errors, HTTP errors, malformed tables
            last_error = f"{type(error).__name__}: {str(error).splitlines()[0]}"
            print(f"  Attempt {attempt}/{retries} failed: {last_error}")
            time.sleep(BACKOFF_SECONDS * attempt)
    return None, last_error


# --------------------------------------------------------------------------
# Checkpoints
# --------------------------------------------------------------------------
def load_checkpoint(path, config):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            state = json.load(f)
        if state.get("config") == config:
            return state
        print("Checkpoint belongs to a different configuration, starting fresh.")
    return None


def save_checkpoint(path, state):
    """Write atomically, so a crash mid-write cannot corrupt the checkpoint."""
    temp = path + ".tmp"
    with open(temp, "w", encoding="utf-8") as f:
        json.dump(state, f)
    os.replace(temp, path)


# --------------------------------------------------------------------------
# Crawl
# --------------------------------------------------------------------------
def scrape_all(get, base_url, terms, max_pages=100, per_page=100,
               checkpoint_path=CHECKPOINT_FILE, resume=False):
    """
    Scrape every search term page by page. A checkpoint is saved after each page.
    One term failing does not stop the others, and the final status is honest.
    """
    config = {"base_url": base_url, "terms": terms, "per_page": per_page}
    state = load_checkpoint(checkpoint_path, config) if resume else None
    if state:
        print(f"Resuming: {len(state['records'])} records already saved.")
    else:
        state = {"config": config, "records": [], "failures": [], "terms": {}, "pages_ok": 0}

    for term in terms:
        label = term or "(all)"
        progress = state["terms"].setdefault(term, {"next_page": 1, "status": "pending", "last_first": None})
        if progress["status"] == "done":
            print(f"Term {label}: already complete, skipping.")
            continue

        state["failures"] = [f for f in state["failures"] if f["search_term"] != term]
        page = progress["next_page"]

        while True:
            if page > max_pages:
                progress["status"] = "limited"
                break
            print(f"Term {label}, page {page}")
            result, error = fetch_with_retries(get, base_url, term, page, per_page)

            if result is None:
                state["failures"].append(
                    {"search_term": label, "page_number": page, "error": error}
                )
                progress["status"] = "failed"
                break

            rows, url = result
            if not rows:
                progress["status"] = "done"
                break
            first_row = list(rows[0].values())
            if progress["last_first"] == first_row:
                progress["status"] = "stalled"   # page number had no effect
                break

            scraped_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
            for row in rows:
                state["records"].append(dict(row, search_term=term, page_number=page,
                                             source_url=url, scraped_at=scraped_at))
            state["pages_ok"] += 1
            progress.update(last_first=first_row, next_page=page + 1, status="running")
            save_checkpoint(checkpoint_path, state)
            page += 1
            time.sleep(DELAY_SECONDS)

        save_checkpoint(checkpoint_path, state)

    statuses = {term: state["terms"][term]["status"] for term in terms}
    bad = [t or "(all)" for t, s in statuses.items() if s in ("failed", "stalled")]
    if bad:
        status = f"PARTIAL - problems with: {', '.join(bad)} (see Failed Pages / check pagination)"
    elif "limited" in statuses.values():
        status = "LIMITED - stopped at --max-pages, more pages exist"
    else:
        status = "COMPLETE"

    return {"records": state["records"], "failures": state["failures"],
            "pages_ok": state["pages_ok"], "status": status, "term_status": statuses}


# --------------------------------------------------------------------------
# Cleaning and reports
# --------------------------------------------------------------------------
def clean(records):
    """Return (clean DataFrame, duplicates removed)."""
    df = pd.DataFrame(records)

    for col in df.columns:
        if col in META_COLUMNS:
            continue
        values = df[col].astype(str).str.strip().str.replace("\u2212", "-", regex=False)
        values = values.replace("", pd.NA)
        converted = pd.to_numeric(values, errors="coerce")
        df[col] = converted if converted.notna().sum() == values.notna().sum() else values

    key = [c for c in ("team_name", "year") if c in df.columns]
    key = key or [c for c in df.columns if c not in META_COLUMNS]
    before = len(df)
    df = df.drop_duplicates(subset=key).reset_index(drop=True)
    return df, before - len(df)


def build_reports(df, duplicates, result, terms, seconds, base_url):
    unique_key = "team_name" if "team_name" in df.columns else df.columns[0]
    summary = pd.DataFrame(
        [
            ("Completion status", result["status"]),
            ("Source", base_url),
            ("Search terms", ", ".join(t or "(all)" for t in terms)),
            ("Pages scraped successfully", result["pages_ok"]),
            ("Pages failed", len(result["failures"])),
            ("Rows collected", len(df)),
            ("Duplicates removed", duplicates),
            (f"Unique {unique_key}", df[unique_key].nunique()),
            ("Run time (seconds)", round(seconds, 1)),
        ],
        columns=["Metric", "Value"],
    )

    by_term = pd.DataFrame(
        [(t or "(all)", result["term_status"][t], int((df["search_term"] == t).sum())) for t in terms],
        columns=["search_term", "status", "rows_kept"],
    )
    failed = pd.DataFrame(result["failures"], columns=["search_term", "page_number", "error"])

    sheets = [("Summary", summary), ("Data", df.drop(columns=["search_term"]).assign(
        search_term=df["search_term"].replace("", "(all)"))), ("By Term", by_term),
        ("Failed Pages", failed)]

    needed = ["team_name", "year", "wins", "losses", "win_pct"]
    if all(c in df.columns for c in needed) and pd.api.types.is_numeric_dtype(df["win_pct"]):
        sheets.insert(3, ("Top Seasons", df.nlargest(10, "win_pct")[needed].reset_index(drop=True)))
    return sheets


# --------------------------------------------------------------------------
# Export
# --------------------------------------------------------------------------
def style_sheet(ws):
    fill = PatternFill("solid", start_color="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    for index, column in enumerate(ws.columns, start=1):
        longest = max(len(str(c.value)) if c.value is not None else 0 for c in column)
        ws.column_dimensions[get_column_letter(index)].width = min(max(longest + 2, 10), 70)
    ws.freeze_panes = "A2"


def export_excel(path, sheets):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, frame in sheets:
            frame.to_excel(writer, sheet_name=name, index=False)
            style_sheet(writer.sheets[name])


def to_sheet_values(df):
    """Header + rows as plain Python values (JSON-safe, blanks for missing)."""
    body = df.astype(object).where(df.notna(), "").values.tolist()
    return [list(df.columns)] + body


def open_google_sheet(sheet_id, credentials_path):
    import gspread

    client = gspread.service_account(filename=credentials_path)
    return client.open_by_key(sheet_id), gspread.exceptions.WorksheetNotFound


def export_google_sheet(spreadsheet, sheets, missing_error=Exception):
    """Write each sheet to a worksheet of the same name (created if missing, cleared first)."""
    for title, df in sheets:
        try:
            worksheet = spreadsheet.worksheet(title)
        except missing_error:
            worksheet = spreadsheet.add_worksheet(
                title=title, rows=max(len(df) + 1, 2), cols=max(len(df.columns), 1))
        worksheet.clear()
        worksheet.update(range_name="A1", values=to_sheet_values(df))


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------
def parse_args():
    p = argparse.ArgumentParser(description="Paginated directory scraper with resume and Google Sheets export.")
    p.add_argument("--base-url", default=BASE_URL)
    p.add_argument("--terms", default="", help="comma-separated search terms; empty = everything")
    p.add_argument("--output", default=OUTPUT_FILE, help="Excel file to create")
    p.add_argument("--max-pages", type=int, default=100, help="page limit per search term")
    p.add_argument("--per-page", type=int, default=100)
    p.add_argument("--resume", action="store_true", help="continue from the checkpoint file")
    p.add_argument("--checkpoint", default=CHECKPOINT_FILE)
    p.add_argument("--sheet-id", help="Google Sheet ID to write to (optional)")
    p.add_argument("--credentials", default="credentials.json", help="Google service account key file")
    return p.parse_args()


def main():
    args = parse_args()
    terms = list(dict.fromkeys(t.strip() for t in args.terms.split(",")))
    started = time.time()

    try:
        result = scrape_all(make_getter(), args.base_url, terms, args.max_pages,
                            args.per_page, args.checkpoint, args.resume)
    except KeyboardInterrupt:
        print(f"\nInterrupted. Progress is saved in {args.checkpoint}. "
              f"Continue with: python scraper.py --resume (same options)")
        return

    if not result["records"]:
        print(f"\nNo data collected. Status: {result['status']}")
        for failure in result["failures"]:
            print(f"  {failure['search_term']} page {failure['page_number']}: {failure['error']}")
        return

    df, duplicates = clean(result["records"])
    sheets = build_reports(df, duplicates, result, terms, time.time() - started, args.base_url)
    export_excel(args.output, sheets)
    print(f"\nStatus: {result['status']}")
    print(f"{len(df)} rows saved to {args.output}")

    if args.sheet_id:
        try:
            spreadsheet, missing = open_google_sheet(args.sheet_id, args.credentials)
            export_google_sheet(spreadsheet, sheets, missing)
            print("Google Sheet updated.")
        except Exception as error:
            print(f"Google Sheets export failed ({type(error).__name__}: {error}). "
                  f"The Excel file was still saved. See README for the Google setup steps.")


if __name__ == "__main__":
    main()
