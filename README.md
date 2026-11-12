# Directory Scraper with Resume and Google Sheets Export

Scrapes a paginated, searchable table across several search terms, combines and de-duplicates the results, and exports to **Excel** and (optionally) **Google Sheets**. If a run is interrupted, it can be **resumed** without losing or duplicating work. Problems are reported, never hidden.

Practice target (a site built for learning scraping): https://www.scrapethissite.com/pages/forms/

## What it does

1. Scrapes each search term page by page, reading the table by its **header names** (so it adapts to different columns)
2. Saves a **checkpoint after every page** (written atomically, so a crash can't corrupt it)
3. Retries failed pages a limited number of times, then records the failure and carries on with the other terms
4. Cleans the data with pandas (numbers converted, blanks kept as missing, duplicates removed across terms)
5. Exports these sheets to `teams_report.xlsx` and, if you provide a sheet ID, to Google Sheets:
   - **Summary**: completion status, pages scraped/failed, rows, duplicates, run time
   - **Data**: every row, tagged with its search term, page number, source URL and capture time
   - **By Term**: status and row count for each search term
   - **Top Seasons**: the 10 best win percentages
   - **Failed Pages**: every failed page with the reason

## Honest status

The Summary never reports a partial run as complete:

- `COMPLETE`: every term finished
- `PARTIAL`: a page failed after all retries, or pagination stopped advancing (for example, the site ignored the page number)
- `LIMITED`: stopped at `--max-pages` while more pages exist

## Setup and run

```bash
pip install -r requirements.txt

python scraper.py                          # everything, Excel report
python scraper.py --terms new,san,los      # three search terms
python scraper.py --max-pages 3            # quick test
python scraper.py --resume                 # continue an interrupted run
```

If a run is interrupted (Ctrl+C, power cut, lost connection), run the same command again with `--resume`. Pages that were already finished are not fetched again.

## Google Sheets setup (one time, optional)

1. In the [Google Cloud Console](https://console.cloud.google.com/), create a project.
2. Enable the **Google Sheets API** and the **Google Drive API** for it.
3. Create a **service account** and download its JSON key. Save it as `credentials.json` in this folder.
4. Create a Google Sheet, then **share it with the service account's email** (the `client_email` inside the JSON file) as **Editor**.
5. Copy the sheet ID from its URL (`docs.google.com/spreadsheets/d/<SHEET_ID>/edit`).
6. Run:

```bash
python scraper.py --sheet-id <SHEET_ID>
```

Never upload `credentials.json` to GitHub. The included `.gitignore` excludes it. If the Google export fails, the Excel file is still saved and the error is printed.

## Tests

```bash
python test_scraper.py
```

12 tests cover parsing, a multi-term crawl, retries, bounded retries, a failing term not stopping the others, `LIMITED` and stalled-pagination labelling, interrupt-and-resume (the resumed result matches an uninterrupted run and finished pages are not re-fetched), cleaning, Excel export and Google Sheets export. They use a fake website and a fake Google Sheet, so they need no internet or credentials. A full run against the live site is done with the commands above.

## Adapting it to another directory

Change `BASE_URL` and the three parameter names at the top of `scraper.py` (`PAGE_PARAM`, `QUERY_PARAM`, `PER_PAGE_PARAM`). The table parser, retries, checkpoints, status and exports stay the same. Always check a site's terms of use and `robots.txt` before scraping it.

## Tech used

Python, requests, BeautifulSoup, pandas, openpyxl, gspread

## Need data collected into a spreadsheet?

Contact me: *add your Fiverr/Upwork/email link here*
