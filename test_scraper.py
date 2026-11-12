"""
Tests for parsing, crawling, retries, failure labelling, resume, cleaning,
Excel export and Google Sheets export. They use a fake website and a fake
Google Sheet, so no internet connection or credentials are needed.

Run:  python test_scraper.py      (or: pytest)
"""

import json
import os
import tempfile

import pandas as pd

import scraper

scraper.BACKOFF_SECONDS = 0
scraper.DELAY_SECONDS = 0

BASE = "https://example.test/forms/"
HEADERS = ["Team Name", "Year", "Wins", "Losses", "OT Losses", "Win %",
           "Goals For (GF)", "Goals Against (GA)", "+ / -"]


def make_team(term, i):
    wins = 20 + i % 15
    return [f"{term.title() or 'All'} Team {i}", str(1990 + i % 20), str(wins), str(40 - wins), "",
            f"{wins / 80:.3f}", str(250 + i), str(240 + i), f"{'+' if i % 2 else '-'}{i % 9}"]


def table_html(rows):
    head = "".join(f"<th>{h}</th>" for h in HEADERS)
    body = "".join("<tr class='team'>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>" for r in rows)
    return f"<table class='table'><tr>{head}</tr>{body}</table>"


class FakeSite:
    """Serves paginated search results like the real directory."""

    def __init__(self, counts, flaky=None, broken=(), ignore_page=False, interrupt_after=None):
        self.teams = {t: [make_team(t, i) for i in range(n)] for t, n in counts.items()}
        self.flaky = dict(flaky or {})     # {(term, page): failures before success}
        self.broken = set(broken)          # {(term, page)} that always fail
        self.ignore_page = ignore_page
        self.interrupt_after = interrupt_after
        self.calls = []

    def __call__(self, url, params):
        term, page, per_page = params.get("q", ""), params["page_num"], params["per_page"]
        if self.interrupt_after is not None and len(self.calls) >= self.interrupt_after:
            raise KeyboardInterrupt
        self.calls.append((term, page))
        if (term, page) in self.broken:
            raise ConnectionError("server unavailable\nmore detail")
        if self.flaky.get((term, page), 0) > 0:
            self.flaky[(term, page)] -= 1
            raise ConnectionError("temporary problem")
        if self.ignore_page:
            page = 1
        start = (page - 1) * per_page
        return table_html(self.teams[term][start:start + per_page])


def run(site, terms, **kwargs):
    folder = kwargs.pop("folder", None) or tempfile.mkdtemp()
    path = os.path.join(folder, "checkpoint.json")
    return scraper.scrape_all(site, BASE, terms, per_page=10, checkpoint_path=path, **kwargs), path


def test_parse_table_uses_header_names():
    rows = scraper.parse_table(table_html([make_team("new", 1)]))
    assert list(rows[0]) == ["team_name", "year", "wins", "losses", "ot_losses",
                             "win_pct", "goals_for_gf", "goals_against_ga", "diff"]
    assert scraper.parse_table("<p>no table here</p>") == []
    assert scraper.parse_table(table_html([])) == []


def test_malformed_row_raises_instead_of_being_dropped():
    bad = table_html([make_team("a", 1)]).replace("<td></td>", "")
    try:
        scraper.parse_table(bad)
    except ValueError as error:
        assert "cells" in str(error)
    else:
        raise AssertionError("malformed row should raise")


def test_full_crawl_with_two_terms_is_complete():
    site = FakeSite({"new": 25, "san": 12})
    result, _ = run(site, ["new", "san"])
    assert result["status"] == "COMPLETE"
    assert len(result["records"]) == 37 and result["pages_ok"] == 3 + 2
    assert result["term_status"] == {"new": "done", "san": "done"}
    assert result["failures"] == []


def test_transient_failure_is_retried():
    result, _ = run(FakeSite({"new": 25}, flaky={("new", 2): 2}), ["new"])
    assert result["status"] == "COMPLETE" and len(result["records"]) == 25


def test_permanent_failure_is_partial_and_other_terms_continue():
    site = FakeSite({"new": 25, "san": 12}, broken={("new", 2)})
    result, _ = run(site, ["new", "san"])
    assert result["status"].startswith("PARTIAL")
    assert result["term_status"] == {"new": "failed", "san": "done"}
    assert result["failures"][0]["search_term"] == "new" and result["failures"][0]["page_number"] == 2
    assert len(result["records"]) == 10 + 12      # page 1 of "new" kept, all of "san"


def test_retries_are_bounded():
    site = FakeSite({"new": 25}, flaky={("new", 2): 99})
    result, _ = run(site, ["new"])
    assert result["status"].startswith("PARTIAL")
    assert site.flaky[("new", 2)] == 99 - scraper.MAX_RETRIES


def test_max_pages_is_labelled_limited():
    result, _ = run(FakeSite({"new": 25}), ["new"], max_pages=2)
    assert result["status"].startswith("LIMITED") and len(result["records"]) == 20


def test_stalled_pagination_is_not_reported_complete():
    result, _ = run(FakeSite({"new": 25}, ignore_page=True), ["new"])
    assert result["term_status"]["new"] == "stalled"
    assert result["status"].startswith("PARTIAL") and len(result["records"]) == 10


def test_resume_matches_an_uninterrupted_run():
    counts, terms = {"new": 25, "san": 12}, ["new", "san"]
    full, _ = run(FakeSite(counts), terms)

    folder = tempfile.mkdtemp()
    interrupted = FakeSite(counts, interrupt_after=3)       # dies during the 4th page
    try:
        run(interrupted, terms, folder=folder)
    except KeyboardInterrupt:
        pass
    else:
        raise AssertionError("run should have been interrupted")

    resumed_site = FakeSite(counts)
    resumed, _ = run(resumed_site, terms, folder=folder, resume=True)

    assert resumed["status"] == "COMPLETE"
    assert not set(interrupted.calls) & set(resumed_site.calls)     # finished pages are not refetched
    a, _ = scraper.clean(full["records"])
    b, _ = scraper.clean(resumed["records"])
    pd.testing.assert_frame_equal(a.drop(columns="scraped_at"), b.drop(columns="scraped_at"))


def test_cleaning_converts_numbers_and_removes_duplicates_across_terms():
    records = []
    for term in ("new", ""):                       # same team found by two searches
        for row in (make_team("new", 1), make_team("new", 2)):
            record = dict(zip(["team_name", "year", "wins", "losses", "ot_losses", "win_pct",
                               "goals_for_gf", "goals_against_ga", "diff"], row))
            records.append(dict(record, search_term=term, page_number=1, source_url="u", scraped_at="t"))
    df, duplicates = scraper.clean(records)
    assert duplicates == 2 and len(df) == 2
    assert pd.api.types.is_numeric_dtype(df["wins"]) and pd.api.types.is_numeric_dtype(df["win_pct"])
    assert df["ot_losses"].isna().all()                            # blanks stay missing
    assert df["team_name"].tolist() == ["New Team 1", "New Team 2"]


def test_excel_export_has_expected_sheets():
    result, _ = run(FakeSite({"new": 25}), ["new"])
    df, duplicates = scraper.clean(result["records"])
    sheets = scraper.build_reports(df, duplicates, result, ["new"], 1.0, BASE)
    with tempfile.TemporaryDirectory() as folder:
        path = os.path.join(folder, "report.xlsx")
        scraper.export_excel(path, sheets)
        read = pd.read_excel(path, sheet_name=None)
    assert list(read) == ["Summary", "Data", "By Term", "Top Seasons", "Failed Pages"]
    assert len(read["Data"]) == 25 and len(read["Top Seasons"]) == 10
    assert read["Summary"].iloc[0, 1] == "COMPLETE"


class FakeWorksheet:
    def __init__(self):
        self.cleared, self.values = 0, None

    def clear(self):
        self.cleared += 1

    def update(self, range_name, values):
        assert range_name == "A1"
        self.values = values


class FakeSpreadsheet:
    def __init__(self, existing=()):
        self.sheets = {name: FakeWorksheet() for name in existing}

    def worksheet(self, title):
        return self.sheets[title]             # raises KeyError when missing

    def add_worksheet(self, title, rows, cols):
        self.sheets[title] = FakeWorksheet()
        return self.sheets[title]


def test_google_sheet_export_creates_clears_and_writes_plain_values():
    result, _ = run(FakeSite({"new": 25}), ["new"])
    df, duplicates = scraper.clean(result["records"])
    sheets = scraper.build_reports(df, duplicates, result, ["new"], 1.0, BASE)

    spreadsheet = FakeSpreadsheet(existing=["Summary"])     # one sheet exists, the rest are created
    scraper.export_google_sheet(spreadsheet, sheets, missing_error=KeyError)

    assert set(spreadsheet.sheets) == {name for name, _ in sheets}
    assert all(ws.cleared == 1 for ws in spreadsheet.sheets.values())
    data = spreadsheet.sheets["Data"].values
    assert data[0][0] == "team_name" and len(data) == 26
    assert any(cell == "" for cell in data[1])               # missing values become blanks
    json.dumps(data)                                         # plain Python types only


if __name__ == "__main__":
    tests = [value for name, value in sorted(globals().items()) if name.startswith("test_")]
    for test in tests:
        test()
        print(f"PASS  {test.__name__}")
    print(f"\n{len(tests)} tests passed")
