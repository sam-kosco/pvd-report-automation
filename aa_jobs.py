"""Job Details table for the daily AA cabin report, read from AA Debriefs.csv.

AA Debriefs.csv is published hourly by aa-cabin-automation from cabin.aa.com, so the
report no longer needs the Job Detail CSV attached to the JotForm or the workbook's
Job_Details table. Shared verbatim by stl-report-automation and pvd-report-automation.

Rules (Sam, 2026-11-17):
  - only the report's station, and only COMPLETED jobs;
  - Ultras, Shrouds, and the Security search on those same aircraft visits are left
    out (they are the Ultra crew's work, tracked in Ultra Debriefs);
  - incorrectly scanned jobs come first, highlighted red. Incorrect = no start scan,
    no end scan, or started/closed from Desktop instead of the mobile app;
  - job dates: the day before the report's Start Date through the day before its End
    Date, matching the workbook's own count formula (B10-1 .. B11-1).
"""
import csv
import datetime as dt
import io

DEBRIEFS_PATH = "/Power Flows/Debriefs/AA Debriefs.csv"
HEADER = ["Job Date", "Start Time", "End Time", "Job Type", "Start Scan", "End Scan",
          "Start User", "Tail", "Status"]
EXCLUDED_TYPES = ("Ultra", "Shroud")


def _as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return dt.date.fromisoformat(str(v)[:10])


def _hhmm(stamp):
    s = (stamp or "").strip()
    return s[11:16] if len(s) >= 16 else s


def _mdy(stamp):
    d = _as_date(stamp)
    return f"{d.month}/{d.day}/{d.year}"


def miscanned(r):
    return (r.get("Start Scan") != "Y" or r.get("End Scan") != "Y"
            or r.get("Start Client Type") != "Mobile" or r.get("End Client Type") != "Mobile")


def build_job_details(debriefs_bytes, station, start_date, end_date):
    """(header+rows, highlighted row numbers, n_jobs, n_scanned_ok) for the report."""
    rows = list(csv.DictReader(io.StringIO(debriefs_bytes.decode("utf-8-sig"))))
    lo = _as_date(start_date) - dt.timedelta(days=1)
    hi = _as_date(end_date) - dt.timedelta(days=1)

    def in_scope(r):
        return (r.get("Station") == station and r.get("Status") == "Completed"
                and lo <= _as_date(r["Job Date"]) <= hi)

    scope = [r for r in rows if in_scope(r)]
    ultra_visits = {r.get("Visit") for r in scope
                    if r.get("Job Type") in EXCLUDED_TYPES and r.get("Visit")}
    jobs = [r for r in scope
            if r.get("Job Type") not in EXCLUDED_TYPES
            and not (r.get("Job Type") == "Security" and r.get("Visit") in ultra_visits)]
    jobs.sort(key=lambda r: (not miscanned(r), r.get("Start Time") or ""))

    table = [HEADER]
    highlight = []
    for r in jobs:
        table.append([_mdy(r["Job Date"]), _hhmm(r.get("Start Time")), _hhmm(r.get("End Time")),
                      r.get("Job Type", ""), r.get("Start Scan", ""), r.get("End Scan", ""),
                      r.get("Start User", ""), r.get("Tail", ""), r.get("Status", "")])
        if miscanned(r):
            highlight.append(len(table) - 1)
    ok = len(jobs) - len(highlight)
    return table, highlight, len(jobs), ok


def job_details_section(title, debriefs_bytes, station, start_date, end_date, as_of=None):
    """Section dict in the shape the report's section_elems() expects ('data' sections)."""
    table, highlight, n, ok = build_job_details(debriefs_bytes, station, start_date, end_date)
    if n == 0:
        text = "No completed jobs found in AA Debriefs for this period."
    elif ok == n:
        text = f"Of the {n} jobs completed, all {ok} of them were scanned properly."
    else:
        text = (f"Of the {n} jobs completed, {ok} of them were scanned properly. "
                f"The {n - ok} that were not are listed first, in red.")
    text += " Ultras, Shrouds and their Security searches are not included."
    if as_of:
        text += f" Job data as of {as_of}."
    return {"title": title, "text": text, "count": n,
            "col_start": None, "row_start": None, "row_end": None, "col_end": None,
            "data": table if n else [], "highlight": highlight}
