"""Job Details table for the daily AA cabin report, read from AA Debriefs.csv.

AA Debriefs.csv is published hourly by aa-cabin-automation from cabin.aa.com, so the
report no longer needs the Job Detail CSV attached to the JotForm or the workbook's
Job_Details table. Shared verbatim by stl-report-automation and pvd-report-automation.

Rules (Sam, 2026-11-17):
  - only the report's station, and only COMPLETED jobs;
  - Ultras, Shrouds, and the Security search on those same aircraft visits are left
    out (they are the Ultra crew's work, tracked in Ultra Debriefs);
  - the "Scanned In" / "Scanned Out" columns show what cabin.aa.com recorded the job
    was started and closed with (Start / End Client Type: Mobile, Desktop, blank, or
    whatever else AA records), verbatim (Sam, 2026-11-22);
  - incorrectly scanned jobs come first, highlighted red. Incorrect = not started AND
    closed in the mobile app;
  - a Security search closed with anything but the mobile app has no electronic
    checklist in AA's system, so a physical security search form must be filled out
    and archived for it; those rows say so in the Job Type column;
  - those searches are also listed on their own in report item 12 ("Physical Security
    Search Forms Required"), INCLUDING ones on Ultra/Shroud visits, with tail, start/end
    times, the tech, and the job the search went with (Sam, 2026-11-22);
  - job dates: the day before the report's Start Date through the day before its End
    Date, matching the workbook's own count formula (B10-1 .. B11-1).
"""
import csv
import datetime as dt
import io

DEBRIEFS_PATH = "/Power Flows/Debriefs/AA Debriefs.csv"
HEADER = ["Job Date", "Start Time", "End Time", "Job Type", "Scanned In", "Scanned Out",
          "Start User", "Tail", "Status"]
PAPER_FORM = "Security (physical form)"
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
    return r.get("Start Client Type") != "Mobile" or r.get("End Client Type") != "Mobile"


def needs_paper_form(r):
    """AA's search checklist is only filled in when the Security job is closed in the
    mobile app; otherwise a physical form is required (checked against 230 searches
    in AA's DFW 11/15 and STL 11/01 search reports)."""
    return r.get("Job Type") == "Security" and r.get("End Client Type") != "Mobile"


def build_job_details(debriefs_bytes, station, start_date, end_date):
    """(header+rows, highlighted row numbers, n_jobs, n_scanned_ok, n_paper_forms)."""
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
                      PAPER_FORM if needs_paper_form(r) else r.get("Job Type", ""),
                      r.get("Start Client Type", ""), r.get("End Client Type", ""),
                      r.get("Start User", ""), r.get("Tail", ""), r.get("Status", "")])
        if miscanned(r):
            highlight.append(len(table) - 1)
    ok = len(jobs) - len(highlight)
    return table, highlight, len(jobs), ok, sum(1 for r in jobs if needs_paper_form(r))


def job_details_section(title, debriefs_bytes, station, start_date, end_date, as_of=None):
    """Section dict in the shape the report's section_elems() expects ('data' sections)."""
    table, highlight, n, ok, paper = build_job_details(debriefs_bytes, station, start_date, end_date)
    if n == 0:
        text = "No completed jobs found in AA Debriefs for this period."
    elif ok == n:
        text = (f"Of the {n} jobs completed, all {ok} of them were scanned in and out "
                "with the mobile app.")
    else:
        text = (f"Of the {n} jobs completed, {ok} of them were scanned in and out with the "
                f"mobile app. The {n - ok} that were not are listed first, in red.")
    if paper:
        s = "" if paper == 1 else "es"
        text += (f" <b>{paper} Security search{s} {'was' if paper == 1 else 'were'} closed out on "
                 "Desktop instead of the mobile app.</b> AA keeps no electronic search checklist "
                 "for a search closed on Desktop, so a physical security search form must be filled "
                 f"out and archived for {'it' if paper == 1 else 'each one'}; "
                 f"{'it is' if paper == 1 else 'they are'} marked \"{PAPER_FORM}\" below "
                 "and listed in item 12.")
    text += " Ultras, Shrouds and their Security searches are not included."
    if as_of:
        text += f" Job data as of {as_of}."
    return {"title": title, "text": text, "count": n,
            "col_start": None, "row_start": None, "row_end": None, "col_end": None,
            "data": table if n else [], "highlight": highlight}


# ── Report item 12: physical security search forms required ────────────────────
FORMS_HEADER = ["Job Date", "Tail", "Start Time", "End Time", "Tech", "Search For"]


def build_physical_forms(debriefs_bytes, station, start_date, end_date):
    """Completed Security searches closed with anything but the mobile app, for the
    report's station and dates. Unlike Job Details this INCLUDES searches on Ultra and
    Shroud visits: every one of them needs a physical form."""
    rows = list(csv.DictReader(io.StringIO(debriefs_bytes.decode("utf-8-sig"))))
    lo = _as_date(start_date) - dt.timedelta(days=1)
    hi = _as_date(end_date) - dt.timedelta(days=1)
    scope = [r for r in rows if r.get("Station") == station and r.get("Status") == "Completed"
             and lo <= _as_date(r["Job Date"]) <= hi]
    # The job(s) a search went with: the other completed jobs on the same aircraft visit.
    with_search = {}
    for r in scope:
        if r.get("Job Type") != "Security" and r.get("Visit"):
            with_search.setdefault(r["Visit"], set()).add(r.get("Job Type", ""))
    forms = sorted((r for r in scope if needs_paper_form(r)), key=lambda r: r.get("End Time") or "")
    table = [FORMS_HEADER]
    for r in forms:
        start_tech = (r.get("Start User") or "").strip()
        end_tech = (r.get("End User") or "").strip()
        tech = start_tech if (not end_tech or end_tech == start_tech) else f"{start_tech} / closed by {end_tech}"
        jobs = ", ".join(sorted(with_search.get(r.get("Visit"), ()))) or "None recorded"
        table.append([_mdy(r["Job Date"]), r.get("Tail", ""), _hhmm(r.get("Start Time")),
                      _hhmm(r.get("End Time")), tech or "-", jobs])
    return table, len(forms)


def physical_forms_section(title, debriefs_bytes, station, start_date, end_date):
    """Section dict for report item 12 ('data' section; empty table = None to Report)."""
    table, n = build_physical_forms(debriefs_bytes, station, start_date, end_date)
    if n == 0:
        text = "None to Report"
    else:
        s = "" if n == 1 else "es"
        text = (f"{n} Security search{s} {'was' if n == 1 else 'were'} closed out on Desktop instead of "
                "the mobile app. AA keeps no electronic search checklist for these, so a physical security "
                f"search form must be filled out and archived for {'it' if n == 1 else 'each one'}.")
    return {"title": title, "text": text, "count": n,
            "col_start": None, "row_start": None, "row_end": None, "col_end": None,
            "data": table if n else []}

