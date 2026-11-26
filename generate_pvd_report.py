"""
PVD AA Cabin Report PDF Generator — Cloud Version
Reads Excel from SharePoint via Microsoft Graph API.
Generates two PDFs and saves them to SharePoint.
Sends email via Microsoft Graph API.
"""

import os
import re
import io
import sys
import csv
import json
import base64
import requests
import openpyxl
from pathlib import Path
from datetime import date, datetime, time, timedelta, timezone
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from zoneinfo import ZoneInfo

from aa_jobs import DEBRIEFS_PATH, job_details_section, physical_forms_section

# ─── MICROSOFT GRAPH AUTH CONFIG (set via GitHub Actions secrets) ─────────────
TENANT_ID     = os.environ.get("TENANT_ID")
CLIENT_ID     = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")

# Dry-run mode: point at the OneDrive-synced Data Hub folder to read every
# input locally, write the PDFs to ./out, and skip upload + email entirely.
LOCAL_DATA_HUB = os.environ.get("LOCAL_DATA_HUB")

# ─── SHAREPOINT CONFIG ────────────────────────────────────────────────────────
# Drive ID for foxtrotaviationcom.sharepoint.com/sites/DataHub
DRIVE_ID = "b!_bzXaIx86kOufgJN3ih-BaDIDthKYuxJkJtLi1Bm5irGjCEnK-VHSpBRRm3_SDKU"

# Path to the PVD Excel file (relative to drive root)
EXCEL_PATH = "/Report Automation/PVD AA Cabin/PVD Tables.xlsx"

# SharePoint folder path where PDFs will be saved
PDF_OUTPUT_FOLDER = "/Report Automation/PVD AA Cabin/Daily Reports"

# --- Dashboard screenshot: JotForm form 252504858606058, upload field 29 ---
# Maren uploads the station dashboard screenshot through this form. The PA
# flow that used to copy it to /Flow Dumps/AA Screenshots/<station>.png was
# deleted by accident (2026-11), so the report reads the form directly. The
# repository_dispatch carries no submission id, so the newest submission for
# THIS station wins; the station field is found by its ANSWER VALUE, never a
# hardcoded qid. A stale dashboard is worse than none: nothing fresh within
# MAX_SCREENSHOT_AGE_DAYS means the email simply goes out without the image.
JOTFORM_KEY = os.environ.get("JOTFORM_KEY", "")
# Foxtrot is on JotForm ENTERPRISE: api.jotform.com answers 401 to its
# keys — the org subdomain is the only valid host (same base the
# platform's engine/safety.py and rampbudget.py use).
JF_BASE = os.environ.get("JOTFORM_BASE", "https://foxtrotaviation.jotform.com/API")
JF_FORM_ID = "252504858606058"
JF_UPLOAD_QID = "29"
JF_STATION = "PVD"
MAX_SCREENSHOT_AGE_DAYS = 3

# ─── EMAIL CONFIG ─────────────────────────────────────────────────────────────
# Set to False to skip sending the email (useful when testing)
SEND_EMAIL = True

EMAIL_FROM = "foxtrot.automation@foxtrotaviation.com"

EMAIL_RECIPIENTS = [
    "jamie.mathers@foxtrotaviation.com",
    "levorn.smallwood@foxtrotaviation.com",
    "lori.desjardin@foxtrotaviation.com",
    "maren.pinpin@foxtrotaviation.com",
    "daniel.digiambattista@foxtrotaviation.com",
    "anthony.pentz@foxtrotaviation.com",
    "chris.stump@foxtrotaviation.com",
    "samuel.kosco@foxtrotaviation.com",
]

# ─── MICROSOFT GRAPH AUTH ─────────────────────────────────────────────────────

def get_access_token():
    """Obtain an OAuth2 access token from Microsoft identity platform."""
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    })
    resp.raise_for_status()
    return resp.json()["access_token"]


# ─── SHAREPOINT FILE OPERATIONS ───────────────────────────────────────────────

def download_excel(token):
    """Download the PVD Excel file from SharePoint into a BytesIO buffer."""
    data = download_source(token, EXCEL_PATH)
    print("PVD Excel file loaded.")
    return io.BytesIO(data)


def download_source(token, path):
    """Download a Data Hub file by drive-relative path (or read the synced
    local copy when LOCAL_DATA_HUB is set)."""
    if LOCAL_DATA_HUB:
        return (Path(LOCAL_DATA_HUB) / path.lstrip("/")).read_bytes()
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:{path}:/content"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    return resp.content


def source_last_modified(token, path):
    """Return the file's last-modified datetime (local mtime in dry-run mode)."""
    if LOCAL_DATA_HUB:
        mtime = (Path(LOCAL_DATA_HUB) / path.lstrip("/")).stat().st_mtime
        return datetime.fromtimestamp(mtime)
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:{path}"
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    stamp = resp.json()["lastModifiedDateTime"]
    return datetime.fromisoformat(stamp.replace("Z", "+00:00"))


def download_screenshot():
    """The station dashboard screenshot from the newest matching JotForm
    submission. Returns PNG bytes, or None when the key is missing, nothing
    fresh matches, or any step fails — the email then omits the image rather
    than embedding a stale one (the old SharePoint copy stopped updating the
    day the PA flow was deleted, which is exactly the failure this avoids)."""
    if not JOTFORM_KEY:
        print("JOTFORM_KEY not set - sending without the dashboard image.")
        return None
    try:
        resp = requests.get(
            f"{JF_BASE}/form/{JF_FORM_ID}/submissions",
            params={"apiKey": JOTFORM_KEY, "limit": "100"}, timeout=30)
        resp.raise_for_status()
        subs = resp.json().get("content") or []
        subs.sort(key=lambda s: s.get("created_at") or "", reverse=True)
        cutoff = datetime.now(timezone.utc) - timedelta(days=MAX_SCREENSHOT_AGE_DAYS)
        for sub in subs:
            created = sub.get("created_at") or ""
            try:  # JotForm stamps are naive US/Eastern-ish strings; compare loosely
                created_dt = datetime.fromisoformat(created).replace(tzinfo=timezone.utc)
            except ValueError:
                created_dt = None
            if created_dt and created_dt < cutoff:
                break  # newest-first: everything after this is older still
            answers = sub.get("answers") or {}
            station = ""
            for qid, a in answers.items():
                if qid == JF_UPLOAD_QID:
                    continue
                v = a.get("answer")
                if isinstance(v, str) and v.strip().upper() in ("STL", "PVD"):
                    station = v.strip().upper()
                    break
            if station != JF_STATION:
                continue
            upload = (answers.get(JF_UPLOAD_QID) or {}).get("answer")
            if isinstance(upload, list):
                upload = upload[0] if upload else None
            if not isinstance(upload, str) or not upload.strip():
                continue
            file_url = upload.strip()
            fresp = requests.get(file_url, params={"apiKey": JOTFORM_KEY}, timeout=30)
            fresp.raise_for_status()
            if not fresp.content or not fresp.content.startswith(b"\x89PNG"):
                print(f"Submission {sub.get('id')} field {JF_UPLOAD_QID} is not "
                      f"a PNG ({len(fresp.content)} bytes) - skipping.")
                continue
            print(f"Dashboard screenshot from JotForm submission "
                  f"{sub.get('id')} ({created}, {station}).")
            return fresp.content
        print(f"No {JF_STATION} submission with a PNG in the last "
              f"{MAX_SCREENSHOT_AGE_DAYS} days - sending without the image.")
        return None
    except Exception as e:  # noqa: BLE001 - the report must still go out
        print(f"Dashboard screenshot fetch failed ({e!r}) - sending without it.")
        return None


def upload_pdf(token, filename, pdf_bytes):
    """Upload a PDF to the SharePoint PVD output folder."""
    upload_path = f"{PDF_OUTPUT_FOLDER}/{filename}"
    url = (f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
           f"/root:{upload_path}:/content")
    resp = requests.put(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type":  "application/pdf",
        },
        data=pdf_bytes,
    )
    resp.raise_for_status()
    print(f"Uploaded to SharePoint: {upload_path}")
    return resp.json().get("webUrl", "")


# ─── EMAIL SENDING ────────────────────────────────────────────────────────────

def send_email(token, subject, body_text, recipients, from_address,
               attachments, screenshot_bytes=None):
    """
    Send HTML email via Microsoft Graph API with PDF attachments.
    If screenshot_bytes is provided it is embedded inline in the email body
    below the message text, using a cid: reference.
    attachments: list of (filename, bytes) tuples (PDFs)
    """
    attachment_list = []

    # Inline screenshot — referenced in HTML body via cid:stationScreenshot
    if screenshot_bytes:
        attachment_list.append({
            "@odata.type":  "#microsoft.graph.fileAttachment",
            "name":         "screenshot.png",
            "contentType":  "image/png",
            "contentBytes": base64.b64encode(screenshot_bytes).decode("utf-8"),
            "contentId":    "stationScreenshot",
            "isInline":     True,
        })

    # PDF file attachments (appear as regular attachments)
    for filename, pdf_bytes in attachments:
        attachment_list.append({
            "@odata.type":  "#microsoft.graph.fileAttachment",
            "name":         filename,
            "contentType":  "application/pdf",
            "contentBytes": base64.b64encode(pdf_bytes).decode("utf-8"),
        })

    # Build HTML body: each non-empty line becomes a paragraph, image appended below
    html_paragraphs = "".join(
        f'<p style="margin:0 0 10px 0;">{line}</p>'
        for line in body_text.split("\n") if line.strip()
    )
    img_tag = (
        '<p><img src="cid:stationScreenshot" '
        'style="max-width:600px;width:100%;border:1px solid #ccc;" '
        'alt="Station Screenshot" /></p>'
    ) if screenshot_bytes else ""
    html_body = (
        '<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;">'
        f"{html_paragraphs}{img_tag}"
        "</body></html>"
    )

    message = {
        "subject": subject,
        "body":    {"contentType": "HTML", "content": html_body},
        "toRecipients": [
            {"emailAddress": {"address": addr}} for addr in recipients
        ],
        "attachments": attachment_list,
    }

    url = f"https://graph.microsoft.com/v1.0/users/{from_address}/sendMail"
    resp = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"message": message, "saveToSentItems": True},
    )
    resp.raise_for_status()
    print(f"Email sent to {len(recipients)} recipients.")


# ─── VALUE FORMATTING ─────────────────────────────────────────────────────────

def fmt_value(val):
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%-m/%-d/%Y") if sys.platform != "win32" else val.strftime("%#m/%#d/%Y")
    if isinstance(val, time):
        return val.strftime("%-I:%M %p") if sys.platform != "win32" else val.strftime("%#I:%M %p")
    return str(val)


def fmt_date_short(val):
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%-m/%-d") if sys.platform != "win32" else val.strftime("%#m/%#d")
    return str(val)


def fmt_pct(x):
    """Match Excel's ROUND(...,2) display: 87.5 not 87.50, 84.25 stays 84.25."""
    return f"{round(x, 2):.2f}".rstrip("0").rstrip(".")


# ─── DIRECT-SOURCE SECTIONS (replaces the workbook's Power Query layer) ───────
# The LMS / Safety101 tables inside PVD Tables.xlsx are Power Query connections
# whose cached values only refresh when a human opens the file in desktop
# Excel — the pivots hadn't refreshed since June and the old queries filtered
# on "PVD AA CABIN", a value that doesn't exist in the source CSVs ("PVD AA
# CAB"), which is where the company-wide LMS numbers and the #DIV/0! in the
# emailed PDF came from. Everything in this block reads the underlying source
# files directly with the correct filters. Tables Power Automate writes into
# the workbook (PTO, delays, skips, flags, tests, punches, breaks, audits,
# jobs) are still read from the workbook — that data lives nowhere else.

LMS_TRANSCRIPT_PATH = "/Flow Dumps/dashboard-transcript/learner_transcript.csv"
S101_EXPIRING_PATH  = ("/Safety101/S101 Compliance/"
                       "Foxtrot Aviation Services Entire Organization"
                       " Expiring Training.csv")
S101_DATA_PATH      = "/Safety101/S101 Compliance/Safety101 Data.csv"

LMS_LOCATIONS  = {"PVD AA CAB"}  # learner_transcript Location
S101_LOCATIONS = {"PVD AA CAB"}  # Safety101 Data Location


def s101_department_match(dept):
    return dept == "PVD AA CAB"


def read_csv_dicts(raw):
    return list(csv.DictReader(io.StringIO(raw.decode("utf-8-sig"))))


def build_lms_sections(lms_rows):
    """LMS by-employee and by-title sections from learner_transcript.csv.
    Sentence mirrors the workbook's B3 formula exactly."""
    rows = [r for r in lms_rows if r["Location"] in LMS_LOCATIONS]
    status    = [r["Status"] for r in rows]
    overdue   = status.count("Overdue")
    due_soon  = status.count("Not Started") + status.count("In Progress")
    completed = status.count("Completed")
    total     = sum(1 for r in rows if r["Employee"])
    compliance = fmt_pct(100 * completed / total) if total else "0"
    text = (f"There are currently {overdue} Overdue Courses on the LMS, and "
            f"{due_soon} Courses Due Soon ({compliance}% Compliance)")

    def grouped(field, header):
        counts = {}
        for r in rows:
            if r["Status"] == "Overdue":
                key = r[field].strip()
                counts[key] = counts.get(key, 0) + 1
        data = [[header, "Overdue"]]
        data += [[k, str(v)] for k, v in
                 sorted(counts.items(), key=lambda kv: kv[0].lower())]
        if counts:
            data.append(["Grand Total", str(sum(counts.values()))])
        return {'text': text, 'count': len(counts),
                'data': data if counts else []}

    return grouped("Employee", "Employee"), grouped("Course Title", "Title")


def build_s101_section(expiring_rows, s101_rows):
    """Safety101 section from the platform's Expiring Training report plus the
    internal compliance calc. Sentence mirrors the workbook's B4 formula."""
    exp = [r for r in expiring_rows if s101_department_match(r["Department"])]
    first_words   = [r["Expiration Status"].split(" ")[0] for r in exp]
    expiring_soon = first_words.count("Expires")
    noncompliant  = len(first_words) - expiring_soon
    comp_vals = [float(r["Comp?"]) for r in s101_rows
                 if r["Location"] in S101_LOCATIONS and r["Comp?"] not in (None, "")]
    compliance = fmt_pct(100 * sum(comp_vals) / len(comp_vals)) if comp_vals else "0"
    text = (f"There are currently {noncompliant} Noncompliant Qualifications on "
            f"Safety101, and {expiring_soon} Qualifications Expiring Soon "
            f"({compliance}% Compliance)")

    counts = {}
    for r in exp:
        key = r["Employee"].strip()
        counts[key] = counts.get(key, 0) + 1
    data = [["Employee", "Noncompliant Courses"]]
    data += [[k, str(v)] for k, v in
             sorted(counts.items(), key=lambda kv: kv[0].lower())]
    if counts:
        data.append(["Grand Total", str(sum(counts.values()))])
    return {'text': text, 'count': len(counts), 'data': data if counts else []}


def build_tests_section(wb, start_date, end_date):
    """Security-tests section read straight from the Tests sheet. The Email
    Generation FILTER formulas for this section broke to #REF! long ago, so
    the emailed report always claimed there were no tests. Sentence mirrors
    the workbook's B9 formula."""
    ws = wb["Tests"]
    it = ws.iter_rows(values_only=True)
    next(it)  # header: Date | Shift | Tester | Test Object | Employee | Outcome | Notes
    rows, fails = [], 0
    for r in it:
        d = r[0]
        if not isinstance(d, datetime):
            continue
        if start_date and end_date and start_date <= d <= end_date:
            rows.append([fmt_value(d), fmt_value(r[2]), fmt_value(r[4]),
                         fmt_value(r[5]), fmt_value(r[6])])
            if str(r[5]).strip() == "Fail":
                fails += 1
    n = len(rows)
    if n == 0:
        text = "There were no security tests in date range"
    else:
        text = (f"Of the {n} security tests conducted in date range, "
                f"{fails} were failed")
    data = ([["Date", "Tester", "Employee", "Outcome", "Notes"]] + rows
            if rows else [])
    return {'text': text, 'count': n, 'data': data}


# ─── TABLE HELPERS ────────────────────────────────────────────────────────────

def get_table_data(ws, row_start, row_end, col_start, col_end):
    data = []
    for row in ws.iter_rows(
        min_row=row_start, max_row=row_end,
        min_col=col_start, max_col=col_end,
        values_only=True,
    ):
        formatted = [fmt_value(v) for v in row]
        if any(v != "" for v in formatted):
            data.append(formatted)
    return data


def make_table(data, col_widths=None,
               header_bg=colors.HexColor('#2F5496'),
               stripe_color=colors.HexColor('#DCE6F1'),
               highlight_rows=()):
    if not data:
        return None
    while data and all(v == "" for v in data[-1]):
        data.pop()
    if not data:
        return None

    n_cols     = len(data[0])
    available  = 7.5 * inch
    if col_widths is None:
        col_widths = [available / n_cols] * n_cols

    header_style = ParagraphStyle(
        'TableHeader', fontName='Helvetica-Bold', fontSize=8,
        textColor=colors.white, alignment=TA_CENTER, leading=10,
    )
    cell_style = ParagraphStyle(
        'TableCell', fontName='Helvetica', fontSize=7.5,
        textColor=colors.black, alignment=TA_LEFT, leading=10,
    )

    # Rows flagged by the caller (e.g. incorrectly scanned jobs) print red.
    miscan_style = ParagraphStyle(
        'TableCellFlagged', parent=cell_style,
        textColor=colors.HexColor('#9C0006'),
    )
    highlight_rows = set(highlight_rows)

    wrapped = []
    for r_idx, row in enumerate(data):
        new_row = []
        for cell in row:
            if isinstance(cell, str):
                s = (header_style if r_idx == 0 else
                     miscan_style if r_idx in highlight_rows else cell_style)
                new_row.append(Paragraph(cell, s))
            else:
                new_row.append(cell)
        wrapped.append(new_row)

    t = Table(wrapped, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ('BACKGROUND', (0, 0), (-1, 0),  header_bg),
        ('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID',       (0, 0), (-1, -1), 0.5, colors.HexColor('#B8CCE4')),
        ('BOX',        (0, 0), (-1, -1), 1,   colors.HexColor('#2F5496')),
        ('TOPPADDING',    (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING',   (0, 0), (-1, -1), 4),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 4),
    ]
    for i in range(1, len(wrapped)):
        if i % 2 == 0:
            style_cmds.append(('BACKGROUND', (0, i), (-1, i), stripe_color))
    for i in highlight_rows:
        style_cmds.append(('BACKGROUND', (0, i), (-1, i),
                           colors.HexColor('#F8D7DA')))
    t.setStyle(TableStyle(style_cmds))
    return t


# ─── PDF GENERATION ───────────────────────────────────────────────────────────

def build_pdf_bytes(ws_email, ws_jobs, date_line, sections, lms_section,
                    abridged=False):
    """Generate a PDF and return its bytes."""
    buf = io.BytesIO()

    style_title = ParagraphStyle(
        'ReportTitle', fontName='Helvetica-Bold', fontSize=14,
        textColor=colors.HexColor('#1F3864'), spaceAfter=4,
    )
    style_section = ParagraphStyle(
        'SectionHeader', fontName='Helvetica-Bold', fontSize=9,
        textColor=colors.HexColor('#2F5496'), spaceBefore=8, spaceAfter=2,
    )
    style_body = ParagraphStyle(
        'Body', fontName='Helvetica', fontSize=8.5, spaceAfter=4, leading=12,
    )
    style_small = ParagraphStyle(
        'SmallNote', fontName='Helvetica-Oblique', fontSize=7.5,
        textColor=colors.HexColor('#555555'), spaceAfter=4,
    )

    flag_widths = [0.5*inch, 0.7*inch, 0.7*inch, 0.6*inch,
                   1.2*inch, 0.5*inch, 3.3*inch]
    mp_widths   = [1.8*inch, 0.8*inch, 1.5*inch, 1.2*inch, 1.2*inch]
    eb_widths   = [1.8*inch, 0.8*inch, 1.7*inch, 1.5*inch]
    audit_widths = [1.2*inch, 1.5*inch, 1.5*inch, 3.3*inch]
    forms_widths = [0.8*inch, 0.6*inch, 0.75*inch, 0.75*inch, 2.5*inch, 2.1*inch]
    qr_widths   = [0.8*inch, 0.75*inch, 0.75*inch, 0.7*inch,
                   0.7*inch, 0.7*inch, 1.4*inch, 0.5*inch, 0.7*inch]

    def section_elems(meta, ws_source=None, col_widths=None, include_table=True):
        rs, re_ = meta['row_start'], meta['row_end']
        cs, ce  = meta['col_start'], meta['col_end']
        count   = meta['count']
        elems = [
            Paragraph(str(meta['title'] or ""), style_section),
            Paragraph(str(meta['text']  or ""), style_body),
        ]
        if include_table:
            if 'data' in meta:
                # Section computed directly from source files, not the workbook
                tbl = make_table([list(r) for r in meta['data']],
                                 highlight_rows=meta.get('highlight', ()),
                                 col_widths=col_widths) if meta['data'] else None
                if tbl:
                    elems.append(tbl)
            else:
                has_table = not (rs == re_ and count == 0)
                if has_table and all(v is not None for v in [rs, re_, cs, ce]):
                    src  = ws_source or ws_email
                    data = get_table_data(src, rs, re_, cs, ce)
                    tbl  = make_table(data, col_widths=col_widths)
                    if tbl:
                        elems.append(tbl)
        elems.append(Spacer(1, 4))
        return elems

    title_text = "PVD AA Cabin Report" + (" (Abridged)" if abridged else "")
    story = [
        Paragraph(title_text, style_title),
        Paragraph(date_line,  style_small),
        Spacer(1, 6),
    ]

    def add(meta, ws_source=None, col_widths=None, is_flags=False):
        show_table = (not abridged) or is_flags
        story.append(KeepTogether(
            section_elems(meta, ws_source=ws_source,
                          col_widths=col_widths, include_table=show_table)
        ))

    add(sections[2])
    add(lms_section)
    add(sections[4])
    add(sections[5])
    add(sections[6])
    add(sections[7])
    add(sections[8],  col_widths=flag_widths, is_flags=True)
    add(sections[9])
    add(sections[14], col_widths=mp_widths)
    add(sections[15], col_widths=eb_widths)
    add(sections[16], col_widths=audit_widths)
    if 17 in sections:
        add(sections[17], col_widths=forms_widths)
    add(sections[12], ws_source=ws_jobs, col_widths=qr_widths)

    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        leftMargin=0.4*inch, rightMargin=0.4*inch,
        topMargin=0.4*inch,  bottomMargin=0.4*inch,
    )
    doc.build(story)
    return buf.getvalue()


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    # 1. Authenticate
    if LOCAL_DATA_HUB:
        token = None
        print(f"LOCAL_DATA_HUB set — reading {LOCAL_DATA_HUB}; "
              f"dry run (no upload, no email).")
    else:
        for var in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET"):
            if not os.environ.get(var):
                sys.exit(f"Missing env var {var} (or set LOCAL_DATA_HUB)")
        print("Authenticating with Microsoft Graph...")
        token = get_access_token()

    # 2. Download PVD Excel from SharePoint
    excel_bytes = download_excel(token)
    wb       = openpyxl.load_workbook(excel_bytes, data_only=True)
    ws_email = wb['Email Generation']
    ws_jobs  = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else None  # retired: Job Details now come from AA Debriefs.csv

    # 3. Read submission ID and dates
    submission_id = str(ws_email['B1'].value or "report")
    if submission_id.startswith('_'):
        submission_id = submission_id[1:]

    start_date_raw = ws_email.cell(row=10, column=2).value
    end_date_raw   = ws_email.cell(row=11, column=2).value
    start_date_str = fmt_value(start_date_raw) if start_date_raw else ""
    end_date_str   = fmt_value(end_date_raw)   if end_date_raw   else ""

    date_line = start_date_str
    if end_date_str and end_date_str != start_date_str:
        date_line += f" \u2013 {end_date_str}"

    start_short = fmt_date_short(start_date_raw) if start_date_raw else ""
    end_short   = fmt_date_short(end_date_raw)   if end_date_raw   else ""
    date_short  = (f"{start_short}-{end_short}"
                   if end_short and end_short != start_short
                   else start_short)

    # 4. Read section metadata (rows 2–15 of Email Generation)
    sections = {}
    for sheet_row in range(2, 17):
        row_data = list(ws_email.iter_rows(
            min_row=sheet_row, max_row=sheet_row,
            min_col=1, max_col=7, values_only=True,
        ))[0]
        title, text, count, d, e, f, g = row_data
        sections[sheet_row] = {
            'title':     title,
            'text':      text,
            'count':     count or 0,
            'col_start': d,
            'row_start': e,
            'row_end':   f,
            'col_end':   g,
        }

    # 4b. Replace the Power-Query-backed sections with direct source reads
    print("Reading source files directly (LMS / Safety101)...")
    lms_rows = read_csv_dicts(download_source(token, LMS_TRANSCRIPT_PATH))
    sec_lms_emp, sec_lms_title = build_lms_sections(lms_rows)
    sections[3].update(sec_lms_emp)
    sections[13].update(sec_lms_title)

    exp_rows  = read_csv_dicts(download_source(token, S101_EXPIRING_PATH))
    s101_rows = read_csv_dicts(download_source(token, S101_DATA_PATH))
    sections[4].update(build_s101_section(exp_rows, s101_rows))

    # AA Learning Hub: no PVD source exists — keep the workbook's static text,
    # suppress the (broken #REF!) table range.
    sections[5]['data'] = []

    # Security tests: the Email Generation FILTER for this section is #REF!;
    # read the Tests sheet directly instead.
    sections[9].update(build_tests_section(wb, start_date_raw, end_date_raw))

    # Item numbers (Sam, 2026-11-22): the workbook numbered both Station Audits and
    # Job Details "11.". Station Audits = 11, the new physical security search
    # forms item = 12, Job Details = 13.
    def renumber(title, n):
        return re.sub(r"^\s*\d+\.\s*", f"{n}. ", str(title or ""))
    sections[16]['title'] = renumber(sections[16]['title'], 11)
    sections[12]['title'] = renumber(sections[12]['title'], 13)

    # Job Details (item 13) and physical forms (item 12): read from AA Debriefs.csv,
    # which aa-cabin-automation publishes hourly from cabin.aa.com, instead of the
    # workbook's Job_Details table fed by the JotForm CSV upload.
    if start_date_raw and end_date_raw:
        as_of = source_last_modified(token, DEBRIEFS_PATH).astimezone(
            ZoneInfo("America/New_York")).strftime("%m/%d %I:%M %p ET")
        sections[12] = job_details_section(
            sections[12]['title'], download_source(token, DEBRIEFS_PATH),
            "PVD", start_date_raw, end_date_raw, as_of=as_of)
        sections[17] = physical_forms_section(
            "12. Physical Security Search Forms Required:",
            download_source(token, DEBRIEFS_PATH),
            "PVD", start_date_raw, end_date_raw)

    lms_section = (
        sections[13] if sections[13]['count'] < sections[3]['count']
        else sections[3]
    )

    # 5. Generate PDFs
    full_filename     = f"PVD Report Document_{submission_id}.pdf"
    abridged_filename = f"PVD Report Document Abridged_{submission_id}.pdf"

    print("Generating full report PDF...")
    full_bytes = build_pdf_bytes(
        ws_email, ws_jobs, date_line, sections, lms_section, abridged=False)

    print("Generating abridged report PDF...")
    abridged_bytes = build_pdf_bytes(
        ws_email, ws_jobs, date_line, sections, lms_section, abridged=True)

    # Dry run: write the PDFs locally and stop before any side effects
    if LOCAL_DATA_HUB:
        out = Path("out")
        out.mkdir(exist_ok=True)
        (out / full_filename).write_bytes(full_bytes)
        (out / abridged_filename).write_bytes(abridged_bytes)
        print(f"Dry run complete — PDFs written to {out.resolve()}")
        return

    # 6. Upload PDFs to SharePoint
    upload_pdf(token, full_filename,     full_bytes)
    upload_pdf(token, abridged_filename, abridged_bytes)

    # 7. Download screenshot for email body
    print("Downloading station screenshot...")
    screenshot_bytes = download_screenshot()

    # 8. Send email with both PDFs attached
    if SEND_EMAIL:
        subject = f"PVD Recap: {date_short}"
        body = (
            f"Hello team,\n\n"
            f"View the attached abridged and full reports for PVD on {date_short}. "
            f"Please reach out to Maren with any questions about the reported "
            f"operations, and reach out to Sam with questions about the report "
            f"creation or formatting.\n\n"
            f"Thanks,\n"
            f"Foxtrot Automation Services"
        )
        send_email(
            token=token,
            subject=subject,
            body_text=body,
            recipients=EMAIL_RECIPIENTS,
            from_address=EMAIL_FROM,
            attachments=[
                (abridged_filename, abridged_bytes),
                (full_filename,     full_bytes),
            ],
            screenshot_bytes=screenshot_bytes,
        )
    else:
        print("Email sending is disabled (SEND_EMAIL = False).")

    print("All done!")


if __name__ == "__main__":
    main()
