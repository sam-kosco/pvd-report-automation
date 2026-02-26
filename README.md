# PVD AA Cabin Report Generator

Automatically generates and emails the PVD AA Cabin daily report PDFs when triggered by Power Automate.

## How it works

1. A JotForm submission triggers a Power Automate cloud flow
2. Power Automate updates the PVD Excel file on SharePoint
3. Power Automate sends a `repository_dispatch` webhook to this GitHub repo
4. GitHub Actions runs `generate_pvd_report.py`, which:
   - Downloads `PVD Tables.xlsx` from SharePoint via Microsoft Graph API
   - Generates two PDFs (full report + abridged)
   - Uploads both PDFs to SharePoint: `Report Automation/PVD AA Cabin/Daily Reports`
   - Emails both PDFs to the PVD distribution list

## Required GitHub Secrets

Set these in **Settings → Secrets and variables → Actions**:

| Secret name     | Value                                      |
|-----------------|--------------------------------------------|
| `TENANT_ID`     | `ede0c57f-549f-4a90-9f8c-7ea130346f95`    |
| `CLIENT_ID`     | `58191600-ab56-4141-bff6-806805fcbff4`    |
| `CLIENT_SECRET` | App registration client secret value       |

Note: These are the same Entra app credentials used by the STL repo (`stl-report-automation`).
The `CLIENT_SECRET` is shared — if you renew it in Entra, update it in **both** repos.

## Manual trigger

Go to the **Actions** tab → **Generate PVD Report** → **Run workflow**.
Useful for testing or regenerating a report after a data correction.

## Key file locations

| Item | Location |
|---|---|
| Excel source | SharePoint: Report Automation → PVD AA Cabin → PVD Tables.xlsx |
| PDF output | SharePoint: Report Automation → PVD AA Cabin → Daily Reports |
| Entra app | entra.microsoft.com → App registrations → Foxtrot Report Automation |
| Sending mailbox | foxtrot.automation@foxtrotaviation.com |

## Credential expiry reminders

| Credential | Expires | Where to renew |
|---|---|---|
| `CLIENT_SECRET` | 24 months from creation | Entra → update secret in this repo AND stl-report-automation |
| GitHub PAT (Power Automate trigger) | 12 months | GitHub → update in Power Automate HTTP action |
