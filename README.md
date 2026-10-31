# RBI Notifications – BRC

Automated register of Reserve Bank of India publications.

The tool downloads RBI **Notifications, Master Directions, Master Circulars and Press Releases**, files each PDF by financial year and date, classifies it by subject and issuing department, and builds an offline HTML dashboard to browse everything.

## What it does

- Reads RBI's RSS feeds and listing pages and downloads only new publications.
- Saves PDFs as `FY 2026-27\2026-10-07\<Category>\<Title>__<file>.pdf` (financial year April–March, RBI publication date).
- Verifies every download is a genuine PDF and skips RBI site-wide documents.
- Classifies each publication by subject (FDI, ODI, ECB, LRS, NBFC, KYC & AML, Payment Systems, and more) and by issuing department (FED, DOR, DPSS, FIDD and others) using its title and the first pages of its PDF.
- Builds `RBI Dashboard.html` in the download folder: monthly timeline, filters by type, subject, department, financial year and period, full-text search, and one-click access to each PDF.
- Optional: emails a summary of each new publication with the PDF attached (switched off by default; summaries can use the Claude API).

## Requirements

- Windows 10/11 with Python 3.10 or later
- `pip install -r requirements.txt`

## Setup

1. Put `rbi_downloader.py` and `setup_rbi_downloader.bat` in `C:\Scripts`.
2. Check the download folder in the CONFIG section of `rbi_downloader.py` (`BASE_DIR`, default `S:\RBI - Notifications`).
3. Run `setup_rbi_downloader.bat`, or install the requirements manually.

Secrets are **never stored in the code**. Set them as Windows environment variables if you use email or AI summaries:

```
setx RBI_SMTP_PASSWORD "your-mail-password"
setx ANTHROPIC_API_KEY "your-claude-api-key"
```

## Usage

| Command | What it does |
|---|---|
| `python rbi_downloader.py` | Download anything new and refresh the dashboard |
| `python rbi_downloader.py --dashboard` | Rebuild the dashboard and open it |
| `python rbi_downloader.py --repair` | Remove broken files and re-download them (resumable) |
| `python rbi_downloader.py --fix-names` | Fix garbled characters in saved file names |
| `python rbi_downloader.py --test-email` | Send a test email (when email is enabled) |

To run automatically every 2 hours:

```
schtasks /create /tn "RBI Alerts" /tr "\"<python folder>\pythonw.exe\" \"C:\Scripts\rbi_downloader.py\"" /sc hourly /mo 2 /st 08:00 /f
```

## Customising subjects

Subject groups and their keywords are defined in `TOPIC_RULES` in `rbi_downloader.py`. Add or edit a line, then run `python rbi_downloader.py --dashboard`.

## Notes

- The tool pauses between requests and backs off automatically if RBI's website refuses connections, then resumes on the next run.
- Classification is keyword-based; occasional items may need a rule adjustment.
- Downloaded PDFs, logs and the dashboard stay on the local/shared drive and are excluded from this repository by `.gitignore`.
