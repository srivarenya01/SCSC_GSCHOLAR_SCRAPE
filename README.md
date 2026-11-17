# SCSC Google Scholar Scraper

An automated tool to scrape Google Scholar profiles for SCSC faculty, tracking citations, h-index, and i10-index (all-time and the rolling Since window Scholar reports) over time.

## Features

- **Profile Scraping**: Visits each faculty member's Google Scholar profile and extracts citation metrics.
- **Google Sheets Source**: Reads the faculty list from a Google Sheet as xlsx, then writes results to csv.
- **Anti-Bot Defenses**: Randomized delays, CAPTCHA/rate-limit detection, and exponential backoff between requests.
- **Windows server**: Double-click `run_scraper.bat`, or install a Task Scheduler job for Monday and Thursday at 08:00.

## Project Structure

- `scholar_scraper.py`: Downloads the Google Sheet as xlsx, scrapes profiles, and writes csv.
- `New Google Scholar List.csv`: Scraped metrics. The xlsx download is source-only and is not committed.
- `run_scraper.bat`: Click to run on a Windows machine. Restores the previous csv if citations come back empty, then pushes csv to GitHub.
- `install_schedule.bat` / `uninstall_schedule.bat`: Create or remove a Task Scheduler job (Monday and Thursday at 08:00).
- `.github/workflows/scraper.yml`: Manual GitHub Action only. Daily GitHub runs are off because Scholar blocks those runners.

## Local Setup

### Prerequisites

- Python 3.8+
- Recommended: A virtual environment

```bash
# Create and activate a virtual environment
python -m venv venv
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate
```

### Installation

Install the required dependencies:

```bash
pip install requests beautifulsoup4 openpyxl
```

## Usage

Run the scraper with default settings:

```bash
python scholar_scraper.py
```

### Optional Arguments

- `--input`: Google Sheets URL or local path to the Excel file (default: Google Sheets URL).
- `--min-delay` / `--max-delay`: Random delay range in seconds between requests (default: 5.0 / 15.0).
- `--limit`: Limit the number of profiles scraped in one run (default: none).
- `--skip-updated`: Skip profiles already updated today.

## Windows server

1. Install Python 3, then `pip install requests beautifulsoup4 openpyxl` (or use `venv` in this folder).
2. Double-click `run_scraper.bat` to scrape once. Watch `logs\scraper_run.log`.
3. Double-click `install_schedule.bat` to run Monday and Thursday at 08:00 while you are logged on (lock the screen is fine; sign-out is not).
4. The machine needs git credentials if you want the csv pushed to GitHub (IMPORTDATA). Set `PUSH_TO_GITHUB=0` at the top of `run_scraper.bat` to skip that.

If a run gets empty citation columns (Scholar block), the bat puts the previous csv back and does not push.

## GitHub Actions

Daily GitHub scrape is off. You can still run the workflow by hand from the Actions tab. Prefer the Windows bat on your server.

## Output Columns (CSV)

The `New Google Scholar List.csv` file includes the following columns:
- `Last Name` / `First Name`: Faculty name
- `Rank`, `adloc`, `Location`: Faculty metadata
- `Link to Google Scholar Profile`: Source profile URL
- `Citations (All)`, `h_index (All)`, `i10_index (All)`: All-time totals
- `Citations (Since <year>)`, `h_index (Since <year>)`, `i10_index (Since <year>)`: Scholar's rolling window (currently Since 2021)
- `Last Updated`: Timestamp of the most recent successful scrape for that row
