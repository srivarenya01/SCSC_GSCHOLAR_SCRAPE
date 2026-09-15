# SCSC Google Scholar Scraper

An automated tool to scrape Google Scholar profiles for SCSC faculty, tracking citations, h-index, and i10-index (all-time and the rolling Since window Scholar reports) over time.

## Features

- **Profile Scraping**: Visits each faculty member's Google Scholar profile and extracts citation metrics.
- **Google Sheets Source**: Reads the faculty list from a Google Sheet as xlsx, then writes results to csv.
- **Anti-Bot Defenses**: Randomized delays, CAPTCHA/rate-limit detection, and exponential backoff between requests.
- **Automated Scheduling**: Includes a GitHub Actions workflow to run the scraper **daily** and commit results back to the repository.

## Project Structure

- `scholar_scraper.py`: Downloads the Google Sheet as xlsx, scrapes profiles, and writes csv.
- `New Google Scholar List.csv`: Scraped metrics. The xlsx download is source-only and is not committed.
- `.github/workflows/scraper.yml`: GitHub Actions configuration for automated daily runs.

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

## Automation (GitHub Actions)

The included GitHub Action is configured to:
1. Run **daily** at 14:00 UTC (08:00 AM Central Time).
2. Download the latest Google Sheet (xlsx), scrape each profile, and write csv.
3. Automatically commit and push the updated `New Google Scholar List.csv` back to the repository.

You can also trigger a run manually via the **Actions** tab in your GitHub repository by selecting the "SCSC Google Scholar Scraper" workflow and clicking "Run workflow".

## Output Columns (CSV)

The `New Google Scholar List.csv` file includes the following columns:
- `Last Name` / `First Name`: Faculty name
- `Rank`, `adloc`, `Location`: Faculty metadata
- `Link to Google Scholar Profile`: Source profile URL
- `Citations (All)`, `h-index (All)`, `i10-index (All)`: All-time totals
- `Citations (Since <year>)`, `h-index (Since <year>)`, `i10-index (Since <year>)`: Scholar's rolling window (currently Since 2021)
- `Last Updated`: Timestamp of the most recent successful scrape for that row
