"""
Google Scholar Profile Scraper

This script parses an Excel spreadsheet containing Google Scholar profile URLs,
fetches citations, h-index, and i10-index statistics, and updates the spreadsheet
directly. 
To bypass ratelimits, this script includes defensive anti-bot protection.
"""

from __future__ import annotations

import os
import re
import sys
import time
import random
import argparse
import logging
from datetime import datetime
from typing import Optional
from urllib.parse import urlparse, parse_qs
import requests
from bs4 import BeautifulSoup
import openpyxl

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Local file that the Google Sheet is downloaded into (and then updated in place).
LOCAL_WORKBOOK_PATH = "New Google Scholar List.xlsx"

# Google Sheet holding the list of Google Scholar profile URLs.
DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1gt2U6_JpOWlNZrsDbaQ3Y8VS7DSHnKBT/edit?usp=sharing&ouid=102587337650408258618&rtpof=true&sd=true"


def resolve_excel_source(input_arg: str) -> str:
    """
    If input_arg is a Google Sheets URL, download it as .xlsx and save it locally
    (so it gets committed to the repo). Otherwise treat input_arg as a local path.
    """
    if "docs.google.com/spreadsheets" not in input_arg:
        return input_arg

    match = re.match(r"(https://docs\.google\.com/spreadsheets/d/[^/]+)", input_arg)
    if not match:
        logger.error(f"Could not parse Google Sheets URL: {input_arg}")
        return input_arg

    export_url = match.group(1) + "/export?format=xlsx"
    logger.info(f"Downloading Google Sheet from {export_url}")
    try:
        response = requests.get(export_url, timeout=30)
        response.raise_for_status()
    except Exception as e:
        logger.error(f"Failed to download Google Sheet: {e}")
        return input_arg

    with open(LOCAL_WORKBOOK_PATH, "wb") as f:
        f.write(response.content)
    logger.info(f"Saved local copy of Google Sheet to {LOCAL_WORKBOOK_PATH}")
    return LOCAL_WORKBOOK_PATH


def parse_args() -> argparse.Namespace:
    """
    Parse command line arguments.
    """
    parser = argparse.ArgumentParser(description="Google Scholar Profile Scraper and Excel Updater")
    parser.add_argument(
        "--input",
        default=DEFAULT_SHEET_URL,
        help="Path to the Excel file, or a Google Sheets URL (default: Google Sheets URL)"
    )
    parser.add_argument(
        "--min-delay",
        type=float,
        default=5.0,
        help="Minimum delay between requests in seconds (default: 5.0)"
    )
    parser.add_argument(
        "--max-delay",
        type=float,
        default=15.0,
        help="Maximum delay between requests in seconds (default: 15.0)"
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Limit the number of profiles to scrape (default: None)"
    )
    parser.add_argument(
        "--skip-updated",
        action="store_true",
        help="Skip profiles that have already been updated today"
    )
    return parser.parse_args()


def normalize_scholar_url(url: str) -> Optional[str]:
    """
    Clean and extract normalized Google Scholar URL.
    """
    if not url:
        return None
    url = url.strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    try:
        parsed = urlparse(url)
        if "scholar.google" not in parsed.netloc or "citations" not in parsed.path:
            return None
        qs = parse_qs(parsed.query)
        user_ids = qs.get("user")
        if not user_ids:
            return None
        user_id = user_ids[0].strip()
        return f"https://scholar.google.com/citations?user={user_id}&hl=en"
    except Exception as e:
        logger.warning(f"Error parsing URL '{url}': {e}")
        return None


def save_workbook(wb: openpyxl.Workbook, excel_path: str) -> str:
    """
    Attempt to save the workbook.
    If writing is blocked (e.g. PermissionError due to file being open in Excel),
    fallback to saving as filename_YYYY-MM-DD.xlsx in the same folder.
    """
    try:
        wb.save(excel_path)
        return excel_path
    except PermissionError:
        dir_name, base_name = os.path.split(excel_path)
        name_part, ext_part = os.path.splitext(base_name)
        date_suffix = datetime.now().strftime("%Y-%m-%d")
        
        # Check if the name already ends with the date suffix to avoid growing the filename
        if not name_part.endswith(date_suffix):
            new_base = f"{name_part}_{date_suffix}{ext_part}"
        else:
            new_base = base_name
            
        new_path = os.path.join(dir_name, new_base)
        logger.warning(
            f"Permission denied writing to '{excel_path}' (it may be open in Excel). "
            f"Saving fallback file to '{new_path}' instead."
        )
        try:
            wb.save(new_path)
            return new_path
        except Exception as e:
            logger.error(f"Failed to save fallback file '{new_path}': {e}")
            raise e
    except Exception as e:
        logger.error(f"Failed to save workbook to '{excel_path}': {e}")
        raise e


def scrape_profile(session: requests.Session, url: str) -> dict:
    """
    Fetch profile HTML and extract metrics.
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": "https://www.google.com/",
        "Sec-Ch-Ua": '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
        "Sec-Ch-Ua-Mobile": "?0",
        "Sec-Ch-Ua-Platform": '"Windows"'
    }

    try:
        response = session.get(url, headers=headers, timeout=15)
    except Exception as e:
        logger.error(f"Request exception for {url}: {e}")
        return {"status": "error", "error": str(e)}

    if response.status_code == 429:
        return {"status": "rate_limit", "reason": "status_429"}

    if "/sorry/" in response.url:
        return {"status": "rate_limit", "reason": "sorry_redirect"}

    if "accounts.google.com" in response.url:
        return {"status": "login_redirect"}

    if response.status_code != 200:
        return {"status": "error", "error": f"HTTP status {response.status_code}"}

    # Detect captcha / bot block in 200 response
    text = response.text.lower()
    if "captcha" in text or "robot check" in text or "unusual traffic" in text:
        return {"status": "rate_limit", "reason": "captcha_in_html"}

    soup = BeautifulSoup(response.text, "html.parser")
    table = soup.find("table", id="gsc_rsb_st")
    if not table:
        title = soup.title.string.lower() if soup.title else ""
        if "unusual traffic" in title or "captcha" in title or "sorry" in title:
            return {"status": "rate_limit", "reason": "table_missing_captcha_title"}
        return {"status": "not_found", "reason": "table_missing"}

    # Parse metrics
    # The gsc_rsb_st table has two value columns per metric row: "All" (cols[1])
    # and "Since <year>" (cols[2], the one to its right). We only want the
    # "Since <year>" column, not the all-time total.
    rows = table.find_all("tr")
    since_label = None
    if rows:
        header_cols = rows[0].find_all(["td", "th"])
        if len(header_cols) >= 3:
            since_label = header_cols[2].text.strip()  # e.g. "Since 2021"

    metrics = {}
    for row in rows[1:]:  # skip header
        cols = row.find_all(["td", "th"])
        if len(cols) >= 3:
            metric_name = cols[0].text.strip().lower()
            since_val_str = cols[2].text.strip()

            if "citation" in metric_name:
                key = "citations"
            elif "h-index" in metric_name:
                key = "h_index"
            elif "i10-index" in metric_name:
                key = "i10_index"
            else:
                continue

            since_val = 0
            if since_val_str and since_val_str != "-":
                try:
                    since_val = int(since_val_str.replace(",", ""))
                except ValueError:
                    since_val = 0
            metrics[key] = since_val

    return {
        "status": "success",
        "metrics": metrics,
        "since_label": since_label
    }


def main() -> int:
    start_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logger.info(f"Started at: {start_time_str}")

    args = parse_args()
    excel_path = resolve_excel_source(args.input)

    if not os.path.exists(excel_path):
        logger.error(f"Excel file not found at: {excel_path}")
        return 1

    logger.info(f"Loading Excel file: {excel_path}")
    try:
        wb = openpyxl.load_workbook(excel_path)
    except Exception as e:
        logger.error(f"Failed to load Excel file: {e}")
        return 1

    sheet = wb.active
    logger.info(f"Active sheet name: {sheet.title}")

    # Read header row
    headers = [cell.value for cell in sheet[1]]
    logger.info(f"Headers: {headers}")

    # Detect column indices (1-based)
    link_col = None
    citations_col = None
    hindex_col = None
    i10index_col = None
    updated_col = None

    for idx, val in enumerate(headers):
        if not val:
            continue
        val_str = str(val).strip().lower()
        if "google scholar" in val_str or "profile" in val_str or "link" in val_str:
            link_col = idx + 1
        elif "citation" in val_str:
            citations_col = idx + 1
        elif "h-index" in val_str:
            hindex_col = idx + 1
        elif "i10-index" in val_str:
            i10index_col = idx + 1
        elif "last updated" in val_str or "updated" in val_str:
            updated_col = idx + 1

    logger.info("Detected columns:")
    logger.info(f"  Link: {link_col}")
    logger.info(f"  Citations: {citations_col}")
    logger.info(f"  h-index: {hindex_col}")
    logger.info(f"  i10-index: {i10index_col}")
    logger.info(f"  Last Updated: {updated_col}")

    if not link_col:
        logger.error("Could not find Google Scholar profile link column in sheet headers.")
        return 1
    if not citations_col or not hindex_col or not i10index_col:
        logger.error("Could not find all metric columns (Citations, h-index, i10-index).")
        return 1

    # Resolve updated_col if not found
    if not updated_col:
        logger.info("Header 'Last Updated' not found. Searching for existing empty-header columns or appending.")
        # Search for column where header is empty but row 2 has a value
        for idx, val in enumerate(headers):
            if not val:
                val_in_row_2 = sheet.cell(row=2, column=idx+1).value
                if val_in_row_2:
                    updated_col = idx + 1
                    sheet.cell(row=1, column=updated_col, value="Last Updated")
                    logger.info(f"  Using empty-header column {updated_col} (contains row 2 value '{val_in_row_2}') as 'Last Updated'")
                    break
        else:
            # Append a new column
            last_non_empty = 0
            for idx, val in enumerate(headers):
                if val:
                    last_non_empty = idx + 1
            updated_col = last_non_empty + 1
            sheet.cell(row=1, column=updated_col, value="Last Updated")
            logger.info(f"  Appending new column {updated_col} as 'Last Updated'")
        # Save change to header
        excel_path = save_workbook(wb, excel_path)

    # Scrape loop setup
    session = requests.Session()

    today_str = datetime.now().strftime("%Y-%m-%d")
    rows_to_process = []
    for row_idx in range(2, sheet.max_row + 1):
        cell_val = sheet.cell(row=row_idx, column=link_col).value
        if cell_val:
            normalized_url = normalize_scholar_url(str(cell_val))
            if normalized_url:
                if args.skip_updated and updated_col:
                    last_updated_val = sheet.cell(row=row_idx, column=updated_col).value
                    if last_updated_val and str(last_updated_val).startswith(today_str):
                        logger.info(f"Row {row_idx}: Already updated today ({last_updated_val}). Skipping.")
                        continue
                rows_to_process.append((row_idx, cell_val, normalized_url))
            else:
                logger.debug(f"Row {row_idx}: URL '{cell_val}' could not be normalized.")

    total_profiles = len(rows_to_process)
    logger.info(f"Found {total_profiles} rows with valid Google Scholar Profile URLs.")

    if args.limit is not None:
        rows_to_process = rows_to_process[:args.limit]
        logger.info(f"Limit applied. Processing first {len(rows_to_process)} profiles.")

    # Counters
    success_count = 0
    fail_count = 0
    skip_count = 0

    backoff_delay = 60.0
    max_backoff = 600.0
    headers_renamed = False

    for idx, (row_idx, raw_url, url) in enumerate(rows_to_process, 1):
        first_name = sheet.cell(row=row_idx, column=2).value or ""
        last_name = sheet.cell(row=row_idx, column=1).value or ""
        logger.info(f"[{idx}/{len(rows_to_process)}] Processing {first_name} {last_name} ({url})")

        retries = 0
        max_retries = 5
        scraped_data = None

        while retries < max_retries:
            scraped_data = scrape_profile(session, url)
            status = scraped_data["status"]

            if status == "success":
                backoff_delay = 60.0  # Reset backoff on success
                break
            elif status == "rate_limit":
                # Save progress immediately before sleep/backoff
                logger.warning("Rate limit or CAPTCHA detected. Saving spreadsheet progress and backing off.")
                excel_path = save_workbook(wb, excel_path)

                reason = scraped_data.get("reason", "unknown")
                logger.warning(f"Rate limit reason: {reason}. Sleeping for {backoff_delay} seconds (retry {retries+1}/{max_retries})...")
                time.sleep(backoff_delay)
                backoff_delay = min(backoff_delay * 2, max_backoff)
                retries += 1
            elif status == "login_redirect":
                logger.warning("Redirected to Google sign-in. Profile may be invalid or deleted. Skipping profile.")
                break
            elif status == "not_found":
                logger.warning("Google Scholar profile elements not found (missing table). Profile might be invalid/empty. Skipping profile.")
                break
            else:
                logger.error(f"Error scraping profile (attempt {retries+1}/{max_retries}): {scraped_data.get('error', 'unknown error')}")
                retries += 1
                if retries < max_retries:
                    logger.info("Retrying after 5 seconds...")
                    time.sleep(5)

        if scraped_data and scraped_data["status"] == "success":
            metrics = scraped_data["metrics"]
            citations = metrics.get("citations", 0)
            h_index = metrics.get("h_index", 0)
            i10_index = metrics.get("i10_index", 0)
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # Rename headers once to reflect the "Since <year>" values now being
            # written, instead of the previous "All" totals.
            if not headers_renamed:
                since_label = scraped_data.get("since_label")
                if since_label:
                    sheet.cell(row=1, column=citations_col, value=f"Citations ({since_label})")
                    sheet.cell(row=1, column=hindex_col, value=f"h-index ({since_label})")
                    sheet.cell(row=1, column=i10index_col, value=f"i10-index ({since_label})")
                    logger.info(f"Renamed metric headers to reflect '{since_label}' values.")
                headers_renamed = True

            # Write values to cells
            sheet.cell(row=row_idx, column=citations_col, value=citations)
            sheet.cell(row=row_idx, column=hindex_col, value=h_index)
            sheet.cell(row=row_idx, column=i10index_col, value=i10_index)
            sheet.cell(row=row_idx, column=updated_col, value=current_time)

            logger.info(f"  Updated: Citations={citations}, h-index={h_index}, i10-index={i10_index}, Date={current_time}")

            # Save spreadsheet immediately
            excel_path = save_workbook(wb, excel_path)
            success_count += 1
        elif scraped_data and scraped_data["status"] in ("login_redirect", "not_found"):
            logger.info("  Skipped: Profile could not be parsed.")
            skip_count += 1
        else:
            logger.error(f"  Failed to scrape after {max_retries} attempts.")
            fail_count += 1

        # Apply random delay between requests (if not the last request)
        if idx < len(rows_to_process):
            delay = random.uniform(args.min_delay, args.max_delay)
            logger.info(f"Sleeping for {delay:.2f} seconds...")
            time.sleep(delay)

    logger.info("Scrape session completed.")
    logger.info(f"Summary: Success={success_count}, Skipped={skip_count}, Failed={fail_count}")

    end_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    logger.info(f"Ended at: {end_time_str}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
