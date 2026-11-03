"""
Stage A (Steps 1-3): Download and label 10-K filings from SEC EDGAR.

What this script does:
    1. Prompts the user for a ticker symbol and how many 10-Ks to download.
    2. Looks up the company's CIK number and its filing history on EDGAR.
    3. Downloads the most recent N 10-K documents.
    4. Saves each file as  TICKER_10K_YYYY-MM-DD.htm  (date = fiscal period end)
       inside  filings/TICKER/  next to this script.

Run with:
    python3 stage_a_download_10k.py
"""

import sys
import time
from pathlib import Path

import requests

# --------------------------------------------------------------------------
# CONFIGURATION -- EDIT THIS LINE
# The SEC requires every automated request to identify the requester.
# Use your real name and school email, e.g. "Jane Smith jsmith@school.edu"
# --------------------------------------------------------------------------
USER_AGENT = "YOUR NAME your.email@school.edu"

# Where files are saved: a "filings" folder next to this script.
BASE_DIR = Path(__file__).resolve().parent
FILINGS_DIR = BASE_DIR / "filings"

# SEC allows at most 10 requests per second; we stay well under that.
REQUEST_DELAY_SECONDS = 0.25

TICKER_LOOKUP_URL = "https://www.sec.gov/files/company_tickers.json"
SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"
SUBMISSIONS_EXTRA_URL = "https://data.sec.gov/submissions/{name}"
ARCHIVE_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{accession}/{doc}"


def sec_get(url):
    """Send a polite GET request to the SEC and return the response.

    Raises:
        requests.HTTPError: if the SEC returns an error status code.
    """
    time.sleep(REQUEST_DELAY_SECONDS)
    response = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=30)
    response.raise_for_status()
    return response


def get_user_inputs():
    """Ask the user for a ticker symbol and the number of 10-Ks to download."""
    ticker = input("Enter the ticker symbol (e.g., AAPL): ").strip().upper()
    if not ticker:
        sys.exit("No ticker entered. Exiting.")

    raw_count = input("How many of the most recent 10-Ks? (default 10): ").strip()
    if raw_count == "":
        count = 10
    elif raw_count.isdigit() and int(raw_count) > 0:
        count = int(raw_count)
    else:
        sys.exit("Please enter a positive whole number. Exiting.")
    return ticker, count


def get_cik(ticker):
    """Convert a ticker symbol to a zero-padded 10-digit CIK string."""
    data = sec_get(TICKER_LOOKUP_URL).json()
    for entry in data.values():
        if entry["ticker"].upper() == ticker:
            return str(entry["cik_str"]).zfill(10)
    sys.exit(f"Ticker '{ticker}' was not found on EDGAR.")


def collect_10k_filings(cik):
    """Return a list of 10-K filings (newest first) for the given CIK.

    Each item is a dict with: accession, document, report_date, filing_date.
    EDGAR stores recent filings in one block and older ones in extra files,
    so both are checked.
    """
    main = sec_get(SUBMISSIONS_URL.format(cik=cik)).json()
    blocks = [main["filings"]["recent"]]

    for extra in main["filings"].get("files", []):
        blocks.append(sec_get(SUBMISSIONS_EXTRA_URL.format(name=extra["name"])).json())

    filings = []
    for block in blocks:
        for i, form in enumerate(block["form"]):
            if form == "10-K":  # excludes 10-K/A amendments
                filings.append(
                    {
                        "accession": block["accessionNumber"][i],
                        "document": block["primaryDocument"][i],
                        "report_date": block["reportDate"][i],
                        "filing_date": block["filingDate"][i],
                    }
                )

    filings.sort(key=lambda f: f["report_date"], reverse=True)
    return filings


def download_filings(ticker, cik, filings):
    """Download each filing and save it as TICKER_10K_YYYY-MM-DD.htm."""
    save_dir = FILINGS_DIR / ticker
    save_dir.mkdir(parents=True, exist_ok=True)

    for filing in filings:
        file_name = f"{ticker}_10K_{filing['report_date']}.htm"
        save_path = save_dir / file_name

        if save_path.exists():
            print(f"  Skipping (already saved): {file_name}")
            continue

        url = ARCHIVE_URL.format(
            cik=int(cik),
            accession=filing["accession"].replace("-", ""),
            doc=filing["document"],
        )
        try:
            response = sec_get(url)
            save_path.write_bytes(response.content)
            print(f"  Saved: {file_name}")
        except requests.RequestException as error:
            print(f"  FAILED: {file_name} ({error})")


def main():
    """Run the full Stage A workflow."""
    if "YOUR NAME" in USER_AGENT:
        sys.exit("Please edit USER_AGENT near the top of the script first.")

    ticker, count = get_user_inputs()

    try:
        print(f"\nLooking up {ticker} on EDGAR...")
        cik = get_cik(ticker)
        print(f"Found CIK: {cik}")

        filings = collect_10k_filings(cik)[:count]
        if not filings:
            sys.exit("No 10-K filings were found for this company.")

        print(f"Downloading {len(filings)} 10-K filing(s)...")
        download_filings(ticker, cik, filings)
    except requests.RequestException as error:
        sys.exit(f"Network or SEC error: {error}")

    print(f"\nDone. Files are in: {FILINGS_DIR / ticker}")


if __name__ == "__main__":
    main()
