"""
Stage B (Step 4): Extract the MD&A section from each downloaded 10-K.

What this script does:
    1. Asks for a ticker symbol (must already be downloaded by Stage A).
    2. Reads every  filings/TICKER/TICKER_10K_YYYY-MM-DD.htm  file.
    3. Converts the HTML to plain text.
    4. Finds Item 7 (Management's Discussion and Analysis) and cuts it out,
       stopping at Item 7A (or Item 8 if 7A is missing).
    5. Saves each result as  mdna/TICKER/TICKER_MDNA_YYYY-MM-DD.txt

Run with:
    python3 stage_b_extract_mdna.py
"""

import re
import sys
from pathlib import Path

from bs4 import BeautifulSoup

BASE_DIR = Path(__file__).resolve().parent
FILINGS_DIR = BASE_DIR / "filings"
MDNA_DIR = BASE_DIR / "mdna"

# Sections shorter than this are probably a mistake (e.g., a table of contents).
MIN_EXPECTED_CHARS = 5000

# Characters allowed between "Item 7" and the title: spaces, periods, dashes.
SEP = r"[\s.:\-\u2013\u2014]*"

# The heading that STARTS the MD&A section.
START_PATTERN = re.compile(
    r"item\s*7" + SEP + r"management.{0,3}s\s+discussion\s+and\s+analysis",
    re.IGNORECASE,
)

# Headings that END the MD&A section.
END_PATTERN = re.compile(
    r"item\s*7a" + SEP + r"quantitative"
    r"|item\s*8" + SEP + r"(consolidated\s+)?financial\s+statements",
    re.IGNORECASE,
)


def html_to_text(html_path):
    """Read a 10-K .htm file and return clean plain text.

    Hidden XBRL header data is removed so it does not pollute the text.
    """
    raw_html = html_path.read_bytes()
    soup = BeautifulSoup(raw_html, "lxml")

    # Remove hidden inline-XBRL metadata blocks, scripts, and styles.
    for tag in soup.find_all(["ix:header", "script", "style"]):
        tag.decompose()

    text = soup.get_text(separator="\n")

    # Normalize whitespace so the heading patterns match reliably.
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    return text


def extract_mdna(text):
    """Return the MD&A section from the full 10-K text, or None if not found.

    The table of contents also contains 'Item 7 ... Management's Discussion',
    so every possible start is tried and the LONGEST section is kept.
    """
    best_section = None

    for start_match in START_PATTERN.finditer(text):
        end_match = END_PATTERN.search(text, start_match.end())
        if end_match is None:
            continue
        section = text[start_match.start():end_match.start()]
        if best_section is None or len(section) > len(best_section):
            best_section = section

    return best_section.strip() if best_section else None


def get_ticker():
    """Ask the user which ticker to process."""
    ticker = input("Enter the ticker symbol (e.g., AAPL): ").strip().upper()
    if not ticker:
        sys.exit("No ticker entered. Exiting.")
    return ticker


def main():
    """Extract the MD&A from every downloaded filing for one ticker."""
    ticker = get_ticker()
    source_dir = FILINGS_DIR / ticker

    if not source_dir.exists():
        sys.exit(f"Folder not found: {source_dir}\nRun Stage A for {ticker} first.")

    html_files = sorted(source_dir.glob(f"{ticker}_10K_*.htm"), reverse=True)
    if not html_files:
        sys.exit(f"No 10-K files found in {source_dir}.")

    output_dir = MDNA_DIR / ticker
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nProcessing {len(html_files)} filing(s) for {ticker}...\n")

    for html_path in html_files:
        report_date = html_path.stem.replace(f"{ticker}_10K_", "")
        output_path = output_dir / f"{ticker}_MDNA_{report_date}.txt"

        try:
            text = html_to_text(html_path)
            mdna = extract_mdna(text)
        except Exception as error:  # keep going if one file fails
            print(f"  {report_date}: ERROR ({error})")
            continue

        if mdna is None:
            print(f"  {report_date}: MD&A NOT FOUND - check this file manually")
            continue

        output_path.write_text(mdna, encoding="utf-8")
        word_count = len(mdna.split())
        warning = "  <-- looks short, please check" if len(mdna) < MIN_EXPECTED_CHARS else ""
        print(f"  {report_date}: saved {word_count:,} words{warning}")

    print(f"\nDone. Files are in: {output_dir}")


if __name__ == "__main__":
    main()
