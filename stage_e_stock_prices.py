"""
Stage E (Step 8): Add the stock price on each balance sheet date.

What this script does:
    1. Asks for a ticker symbol (Stages C and D must be done first).
    2. Reads the ten balance sheet dates from the "Balance Sheet" tab of
       output/TICKER_financials.xlsx.
    3. Downloads the daily closing prices from Yahoo Finance (yfinance).
    4. For each balance sheet date, takes the closing price on that day. If
       the market was closed (weekend or holiday), it uses the LAST TRADING
       DAY BEFORE that date, and shows which day it used.
    5. Adds a "Stock Price" tab to the workbook.

Prices are adjusted for stock splits (for example Apple's 4-for-1 split in
2020), so prices from different years can be compared directly.

First time only, install the package:
    pip3 install yfinance

Run with:
    python3 stage_e_stock_prices.py
"""

import sys
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

try:
    import yfinance as yf
except ImportError:
    sys.exit("The yfinance package is not installed.\n"
             "Run this first:  pip3 install yfinance")

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"

SOURCE_SHEET = "Balance Sheet"
SHEET_NAME = "Stock Price"
PRICE_ROW_LABEL = "Closing Stock Price ($ per share)"
USED_ROW_LABEL = "Trading Day Used"

PRICE_FORMAT = '"$"#,##0.00_);("$"#,##0.00)'


def read_balance_sheet_dates(workbook):
    """Return the balance sheet dates (YYYY-MM-DD text) from row 1."""
    if SOURCE_SHEET not in workbook.sheetnames:
        sys.exit(f"No '{SOURCE_SHEET}' tab found. Run Stage C first.")
    sheet = workbook[SOURCE_SHEET]
    dates = [str(cell.value) for cell in sheet[1][1:] if cell.value]
    if not dates:
        sys.exit("No dates found in the Balance Sheet tab.")
    return dates


def download_prices(ticker, dates):
    """Return a table of daily closing prices covering all the dates."""
    first = datetime.strptime(min(dates), "%Y-%m-%d") - timedelta(days=14)
    last = datetime.strptime(max(dates), "%Y-%m-%d") + timedelta(days=3)
    try:
        history = yf.Ticker(ticker).history(
            start=first.strftime("%Y-%m-%d"),
            end=last.strftime("%Y-%m-%d"),
            auto_adjust=False,  # split-adjusted only, not dividend-adjusted
        )
    except Exception as error:  # network problems, blocked requests, etc.
        sys.exit(f"Could not download prices for {ticker}: {error}")
    if history.empty:
        sys.exit(f"Yahoo Finance returned no prices for {ticker}. "
                 "Check your internet connection and try again.")
    index = history.index
    if getattr(index, "tz", None) is not None:
        index = index.tz_localize(None)
    history.index = index.normalize()
    return history["Close"]


def price_on_or_before(closes, date_text):
    """Return (trading_day_text, close) for the last trading day <= date."""
    target = datetime.strptime(date_text, "%Y-%m-%d")
    available = closes.loc[:target].dropna()
    if available.empty:
        return None, None
    return available.index[-1].strftime("%Y-%m-%d"), float(available.iloc[-1])


def write_price_tab(workbook, dates, results, workbook_path):
    """Add (or replace) the Stock Price tab."""
    if SHEET_NAME in workbook.sheetnames:
        del workbook[SHEET_NAME]
    sheet = workbook.create_sheet(SHEET_NAME)

    sheet.cell(row=1, column=1, value=SHEET_NAME)
    for col, date in enumerate(dates, start=2):
        sheet.cell(row=1, column=col, value=date)

    sheet.cell(row=2, column=1, value=PRICE_ROW_LABEL)
    sheet.cell(row=3, column=1, value=USED_ROW_LABEL)
    for col, date in enumerate(dates, start=2):
        used, price = results[date]
        price_cell = sheet.cell(row=2, column=col, value=price)
        price_cell.number_format = PRICE_FORMAT
        used_cell = sheet.cell(row=3, column=col, value=used)
        used_cell.alignment = Alignment(horizontal="right")

    note = ("Closing price on the balance sheet date, or on the last trading "
            "day before it if the market was closed. Adjusted for stock "
            "splits, in US dollars. Source: Yahoo Finance (yfinance).")
    sheet.cell(row=5, column=1, value=note).font = Font(
        italic=True, color="595959")

    sheet.freeze_panes = "B2"
    sheet.column_dimensions["A"].width = 42
    for col in range(2, len(dates) + 2):
        sheet.column_dimensions[get_column_letter(col)].width = 15
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    sheet["A1"].alignment = Alignment(horizontal="left")

    workbook.save(workbook_path)


def main():
    """Look up the price on each balance sheet date and save the tab."""
    ticker = input("Enter the ticker symbol (e.g., AAPL): ").strip().upper()
    if not ticker:
        sys.exit("No ticker entered. Exiting.")

    workbook_path = OUTPUT_DIR / f"{ticker}_financials.xlsx"
    if not workbook_path.exists():
        sys.exit(f"{workbook_path} not found. Run Stage C first.")

    workbook = load_workbook(workbook_path)
    dates = read_balance_sheet_dates(workbook)
    print(f"\nGetting prices for {ticker} on {len(dates)} balance sheet dates...\n")

    closes = download_prices(ticker, dates)

    results = {}
    for date in dates:
        used, price = price_on_or_before(closes, date)
        results[date] = (used, price)
        if price is None:
            print(f"  {date}: NO PRICE FOUND")
            continue
        note = "" if used == date else f"  (market closed; used {used})"
        print(f"  {date}: ${price:,.2f}{note}")

    write_price_tab(workbook, dates, results, workbook_path)
    print(f"\n'{SHEET_NAME}' tab added to: {workbook_path}")


if __name__ == "__main__":
    main()
