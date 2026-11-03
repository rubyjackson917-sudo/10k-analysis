"""
Stage D (Steps 6-7): Clean the MD&A text and score it with the
Loughran-McDonald (LM) Master Dictionary.

What this script does:
    1. Asks for a ticker symbol (Stages A, B and C must be done first).
    2. Cleans each  mdna/TICKER/TICKER_MDNA_YYYY-MM-DD.txt  file:
       lowercase, remove numbers/punctuation/URLs, drop one-letter tokens.
       Cleaned text is saved to  cleaned/TICKER/TICKER_CLEAN_YYYY-MM-DD.txt
    3. Counts LM words (positive, negative, uncertainty, litigious, etc.).
    4. Adds a "Sentiment" tab to  output/TICKER_financials.xlsx
       (years side by side, oldest on the left, like the other tabs).

Needs the LM Master Dictionary CSV in this folder (the file name must
contain "Loughran" and "MasterDictionary").

Run with:
    python3 stage_d_sentiment.py
"""

import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
MDNA_DIR = BASE_DIR / "mdna"
CLEANED_DIR = BASE_DIR / "cleaned"
OUTPUT_DIR = BASE_DIR / "output"

SHEET_NAME = "Sentiment"

# LM dictionary columns we score (a word counts if its column is non-zero).
CATEGORIES = [
    "Positive", "Negative", "Uncertainty", "Litigious",
    "Constraining", "Strong_Modal", "Weak_Modal",
]


# --------------------------------------------------------------------------
# Step 6: cleaning
# --------------------------------------------------------------------------
def clean_text(text):
    """Return a list of clean lowercase word tokens from raw MD&A text."""
    text = text.lower()
    text = re.sub(r"https?://\S+|www\.\S+", " ", text)   # web addresses
    text = re.sub(r"[\u2019']s\b", "", text)             # possessives: apple's
    text = text.replace("\u2019", "").replace("'", "")   # other apostrophes
    text = re.sub(r"[^a-z\s]", " ", text)                # numbers, punctuation
    return [token for token in text.split() if len(token) > 1]


def save_cleaned(tokens, path, words_per_line=20):
    """Save tokens to a text file, 20 words per line for easy reading."""
    lines = [" ".join(tokens[i:i + words_per_line])
             for i in range(0, len(tokens), words_per_line)]
    path.write_text("\n".join(lines), encoding="utf-8")


# --------------------------------------------------------------------------
# Step 7: Loughran-McDonald scoring
# --------------------------------------------------------------------------
def find_dictionary():
    """Locate the LM Master Dictionary CSV in the project folder."""
    for path in BASE_DIR.glob("*.csv"):
        name = path.name.lower()
        if "loughran" in name and "master" in name:
            return path
    sys.exit(
        "Could not find the Loughran-McDonald Master Dictionary CSV.\n"
        f"Put it in {BASE_DIR} (file name should contain "
        "'Loughran' and 'MasterDictionary')."
    )


def load_word_sets(dictionary_path):
    """Return {category: set of lowercase words} from the LM CSV."""
    frame = pd.read_csv(dictionary_path)
    frame["Word"] = frame["Word"].astype(str).str.lower()

    word_sets = {}
    for category in CATEGORIES:
        if category in frame.columns:
            word_sets[category] = set(frame.loc[frame[category] > 0, "Word"])
        else:
            print(f"  Note: column '{category}' not in dictionary; skipped.")
    return word_sets


def score_tokens(tokens, word_sets):
    """Return a dict of sentiment metrics for one document."""
    total = len(tokens)
    counts = {name: sum(1 for t in tokens if t in words)
              for name, words in word_sets.items()}

    positive = counts.get("Positive", 0)
    negative = counts.get("Negative", 0)

    metrics = {"Total Word Count": total}
    for name, count in counts.items():
        metrics[f"{name.replace('_', ' ')} Words"] = count
    metrics["Positive % of Words"] = positive / total if total else None
    metrics["Negative % of Words"] = negative / total if total else None
    metrics["P/N Ratio (Positive / Negative)"] = (
        positive / negative if negative else None)
    metrics["Net Sentiment ((Pos - Neg) / Total)"] = (
        (positive - negative) / total if total else None)
    return metrics


# --------------------------------------------------------------------------
# Writing the Excel tab
# --------------------------------------------------------------------------
def write_sentiment_tab(results, workbook_path):
    """Add (or replace) the Sentiment tab in the company's workbook.

    Rows are grouped under shaded headings, key rows are bold, and a
    footnote defines the metrics. Row labels are unchanged, so later steps
    can still find each metric by name.
    """
    if workbook_path.exists():
        workbook = load_workbook(workbook_path)
    else:
        workbook = Workbook()
        workbook.remove(workbook.active)

    if SHEET_NAME in workbook.sheetnames:
        del workbook[SHEET_NAME]
    sheet = workbook.create_sheet(SHEET_NAME)

    dates = sorted(results)
    metric_names = list(results[dates[0]].keys())

    share_rows = [n for n in metric_names if "%" in n or "Net Sentiment" in n]
    ratio_rows = [n for n in metric_names if "Ratio" in n]
    count_rows = [n for n in metric_names
                  if n not in share_rows and n not in ratio_rows]
    groups = [("Word counts", count_rows),
              ("Share of total words", share_rows),
              ("Ratio", ratio_rows)]

    heading_fill = PatternFill("solid", fgColor="E7E6E6")
    bold_rows = {"Total Word Count", "P/N Ratio (Positive / Negative)"}

    sheet.cell(row=1, column=1, value="Sentiment")
    for col, date in enumerate(dates, start=2):
        sheet.cell(row=1, column=col, value=date)

    row = 2
    row_of = {}
    for title, names in groups:
        if not names:
            continue
        for col in range(1, len(dates) + 2):  # shaded heading row
            sheet.cell(row=row, column=col).fill = heading_fill
        heading = sheet.cell(row=row, column=1, value=title)
        heading.font = Font(bold=True)
        row += 1
        for name in names:
            row_of[name] = row
            label = sheet.cell(row=row, column=1, value=name)
            label.alignment = Alignment(indent=1)
            bold = name in bold_rows
            if bold:
                label.font = Font(bold=True)
            for col, date in enumerate(dates, start=2):
                cell = sheet.cell(row=row, column=col,
                                  value=results[date][name])
                if "%" in name or "Net Sentiment" in name:
                    cell.number_format = "0.00%"
                elif "Ratio" in name:
                    cell.number_format = "0.00"
                else:
                    cell.number_format = "#,##0"
                if bold:
                    cell.font = Font(bold=True)
            row += 1

    # Subtle red-to-green colour scale on the two "how positive" rows.
    last_col = get_column_letter(len(dates) + 1)
    for name in share_rows + ratio_rows:
        if name in row_of and ("Ratio" in name or "Net Sentiment" in name):
            cell_range = f"B{row_of[name]}:{last_col}{row_of[name]}"
            sheet.conditional_formatting.add(cell_range, ColorScaleRule(
                start_type="min", start_color="F8CBAD",
                mid_type="percentile", mid_value=50, mid_color="FFFFFF",
                end_type="max", end_color="C6E0B4"))

    note = ("Words are counted after cleaning (lowercase; numbers and "
            "punctuation removed). Word categories come from the "
            "Loughran-McDonald Master Dictionary. P/N Ratio = positive words "
            "/ negative words. Net Sentiment = (positive - negative) / total "
            "words.")
    sheet.cell(row=row + 1, column=1, value=note).font = Font(
        italic=True, color="595959")

    sheet.freeze_panes = "B2"
    sheet.column_dimensions["A"].width = 42
    for col in range(2, len(dates) + 2):
        sheet.column_dimensions[get_column_letter(col)].width = 15
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    sheet["A1"].alignment = Alignment(horizontal="left")

    OUTPUT_DIR.mkdir(exist_ok=True)
    workbook.save(workbook_path)


def main():
    """Clean and score every MD&A file for one ticker."""
    ticker = input("Enter the ticker symbol (e.g., AAPL): ").strip().upper()
    if not ticker:
        sys.exit("No ticker entered. Exiting.")

    mdna_files = sorted((MDNA_DIR / ticker).glob(f"{ticker}_MDNA_*.txt"))
    if not mdna_files:
        sys.exit(f"No MD&A files found in {MDNA_DIR / ticker}. Run Stage B first.")

    dictionary_path = find_dictionary()
    print(f"\nUsing dictionary: {dictionary_path.name}")
    word_sets = load_word_sets(dictionary_path)

    cleaned_dir = CLEANED_DIR / ticker
    cleaned_dir.mkdir(parents=True, exist_ok=True)

    results = {}
    print(f"Cleaning and scoring {len(mdna_files)} file(s) for {ticker}...\n")

    for path in mdna_files:
        date = path.stem.replace(f"{ticker}_MDNA_", "")
        tokens = clean_text(path.read_text(encoding="utf-8"))
        save_cleaned(tokens, cleaned_dir / f"{ticker}_CLEAN_{date}.txt")
        results[date] = score_tokens(tokens, word_sets)

        m = results[date]
        ratio = m["P/N Ratio (Positive / Negative)"]
        ratio_text = f"{ratio:.2f}" if ratio is not None else "n/a"
        print(f"  {date}: {m['Total Word Count']:>6,} words | "
              f"pos {m['Positive Words']:>4} | neg {m['Negative Words']:>4} | "
              f"P/N {ratio_text}")

    workbook_path = OUTPUT_DIR / f"{ticker}_financials.xlsx"
    write_sentiment_tab(results, workbook_path)
    print(f"\nCleaned text saved in: {cleaned_dir}")
    print(f"'{SHEET_NAME}' tab added to: {workbook_path}")


if __name__ == "__main__":
    main()
