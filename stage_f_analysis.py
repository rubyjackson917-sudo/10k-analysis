"""
Stage F (Step 9): Ratios, sentiment summary, and the price-vs-sentiment chart.

What this script does:
    1. Asks for a ticker symbol. Stages C, D and E must be done first, so the
       workbook already has the statements, Sentiment and Stock Price tabs.
    2. Adds an "Analysis" tab to output/TICKER_financials.xlsx with:
         * five financial ratios for each of the ten years (liquidity,
           solvency, efficiency and two profitability measures),
         * a sentiment summary (word counts, positive / negative words and
           the P/N ratio),
         * the closing stock price,
         * the correlation between stock price and the P/N ratio,
         * a chart of stock price versus the P/N ratio.
    3. Every number in the Analysis tab is an Excel FORMULA that points at
       the other tabs, so you can click a cell and see exactly where it
       comes from.
    4. Prints the ratios in the Terminal so you can check them.

Run with:
    python3 stage_f_analysis.py
"""

import re
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
SHEET_NAME = "Analysis"
REQUIRED_SHEETS = ["Income Statement", "Balance Sheet", "Sentiment", "Stock Price"]

# Row labels to look up (regular expressions, matched against column A).
ROW_PATTERNS = {
    "revenue": ("Income Statement", r"^total (net sales|revenues?)"),
    "net_income": ("Income Statement", r"^net income"),
    "current_assets": ("Balance Sheet", r"^total current assets"),
    "current_liabilities": ("Balance Sheet", r"^total current liabilities"),
    "total_assets": ("Balance Sheet", r"^total assets"),
    "total_liabilities": ("Balance Sheet", r"^total liabilities$"),
    "equity": ("Balance Sheet", r"^total (share|stock)holders.{0,2}equity"),
    "price": ("Stock Price", r"^closing stock price"),
    "pn_ratio": ("Sentiment", r"^p/n ratio"),
}

# (name shown, top row, bottom row, number format, plain-English meaning)
RATIOS = [
    ("Liquidity: Current Ratio", "current_assets", "current_liabilities",
     "0.00", "Total current assets / Total current liabilities"),
    ("Solvency: Debt-to-Equity", "total_liabilities", "equity",
     "0.00", "Total liabilities / Total shareholders' equity"),
    ("Efficiency: Asset Turnover", "revenue", "total_assets",
     "0.00", "Total revenue / Total assets (year-end)"),
    ("Profitability: Net Profit Margin", "net_income", "revenue",
     "0.0%", "Net income / Total revenue"),
    ("Profitability: Return on Equity", "net_income", "equity",
     "0.0%", "Net income / Total shareholders' equity"),
]

# Sentiment rows copied into the summary (label as written in Sentiment tab).
SENTIMENT_ROWS = [
    "Total Word Count", "Positive Words", "Negative Words",
    "Uncertainty Words", "Litigious Words", "Constraining Words",
    "Strong Modal Words", "Weak Modal Words",
    "P/N Ratio (Positive / Negative)",
]

HEADING_FILL = PatternFill("solid", fgColor="E7E6E6")


def find_row(sheet, pattern):
    """Return the row number whose column-A text matches, or None."""
    for row in range(2, sheet.max_row + 1):
        text = str(sheet.cell(row=row, column=1).value or "").strip()
        if re.search(pattern, text, re.IGNORECASE):
            return row
    return None


def read_dates(sheet):
    """Return the date headers in row 1 (columns B onward)."""
    return [str(c.value) for c in sheet[1][1:] if c.value]


def locate_rows(workbook):
    """Find every needed row and check that all tabs share the same dates."""
    for name in REQUIRED_SHEETS:
        if name not in workbook.sheetnames:
            sys.exit(f"The '{name}' tab is missing. Run the earlier stages first.")

    dates = read_dates(workbook["Balance Sheet"])
    for name in REQUIRED_SHEETS:
        if read_dates(workbook[name]) != dates:
            sys.exit(f"The dates in the '{name}' tab do not match the Balance "
                     "Sheet tab. Re-run Stages C, D and E in order.")

    rows = {}
    for key, (sheet_name, pattern) in ROW_PATTERNS.items():
        row = find_row(workbook[sheet_name], pattern)
        if row is None:
            sys.exit(f"Could not find the '{key}' row in the {sheet_name} tab.")
        rows[key] = (sheet_name, row)
    return dates, rows


def cell_ref(rows, key, col_letter):
    """Return an Excel reference such as 'Balance Sheet'!B8 for a row key."""
    sheet_name, row = rows[key]
    return f"'{sheet_name}'!{col_letter}{row}"


def value_at(workbook, rows, key, col):
    """Read one number straight from the source tab (for the Terminal check)."""
    sheet_name, row = rows[key]
    value = workbook[sheet_name].cell(row=row, column=col).value
    return value if isinstance(value, (int, float)) else None


def correlation(xs, ys):
    """Pearson correlation of two lists, or None if it cannot be computed."""
    pairs = [(x, y) for x, y in zip(xs, ys) if x is not None and y is not None]
    if len(pairs) < 3:
        return None
    n = len(pairs)
    mean_x = sum(p[0] for p in pairs) / n
    mean_y = sum(p[1] for p in pairs) / n
    cov = sum((x - mean_x) * (y - mean_y) for x, y in pairs)
    var_x = sum((x - mean_x) ** 2 for x, _ in pairs)
    var_y = sum((y - mean_y) ** 2 for _, y in pairs)
    if var_x == 0 or var_y == 0:
        return None
    return cov / (var_x ** 0.5 * var_y ** 0.5)


def add_heading(sheet, row, title, n_cols):
    """Write a shaded section heading across the table."""
    for col in range(1, n_cols + 2):
        sheet.cell(row=row, column=col).fill = HEADING_FILL
    sheet.cell(row=row, column=1, value=title).font = Font(bold=True)


def build_sheet(workbook, ticker, dates, rows):
    """Create the Analysis tab and return the row numbers used."""
    if SHEET_NAME in workbook.sheetnames:
        del workbook[SHEET_NAME]
    sheet = workbook.create_sheet(SHEET_NAME)
    n = len(dates)
    letters = [get_column_letter(c) for c in range(2, n + 2)]

    sheet.cell(row=1, column=1, value=SHEET_NAME)
    for col, date in enumerate(dates, start=2):
        sheet.cell(row=1, column=col, value=date)

    used = {}
    row = 2
    add_heading(sheet, row, "Financial ratios", n)
    row += 1
    for name, top, bottom, fmt, _ in RATIOS:
        sheet.cell(row=row, column=1, value=name).alignment = Alignment(indent=1)
        for col, letter in enumerate(letters, start=2):
            formula = (f"=IFERROR({cell_ref(rows, top, letter)}/"
                       f"{cell_ref(rows, bottom, letter)},\"\")")
            sheet.cell(row=row, column=col, value=formula).number_format = fmt
        row += 1

    add_heading(sheet, row, "Sentiment summary (MD&A)", n)
    row += 1
    sentiment_sheet = workbook["Sentiment"]
    for label in SENTIMENT_ROWS:
        source_row = find_row(sentiment_sheet, "^" + re.escape(label) + "$")
        if source_row is None:
            continue
        is_ratio = label.startswith("P/N")
        cell = sheet.cell(row=row, column=1, value=label)
        cell.alignment = Alignment(indent=1)
        if is_ratio or label == "Total Word Count":
            cell.font = Font(bold=True)
        for col, letter in enumerate(letters, start=2):
            ref = f"='Sentiment'!{letter}{source_row}"
            out = sheet.cell(row=row, column=col, value=ref)
            out.number_format = "0.00" if is_ratio else "#,##0"
            if is_ratio or label == "Total Word Count":
                out.font = Font(bold=True)
        if is_ratio:
            used["pn"] = row
        row += 1

    add_heading(sheet, row, "Stock price", n)
    row += 1
    sheet.cell(row=row, column=1,
               value="Closing Stock Price ($ per share)").alignment = Alignment(indent=1)
    for col, letter in enumerate(letters, start=2):
        out = sheet.cell(row=row, column=col,
                         value=f"={cell_ref(rows, 'price', letter)}")
        out.number_format = '"$"#,##0.00_);("$"#,##0.00)'
    used["price"] = row
    row += 1

    sheet.cell(row=row, column=1,
               value="Correlation: stock price vs P/N ratio").alignment = Alignment(indent=1)
    last = letters[-1]
    corr = (f"=IFERROR(CORREL(B{used['price']}:{last}{used['price']},"
            f"B{used['pn']}:{last}{used['pn']}),\"\")")
    sheet.cell(row=row, column=2, value=corr).number_format = "0.00"
    used["corr"] = row
    row += 2

    notes = ["How each ratio is calculated (all from the 10-K statements):"]
    notes += [f"  {name}: {meaning}" for name, _, _, _, meaning in RATIOS]
    notes.append("Correlation runs from -1 to +1: near +1 means the price and "
                 "the P/N ratio rise and fall together; near -1, opposite "
                 "directions; near 0, no clear link.")
    for text in notes:
        sheet.cell(row=row, column=1, value=text).font = Font(
            italic=True, color="595959")
        row += 1
    used["chart_row"] = row + 1

    sheet.page_setup.orientation = "landscape"  # prints on one page wide
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    sheet.freeze_panes = "B2"
    sheet.column_dimensions["A"].width = 44
    for col in range(2, n + 2):
        sheet.column_dimensions[get_column_letter(col)].width = 15
    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    sheet["A1"].alignment = Alignment(horizontal="left")
    return sheet, used


def add_chart(sheet, ticker, used, n):
    """Add a line chart: stock price (left axis) and P/N ratio (right axis)."""
    categories = Reference(sheet, min_col=2, max_col=n + 1, min_row=1, max_row=1)

    price_chart = LineChart()
    price_chart.title = f"{ticker}: Stock Price vs. Sentiment P/N Ratio"
    price_chart.height = 13
    price_chart.width = 26
    price_chart.x_axis.title = "Balance sheet date"
    price_chart.y_axis.title = "Closing price ($)"
    price_chart.add_data(Reference(sheet, min_col=1, max_col=n + 1,
                                   min_row=used["price"], max_row=used["price"]),
                         from_rows=True, titles_from_data=True)
    price_chart.set_categories(categories)
    price_chart.x_axis.delete = False
    price_chart.y_axis.delete = False
    price_chart.y_axis.majorGridlines.spPr = GraphicalProperties(
        ln=LineProperties(solidFill="D9D9D9"))
    price_series = price_chart.series[0]
    price_series.graphicalProperties.line.solidFill = "1F4E79"
    price_series.marker.symbol = "circle"
    price_series.smooth = False

    pn_chart = LineChart()
    pn_chart.add_data(Reference(sheet, min_col=1, max_col=n + 1,
                                min_row=used["pn"], max_row=used["pn"]),
                      from_rows=True, titles_from_data=True)
    pn_chart.y_axis.axId = 200
    pn_chart.y_axis.title = "P/N ratio"
    pn_chart.y_axis.crosses = "max"
    pn_chart.y_axis.delete = False
    pn_chart.y_axis.majorGridlines = None  # one set of grid lines only
    pn_series = pn_chart.series[0]
    pn_series.graphicalProperties.line.solidFill = "C55A11"
    pn_series.marker.symbol = "diamond"
    pn_series.smooth = False

    price_chart += pn_chart
    price_chart.legend.position = "t"  # top, so it never covers the dates
    price_chart.legend.overlay = False
    # Stop the title, axis titles and legend from overlapping the plot
    price_chart.title.overlay = False
    price_chart.x_axis.title.overlay = False
    price_chart.y_axis.title.overlay = False
    pn_chart.y_axis.title.overlay = False
    price_chart.layout = Layout(manualLayout=ManualLayout(
        layoutTarget="inner", xMode="edge", yMode="edge",
        x=0.10, y=0.20, w=0.79, h=0.60))
    sheet.add_chart(price_chart, f"A{used['chart_row']}")


def print_check(workbook, rows, dates):
    """Print the ratios in the Terminal, computed in Python, for checking."""
    n = len(dates)
    print(f"\n{'':34}" + "".join(f"{d[:4]:>9}" for d in dates))
    for name, top, bottom, fmt, _ in RATIOS:
        cells = []
        for col in range(2, n + 2):
            a = value_at(workbook, rows, top, col)
            b = value_at(workbook, rows, bottom, col)
            if a is None or not b:
                cells.append(f"{'n/a':>9}")
            elif "%" in fmt:
                cells.append(f"{a / b:>8.1%} ")
            else:
                cells.append(f"{a / b:>9.2f}")
        print(f"{name:34}" + "".join(cells))

    prices = [value_at(workbook, rows, "price", c) for c in range(2, n + 2)]
    pn = [value_at(workbook, rows, "pn_ratio", c) for c in range(2, n + 2)]
    print(f"{'Stock price ($)':34}" + "".join(
        f"{p:>9.2f}" if p is not None else f"{'n/a':>9}" for p in prices))
    print(f"{'P/N ratio':34}" + "".join(
        f"{p:>9.2f}" if p is not None else f"{'n/a':>9}" for p in pn))
    corr = correlation(prices, pn)
    if corr is not None:
        print(f"\nCorrelation between stock price and P/N ratio: {corr:.2f}")


def main():
    """Build the Analysis tab for one ticker."""
    ticker = input("Enter the ticker symbol (e.g., AAPL): ").strip().upper()
    if not ticker:
        sys.exit("No ticker entered. Exiting.")

    workbook_path = OUTPUT_DIR / f"{ticker}_financials.xlsx"
    if not workbook_path.exists():
        sys.exit(f"{workbook_path} not found. Run the earlier stages first.")

    workbook = load_workbook(workbook_path)
    dates, rows = locate_rows(workbook)
    print(f"\nBuilding the Analysis tab for {ticker} ({len(dates)} years)...")

    sheet, used = build_sheet(workbook, ticker, dates, rows)
    add_chart(sheet, ticker, used, len(dates))
    print_check(workbook, rows, dates)

    workbook.save(workbook_path)
    print(f"\n'{SHEET_NAME}' tab added to: {workbook_path}")


if __name__ == "__main__":
    main()
