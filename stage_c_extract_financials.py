"""
Stage C (Step 5): Extract the three financial statements into Excel.

What this script does:
    1. Asks for a ticker symbol (must already be downloaded by Stage A).
    2. Reads every  filings/TICKER/TICKER_10K_YYYY-MM-DD.htm  file.
    3. Finds the Income Statement, Balance Sheet, and Cash Flow Statement
       tables inside each filing.
    4. Reads the fiscal years printed in each table's column headers and
       keeps the MOST RECENT year's column from each filing (proved by the
       header, not assumed), so ten filings give ten fiscal years.
    5. Lines up line items that were renamed over the years, so each item
       sits on ONE row for all ten years.
    6. Puts every statement in the traditional order:
         Balance Sheet : current assets, non-current assets, current
                         liabilities, non-current liabilities, equity.
         Cash Flow     : operating, investing, financing, then the net
                         change in cash, beginning cash and ending cash.
    7. Writes output/TICKER_financials.xlsx with three tabs. Cell A1 of
       each tab holds the statement name, and the years run across the top
       (oldest on the left, newest on the right).

Run with:
    python3 stage_c_extract_financials.py
"""

import io
import re
import sys
import warnings
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

BASE_DIR = Path(__file__).resolve().parent
FILINGS_DIR = BASE_DIR / "filings"
OUTPUT_DIR = BASE_DIR / "output"

# Smallest table (in rows with numbers) that can be a full statement.
MIN_STATEMENT_ROWS = 10

# Labels that mean different things in different sections (for example
# "Other" under Operating vs. Investing). They always get the section
# name added in front so they can never be mixed up.
GENERIC_LABELS = {"other", "other, net", "basic", "diluted", "products", "services"}

# Header rows such as "Year Ended June 30," are not data and are skipped.
MONTHS = ("january|february|march|april|may|june|july|august|september|"
          "october|november|december")
DATE_HEADER = re.compile(
    rf"^(fiscal )?years? ended\b|^({MONTHS})\s+\d{{1,2}},?$", re.IGNORECASE)

# Per-share rows always show two decimals (for example 18.00, not 18).
PER_SHARE = re.compile(r"per (common )?share", re.IGNORECASE)

# Fiscal-year column headers: "2024" or "September 28, 2024" (month names
# may be abbreviated). They prove which column holds which fiscal year.
YEAR_CELL = re.compile(r"^(19|20)\d{2}$")
DATE_YEAR_CELL = re.compile(
    r"\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2},?\s+"
    r"((?:19|20)\d{2})\s*$", re.IGNORECASE)
HEADER_WORDS = re.compile(r"\b(ended|fiscal|years?)\b", re.IGNORECASE)
UNITS_TEXT = re.compile(r"in\s+(millions|thousands|billions)", re.IGNORECASE)

# Excel number formats in US dollars.
MONEY_FORMAT = '"$"#,##0_);("$"#,##0)'
MONEY_FORMAT_2 = '"$"#,##0.00_);("$"#,##0.00)'
SHARES_FORMAT = "#,##0_);(#,##0)"

# Blank cells are filled from a later 10-K's prior-year column ONLY for these
# statements. On the balance sheet and cash flow, later filings regroup and
# restate line items, so mixing them with the original filing would break the
# totals; those cells stay blank (exactly as the company first reported).
BACKFILL_STATEMENTS = {"Income Statement"}

# Rows shown in bold: subtotals and totals.
BOLD_ROWS = re.compile(
    r"^(total |gross margin|operating income|net income|net cash |cash generated"
    r"|net change in cash|increase/\(decrease\) in cash)", re.IGNORECASE)

# Line items that companies renamed over time are merged onto one row.
# Rules are listed per statement so a name used in one statement can never
# change a row in another. Each rule is (pattern matched against the
# lowercase row key, final display name). The first matching rule wins.
ALIAS_RULES = {
    "Income Statement": [
        (r"^net sales$", "Total net sales"),
        (r"^cost of sales$", "Total cost of sales"),
        (r"^(net )?revenues?$", "Total revenue"),
        (r"^other income( \(expense\)|/\(expense\))?, net$",
         "Other income/(expense), net"),
        (r"^(impairment.*restructuring|restructuring)$",
         "Impairment, integration, and restructuring"),
    ],
    "Balance Sheet": [
        (r"^(short-term marketable securities|current assets: marketable securities)$",
         "Marketable securities (current)"),
        (r"^(long-term marketable securities|non-?current assets: marketable securities)$",
         "Marketable securities (non-current)"),
        (r"^accrued expenses$", "Other current liabilities"),
        (r"^current assets: other$", "Other current assets"),
        (r"^current liabilities: other$", "Other current liabilities"),
        (r"^(current portion of long-term debt|current liabilities: term debt)$",
         "Term debt (current)"),
        (r"^(long-term debt|non-?current liabilities: term debt)$",
         "Term debt (non-current)"),
        (r"^current liabilities: deferred revenue$", "Deferred revenue"),
        (r"^(deferred revenue, non-?current|non-?current liabilities: deferred revenue)$",
         "Deferred revenue (non-current)"),
        (r"^(income taxes|short-term income taxes)$", "Short-term income taxes"),
        (r"^equity (and other )?investments$", "Equity and other investments"),
        (r"^(retained earnings.*|accumulated deficit)$",
         "Retained earnings / (accumulated deficit)"),
        (r"^accumulated other comprehensive (income|loss).*$",
         "Accumulated other comprehensive income / (loss)"),
    ],
    "Cash Flow": [
        (r"^cash.*beginning (of (the )?(year|period)|balances)$",
         "Cash, cash equivalents and restricted cash, beginning of year"),
        (r"^cash.*(end of (the )?(year|period)|ending balances)$",
         "Cash, cash equivalents and restricted cash, end of year"),
        (r"^(increase|decrease).*in cash",
         "Increase/(decrease) in cash, cash equivalents and restricted cash"),
        (r"^deferred income tax", "Deferred income tax expense/(benefit)"),
        (r"^(goodwill and )?asset impairments$", "Goodwill and asset impairments"),
        (r"^net recognized (gains|losses).*investments and derivatives$",
         "Net recognized (gains)/losses on investments and derivatives"),
        (r"90 days? or less",
         "Short-term debt, maturities of 90 days or less, net"),
        (r"^acquisition of companies",
         "Acquisition of companies, net of cash acquired, and purchases of "
         "intangible and other assets"),
        (r"^payments for strategic investments",
         "Payments for strategic investments, net"),
        (r".*commercial paper, net$",
         "Proceeds from/(repayments of) commercial paper, net"),
        (r"^cash (generated|used).*operating activities$",
         "Cash generated by operating activities"),
        (r"^cash (generated|used).*investing activities$",
         "Cash generated by/(used in) investing activities"),
        (r"^cash (generated|used).*financing activities$",
         "Cash generated by/(used in) financing activities"),
        (r"^net cash .*operat", "Net cash from operations"),
        (r"^net cash .*investing", "Net cash from/(used in) investing"),
        (r"^net cash .*financing", "Net cash from/(used in) financing"),
        (r"^net (change|increase|decrease).* in cash",
         "Net change in cash, cash equivalents and restricted cash"),
    ],
}
ALIAS_RULES = {
    statement: [(re.compile(pattern), name) for pattern, name in rules]
    for statement, rules in ALIAS_RULES.items()
}

# The traditional order of each statement. Each entry is
# (block code, heading shown in column A or None for no heading).
BLOCK_LAYOUT = {
    "Income Statement": [("all", None)],
    "Balance Sheet": [
        ("ca", "Current assets"),
        ("nca", "Non-current assets"),
        ("cl", "Current liabilities"),
        ("ncl", "Non-current liabilities"),
        ("eq", "Shareholders' equity"),
        ("total", None),
    ],
    "Cash Flow": [
        ("op", "Operating activities"),
        ("inv", "Investing activities"),
        ("fin", "Financing activities"),
        ("recon", None),
        ("supp", "Supplemental cash flow information"),
    ],
}

# Cash flow total rows that close each section.
CASH_FLOW_TOTALS = {
    "op": re.compile(r"^(net )?cash (generated|from|provided|used).*operat"),
    "inv": re.compile(r"^(net )?cash .*investing"),
    "fin": re.compile(r"^(net )?cash .*financing"),
}

# The closing rows of a cash flow statement, in their traditional order.
CASH_FLOW_CLOSING = [
    re.compile(r"effect of (foreign )?exchange"),
    re.compile(r"^(net change|increase/\(decrease\)).* in cash"),
    re.compile(r"beginning of (the )?(year|period)"),
    re.compile(r"end of (the )?(year|period)"),
]


# --------------------------------------------------------------------------
# Reading numbers and labels out of HTML tables
# --------------------------------------------------------------------------
def parse_number(text):
    """Convert text like '$ 1,234', '(56)', or '—' to a float, else None."""
    cleaned = text.replace("$", "").replace(",", "").replace("\xa0", " ").strip()
    if cleaned in ("", ")", "%", ")%"):
        return None
    if cleaned in ("—", "-", "–"):
        return 0.0
    negative = cleaned.startswith("(")
    cleaned = cleaned.strip("()%").strip()
    try:
        value = float(cleaned)
    except ValueError:
        return None
    return -value if negative else value


def clean_label(label):
    """Tidy a row label.

    Removes footnote marks, colons, extra spaces, and year-specific details
    that would otherwise put the same item on a different row every year
    (allowance amounts, accumulated depreciation, par value, share counts,
    and securities-loaned amounts).
    """
    label = label.replace("\xa0", " ").replace("’", "'")
    label = re.sub(r"\s+", " ", label).strip()
    label = re.sub(r"\s*\(\d\)$", "", label)  # trailing footnote like (1)
    label = re.sub(r"\s*\(including securities loaned.*\)$", "", label, flags=re.I)
    label = re.sub(r",\s*less allowances?\b.*$", ", net", label, flags=re.I)
    label = re.sub(r",\s*net of (allowance|accumulated)\b.*$", ", net", label,
                   flags=re.I)
    label = re.sub(r",\s*\$?\s*[\d.]+\s*par value\b.*$", "", label, flags=re.I)
    label = re.sub(r"\s*[–—-]\s*shares authorized.*$", "", label,
                   flags=re.I)
    return label.rstrip(":").strip()


def extract_years(cells):
    """Return the 4-digit years named in header cells, left to right."""
    years = []
    for cell in cells:
        text = cell.strip()
        if YEAR_CELL.match(text):
            years.append(int(text))
            continue
        match = DATE_YEAR_CELL.search(text)
        if match:
            years.append(int(match.group(1)))
    return years


def is_header_cell(cell):
    """True for cells that belong to a column-header row."""
    text = cell.strip()
    return bool(YEAR_CELL.match(text) or DATE_YEAR_CELL.search(text)
                or DATE_HEADER.match(text) or HEADER_WORDS.search(text))


def parse_table(table):
    """Return (rows, years) for one HTML table.

    rows  : list of (label, values, section) for rows that contain numbers.
            'values' holds EVERY number in the row, left to right.
            'section' is the latest heading row without numbers (for example
            'Investing activities'), used to tell apart rows sharing a label.
    years : fiscal years printed in the column headers, left to right
            (for example [2024, 2023, 2022]); empty if none were found.
    """
    rows = []
    years = []
    section = ""
    for tr in table.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        cells = [c for c in cells if c.strip()]
        if not cells:
            continue

        # A column-header row such as "2024 2023 2022" or "September 28, 2024".
        row_years = extract_years(cells)
        if row_years and all(is_header_cell(c) for c in cells):
            if not rows and not years:
                years = row_years
            continue

        label = None
        value_cells = []
        for position, cell in enumerate(cells):
            if re.search(r"[A-Za-z]", cell):
                label = clean_label(cell)
                value_cells = cells[position + 1:]
                break
        if not label or DATE_HEADER.match(label):
            continue  # blank row, or a header like "Year Ended June 30,"

        values = [parse_number(cell) for cell in value_cells]
        values = [value for value in values if value is not None]

        if not values:
            section = label  # a heading row such as "Investing activities:"
        else:
            rows.append((label, values, section))
    return rows, years


def detect_units(table):
    """Return 'millions', 'thousands' or 'billions' for the dollar amounts.

    Apple's note reads "(In millions, except number of shares, which are
    reflected in thousands ...)". The table's own first lines and the text
    just above it are searched together, and if "millions" (or "billions")
    appears anywhere there, it is the unit of the dollar amounts; a
    "thousands" mention is then just the share counts.
    """
    nearby = [text.strip() for text in table.find_all_previous(string=True, limit=80)
              if text.strip()][:12]
    text = " ".join(reversed(nearby)) + " " + table.get_text(" ", strip=True)[:600]
    found = [m.group(1).lower() for m in UNITS_TEXT.finditer(text)]
    for unit in ("millions", "billions", "thousands"):
        if unit in found:
            return unit
    return None


def load_tables(html_path):
    """Parse every large table in a 10-K file."""
    soup = BeautifulSoup(html_path.read_bytes(), "lxml")
    for tag in soup.find_all(["ix:header", "script", "style"]):
        tag.decompose()

    parsed = []
    for table in soup.find_all("table"):
        rows, years = parse_table(table)
        if len(rows) >= MIN_STATEMENT_ROWS:
            parsed.append({
                "rows": rows,
                "years": years,
                "units": detect_units(table),
                "dollars": "$" in table.get_text(),
            })
    return parsed


def resolve_columns(table):
    """Pick each row's value for the newest fiscal year, using the headers.

    Returns (rows, unaligned). Each row is (label, current_value, section,
    by_year) where by_year maps every header year to that row's value, so the
    older comparative columns can be cross-checked against earlier filings.
    A row whose number of values differs from the number of header years
    cannot be matched to a year safely: it falls back to its first number
    and is counted in 'unaligned'.
    """
    years = table["years"]
    usable = bool(years) and len(set(years)) == len(years)
    latest = max(years) if usable else None
    resolved, unaligned = [], 0
    for label, values, section in table["rows"]:
        if usable and len(values) == len(years):
            by_year = dict(zip(years, values))
            current = by_year[latest]
        else:
            by_year = {}
            current = values[0]
            unaligned += 1
        resolved.append((label, current, section, by_year))
    return resolved, unaligned


# --------------------------------------------------------------------------
# Deciding which table is which statement
# --------------------------------------------------------------------------
def has(labels, text):
    """True if any row label contains the given text."""
    return any(text in label for label in labels)


def is_income_statement(labels):
    """Income statement: revenue, net income, and diluted EPS; no assets."""
    revenue = has(labels, "revenue") or has(labels, "net sales")
    return (revenue and has(labels, "net income") and has(labels, "diluted")
            and not has(labels, "total assets") and not has(labels, "investing"))


def is_balance_sheet(labels):
    """Balance sheet: has both total assets and total liabilities and equity."""
    return has(labels, "total assets") and has(labels, "total liabilities and")


def is_cash_flow(labels):
    """Cash flow: has operating, investing, and financing sections."""
    operating = has(labels, "operating") or has(labels, "operations")
    return (operating and has(labels, "investing") and has(labels, "financing")
            and not has(labels, "total assets"))


STATEMENTS = {
    "Income Statement": is_income_statement,
    "Balance Sheet": is_balance_sheet,
    "Cash Flow": is_cash_flow,
}


def find_statement(tables, matcher):
    """Return the largest table that matches the statement test, or None."""
    best = None
    for table in tables:
        labels = [row[0].lower() for row in table["rows"]]
        if matcher(labels) and (best is None
                                or len(table["rows"]) > len(best["rows"])):
            best = table
    return best


# --------------------------------------------------------------------------
# Giving every line item one consistent name across years
# --------------------------------------------------------------------------
def base_key(label):
    """Lowercase a label so small wording differences do not matter."""
    return label.lower().replace("stockholders", "shareholders")


def short_section(section):
    """Shorten long cash-flow headings to Operating/Investing/Financing."""
    lowered = section.lower()
    if "activities" in lowered:
        for word, name in (("investing", "Investing activities"),
                           ("financing", "Financing activities"),
                           ("operating", "Operating activities")):
            if word in lowered:
                return name
    return section


def apply_alias(statement, key, display):
    """Replace a key and display name using this statement's rename rules."""
    for pattern, name in ALIAS_RULES.get(statement, []):
        if pattern.search(key):
            return name.lower(), name
    return key, display


def make_unique_keys(rows, statement):
    """Return [(key, display_label, value, by_year)] for one filing's statement.

    * A label that appears more than once in a statement, and generic
      labels such as 'Other', get their section heading added in front.
    * Renamed line items are merged using ALIAS_RULES for this statement.
    """
    counts = {}
    for label, _, _, _ in rows:
        key = base_key(label)
        counts[key] = counts.get(key, 0) + 1

    used = set()
    result = []
    for label, value, section, by_year in rows:
        base = base_key(label)
        section_name = short_section(section)
        key, display = base, label

        if section_name and (counts[base] > 1 or base in GENERIC_LABELS):
            key = f"{section_name.lower()}: {base}"
            display = f"{section_name}: {label}"

        key, display = apply_alias(statement, key, display)

        original, number = key, 1
        while key in used:  # two rows ended up with the same name
            number += 1
            key = f"{original} #{number}"
        if number > 1:
            display = f"{display} ({number})"
        used.add(key)
        result.append((key, display, value, by_year))
    return result


# --------------------------------------------------------------------------
# Sorting rows into the traditional sections of each statement
# --------------------------------------------------------------------------
def balance_sheet_blocks(displays):
    """Label each balance sheet row ca / nca / cl / ncl / eq / total.

    The section changes after each subtotal: Total current assets, Total
    assets, Total current liabilities, Total liabilities, Total equity.
    """
    blocks = []
    current = "ca"
    for display in displays:
        low = display.lower()
        blocks.append("total" if low.startswith("total liabilities and") else current)
        if low.startswith("total current assets"):
            current = "nca"
        elif low.startswith("total assets"):
            current = "cl"
        elif low.startswith("total current liabilities"):
            current = "ncl"
        elif low == "total liabilities":
            current = "eq"
        elif re.match(r"total (stockholders|shareholders)", low):
            current = "total"
    return blocks


def cash_flow_blocks(displays):
    """Label each cash flow row op / inv / fin / recon / supp.

    Rows are grouped up to the total that closes their section, so this
    works whether a company prints Investing before Financing (Apple) or
    after it (Microsoft).
    """
    blocks = [None] * len(displays)
    pending = []
    for index, display in enumerate(displays):
        low = display.lower()
        if any(pattern.search(low) for pattern in CASH_FLOW_CLOSING):
            blocks[index] = "recon"
            continue
        pending.append(index)
        for block, pattern in CASH_FLOW_TOTALS.items():
            if pattern.search(low):
                for i in pending:
                    blocks[i] = block
                pending = []
                break
    for i in pending:  # rows after the last total, e.g. cash paid for taxes
        blocks[i] = "supp"
    return blocks


def add_blocks(statement, items):
    """Return [(key, display, value, block, by_year)] for one filing."""
    displays = [item[1] for item in items]
    if statement == "Balance Sheet":
        blocks = balance_sheet_blocks(displays)
    elif statement == "Cash Flow":
        blocks = cash_flow_blocks(displays)
    else:
        blocks = ["all"] * len(items)
    return [(k, d, v, b, y) for (k, d, v, y), b in zip(items, blocks)]


def sort_rank(statement, block, display):
    """Order rows inside a section: normal rows, subtotals, then totals."""
    low = display.lower()
    if statement == "Cash Flow":
        if block == "recon":
            for rank, pattern in enumerate(CASH_FLOW_CLOSING):
                if pattern.search(low):
                    return rank
            return len(CASH_FLOW_CLOSING)
        return 1 if any(p.search(low) for p in CASH_FLOW_TOTALS.values()) else 0
    if statement == "Balance Sheet":
        if low.startswith("total assets") or low == "total liabilities":
            return 2
        return 1 if low.startswith("total ") else 0
    return 0


# --------------------------------------------------------------------------
# Combining years side by side
# --------------------------------------------------------------------------
def display_name(key, year_data):
    """Return the lowercase display label of a row key (newest filing first)."""
    for date in sorted(year_data, reverse=True):
        for item in year_data[date]:
            if item[0] == key:
                return item[1].lower()
    return ""


def build_dataframe(year_data, statement, fiscal_years):
    """Combine {date: [(key, label, value, block, by_year)]} into one table.

    Years are columns. Rows follow the traditional order of the statement,
    with a heading row at the start of each section. A blank cell is filled
    from the prior-year comparative column of a LATER filing when that
    filing reports the item (for example Apple's Products/Services split,
    first printed in a later 10-K; Income Statement only). Returns the table, a list of True/False
    flags marking heading rows, and the set of (row label, date) cells that
    were filled this way.
    """
    dates_newest_first = sorted(year_data, reverse=True)

    master = []
    for date in dates_newest_first:
        keys = [item[0] for item in year_data[date]]
        if not master:
            master = list(keys)
            continue
        last_index = -1
        for key in keys:
            if key in master:
                last_index = master.index(key)
            else:
                master.insert(last_index + 1, key)
                last_index += 1

    display, blocks = {}, {}
    for date in dates_newest_first:  # the newest filing decides the section
        for key, label, _, block, _ in year_data[date]:
            blocks.setdefault(key, block)
    for date in sorted(year_data):  # newest filing's wording wins last
        for key, label, _, _, _ in year_data[date]:
            display[key] = label

    values = {date: {item[0]: item[2] for item in year_data[date]}
              for date in sorted(year_data)}
    frame = pd.DataFrame(values, index=master)

    # Fill blanks from later filings' prior-year columns, but only for the
    # statements in BACKFILL_STATEMENTS, and only from a later filing that
    # agrees with the original filing on net income for that year (so a
    # restated year is never mixed in).
    filled = set()
    net_income_key = next((k for k in master if display_name(k, year_data) == "net income"), None)
    comparatives = {date: {item[0]: item[4] for item in year_data[date]}
                    for date in year_data}
    later_first = sorted(year_data, key=lambda d: fiscal_years[d])
    for date in frame.columns:
        if statement not in BACKFILL_STATEMENTS:
            break
        for key in master:
            if pd.notna(frame.at[key, date]):
                continue
            for later in later_first:
                if fiscal_years[later] <= fiscal_years[date]:
                    continue
                year = fiscal_years[date]
                if net_income_key:
                    later_ni = comparatives[later].get(net_income_key, {}).get(year)
                    original_ni = frame.at[net_income_key, date]
                    if later_ni is None or pd.isna(original_ni) or abs(later_ni - original_ni) > 0.5:
                        continue
                found = comparatives[later].get(key, {}).get(year)
                if found is not None:
                    frame.at[key, date] = found
                    filled.add((key, date))
                    break

    layout = BLOCK_LAYOUT[statement]
    block_order = {code: position for position, (code, _) in enumerate(layout)}
    titles = dict(layout)
    ordered = sorted(master, key=lambda key: (
        block_order.get(blocks[key], 99),
        sort_rank(statement, blocks[key], display[key]),
        master.index(key)))

    labels, rows, flags = [], [], []
    filled_cells = set()
    last_block = None
    for key in ordered:
        if blocks[key] != last_block:
            last_block = blocks[key]
            if titles.get(last_block):
                labels.append(titles[last_block])
                rows.append([None] * len(frame.columns))
                flags.append(True)
        labels.append(display[key])
        rows.append(list(frame.loc[key]))
        flags.append(False)
        for date in frame.columns:
            if (key, date) in filled:
                filled_cells.add((display[key], date))

    result = pd.DataFrame(rows, index=labels, columns=frame.columns)
    result.index.name = statement  # cell A1 holds just the statement name
    return result, flags, filled_cells


def find_row(frame, prefix):
    """Return the first row name that starts with the given text, or None."""
    for name in frame.index:
        if name.lower().startswith(prefix):
            return name
    return None


def find_row_matching(frame, pattern):
    """Return the first row name whose lowercase text matches the pattern."""
    for name in frame.index:
        if re.search(pattern, name.lower()):
            return name
    return None


def check_balance_sheet(frame):
    """Print a warning if Total assets != Total liabilities and equity."""
    assets = find_row(frame, "total assets")
    liab_eq = find_row(frame, "total liabilities and")
    if not assets or not liab_eq:
        print("  Balance check skipped (could not find both total rows).")
        return
    problems = []
    for date in frame.columns:
        a, b = frame.loc[assets, date], frame.loc[liab_eq, date]
        if pd.notna(a) and pd.notna(b) and abs(a - b) > 1:
            problems.append(date)
    if problems:
        print(f"  WARNING: balance sheet does not balance for: {problems}")
    else:
        print("  Balance check passed: Total assets = Total liabilities + equity "
              "in every year.")


def check_cash_flow(frame):
    """Check that operating + investing + financing (+ FX) = change in cash.

    Works for Apple-style and Microsoft-style row names. Microsoft also
    reports an 'effect of foreign exchange rates' line, which is included.
    """
    patterns = {
        "operating total": r"^(net )?cash (generated|from|provided|used).*operat",
        "investing total": r"^(net )?cash .*investing",
        "financing total": r"^(net )?cash .*financing",
        "change in cash": r"net change in cash|(increase|decrease).* in cash",
    }
    found = {key: find_row_matching(frame, pattern)
             for key, pattern in patterns.items()}
    missing = [key for key, name in found.items() if name is None]
    if missing:
        print("  Cash flow check skipped (could not find: "
              + ", ".join(missing) + ").")
        return

    fx_row = find_row_matching(frame, r"effect of (foreign )?exchange")
    problems = []
    for date in frame.columns:
        values = [frame.loc[found[key], date] for key in patterns]
        if any(pd.isna(v) for v in values):
            continue
        fx = frame.loc[fx_row, date] if fx_row else 0
        fx = 0 if pd.isna(fx) else fx
        if abs(values[0] + values[1] + values[2] + fx - values[3]) > 1:
            problems.append(date)

    if problems:
        print(f"  WARNING: cash flow totals do not add up for: {problems}")
    else:
        print("  Cash flow check passed: operating + investing + financing "
              "= change in cash in every year.")



def fiscal_year_for(date, years):
    """Return (fiscal_year, problem) for one filing, proven by its headers.

    The newest column header should match the report date's year (or the
    year before it for a fiscal year that ends in January or February).
    """
    expected = int(date[:4])
    month = int(date[5:7])
    if not years:
        return expected, "no fiscal years could be read from the column headers"
    latest = max(years)
    if latest == expected or (latest == expected - 1 and month <= 2):
        return latest, None
    return latest, (f"newest column header says {latest} but the report date "
                    f"is {date}")


def tie_out(year_data, fiscal_years):
    """Compare each filing's prior-year columns with the earlier filing.

    Returns (compared, matched, differences). A difference is a figure that
    the later filing shows differently for a year (a restatement or a
    reclassification), as (label, year, earlier value, later value).
    """
    first_values = {date: {item[0]: item[2] for item in year_data[date]}
                    for date in year_data}
    date_of_year = {fiscal_years[date]: date for date in year_data}
    compared = matched = 0
    differences = []
    for date in sorted(year_data):
        for key, label, _, _, by_year in year_data[date]:
            for year, later_value in by_year.items():
                if year == fiscal_years[date] or year not in date_of_year:
                    continue
                earlier = first_values[date_of_year[year]].get(key)
                if earlier is None:
                    continue
                compared += 1
                if abs(earlier - later_value) <= 0.5:
                    matched += 1
                else:
                    differences.append((label, year, earlier, later_value))
    return compared, matched, differences


def check_scale(frame, pattern):
    """Warn if a key total jumps by more than 5x between neighbouring years."""
    name = find_row_matching(frame, pattern)
    if not name:
        return
    series = frame.loc[name].dropna()
    previous = None
    for date, value in series.items():
        if previous and value and not 0.2 <= value / previous <= 5:
            print(f"  WARNING: '{name}' jumps from {previous:,.0f} to "
                  f"{value:,.0f} at {date} - the units may differ.")
        previous = value

# --------------------------------------------------------------------------
# Writing the Excel file
# --------------------------------------------------------------------------
def is_share_count_row(label):
    """True for share counts (weighted shares), which are not dollar amounts."""
    lowered = label.lower()
    return (lowered.startswith(("shares used", "weighted"))
            or "shares outstanding" in lowered)


def is_per_share_row(label):
    """True for per-share amounts (EPS, dividends) but not share counts."""
    lowered = label.lower()
    if is_share_count_row(lowered):
        return False
    return bool(PER_SHARE.search(lowered))


def write_excel(frames, heading_flags, output_path, notes, filled):
    """Write one tab per statement; money cells use a US-dollar format."""
    heading_fill = PatternFill("solid", fgColor="E7E6E6")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, frame in frames.items():
            frame.to_excel(writer, sheet_name=sheet_name)
            sheet = writer.sheets[sheet_name]
            dates = list(frame.columns)
            filled_here = filled.get(sheet_name, set())

            sheet.freeze_panes = "B2"
            sheet.column_dimensions["A"].width = 62
            for col in range(2, frame.shape[1] + 2):
                sheet.column_dimensions[get_column_letter(col)].width = 15

            for cell in sheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
            sheet["A1"].alignment = Alignment(horizontal="left")

            flags = heading_flags[sheet_name]
            indent = 1 if sheet_name != "Income Statement" else 0
            for position, row in enumerate(sheet.iter_rows(min_row=2)):
                label = str(row[0].value or "")
                if flags[position]:  # a section heading such as "Current assets"
                    for cell in row:
                        cell.font = Font(bold=True)
                        cell.fill = heading_fill
                    continue

                row[0].alignment = Alignment(indent=indent)
                bold = bool(BOLD_ROWS.match(label))
                if bold:
                    row[0].font = Font(bold=True)

                share_count = is_share_count_row(label)
                per_share = is_per_share_row(label)
                for offset, cell in enumerate(row[1:]):
                    if not isinstance(cell.value, (int, float)):
                        continue
                    if (label, dates[offset]) in filled_here:
                        cell.font = Font(italic=True, bold=bold, color="1F4E9E")
                    elif bold:
                        cell.font = Font(bold=True)

                    if share_count:
                        cell.number_format = SHARES_FORMAT
                    elif per_share:
                        cell.number_format = MONEY_FORMAT_2
                    elif float(cell.value).is_integer():
                        cell.number_format = MONEY_FORMAT
                    else:
                        cell.number_format = MONEY_FORMAT_2

            note_cell = sheet.cell(row=sheet.max_row + 2, column=1,
                                   value=notes.get(sheet_name, ""))
            note_cell.font = Font(italic=True, color="595959")


class Tee:
    """Send printed text to the screen and to a buffer (for the log file)."""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, text):
        for stream in self.streams:
            stream.write(text)

    def flush(self):
        for stream in self.streams:
            stream.flush()


def main():
    """Extract the three statements from every filing for one ticker."""
    ticker = input("Enter the ticker symbol (e.g., AAPL): ").strip().upper()
    if not ticker:
        sys.exit("No ticker entered. Exiting.")

    source_dir = FILINGS_DIR / ticker
    html_files = sorted(source_dir.glob(f"{ticker}_10K_*.htm"))
    if not html_files:
        sys.exit(f"No 10-K files found in {source_dir}. Run Stage A first.")

    buffer = io.StringIO()
    sys.stdout = Tee(sys.__stdout__, buffer)
    try:
        run(ticker, html_files, buffer)
    finally:
        sys.stdout = sys.__stdout__


def run(ticker, html_files, buffer):
    """Do the extraction, the accuracy checks, and write the files."""
    collected = {name: {} for name in STATEMENTS}
    fiscal_years = {name: {} for name in STATEMENTS}
    units_seen = {name: {} for name in STATEMENTS}
    print(f"\nReading {len(html_files)} filing(s) for {ticker}...\n")

    for html_path in html_files:
        date = html_path.stem.replace(f"{ticker}_10K_", "")
        try:
            tables = load_tables(html_path)
        except Exception as error:  # keep going if one file fails
            print(f"  {date}: ERROR ({error})")
            continue

        report, problems = [], []
        for name, matcher in STATEMENTS.items():
            table = find_statement(tables, matcher)
            if table is None:
                report.append(f"{name}: NOT FOUND")
                continue
            rows, unaligned = resolve_columns(table)
            year, problem = fiscal_year_for(date, table["years"])
            fiscal_years[name][date] = year
            units_seen[name][date] = table["units"]
            collected[name][date] = add_blocks(name, make_unique_keys(rows, name))
            report.append(f"{name}: {len(rows)} rows")

            if problem:
                problems.append(f"{name}: {problem}")
            if unaligned and table["years"]:
                problems.append(f"{name}: {unaligned} row(s) did not line up "
                                "with the year headers (first number used)")
            if not table["dollars"]:
                problems.append(f"{name}: no $ sign found - check the currency")
        print(f"  {date}: " + " | ".join(report))
        for problem in problems:
            print(f"      ! {problem}")

    frames, heading_flags, filled = {}, {}, {}
    for name, year_data in collected.items():
        if year_data:
            frames[name], heading_flags[name], filled[name] = build_dataframe(
                year_data, name, fiscal_years[name])
    if not frames:
        sys.exit("No statements were extracted.")

    print("\n--- Accuracy checks ---")
    if "Balance Sheet" in frames:
        check_balance_sheet(frames["Balance Sheet"])
    if "Cash Flow" in frames:
        check_cash_flow(frames["Cash Flow"])

    all_differences = {}
    tie_rates = {}
    for name, year_data in collected.items():
        compared, matched, differences = tie_out(year_data, fiscal_years[name])
        all_differences[name] = differences
        if not compared:
            print(f"  Year tie-out ({name}): nothing to compare.")
            continue
        percent = 100 * matched / compared
        tie_rates[name] = percent
        print(f"  Year tie-out ({name}): {matched} of {compared} prior-year "
              f"figures equal the earlier filing's figure ({percent:.1f}%).")
        if percent < 80:
            print("  WARNING: low match rate - columns may be misplaced.")
        for label, year, earlier, later in differences[:5]:
            print(f"      FY{year} {label}: first reported {earlier:,.2f}, "
                  f"later filing shows {later:,.2f}")
        if len(differences) > 5:
            print(f"      ... and {len(differences) - 5} more (see the log file)")

    if "Balance Sheet" in frames:
        check_scale(frames["Balance Sheet"], r"^total assets")
    if "Income Statement" in frames:
        check_scale(frames["Income Statement"], r"^total (net sales|revenues?)")

    notes = {}
    for name in frames:
        units = {u for u in units_seen[name].values() if u}
        if len(units) > 1:
            if tie_rates.get(name, 0) >= 80:
                print(f"  Note ({name}): some filings word the units differently "
                      f"{sorted(units)}, but the figures match across filings, "
                      "so the amounts are in the same units.")
            else:
                print(f"  WARNING ({name}): filings use different units: "
                      f"{sorted(units)}")
                for unit_date in sorted(units_seen[name]):
                    print(f"      {unit_date}: {units_seen[name][unit_date]}")
        counts = {}
        for unit in units_seen[name].values():
            if unit:
                counts[unit] = counts.get(unit, 0) + 1
        newest = max(counts, key=counts.get) if counts else None
        unit_text = f", in {newest}" if newest else ", as reported"
        notes[name] = (f"Amounts in US dollars ($){unit_text}, except per-share "
                       "amounts ($ per share) and share counts (as reported).")
        if filled.get(name):
            notes[name] += (" Blue italic = taken from the prior-year column "
                            "of a later 10-K because that item was not "
                            "reported separately in that year's own 10-K.")
        if filled.get(name):
            print(f"  {name}: {len(filled[name])} blank cell(s) filled from a "
                  "later 10-K's prior-year column (blue italic).")

    OUTPUT_DIR.mkdir(exist_ok=True)
    output_path = OUTPUT_DIR / f"{ticker}_financials.xlsx"
    write_excel(frames, heading_flags, output_path, notes, filled)
    print(f"\nDone. Workbook saved to: {output_path}")

    log_path = OUTPUT_DIR / f"{ticker}_extraction_log.txt"
    lines = [buffer.getvalue(), "\nALL RESTATED / RECLASSIFIED FIGURES\n"]
    for name, differences in all_differences.items():
        for label, year, earlier, later in differences:
            lines.append(f"{name} | FY{year} | {label} | first reported "
                         f"{earlier:,.2f} | later filing {later:,.2f}\n")
    log_path.write_text("".join(lines), encoding="utf-8")
    print(f"Log saved to: {log_path}")


if __name__ == "__main__":
    main()
