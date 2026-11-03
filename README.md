# 10-K Analysis: Apple vs. Microsoft

Interactive page and Python code for a ten-year comparison of Apple (AAPL) and Microsoft (MSFT) using their SEC 10-K filings.

- `index.html` is the page (open it directly, or view it through GitHub Pages).
- `stage_a` to `stage_f` are the Python programs, run in order: download 10-Ks, extract the MD&A, extract the financial statements, score the MD&A with the Loughran-McDonald dictionary, add stock prices, and build the analysis tab and chart.

Data sources: SEC EDGAR (10-K filings) and Yahoo Finance (prices, via yfinance). Prices are split-adjusted, not dividend-adjusted.
