# tax-utils

## zerodha-report.py

Summarizes a Zerodha tax P&L report (`.xlsx`): dividends and capital gains
by quarter.

### Requirements

- Python 3.7+
- No packages to install.

### Usage

```bash
python3 zerodha-report.py
```

Enter the path to the xlsx file when prompted. Writes an HTML report next
to the input file and opens it in your browser.
