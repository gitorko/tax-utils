#!/usr/bin/env python3
"""
Summarize a Zerodha tax P&L xlsx: Equity Dividends by advance-tax quarter,
and Tradewise Exits (short/long term capital gains) by quarter and totals.
India FY: Apr 1 - Mar 31. Produces a single HTML report.

Usage:
    python3 zerodha-report.py

No third-party dependencies - reads the xlsx as a zip of XML files using
only the Python standard library.
"""

import sys
import zipfile
import re
import html
import webbrowser
from pathlib import Path
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict
import xml.etree.ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
R_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

DIVIDENDS_SHEET = "equity dividends"
TRADES_SHEET_PREFIX = "tradewise exits"
OTHER_DEBITS_SHEET = "other debits and credits"

EQ_STT_PERCENTAGE = Decimal("0.00102")
DP_CHARGES = Decimal("15.34")
MF_STT_PERCENTAGE = Decimal("0.00006")


def load_shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
    strings = []
    for si in sst.findall("m:si", NS):
        strings.append("".join(t.text or "" for t in si.findall(".//m:t", NS)))
    return strings


def find_sheet_path(z, matcher):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    relmap = {rel.get("Id"): rel.get("Target") for rel in rels.findall("r:Relationship", REL_NS)}

    for sheet in wb.findall(".//m:sheets/m:sheet", NS):
        if matcher(sheet.get("name", "").strip().lower()):
            rid = sheet.get(R_ATTR)
            target = relmap[rid]
            return "xl/" + target if not target.startswith("xl/") else target
    raise ValueError("Matching sheet not found in workbook")


def col_letters(cell_ref):
    return re.match(r"[A-Z]+", cell_ref).group()


def read_rows(z, sheet_path, strings, numeric=Decimal):
    sheet = ET.fromstring(z.read(sheet_path))
    rows = []
    for row in sheet.findall(".//m:row", NS):
        cells = {}
        for c in row.findall("m:c", NS):
            ref = c.get("r")
            t = c.get("t")
            v = c.find("m:v", NS)
            val = v.text if v is not None else None
            if val is not None:
                if t == "s":
                    val = strings[int(val)]
                elif t not in ("str", "b"):
                    val = numeric(val)
            cells[col_letters(ref)] = val
        rows.append((int(row.get("r")), cells))
    return rows


def parse_fy_start_year(rows):
    """Look for a title like 'Equity Dividends from 2025-04-01 to 2026-03-31'."""
    for _, row in rows:
        for val in row.values():
            if isinstance(val, str):
                m = re.search(r"from\s+(\d{4})-\d{2}-\d{2}", val, re.I)
                if m:
                    return int(m.group(1))
    return None


def quarter_bounds(fy_start_year, prefix):
    y = fy_start_year
    return [
        (f"{prefix}-Q1 (1-April to 15-June)", date(y, 4, 1), date(y, 6, 15)),
        (f"{prefix}-Q2 (16-Jun to 15-Sep)", date(y, 6, 16), date(y, 9, 15)),
        (f"{prefix}-Q3 (16-Sep to 15-Dec)", date(y, 9, 16), date(y, 12, 15)),
        (f"{prefix}-Q4 (16-Dec to 15-Mar)", date(y, 12, 16), date(y + 1, 3, 15)),
        (f"{prefix}-Q5 (16-Mar to 31-Mar)", date(y + 1, 3, 16), date(y + 1, 3, 31)),
    ]


def which_quarter(d, quarters):
    for i, (_, start, end) in enumerate(quarters):
        if start <= d <= end:
            return i
    return None


def format_inr(value, truncate=False, decimals=0):
    negative = value < 0
    quantum = Decimal(1).scaleb(-decimals)
    rounding = "ROUND_DOWN" if truncate else ROUND_HALF_UP
    scaled = value.copy_abs().quantize(quantum, rounding=rounding)
    whole, _, frac = str(scaled).partition(".")
    if len(whole) > 3:
        last3, rest = whole[-3:], whole[:-3]
        groups = []
        while len(rest) > 2:
            groups.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            groups.insert(0, rest)
        whole = ",".join(groups) + "," + last3
    suffix = f".{frac}" if decimals else ".00"
    return ("-" if negative else "") + whole + suffix


# ---------------------------------------------------------------------------
# Equity Dividends
# ---------------------------------------------------------------------------


def find_dividend_data(rows):
    header_idx = None
    for i, (_, row) in enumerate(rows):
        values = [str(v).strip().lower() for v in row.values() if v is not None]
        if "ex-date" in values and any("net dividend" in v for v in values):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not find header row with 'Ex-date' / 'Net Dividend Amount' columns")

    header = rows[header_idx][1]
    col_for = {}
    for col, val in header.items():
        if val is None:
            continue
        label = str(val).strip().lower()
        if label == "ex-date":
            col_for["date"] = col
        elif "net dividend" in label:
            col_for["amount"] = col

    data = []
    for _, row in rows[header_idx + 1 :]:
        if not row:
            continue
        first_val = next(iter(row.values()), None)
        if isinstance(first_val, str) and first_val.strip().lower().startswith("total"):
            break
        date_val = row.get(col_for["date"])
        amount_val = row.get(col_for["amount"])
        if date_val is None or amount_val is None:
            continue
        if isinstance(date_val, str):
            ex_date = datetime.strptime(date_val.strip(), "%Y-%m-%d").date()
        else:
            continue
        data.append((ex_date, Decimal(amount_val)))
    return data


def build_dividend_rows(rows, fy_start_year):
    data = find_dividend_data(rows)
    quarters = quarter_bounds(fy_start_year, "DIV")
    totals = [Decimal(0)] * len(quarters)
    for ex_date, amount in data:
        totals[which_quarter(ex_date, quarters)] += amount

    headers = ["Quarter", "Amount"]
    rows_out = [[label, str(int(t))] for (label, _, _), t in zip(quarters, totals)]
    rows_out.append(["Total", str(int(sum(totals)))])
    return headers, rows_out


# ---------------------------------------------------------------------------
# Tradewise Exits (STCG / LTCG)
# ---------------------------------------------------------------------------

TRADE_COLS = {
    "Symbol": "B",
    "Entry": "D",
    "Exit": "E",
    "Buy": "G",
    "Sell": "H",
    "Brokerage": "N",
    "ExchTxn": "O",
    "IPFT": "P",
    "SEBI": "Q",
    "CGST": "R",
    "SGST": "S",
    "IGST": "T",
    "Stamp": "U",
}


def find_section(rows, section_title):
    """Return the data rows for a 'Equity - X' section."""
    start = None
    for i, (_, row) in enumerate(rows):
        b = row.get("B")
        if isinstance(b, str) and b.strip().lower() == section_title.lower():
            start = i
            break
    if start is None:
        raise ValueError(f"Section '{section_title}' not found")

    # header row (Symbol/ISIN/...) is the next non-blank row after the title
    header_idx = None
    for i in range(start + 1, len(rows)):
        _, row = rows[i]
        if row.get("B") == "Symbol":
            header_idx = i
            break

    data_rows = []
    for i in range(header_idx + 1, len(rows)):
        _, row = rows[i]
        if not row or row.get("B") is None:
            break
        data_rows.append(row)
    return data_rows


def extract_trades(data_rows):
    trades = []
    for row in data_rows:
        d = {k: row.get(col) for k, col in TRADE_COLS.items()}
        d["Entry"] = datetime.strptime(d["Entry"].strip(), "%Y-%m-%d").date()
        d["Exit"] = datetime.strptime(d["Exit"].strip(), "%Y-%m-%d").date()
        trades.append(d)
    return trades


def group_transactions(trades):
    """A unique transaction is all lots bought on the same entry date and sold
    on the same exit date for one symbol (a single buy order matched with a
    single sell order, even if filled across multiple lot lines). Each unique
    transaction incurs one DP charge plus STT on its buy and sell value."""
    by_transaction = defaultdict(lambda: {"buy": Decimal(0), "sell": Decimal(0)})
    for t in trades:
        key = (t["Symbol"], t["Entry"], t["Exit"])
        by_transaction[key]["buy"] += t["Buy"]
        by_transaction[key]["sell"] += t["Sell"]

    transactions = []
    for (symbol, entry_date, exit_date), totals in sorted(by_transaction.items(), key=lambda kv: kv[0][2]):
        charge = DP_CHARGES + totals["buy"] * EQ_STT_PERCENTAGE + totals["sell"] * EQ_STT_PERCENTAGE
        transactions.append((symbol, entry_date, exit_date, totals["buy"], totals["sell"], charge))
    return transactions


def transaction_charges(trades):
    return sum((t[5] for t in group_transactions(trades)), Decimal(0))


def build_transactions_rows(trades, fy_start_year, prefix):
    quarters = quarter_bounds(fy_start_year, prefix)
    transactions = group_transactions(trades)
    headers = ["Symbol", "Entry Date", "Exit Date", "Quarter", "Duration (days)", "Buy", "Sell", "Profit/Loss", "Charge"]
    rows_out = []
    for symbol, entry_date, exit_date, buy, sell, charge in transactions:
        quarter_label = quarters[which_quarter(exit_date, quarters)][0]
        duration = (exit_date - entry_date).days
        profit_loss = sell - buy
        rows_out.append(
            [
                symbol,
                entry_date.isoformat(),
                exit_date.isoformat(),
                quarter_label,
                str(duration),
                format_inr(buy, decimals=2),
                format_inr(sell, decimals=2),
                format_inr(profit_loss, decimals=2),
                format_inr(charge, decimals=2),
            ]
        )
    total_buy = sum((t[3] for t in transactions), Decimal(0))
    total_sell = sum((t[4] for t in transactions), Decimal(0))
    total_charge = sum((t[5] for t in transactions), Decimal(0))
    total_profit_loss = total_sell - total_buy
    rows_out.append(
        [
            "Total",
            "",
            "",
            "",
            "",
            format_inr(total_buy, decimals=2),
            format_inr(total_sell, decimals=2),
            format_inr(total_profit_loss, decimals=2),
            format_inr(total_charge, decimals=2),
        ]
    )
    return headers, rows_out


def build_quarter_breakup_rows(trades, fy_start_year, prefix):
    quarters = quarter_bounds(fy_start_year, prefix)
    transactions = group_transactions(trades)

    buckets = [{"count": 0, "buy": Decimal(0), "sell": Decimal(0), "charge": Decimal(0)} for _ in quarters]
    for _, _, exit_date, buy, sell, charge in transactions:
        bucket = buckets[which_quarter(exit_date, quarters)]
        bucket["count"] += 1
        bucket["buy"] += buy
        bucket["sell"] += sell
        bucket["charge"] += charge

    headers = ["Quarter", "Transactions", "Buy", "Sell", "Profit/Loss", "Charge"]
    rows_out = []
    for (label, _, _), bucket in zip(quarters, buckets):
        profit_loss = bucket["sell"] - bucket["buy"]
        rows_out.append(
            [
                label,
                str(bucket["count"]),
                format_inr(bucket["buy"], decimals=2),
                format_inr(bucket["sell"], decimals=2),
                format_inr(profit_loss, decimals=2),
                format_inr(bucket["charge"], decimals=2),
            ]
        )

    total_count = sum(b["count"] for b in buckets)
    total_buy = sum((b["buy"] for b in buckets), Decimal(0))
    total_sell = sum((b["sell"] for b in buckets), Decimal(0))
    total_charge = sum((b["charge"] for b in buckets), Decimal(0))
    rows_out.append(
        [
            "Total",
            str(total_count),
            format_inr(total_buy, decimals=2),
            format_inr(total_sell, decimals=2),
            format_inr(total_sell - total_buy, decimals=2),
            format_inr(total_charge, decimals=2),
        ]
    )
    return headers, rows_out


def find_amc_total(z, strings):
    try:
        sheet_path = find_sheet_path(z, lambda n: n == OTHER_DEBITS_SHEET)
    except ValueError:
        return Decimal(0)
    rows = read_rows(z, sheet_path, strings)
    total = Decimal(0)
    for _, row in rows:
        b = row.get("B")
        if isinstance(b, str) and b.startswith("AMC for Demat"):
            debit = row.get("D")
            if isinstance(debit, Decimal):
                total += debit
    return total


def build_gains_quarter_rows(trades, fy_start_year, prefix):
    quarters = quarter_bounds(fy_start_year, prefix)
    totals = [Decimal(0)] * len(quarters)
    for t in trades:
        totals[which_quarter(t["Exit"], quarters)] += t["Sell"] - t["Buy"]

    headers = ["Quarter", "Amount"]
    rows_out = [[label, format_inr(v, truncate=True)] for (label, _, _), v in zip(quarters, totals)]
    return headers, rows_out


def section_totals(trades):
    gain = sum((t["Sell"] - t["Buy"] for t in trades), Decimal(0))
    full_value = sum((t["Sell"] for t in trades), Decimal(0))
    cost = sum((t["Buy"] for t in trades), Decimal(0))
    charges = transaction_charges(trades)
    return gain, full_value, cost, charges


def build_summary_rows(st_trades, lt_trades, amc_total):
    amc_each = amc_total / 2

    st_gain, st_full, st_cost, st_charges = section_totals(st_trades)
    lt_gain, lt_full, lt_cost, lt_charges = section_totals(lt_trades)

    st_total_charges = st_charges + amc_each
    lt_total_charges = lt_charges + amc_each
    st_final = st_gain - st_total_charges
    lt_final = lt_gain - lt_total_charges

    headers = ["", "Gain", "Charges", "AMC Charges", "Full value of consideration", "Cost of acquisition", "Total Charges", "Final Gain"]
    rows_out = [
        ["EQ STCG", format_inr(st_gain), format_inr(st_charges), format_inr(amc_each), format_inr(st_full, truncate=True), format_inr(st_cost, truncate=True), format_inr(st_total_charges), format_inr(st_final)],
        ["EQ LTCG", format_inr(lt_gain), format_inr(lt_charges), format_inr(amc_each), format_inr(lt_full, truncate=True), format_inr(lt_cost, truncate=True), format_inr(lt_total_charges), format_inr(lt_final)],
    ]
    return headers, rows_out


# ---------------------------------------------------------------------------
# HTML rendering
# ---------------------------------------------------------------------------

HTML_STYLE = """
body { font-family: -apple-system, Helvetica, Arial, sans-serif; margin: 2rem; color: #222; }
h1 { font-size: 1.4rem; }
h2 { font-size: 1.1rem; margin-top: 2.5rem; border-bottom: 1px solid #ccc; padding-bottom: 0.3rem; }
table { border-collapse: collapse; margin-top: 0.75rem; }
th, td { border: 1px solid #ddd; padding: 4px 10px; font-size: 0.9rem; white-space: nowrap; }
th { background: #f2f2f2; text-align: left; }
td.num { text-align: right; font-variant-numeric: tabular-nums; }
tr.total td { font-weight: bold; background: #fafafa; }
td.pos { color: #1a7f37; }
td.neg { color: #cf222e; }
"""

PROFIT_LOSS_COLUMNS = {"Profit/Loss", "Gain", "Final Gain"}


def esc(value):
    return html.escape(str(value))


def render_table(headers, rows, right_align_from=1, total_label="Total"):
    parts = ["<table>", "<tr>"]
    for h in headers:
        parts.append(f"<th>{esc(h)}</th>")
    parts.append("</tr>")
    for row in rows:
        row_class = ' class="total"' if row and row[0] == total_label else ""
        parts.append(f"<tr{row_class}>")
        for i, cell in enumerate(row):
            classes = []
            if i >= right_align_from:
                classes.append("num")
            if i < len(headers) and headers[i] in PROFIT_LOSS_COLUMNS and cell != "":
                classes.append("neg" if str(cell).startswith("-") else "pos")
            cls = f' class="{" ".join(classes)}"' if classes else ""
            parts.append(f"<td{cls}>{esc(cell)}</td>")
        parts.append("</tr>")
    parts.append("</table>")
    return "\n".join(parts)


def build_report_html(dividend_table, stg_table, ltg_table, summary_table, stcg_txns, stcg_breakup, ltcg_txns, ltcg_breakup, ne_table, ne_txns, ne_breakup):
    sections = [
        ("Config", ([["EQ STT Percentage", str(EQ_STT_PERCENTAGE)], ["DP Charges", str(DP_CHARGES)], ["MF STT Percentage", str(MF_STT_PERCENTAGE)]])),
    ]

    parts = [
        "<!DOCTYPE html>",
        "<html><head><meta charset='utf-8'><title>Zerodha Tax Report</title>",
        f"<style>{HTML_STYLE}</style></head><body>",
        "<h1>Zerodha Tax Report</h1>",
        "<h2>Config</h2>",
        render_table(["Setting", "Value"], sections[0][1]),
        "<h2>Equity Dividends</h2>",
        render_table(*dividend_table),
        "<h2>EQ Short Term Gains by Quarter</h2>",
        render_table(*stg_table),
        "<h2>EQ Long Term Gains by Quarter</h2>",
        render_table(*ltg_table),
        "<h2>EQ STCG Transactions</h2>",
        render_table(*stcg_txns),
        "<h2>EQ STCG Transactions - Quarter Breakup</h2>",
        render_table(*stcg_breakup),
        "<h2>EQ LTCG Transactions</h2>",
        render_table(*ltcg_txns),
        "<h2>EQ LTCG Transactions - Quarter Breakup</h2>",
        render_table(*ltcg_breakup),
        "<h2>Non Equity Gains by Quarter</h2>",
        render_table(*ne_table),
        "<h2>Non Equity Transactions</h2>",
        render_table(*ne_txns),
        "<h2>Non Equity Transactions - Quarter Breakup</h2>",
        render_table(*ne_breakup),
        "<h2>EQ STCG / LTCG Summary</h2>",
        render_table(*summary_table, right_align_from=1, total_label=None),
        "</body></html>",
    ]
    return "\n".join(parts)


def main():
    path = input("Enter path to the xlsx file: ").strip().strip('"').strip("'")
    if not path:
        print("No file path provided.")
        sys.exit(1)

    with zipfile.ZipFile(path) as z:
        strings = load_shared_strings(z)

        dividends_path = find_sheet_path(z, lambda n: n == DIVIDENDS_SHEET)
        dividend_rows = read_rows(z, dividends_path, strings)

        trades_path = find_sheet_path(z, lambda n: n.startswith(TRADES_SHEET_PREFIX))
        trade_rows = read_rows(z, trades_path, strings)

        amc_total = find_amc_total(z, strings)

    fy_start_year = parse_fy_start_year(dividend_rows) or parse_fy_start_year(trade_rows)

    dividend_table = build_dividend_rows(dividend_rows, fy_start_year)

    st_data = extract_trades(find_section(trade_rows, "Equity - Short Term"))
    lt_data = extract_trades(find_section(trade_rows, "Equity - Long Term"))

    stg_table = build_gains_quarter_rows(st_data, fy_start_year, "EQ-STG")
    ltg_table = build_gains_quarter_rows(lt_data, fy_start_year, "EQ-LTG")
    summary_table = build_summary_rows(st_data, lt_data, amc_total)
    stcg_txns = build_transactions_rows(st_data, fy_start_year, "EQ-STG")
    ltcg_txns = build_transactions_rows(lt_data, fy_start_year, "EQ-LTG")
    stcg_breakup = build_quarter_breakup_rows(st_data, fy_start_year, "EQ-STG")
    ltcg_breakup = build_quarter_breakup_rows(lt_data, fy_start_year, "EQ-LTG")

    ne_data = extract_trades(find_section(trade_rows, "Non Equity"))
    ne_table = build_gains_quarter_rows(ne_data, fy_start_year, "NE")
    ne_txns = build_transactions_rows(ne_data, fy_start_year, "NE")
    ne_breakup = build_quarter_breakup_rows(ne_data, fy_start_year, "NE")

    report_html = build_report_html(
        dividend_table, stg_table, ltg_table, summary_table, stcg_txns, stcg_breakup, ltcg_txns, ltcg_breakup, ne_table, ne_txns, ne_breakup
    )

    output_path = Path(path).with_name(Path(path).stem + "-report.html")
    output_path.write_text(report_html, encoding="utf-8")

    print(f"Report written to {output_path}")
    webbrowser.open(output_path.resolve().as_uri())


if __name__ == "__main__":
    main()
