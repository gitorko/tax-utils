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

DP_CHARGES = Decimal("15.34")


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


def round_inr(value, truncate=False):
    rounding = "ROUND_DOWN" if truncate else ROUND_HALF_UP
    return value.quantize(Decimal(1), rounding=rounding)


def format_inr(value, truncate=False):
    value = round_inr(value, truncate=truncate)
    negative = value < 0
    whole = str(value.copy_abs())
    if len(whole) > 3:
        last3, rest = whole[-3:], whole[:-3]
        groups = []
        while len(rest) > 2:
            groups.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            groups.insert(0, rest)
        whole = ",".join(groups) + "," + last3
    return ("-" if negative else "") + whole


def find_client_info(rows):
    info = {}
    for _, row in rows:
        label = row.get("B")
        value = row.get("C")
        if isinstance(label, str) and label.strip() in ("Client ID", "Client Name", "PAN"):
            info[label.strip()] = value
    return info


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
        if label == "symbol":
            col_for["symbol"] = col
        elif label == "ex-date":
            col_for["date"] = col
        elif label == "quantity":
            col_for["qty"] = col
        elif "dividend per share" in label:
            col_for["dps"] = col
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
        data.append(
            {
                "symbol": row.get(col_for.get("symbol")),
                "date": ex_date,
                "qty": Decimal(row.get(col_for.get("qty")) or 0),
                "dps": Decimal(row.get(col_for.get("dps")) or 0),
                "amount": Decimal(amount_val),
            }
        )
    return data


def build_dividend_rows(rows, fy_start_year):
    data = find_dividend_data(rows)
    quarters = quarter_bounds(fy_start_year, "DIV")
    totals = [Decimal(0)] * len(quarters)
    for d in data:
        totals[which_quarter(d["date"], quarters)] += d["amount"]

    headers = ["Quarter", "Amount"]
    rows_out = [[label, format_inr(t)] for (label, _, _), t in zip(quarters, totals)]
    rows_out.append(["Total", format_inr(sum(totals))])
    return headers, rows_out


def build_dividend_transactions_rows(rows, fy_start_year):
    data = find_dividend_data(rows)
    quarters = quarter_bounds(fy_start_year, "DIV")

    headers = ["Symbol", "Ex-Date", "Quarter", "Quantity", "Dividend Per Share", "Net Amount"]
    rows_out = []
    for d in sorted(data, key=lambda d: d["date"]):
        quarter_label = quarters[which_quarter(d["date"], quarters)][0]
        rows_out.append(
            [
                d["symbol"],
                d["date"].isoformat(),
                quarter_label,
                format_inr(d["qty"]),
                format_inr(d["dps"]),
                format_inr(d["amount"]),
            ]
        )

    total_qty = sum((d["qty"] for d in data), Decimal(0))
    total_amount = sum((d["amount"] for d in data), Decimal(0))
    rows_out.append(["Total", "", "", format_inr(total_qty), "", format_inr(total_amount)])
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


def split_by_holding_period(trades, threshold_days=365):
    """Held for 12 months (365 days) or less: short-term. More than 12 months: long-term."""
    short_term = [t for t in trades if (t["Exit"] - t["Entry"]).days <= threshold_days]
    long_term = [t for t in trades if (t["Exit"] - t["Entry"]).days > threshold_days]
    return short_term, long_term


# ---------------------------------------------------------------------------
# Mutual Funds - Debt
# ---------------------------------------------------------------------------

MF_TRADE_COLS = {
    "Symbol": "B",
    "Entry": "D",
    "Exit": "E",
    "Buy": "G",
    "Sell": "H",
}

DEBT_FUND_RULE_CHANGE_DATE = date(2024, 7, 23)
DEBT_FUND_GRANDFATHER_CUTOFF = date(2023, 4, 1)


def extract_mf_trades(data_rows):
    trades = []
    for row in data_rows:
        d = {k: row.get(col) for k, col in MF_TRADE_COLS.items()}
        d["Entry"] = datetime.strptime(d["Entry"].strip(), "%Y-%m-%d").date()
        d["Exit"] = datetime.strptime(d["Exit"].strip(), "%Y-%m-%d").date()
        trades.append(d)
    return trades


def months_held(entry_date, exit_date):
    months = (exit_date.year - entry_date.year) * 12 + (exit_date.month - entry_date.month)
    if exit_date.day < entry_date.day:
        months -= 1
    return months


def classify_debt_fund_trades(trades):
    """India debt mutual fund tax treatment:
    - Bought before 1-Apr-2023, sold before 23-Jul-2024: STCG (slab) if held <=36 months,
      else LTCG at 20% with indexation.
    - Bought before 1-Apr-2023, sold on/after 23-Jul-2024: STCG (slab) if held <=24 months,
      else LTCG at 12.5% without indexation.
    - Bought on/after 1-Apr-2023: always taxed at slab rate, regardless of holding period."""
    buckets = {
        "old_rule_stcg": [],
        "old_rule_ltcg": [],
        "new_rule_stcg": [],
        "new_rule_ltcg": [],
        "post_2023_purchase": [],
    }
    for t in trades:
        if t["Entry"] >= DEBT_FUND_GRANDFATHER_CUTOFF:
            buckets["post_2023_purchase"].append(t)
        elif t["Exit"] < DEBT_FUND_RULE_CHANGE_DATE:
            key = "old_rule_ltcg" if months_held(t["Entry"], t["Exit"]) > 36 else "old_rule_stcg"
            buckets[key].append(t)
        else:
            key = "new_rule_ltcg" if months_held(t["Entry"], t["Exit"]) > 24 else "new_rule_stcg"
            buckets[key].append(t)
    return buckets


def group_mf_transactions(trades):
    """Mutual fund redemptions carry no STT/DP/brokerage charges."""
    by_transaction = defaultdict(lambda: {"buy": Decimal(0), "sell": Decimal(0)})
    for t in trades:
        key = (t["Symbol"], t["Entry"], t["Exit"])
        by_transaction[key]["buy"] += t["Buy"]
        by_transaction[key]["sell"] += t["Sell"]

    transactions = []
    for (symbol, entry_date, exit_date), totals in sorted(by_transaction.items(), key=lambda kv: kv[0][2]):
        transactions.append((symbol, entry_date, exit_date, totals["buy"], totals["sell"], Decimal(0)))
    return transactions


def group_transactions(trades):
    """A unique transaction is all lots of one symbol sold on the same exit
    date - lots bought on different entry dates are merged into a single row,
    since holding-period (short/long term) is already implied by the section.
    STT is not included here: it's excluded from capital gains cost/charges
    per the proviso to Section 48 (Income Tax Act) for STCG 111A / LTCG 112A
    trades. (DP charges are handled separately - see compute_dp_allocation -
    since they're billed once per (symbol, exit date) regardless of how many
    different entry dates/sections that sale spans.)"""
    by_transaction = defaultdict(lambda: {"buy": Decimal(0), "sell": Decimal(0)})
    for t in trades:
        key = (t["Symbol"], t["Exit"])
        by_transaction[key]["buy"] += t["Buy"]
        by_transaction[key]["sell"] += t["Sell"]

    transactions = []
    for (symbol, exit_date), totals in sorted(by_transaction.items(), key=lambda kv: kv[0][1]):
        transactions.append((symbol, None, exit_date, totals["buy"], totals["sell"], Decimal(0)))
    return transactions


def compute_dp_allocation(labeled_trades):
    """DP charges are billed once per (symbol, exit date) across the whole
    report, not per section or per entry date. When a sale on a given date
    spans multiple sections (e.g. some lots ST, some LT), split that one flat
    fee across the sections proportionally to sell value."""
    global_sell = defaultdict(Decimal)
    section_sell = defaultdict(lambda: defaultdict(Decimal))
    for label, trades in labeled_trades:
        for t in trades:
            key = (t["Symbol"], t["Exit"])
            global_sell[key] += t["Sell"]
            section_sell[label][key] += t["Sell"]

    dp_alloc = defaultdict(lambda: defaultdict(Decimal))
    for key, total_sell in global_sell.items():
        sections_with_key = [label for label, _ in labeled_trades if key in section_sell[label]]
        if total_sell == 0:
            share = DP_CHARGES / len(sections_with_key) if sections_with_key else Decimal(0)
            for label in sections_with_key:
                dp_alloc[label][key] = share
            continue
        for label in sections_with_key:
            dp_alloc[label][key] = DP_CHARGES * section_sell[label][key] / total_sell
    return dp_alloc


def finalize_transactions(transactions, dp_alloc_for_section):
    """Attach each unique transaction's share of its section's DP charge
    allocation (split proportionally to sell value among transactions with
    the same symbol+exit date) on top of its STT charge."""
    key_sell_totals = defaultdict(Decimal)
    for symbol, _, exit_date, _, sell, _ in transactions:
        key_sell_totals[(symbol, exit_date)] += sell

    finalized = []
    for symbol, entry_date, exit_date, buy, sell, stt_charge in transactions:
        key = (symbol, exit_date)
        section_dp_total = dp_alloc_for_section.get(key, Decimal(0))
        key_total_sell = key_sell_totals[key]
        dp_share = section_dp_total * sell / key_total_sell if key_total_sell else Decimal(0)
        finalized.append((symbol, entry_date, exit_date, buy, sell, stt_charge + dp_share))
    return finalized


def transaction_charges(transactions):
    return sum((t[5] for t in transactions), Decimal(0))


def build_transactions_rows(transactions, fy_start_year, prefix, include_entry_date=True):
    quarters = quarter_bounds(fy_start_year, prefix)
    if include_entry_date:
        headers = ["Symbol", "Entry Date", "Exit Date", "Quarter", "Duration (days)", "Buy", "Sell", "Profit/Loss", "Charge"]
    else:
        headers = ["Symbol", "Exit Date", "Quarter", "Buy", "Sell", "Profit/Loss", "Charge"]
    rows_out = []
    total_buy = total_sell = total_profit_loss = total_charge = Decimal(0)
    for symbol, entry_date, exit_date, buy, sell, charge in transactions:
        quarter_label = quarters[which_quarter(exit_date, quarters)][0]
        r_buy, r_sell, r_profit_loss, r_charge = round_inr(buy), round_inr(sell), round_inr(sell - buy), round_inr(charge)
        total_buy += r_buy
        total_sell += r_sell
        total_profit_loss += r_profit_loss
        total_charge += r_charge
        if include_entry_date:
            duration = (exit_date - entry_date).days
            rows_out.append(
                [
                    symbol,
                    entry_date.isoformat(),
                    exit_date.isoformat(),
                    quarter_label,
                    str(duration),
                    format_inr(r_buy),
                    format_inr(r_sell),
                    format_inr(r_profit_loss),
                    format_inr(r_charge),
                ]
            )
        else:
            rows_out.append(
                [
                    symbol,
                    exit_date.isoformat(),
                    quarter_label,
                    format_inr(r_buy),
                    format_inr(r_sell),
                    format_inr(r_profit_loss),
                    format_inr(r_charge),
                ]
            )
    if include_entry_date:
        total_row = [
            "Total",
            "",
            "",
            "",
            "",
            format_inr(total_buy),
            format_inr(total_sell),
            format_inr(total_profit_loss),
            format_inr(total_charge),
        ]
    else:
        total_row = [
            "Total",
            "",
            "",
            format_inr(total_buy),
            format_inr(total_sell),
            format_inr(total_profit_loss),
            format_inr(total_charge),
        ]
    rows_out.append(total_row)
    return headers, rows_out


def build_quarter_breakup_rows(transactions, fy_start_year, prefix):
    quarters = quarter_bounds(fy_start_year, prefix)

    buckets = [{"count": 0, "buy": Decimal(0), "sell": Decimal(0), "charge": Decimal(0)} for _ in quarters]
    for _, _, exit_date, buy, sell, charge in transactions:
        bucket = buckets[which_quarter(exit_date, quarters)]
        bucket["count"] += 1
        bucket["buy"] += buy
        bucket["sell"] += sell
        bucket["charge"] += charge

    headers = ["Quarter", "Transactions", "Buy", "Sell", "Profit/Loss", "Charge", "Final Gain"]
    rows_out = []
    total_count = 0
    total_buy = total_sell = total_profit_loss = total_charge = total_final_gain = Decimal(0)
    for (label, _, _), bucket in zip(quarters, buckets):
        r_buy, r_sell, r_charge = round_inr(bucket["buy"]), round_inr(bucket["sell"]), round_inr(bucket["charge"])
        r_profit_loss = round_inr(bucket["sell"] - bucket["buy"])
        r_final_gain = r_profit_loss - r_charge
        total_count += bucket["count"]
        total_buy += r_buy
        total_sell += r_sell
        total_profit_loss += r_profit_loss
        total_charge += r_charge
        total_final_gain += r_final_gain
        rows_out.append(
            [
                label,
                str(bucket["count"]),
                format_inr(r_buy),
                format_inr(r_sell),
                format_inr(r_profit_loss),
                format_inr(r_charge),
                format_inr(r_final_gain),
            ]
        )

    rows_out.append(
        [
            "Total",
            str(total_count),
            format_inr(total_buy),
            format_inr(total_sell),
            format_inr(total_profit_loss),
            format_inr(total_charge),
            format_inr(total_final_gain),
        ]
    )
    return headers, rows_out


def find_dp_charge_rows(z, strings):
    """Actual DP charge line items billed by the broker, one per (symbol, sale
    date), used to cross-check the DP charge count assumed by group_transactions."""
    try:
        sheet_path = find_sheet_path(z, lambda n: n == OTHER_DEBITS_SHEET)
    except ValueError:
        return []
    rows = read_rows(z, sheet_path, strings)
    entries = []
    for _, row in rows:
        b = row.get("B")
        if isinstance(b, str) and b.startswith("DP Charges for Sale of"):
            debit = row.get("D")
            posting_date = row.get("C")
            if isinstance(debit, Decimal) and isinstance(posting_date, str):
                entries.append((b, datetime.strptime(posting_date.strip(), "%Y-%m-%d").date(), debit))
    return sorted(entries, key=lambda e: e[1])


def build_dp_charges_rows(dp_entries):
    headers = ["Particulars", "Posting Date", "Debit"]
    rounded_debits = [round_inr(debit) for _, _, debit in dp_entries]
    rows_out = [[particulars, posting_date.isoformat(), format_inr(r_debit)] for (particulars, posting_date, _), r_debit in zip(dp_entries, rounded_debits)]
    rows_out.append(["Total", "", format_inr(sum(rounded_debits, Decimal(0)))])
    return headers, rows_out


def build_dp_reconciliation_rows(dp_entries, unique_sale_count, section_transaction_counts):
    sheet_count = len(dp_entries)
    sheet_total = sum((debit for _, _, debit in dp_entries), Decimal(0))
    computed_total = DP_CHARGES * unique_sale_count

    headers = ["", "Count", "Amount"]
    rows_out = [
        ["Other Debits and Credits sheet", str(sheet_count), format_inr(sheet_total)],
        ["Computed (unique symbol+exit-date sales x DP Charges)", str(unique_sale_count), format_inr(computed_total)],
        ["Difference", str(sheet_count - unique_sale_count), format_inr(sheet_total - computed_total)],
    ]
    for label, count in section_transaction_counts.items():
        rows_out.append([f"  {label} transactions (entry+exit lots)", str(count), ""])
    return headers, rows_out


def build_gains_quarter_rows(trades, fy_start_year, prefix):
    quarters = quarter_bounds(fy_start_year, prefix)
    totals = [Decimal(0)] * len(quarters)
    for t in trades:
        totals[which_quarter(t["Exit"], quarters)] += t["Sell"] - t["Buy"]

    headers = ["Quarter", "Amount"]
    rows_out = [[label, format_inr(v, truncate=True)] for (label, _, _), v in zip(quarters, totals)]
    return headers, rows_out


def section_totals(trades, transactions):
    gain = sum((t["Sell"] - t["Buy"] for t in trades), Decimal(0))
    full_value = sum((t["Sell"] for t in trades), Decimal(0))
    cost = sum((t["Buy"] for t in trades), Decimal(0))
    charges = transaction_charges(transactions)
    return gain, full_value, cost, charges


def build_summary_rows(labeled_trades):
    """labeled_trades: list of (label, trades, transactions)."""
    headers = ["", "Gain", "Charges", "Full value of consideration", "Cost of acquisition", "Final Gain"]
    rows_out = []
    for label, trades, transactions in labeled_trades:
        gain, full_value, cost, charges = section_totals(trades, transactions)
        final_gain = gain - charges
        rows_out.append(
            [
                label,
                format_inr(gain),
                format_inr(charges),
                format_inr(full_value, truncate=True),
                format_inr(cost, truncate=True),
                format_inr(final_gain),
            ]
        )
    return headers, rows_out


# ---------------------------------------------------------------------------
# HTML rendering
# ---------------------------------------------------------------------------

HTML_STYLE = """
body { font-family: -apple-system, Helvetica, Arial, sans-serif; margin: 2rem; color: #222; }
h1 { font-size: 1.4rem; }
h2 { font-size: 1.1rem; margin-top: 2.5rem; border-bottom: 1px solid #ccc; padding-bottom: 0.3rem; }
h3 { font-size: 0.95rem; margin-top: 1.25rem; color: #444; }
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


def render_trades_section(trades, transactions, fy_start_year, prefix, title, include_quarter_table=True, include_entry_date=True):
    parts = [f"<h2>{esc(title)}</h2>"]
    if not trades:
        parts.append("<p><em>No transactions.</em></p>")
        return "\n".join(parts)

    if include_quarter_table:
        parts.append("<h3>Gains by Quarter</h3>")
        parts.append(render_table(*build_gains_quarter_rows(trades, fy_start_year, prefix)))
    parts.append("<h3>Transactions</h3>")
    parts.append(render_table(*build_transactions_rows(transactions, fy_start_year, prefix, include_entry_date=include_entry_date)))
    parts.append("<h3>Transactions - Quarter Breakup</h3>")
    parts.append(render_table(*build_quarter_breakup_rows(transactions, fy_start_year, prefix)))
    return "\n".join(parts)


def render_dividend_section(dividend_data, fy_start_year, title):
    parts = [f"<h2>{esc(title)}</h2>"]
    if not dividend_data:
        parts.append("<p><em>No transactions.</em></p>")
        return "\n".join(parts)

    quarters = quarter_bounds(fy_start_year, "DIV")
    totals = [Decimal(0)] * len(quarters)
    for d in dividend_data:
        totals[which_quarter(d["date"], quarters)] += d["amount"]
    quarter_headers = ["Quarter", "Amount"]
    quarter_rows = [[label, format_inr(t)] for (label, _, _), t in zip(quarters, totals)] + [["Total", format_inr(sum(totals))]]

    txn_headers = ["Symbol", "Ex-Date", "Quarter", "Quantity", "Dividend Per Share", "Net Amount"]
    txn_rows = []
    for d in sorted(dividend_data, key=lambda d: d["date"]):
        quarter_label = quarters[which_quarter(d["date"], quarters)][0]
        txn_rows.append(
            [
                d["symbol"],
                d["date"].isoformat(),
                quarter_label,
                format_inr(d["qty"]),
                format_inr(d["dps"]),
                format_inr(d["amount"]),
            ]
        )
    total_qty = sum((d["qty"] for d in dividend_data), Decimal(0))
    total_amount = sum((d["amount"] for d in dividend_data), Decimal(0))
    txn_rows.append(["Total", "", "", format_inr(total_qty), "", format_inr(total_amount)])

    parts.append("<h3>Gains by Quarter</h3>")
    parts.append(render_table(quarter_headers, quarter_rows))
    parts.append("<h3>Transactions</h3>")
    parts.append(render_table(txn_headers, txn_rows))
    return "\n".join(parts)


TAX_RATES_REFERENCE = [
    ["Dividends", "N/A", "Slab rate"],
    ["Listed Equity Shares", "≤ 12 months", "20% (STCG under Section 111A)"],
    ["Listed Equity Shares", "> 12 months", "12.5% on LTCG (₹1.25 lakh annual exemption available)"],
    ["Equity Mutual Funds", "≤ 12 months", "20%"],
    ["Equity Mutual Funds", "> 12 months", "12.5% (₹1.25 lakh exemption)"],
    ["Gold ETF", "≤ 12 months", "Slab rate"],
    ["Gold ETF", "> 12 months", "12.5% (No ₹1.25 lakh exemption, no indexation)"],
    ["Silver ETF", "≤ 12 months", "Slab rate"],
    ["Silver ETF", "> 12 months", "12.5%"],
    ["International Equity Funds", "≤ 24 months", "Slab rate"],
    ["International Equity Funds", "> 24 months", "12.5%"],
    ["Debt Mutual Funds (units bought on/after 1 Apr 2023)", "Any period", "Slab rate"],
    ["Debt Mutual Funds (units bought before 1 Apr 2023)", "> 24 months", "12.5% (subject to grandfathering provisions)"],
    ["Physical Gold", "≤ 24 months", "Slab rate"],
    ["Physical Gold", "> 24 months", "12.5%"],
    ["Real Estate", "≤ 24 months", "Slab rate"],
    ["Real Estate", "> 24 months", "12.5% (subject to applicable provisions)"],
]


def build_report_html(
    client_info,
    dividend_data,
    summary_table,
    st_data,
    st_transactions,
    lt_data,
    lt_transactions,
    ne_st_data,
    ne_st_transactions,
    fy_start_year,
    debt_fund_sections,
    dp_charges_table,
    dp_reconciliation_table,
):
    sections = [
        ("Config", ([["DP Charges", str(DP_CHARGES)]])),
    ]

    parts = [
        "<!DOCTYPE html>",
        "<html><head><meta charset='utf-8'><title>Zerodha Tax Report</title>",
        f"<style>{HTML_STYLE}</style></head><body>",
        "<h1>Zerodha Tax Report</h1>",
        render_table(
            ["Field", "Value"],
            [[label, client_info.get(label, "")] for label in ("Client ID", "Client Name", "PAN")],
            right_align_from=99,
            total_label=None,
        ),
        "<h2>Capital Gains Tax Rates (FY 2025-26 / AY 2026-27)</h2>",
        render_table(["Asset", "Holding Period", "Tax Rate"], TAX_RATES_REFERENCE, right_align_from=99, total_label=None),
        "<h2>Config</h2>",
        render_table(["Setting", "Value"], sections[0][1]),
        render_dividend_section(dividend_data, fy_start_year, "Equity Dividends (Slab rate)"),
        render_trades_section(st_data, st_transactions, fy_start_year, "EQ-STG", "EQ STCG (20%)", include_quarter_table=False, include_entry_date=False),
        render_trades_section(lt_data, lt_transactions, fy_start_year, "EQ-LTG", "EQ LTCG (12.5%, ₹1.25 lakh exemption)", include_quarter_table=False, include_entry_date=False),
        render_trades_section(ne_st_data, ne_st_transactions, fy_start_year, "DEBT-ST", "Debt Short Term", include_entry_date=False),
        *debt_fund_sections,
        "<h2>Overall Summary</h2>",
        render_table(*summary_table, right_align_from=1, total_label=None),
        "<h2>DP Charges (Other Debits and Credits)</h2>",
        render_table(*dp_charges_table),
        "<h2>DP Charges Reconciliation</h2>",
        render_table(*dp_reconciliation_table, total_label=None),
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

        dp_entries = find_dp_charge_rows(z, strings)

    fy_start_year = parse_fy_start_year(dividend_rows) or parse_fy_start_year(trade_rows)
    client_info = find_client_info(dividend_rows)

    dividend_data = find_dividend_data(dividend_rows)

    st_data = extract_trades(find_section(trade_rows, "Equity - Short Term"))
    lt_data = extract_trades(find_section(trade_rows, "Equity - Long Term"))

    ne_data = extract_trades(find_section(trade_rows, "Non Equity"))
    ne_st_data, ne_lt_data = split_by_holding_period(ne_data)

    sections = [
        ("EQ-STG", st_data),
        ("EQ-LTG", lt_data),
        ("NE-STCG", ne_st_data),
        ("NE-LTCG", ne_lt_data),
    ]
    dp_alloc = compute_dp_allocation(sections)

    raw_transactions = {label: group_transactions(trades) for label, trades in sections}
    finalized_transactions = {label: finalize_transactions(raw_transactions[label], dp_alloc[label]) for label, _ in sections}

    mf_data = extract_mf_trades(find_section(trade_rows, "Mutual Funds"))
    debt_buckets = classify_debt_fund_trades(mf_data)

    # Gold ETF long-term trades (Non-Equity section, >12 months, 12.5% no
    # indexation) land in the same ITR2 Schedule CG row as debt funds bought
    # before 1-Apr-2023 and sold on/after 23-Jul-2024 (>24 months, same rate),
    # so they're clubbed into one combined section/summary row. Entry date is
    # kept (unlike the other sections) so Duration can disambiguate the two:
    # anything under ~24 months must be Gold ETF, since debt fund LTCG here
    # requires >24 months.
    ne_lt_transactions_with_entry_date = finalize_transactions(group_mf_transactions(ne_lt_data), dp_alloc["NE-LTCG"])
    combined_new_rule_ltcg_trades = ne_lt_data + debt_buckets["new_rule_ltcg"]
    combined_new_rule_ltcg_transactions = sorted(
        ne_lt_transactions_with_entry_date + group_mf_transactions(debt_buckets["new_rule_ltcg"]),
        key=lambda t: t[2],
    )

    # Debt funds bought on/after 1-Apr-2023 are always slab rate regardless of
    # holding period (Section 50AA) - same treatment as Debt Short Term - so
    # they're clubbed into that section/summary row instead of standing alone.
    combined_debt_st_trades = ne_st_data + debt_buckets["post_2023_purchase"]
    combined_debt_st_transactions = sorted(
        finalized_transactions["NE-STCG"] + group_mf_transactions(debt_buckets["post_2023_purchase"]),
        key=lambda t: t[2],
    )

    debt_fund_sections = [
        render_trades_section(
            combined_new_rule_ltcg_trades,
            combined_new_rule_ltcg_transactions,
            fy_start_year,
            "DEBT-NEW-LTCG",
            "LTCG 12.5% (no indexation) - Gold ETF (held > 12 months) & Debt Fund bought before 1-Apr-2023, sold on/after 23-Jul-2024 (held > 24 months)",
        ),
    ]

    summary_table = build_summary_rows(
        [
            ("EQ STCG (20%)", st_data, finalized_transactions["EQ-STG"]),
            ("EQ LTCG (12.5%, ₹1.25 lakh exemption)", lt_data, finalized_transactions["EQ-LTG"]),
            ("Debt Short Term", combined_debt_st_trades, combined_debt_st_transactions),
            ("Gold ETF & Debt Fund LTCG (12.5%)", combined_new_rule_ltcg_trades, combined_new_rule_ltcg_transactions),
        ]
    )

    unique_sale_count = len({(t["Symbol"], t["Exit"]) for _, trades in sections for t in trades})
    dp_charges_table = build_dp_charges_rows(dp_entries)
    dp_reconciliation_table = build_dp_reconciliation_rows(
        dp_entries,
        unique_sale_count,
        {label: len(raw_transactions[label]) for label, _ in sections},
    )

    report_html = build_report_html(
        client_info,
        dividend_data,
        summary_table,
        st_data,
        finalized_transactions["EQ-STG"],
        lt_data,
        finalized_transactions["EQ-LTG"],
        combined_debt_st_trades,
        combined_debt_st_transactions,
        fy_start_year,
        debt_fund_sections,
        dp_charges_table,
        dp_reconciliation_table,
    )

    output_path = Path(path).with_name(Path(path).stem + "-report.html")
    output_path.write_text(report_html, encoding="utf-8")

    print(f"Report written to {output_path}")
    webbrowser.open(output_path.resolve().as_uri())


if __name__ == "__main__":
    main()
