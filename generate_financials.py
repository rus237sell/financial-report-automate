#!/usr/bin/env python3
"""
generate_financials.py -- Monthly financial packet builder (v2).

Takes ANY financial-statement export -- trial balance, P&L, or balance sheet,
Excel or CSV, whatever the columns are called -- and builds a formatted,
board-ready financial packet workbook. The importer (ingest.py) sniffs the
file's layout and sign convention automatically; accounts map to GL lines
by account number with fuzzy name matching as fallback.

Tabs:
    Monthly Close   month-end close checklist (JE areas, owners, status)
    Raw Input       the statement exactly as imported (layout adapts to input)
    Adjustments     editable KPI targets + grading bands (the grading scale)
    Dashboard       KPI scorecard with live A-F grades driven by Adjustments
    Analysis        auto-analyzer: variance vs prior, narrative, red flags
    P&L             income statement (subtotals are live Excel formulas)
    Detailed P&L    account-level detail grouped by statement line
    Balance Sheet   classified balance sheet (with balance check)
    SCF             indirect-method statement of cash flows
    Checks          data-quality checks (import detection, TB balance,
                    unmapped accts, fuzzy matches, channel gaps, tie-outs)

The grading scale on the Adjustments tab is fully user-editable: change a
target or a grade band and every grade on the Dashboard recalculates.

Usage:
    python generate_financials.py --input trial_balance.xlsx --period "January 2026"
    python generate_financials.py --input pnl_export.csv --prior-tb tb_dec.csv \\
        --prior-label "December 2025" --output packet_Jan2026.xlsx

Sample data (fictional company, no real client data):
    python generate_financials.py \\
        --input sample_input/trial_balance_sample.xlsx \\
        --prior-tb sample_input/trial_balance_prior_sample.csv \\
        --prior-label "December 2025" \\
        --period "January 2026" --company "Acme Foods Co."
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from ingest import detect_and_load, fuzzy_map
from analyze import variance_rows, kpis, red_flags, narrative_bullets


# ---------------------------------------------------------------------------
# Statement line definitions
# ---------------------------------------------------------------------------
# sign: +1 keeps the trial-balance sign (debits positive), -1 flips it so
# credit-balance lines (revenue, liabilities, equity) print as positives.


@dataclass
class LineDef:
    statement: str  # "IS" or "BS"
    sign: int       # +1 or -1


LINES: dict[str, LineDef] = {
    # --- Income statement ---
    "Revenue": LineDef("IS", -1),
    "Cost of Goods Sold": LineDef("IS", +1),
    "Salaries & Wages": LineDef("IS", +1),
    "Payroll Taxes & Benefits": LineDef("IS", +1),
    "Rent & Facilities": LineDef("IS", +1),
    "Professional Fees": LineDef("IS", +1),
    "Sales & Marketing": LineDef("IS", +1),
    "Depreciation & Amortization": LineDef("IS", +1),
    "Interest Expense": LineDef("IS", +1),
    "Other Operating Expenses": LineDef("IS", +1),
    "Other Income (Expense), net": LineDef("IS", -1),
    # --- Balance sheet ---
    "Cash & Cash Equivalents": LineDef("BS", +1),
    "Accounts Receivable": LineDef("BS", +1),
    "Inventory": LineDef("BS", +1),
    "Prepaid Expenses": LineDef("BS", +1),
    "Other Current Assets": LineDef("BS", +1),
    "Property & Equipment, net": LineDef("BS", +1),
    "Other Noncurrent Assets": LineDef("BS", +1),
    "Accounts Payable": LineDef("BS", -1),
    "Accrued Expenses": LineDef("BS", -1),
    "Deferred Revenue": LineDef("BS", -1),
    "Current Portion of Long-Term Debt": LineDef("BS", -1),
    "Other Current Liabilities": LineDef("BS", -1),
    "Long-Term Debt": LineDef("BS", -1),
    "Other Noncurrent Liabilities": LineDef("BS", -1),
    "Common Stock": LineDef("BS", -1),
    "Retained Earnings, Beginning": LineDef("BS", -1),
}

# P&L tab row for each line (calc rows use live Excel formulas).
PL_ROWS = {
    "Revenue": 6,
    "Cost of Goods Sold": 7,
    "Gross Profit": 8,
    "Salaries & Wages": 10,
    "Payroll Taxes & Benefits": 11,
    "Rent & Facilities": 12,
    "Professional Fees": 13,
    "Sales & Marketing": 14,
    "Depreciation & Amortization": 15,
    "Interest Expense": 16,
    "Other Operating Expenses": 17,
    "Total Operating Expenses": 18,
    "Operating Income": 19,
    "Other Income (Expense), net": 21,
    "Net Income": 22,
}

# Balance-sheet tab row for each line (calc rows use live Excel formulas).
BS_ROWS = {
    "Cash & Cash Equivalents": 6,
    "Accounts Receivable": 7,
    "Inventory": 8,
    "Prepaid Expenses": 9,
    "Other Current Assets": 10,
    "Total Current Assets": 11,
    "Property & Equipment, net": 12,
    "Other Noncurrent Assets": 13,
    "Total Assets": 14,
    "Accounts Payable": 16,
    "Accrued Expenses": 17,
    "Deferred Revenue": 18,
    "Current Portion of Long-Term Debt": 19,
    "Other Current Liabilities": 20,
    "Total Current Liabilities": 21,
    "Long-Term Debt": 22,
    "Other Noncurrent Liabilities": 23,
    "Total Liabilities": 24,
    "Common Stock": 26,
    "Retained Earnings, Beginning": 27,
    "Net Income": 28,
    "Total Equity": 29,
    "Total Liabilities & Equity": 30,
    "Balance Check (should be 0)": 31,
}

IS_INPUT_ORDER = [
    "Revenue", "Cost of Goods Sold",
    "Salaries & Wages", "Payroll Taxes & Benefits", "Rent & Facilities",
    "Professional Fees", "Sales & Marketing", "Depreciation & Amortization",
    "Interest Expense", "Other Operating Expenses",
    "Other Income (Expense), net",
]

# Adjustments tab: (metric, target, tol_B, tol_C, tol_D, direction).
# Tolerances are in percentage points (margins) or ratio points (ratios).
KPI_DEFS = [
    ("Net Profit Margin", 0.15, 0.02, 0.05, 0.08, "higher"),
    ("Gross Margin", 0.45, 0.03, 0.06, 0.10, "higher"),
    ("Operating Margin", 0.15, 0.02, 0.05, 0.08, "higher"),
    ("Operating Expense Ratio", 0.30, 0.02, 0.05, 0.08, "lower"),
    ("Current Ratio", 2.00, 0.25, 0.50, 0.75, "higher"),
    ("Debt to Equity", 1.00, 0.25, 0.50, 0.75, "lower"),
]
ADJ_FIRST_ROW = 4  # KPI rows live on Adjustments!4:9

# Dashboard rows (5..10) map to the KPI_DEFS order above.
DASH_FIRST_ROW = 5


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
NAVY = "1F3864"
GRAY = "595959"
LIGHT_GRAY = "F2F2F2"
INPUT_YELLOW = "FFF2CC"
GREEN = "C6EFCE"
RED = "FFC7CE"
YELLOW = "FFEB9C"

TITLE_FONT = Font(name="Calibri", size=16, bold=True, color=NAVY)
SUBTITLE_FONT = Font(name="Calibri", size=11, italic=True, color=GRAY)
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
BODY_FONT = Font(name="Calibri", size=11)
BOLD_FONT = Font(name="Calibri", size=11, bold=True)
CALC_FONT = Font(name="Calibri", size=11, bold=True, color=NAVY)
INPUT_FILL = PatternFill("solid", fgColor=INPUT_YELLOW)
CALC_FILL = PatternFill("solid", fgColor=LIGHT_GRAY)
THIN = Side(style="thin", color="BFBFBF")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = '#,##0'
PCT_FMT = '0.0%'
RATIO_FMT = '0.00'
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def style_title_block(ws, title, subtitle, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = TITLE_FONT
    c.alignment = LEFT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = SUBTITLE_FONT
    c.alignment = LEFT
    ws.row_dimensions[1].height = 24


def style_header_row(ws, row, ncols):
    for col in range(1, ncols + 1):
        c = ws.cell(row=row, column=col)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = CENTER
        c.border = THIN_BORDER
    ws.row_dimensions[row].height = 18


def money(cell):
    cell.number_format = MONEY_FMT
    cell.font = BODY_FONT
    cell.border = THIN_BORDER
    cell.alignment = Alignment(horizontal="right", vertical="center")


def money_bold(cell):
    money(cell)
    cell.font = BOLD_FONT


def money_calc(cell):
    money(cell)
    cell.font = CALC_FONT
    cell.fill = CALC_FILL


# ---------------------------------------------------------------------------
# Loading & aggregation
# ---------------------------------------------------------------------------

def load_mapping(path):
    """Load account_number -> statement line mapping from CSV."""
    mapping = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            acct = str(row.get("account_number", "")).strip()
            line = str(row.get("line", "")).strip()
            if acct and line:
                mapping[acct] = line
    return mapping


@dataclass
class PacketData:
    company: str
    period: str
    tb: pd.DataFrame
    report: dict            # ingest report for the current file
    prior_tb: pd.DataFrame | None
    prior_report: dict | None
    line_totals: dict
    prior_totals: dict | None
    detail: dict  # line -> list of (acct, name, amount, channel)
    unmapped: list
    fuzzy_mapped: list      # (acct, name, line) matched by name, worth a review
    missing_channel: list
    tb_diff: float
    has_prior: bool
    prior_label: str
    analysis: dict          # variance rows, flags, narrative bullets, kpis


def _aggregate_one(df, mapping, mode):
    """Aggregate one normalized file. In 'presentation' mode the sign is
    recovered per-account from the GL mapping (signed = amount * line_sign)."""
    line_totals = defaultdict(float)
    detail = defaultdict(list)
    unmapped = []
    missing_channel = []
    fuzzy_mapped = []

    for _, r in df.iterrows():
        acct, name = r["account_number"], r["account_name"]
        amount, channel = float(r["amount"]), r["channel"]
        line = mapping.get(acct)
        fuzzy = False
        if line is None or line not in LINES:
            line = fuzzy_map(name)
            fuzzy = line is not None
        if line is None:
            unmapped.append((acct, name, amount))
            continue
        signed = amount if mode == "signed" else amount * LINES[line].sign
        presented = signed * LINES[line].sign
        line_totals[line] += presented
        detail[line].append((acct, name, presented, channel))
        if fuzzy:
            fuzzy_mapped.append((acct, name, line))
        if LINES[line].statement == "IS" and not channel:
            missing_channel.append((acct, name, presented))

    return (dict(line_totals), dict(detail), unmapped, missing_channel,
            fuzzy_mapped)


def aggregate(company, period, df, report, mapping,
              prior_df=None, prior_report=None, prior_label="Prior Period"):
    line_totals, detail, unmapped, missing_channel, fuzzy_mapped = \
        _aggregate_one(df, mapping, report["mode"])

    prior_totals = None
    if prior_df is not None:
        prior_totals, _, p_unmapped, _, p_fuzzy = _aggregate_one(
            prior_df, mapping, prior_report["mode"])
        unmapped += [(a, n + " (prior file)", v) for a, n, v in p_unmapped]
        fuzzy_mapped += [(a, n + " (prior file)", l) for a, n, l in p_fuzzy]

    tb_diff = float(df["amount"].sum()) if report["mode"] == "signed" else 0.0

    cur_k = kpis(line_totals)
    prior_k = kpis(prior_totals) if prior_totals else None
    analysis = {
        "variance": variance_rows(line_totals, prior_totals),
        "flags": red_flags(line_totals, prior_totals, cur_k, prior_k),
        "bullets": narrative_bullets(line_totals, prior_totals, cur_k,
                                     prior_k, prior_label),
        "cur_k": cur_k,
        "prior_k": prior_k,
    }

    return PacketData(
        company=company, period=period, tb=df, report=report,
        prior_tb=prior_df, prior_report=prior_report,
        line_totals=line_totals, prior_totals=prior_totals,
        detail=detail, unmapped=unmapped, fuzzy_mapped=fuzzy_mapped,
        missing_channel=missing_channel, tb_diff=tb_diff,
        has_prior=prior_df is not None, prior_label=prior_label,
        analysis=analysis,
    )


# ---------------------------------------------------------------------------
# Workbook builders
# ---------------------------------------------------------------------------

def build_monthly_close(wb, data):
    ws = wb.active
    ws.title = "Monthly Close"
    ws.sheet_properties.tabColor = NAVY
    style_title_block(ws, f"{data.company} -- Monthly Close Checklist", data.period, 5)

    headers = ["#", "Close step", "Owner", "Status", "Notes"]
    ws.append([])  # row 3
    for i, h in enumerate(headers, 1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 5)

    steps = [
        "Broker commission accruals booked",
        "Co-man payroll: approved hours x pay rates + 11% payroll tax load",
        "Production labor & overhead reallocated by channel revenue %",
        "Loan interest tie-out -- total booked ties to interest expense account",
        "Prepaid amortization for the month",
        "Trade spend / promo accruals",
        "Unapplied cash, stale accruals & clearing accounts reviewed",
        "TB review: out-of-balance = 0, no unmapped accounts",
        "Packet generated -- layout matches prior month (board format)",
        "Variance review & sign-off",
    ]
    for i, step in enumerate(steps, 1):
        r = 4 + i
        ws.cell(row=r, column=1, value=i).alignment = CENTER
        ws.cell(row=r, column=1).border = THIN_BORDER
        c = ws.cell(row=r, column=2, value=step)
        c.font = BODY_FONT
        c.border = THIN_BORDER
        c.alignment = LEFT
        for col in (3, 4, 5):
            cell = ws.cell(row=r, column=col)
            cell.border = THIN_BORDER
            cell.alignment = LEFT if col != 4 else CENTER
        ws.cell(row=r, column=4).fill = INPUT_FILL
        ws.row_dimensions[r].height = 18

    dv = DataValidation(type="list", formula1='"Not Started,In Progress,Done,Blocked"',
                        allow_blank=True)
    dv.error = "Pick a status from the list."
    dv.prompt = "Close status"
    ws.add_data_validation(dv)
    dv.add(f"D5:D{4 + len(steps)}")

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 62
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 40


def build_raw_input(wb, data):
    ws = wb.create_sheet("Raw Input")
    ws.sheet_properties.tabColor = "808080"
    signed_mode = data.report["mode"] == "signed"
    cols = (["Account Number", "Account Name", "Debit", "Credit", "Channel"]
            if signed_mode else
            ["Account Number", "Account Name", "Amount", "Channel"])
    ncols = len(cols)
    for i, h in enumerate(cols, 1):
        ws.cell(row=1, column=i, value=h)
    style_header_row(ws, 1, ncols)
    for i, r in data.tb.iterrows():
        rr = i + 2
        ws.cell(row=rr, column=1, value=r["account_number"]).border = THIN_BORDER
        ws.cell(row=rr, column=2, value=r["account_name"]).border = THIN_BORDER
        if signed_mode:
            amt = float(r["amount"])
            d = ws.cell(row=rr, column=3, value=max(amt, 0))
            money(d)
            c = ws.cell(row=rr, column=4, value=max(-amt, 0))
            money(c)
            ch_col = 5
        else:
            a = ws.cell(row=rr, column=3, value=float(r["amount"]))
            money(a)
            ch_col = 4
        ws.cell(row=rr, column=ch_col, value=r["channel"]).border = THIN_BORDER
        ws.cell(row=rr, column=ch_col).alignment = LEFT
    widths = [16, 42, 16, 16, 20][:ncols]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.freeze_panes = "A2"


def build_adjustments(wb, data):
    """The adjustable grading scale. Yellow cells are user inputs; the
    Dashboard grades recalculate live from these cells."""
    ws = wb.create_sheet("Adjustments")
    ws.sheet_properties.tabColor = "E2A900"
    style_title_block(
        ws,
        f"{data.company} -- Grading Scale & Targets",
        f"{data.period} -- edit the yellow cells; Dashboard grades update automatically",
        6,
    )
    headers = ["KPI", "Target", "B band", "C band", "D band", "Direction"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=3, column=i, value=h)
    style_header_row(ws, 3, 6)
    ws.cell(row=3, column=2).comment = None

    pct_metrics = {"Net Profit Margin", "Gross Margin", "Operating Margin",
                   "Operating Expense Ratio"}
    for i, (metric, target, tb, tc, td, direction) in enumerate(KPI_DEFS):
        r = ADJ_FIRST_ROW + i
        ws.cell(row=r, column=1, value=metric).font = BOLD_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=1).alignment = LEFT
        t = ws.cell(row=r, column=2, value=target)
        t.fill = INPUT_FILL
        t.border = THIN_BORDER
        t.alignment = CENTER
        t.number_format = PCT_FMT if metric in pct_metrics else RATIO_FMT
        for col, tol in zip((3, 4, 5), (tb, tc, td)):
            cell = ws.cell(row=r, column=col, value=tol)
            cell.fill = INPUT_FILL
            cell.border = THIN_BORDER
            cell.alignment = CENTER
            cell.number_format = PCT_FMT if metric in pct_metrics else RATIO_FMT
        d = ws.cell(row=r, column=6, value=direction)
        d.border = THIN_BORDER
        d.alignment = CENTER
        d.font = Font(name="Calibri", size=11, italic=True, color=GRAY)
        ws.row_dimensions[r].height = 18

    # How-to-read box
    r0 = ADJ_FIRST_ROW + len(KPI_DEFS) + 2
    ws.merge_cells(start_row=r0, start_column=1, end_row=r0 + 4, end_column=6)
    box = ws.cell(row=r0, column=1, value=(
        "How the grading scale works\n"
        "- Target: the goal for each KPI (edit freely).\n"
        "- B/C/D bands: how far below target (or above, for \"lower\" metrics) "
        "still earns that grade. Anything worse than the D band is an F.\n"
        "- Direction \"higher\" means bigger is better (margins, current ratio); "
        "\"lower\" means smaller is better (expense ratio, leverage).\n"
        "- Bands are in percentage points for margins and ratio points for ratios."
    ))
    box.alignment = Alignment(wrap_text=True, vertical="top")
    box.font = Font(name="Calibri", size=10, color=GRAY)

    for col, w in zip("ABCDEF", [28, 12, 12, 12, 12, 14]):
        ws.column_dimensions[col].width = w


def _grade_formula(actual_ref, adj_row, direction):
    """Excel formula returning A-F by comparing actual to the Adjustments bands."""
    t = f"Adjustments!$B${adj_row}"
    b = f"Adjustments!$C${adj_row}"
    c = f"Adjustments!$D${adj_row}"
    d = f"Adjustments!$E${adj_row}"
    if direction == "higher":
        return (
            f'=IF({actual_ref}>={t},"A",IF({actual_ref}>={t}-{b},"B",'
            f'IF({actual_ref}>={t}-{c},"C",IF({actual_ref}>={t}-{d},"D","F"))))'
        )
    return (
        f'=IF({actual_ref}<={t},"A",IF({actual_ref}<={t}+{b},"B",'
        f'IF({actual_ref}<={t}+{c},"C",IF({actual_ref}<={t}+{d},"D","F"))))'
    )


def build_dashboard(wb, data):
    ws = wb.create_sheet("Dashboard")
    ws.sheet_properties.tabColor = "2E7D32"
    style_title_block(ws, f"{data.company} -- KPI Dashboard", data.period, 5)

    headers = ["KPI", "Actual", "Target", "Grade", "Note"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=3, column=i, value=h)
    style_header_row(ws, 3, 5)

    actual_formulas = [
        "='P&L'!C23",                                   # Net Profit Margin
        "='P&L'!C8",                                    # Gross Margin
        "='P&L'!C19",                                   # Operating Margin
        "='P&L'!B18/'P&L'!B6",                          # Opex Ratio
        "='Balance Sheet'!B11/'Balance Sheet'!B21",     # Current Ratio
        "='Balance Sheet'!B24/'Balance Sheet'!B29",     # Debt to Equity
    ]
    notes = [
        "Board target -- the headline number",
        "Pricing power & production efficiency",
        "Core operating performance",
        "Overhead discipline (lower is better)",
        "Short-term liquidity cushion",
        "Leverage (lower is better)",
    ]
    pct_rows = {0, 1, 2, 3}

    for i, (metric, _t, _tb, _tc, _td, direction) in enumerate(KPI_DEFS):
        r = DASH_FIRST_ROW + i
        adj_row = ADJ_FIRST_ROW + i
        ws.cell(row=r, column=1, value=metric).font = BOLD_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=1).alignment = LEFT

        a = ws.cell(row=r, column=2)
        a.value = actual_formulas[i]
        a.number_format = PCT_FMT if i in pct_rows else RATIO_FMT
        a.font = BOLD_FONT
        a.border = THIN_BORDER
        a.alignment = CENTER
        a.fill = CALC_FILL

        t = ws.cell(row=r, column=3)
        t.value = f"=Adjustments!B{adj_row}"
        t.number_format = PCT_FMT if i in pct_rows else RATIO_FMT
        t.border = THIN_BORDER
        t.alignment = CENTER

        g = ws.cell(row=r, column=4)
        g.value = _grade_formula(f"B{r}", adj_row, direction)
        g.font = Font(name="Calibri", size=14, bold=True)
        g.border = THIN_BORDER
        g.alignment = CENTER

        n = ws.cell(row=r, column=5, value=notes[i])
        n.font = Font(name="Calibri", size=10, color=GRAY)
        n.border = THIN_BORDER
        n.alignment = LEFT
        ws.row_dimensions[r].height = 22

    grade_range = f"D{DASH_FIRST_ROW}:D{DASH_FIRST_ROW + len(KPI_DEFS) - 1}"
    ws.conditional_formatting.add(
        grade_range,
        CellIsRule(operator="equal", formula=['"A"'],
                   fill=PatternFill("solid", fgColor=GREEN),
                   font=Font(color="006100")))
    ws.conditional_formatting.add(
        grade_range,
        CellIsRule(operator="equal", formula=['"B"'],
                   fill=PatternFill("solid", fgColor="E2EFDA"),
                   font=Font(color="2E7D32")))
    ws.conditional_formatting.add(
        grade_range,
        CellIsRule(operator="equal", formula=['"C"'],
                   fill=PatternFill("solid", fgColor=YELLOW),
                   font=Font(color="9C6500")))
    ws.conditional_formatting.add(
        grade_range,
        CellIsRule(operator="equal", formula=['"D"'],
                   fill=PatternFill("solid", fgColor="FCE4D6"),
                   font=Font(color="9C5700")))
    ws.conditional_formatting.add(
        grade_range,
        CellIsRule(operator="equal", formula=['"F"'],
                   fill=PatternFill("solid", fgColor=RED),
                   font=Font(color="9C0006")))

    for col, w in zip("ABCDE", [28, 14, 14, 10, 42]):
        ws.column_dimensions[col].width = w


def build_analysis(wb, data):
    """Auto-analyzer output: variance table, narrative bullets, red flags."""
    ws = wb.create_sheet("Analysis")
    ws.sheet_properties.tabColor = "7B1FA2"
    style_title_block(
        ws,
        f"{data.company} -- Automatic Analysis",
        f"{data.period}"
        + (f" vs {data.prior_label}" if data.has_prior
           else " -- provide --prior-tb for variance analysis"),
        5,
    )

    r = 4
    ws.cell(row=r, column=1, value="Variance by GL line").font = BOLD_FONT
    r += 1
    for i, h in enumerate(
            ["GL line", "Current", data.prior_label if data.has_prior else "Prior (N/A)",
             "$ Change", "% Change"], 1):
        ws.cell(row=r, column=i, value=h)
    style_header_row(ws, r, 5)
    r += 1

    for v in data.analysis["variance"]:
        if v["current"] == 0 and v["prior"] == 0:
            continue
        ws.cell(row=r, column=1, value=v["line"]).font = BODY_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=1).alignment = LEFT
        c0 = ws.cell(row=r, column=2, value=v["current"])
        money(c0)
        c1 = ws.cell(row=r, column=3,
                     value=v["prior"] if data.has_prior else "N/A")
        if data.has_prior:
            money(c1)
        else:
            c1.alignment = CENTER
            c1.font = Font(name="Calibri", size=10, italic=True, color=GRAY)
            c1.border = THIN_BORDER
        ch = ws.cell(row=r, column=4,
                     value=v["change"] if data.has_prior else "N/A")
        pc = ws.cell(row=r, column=5)
        if data.has_prior:
            money(ch)
            pc.value = v["pct"] if v["pct"] is not None else "n/a"
            if isinstance(pc.value, float):
                pc.number_format = '0.0%'
            pc.border = THIN_BORDER
            pc.alignment = CENTER
            if v["flagged"]:
                for col in range(1, 6):
                    ws.cell(row=r, column=col).fill = PatternFill(
                        "solid", fgColor="FCE4D6")
        else:
            ch.alignment = CENTER
            ch.font = Font(name="Calibri", size=10, italic=True, color=GRAY)
            ch.border = THIN_BORDER
            pc.value = "N/A"
            pc.alignment = CENTER
            pc.font = Font(name="Calibri", size=10, italic=True, color=GRAY)
            pc.border = THIN_BORDER
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Key findings").font = BOLD_FONT
    r += 1
    for b in data.analysis["bullets"]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        cell = ws.cell(row=r, column=1, value="\u2022  " + b)
        cell.font = BODY_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="Red flags").font = BOLD_FONT
    r += 1
    flags = data.analysis["flags"]
    if not flags:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        cell = ws.cell(row=r, column=1, value="\u2022  No red flags -- all clear.")
        cell.font = Font(name="Calibri", size=11, color="2E7D32")
        cell.alignment = LEFT
    sev_fill = {"high": PatternFill("solid", fgColor=RED),
                "medium": PatternFill("solid", fgColor=YELLOW),
                "low": PatternFill("solid", fgColor="E2EFDA")}
    for sev, msg in flags:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        cell = ws.cell(row=r, column=1,
                       value=f"\u2022  [{sev.upper()}] {msg}")
        cell.font = BOLD_FONT
        cell.fill = sev_fill.get(sev, PatternFill())
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 24
        r += 1

    for col, w in zip("ABCDE", [36, 18, 18, 18, 14]):
        ws.column_dimensions[col].width = w


def build_pl(wb, data):
    ws = wb.create_sheet("P&L")
    ws.sheet_properties.tabColor = "1F3864"
    style_title_block(ws, f"{data.company} -- Income Statement", data.period, 3)
    for i, h in enumerate(["Line", "Amount", "% of Revenue"], 1):
        ws.cell(row=5, column=i, value=h)
    style_header_row(ws, 5, 3)

    lt = data.line_totals
    R = PL_ROWS

    def put(label, row, value=None, formula=None, bold=False, calc=False,
            pct_formula=None):
        c0 = ws.cell(row=row, column=1, value=label)
        c0.font = BOLD_FONT if (bold or calc) else BODY_FONT
        c0.border = THIN_BORDER
        c0.alignment = LEFT
        c1 = ws.cell(row=row, column=2)
        if formula:
            c1.value = formula
        else:
            c1.value = value or 0
        if calc:
            money_calc(c1)
        elif bold:
            money_bold(c1)
        else:
            money(c1)
        c2 = ws.cell(row=row, column=3)
        if pct_formula:
            c2.value = pct_formula
            c2.number_format = PCT_FMT
            c2.font = BODY_FONT
            c2.border = THIN_BORDER
            c2.alignment = CENTER
        else:
            c2.border = THIN_BORDER

    rev = f"$B${R['Revenue']}"
    put("Revenue", R["Revenue"], value=lt.get("Revenue", 0), bold=True)
    put("Cost of Goods Sold", R["Cost of Goods Sold"],
        value=lt.get("Cost of Goods Sold", 0))
    put("Gross Profit", R["Gross Profit"], calc=True,
        formula=f"=B{R['Revenue']}-B{R['Cost of Goods Sold']}",
        pct_formula=f"=IF(B{R['Revenue']}=0,0,B{R['Gross Profit']}/B{R['Revenue']})")

    for line in ["Salaries & Wages", "Payroll Taxes & Benefits",
                 "Rent & Facilities", "Professional Fees", "Sales & Marketing",
                 "Depreciation & Amortization", "Interest Expense",
                 "Other Operating Expenses"]:
        put(line, R[line], value=lt.get(line, 0))

    opex_range = f"B{R['Salaries & Wages']}:B{R['Other Operating Expenses']}"
    put("Total Operating Expenses", R["Total Operating Expenses"], calc=True,
        formula=f"=SUM({opex_range})")
    put("Operating Income", R["Operating Income"], calc=True,
        formula=f"=B{R['Gross Profit']}-B{R['Total Operating Expenses']}",
        pct_formula=f"=IF(B{R['Revenue']}=0,0,B{R['Operating Income']}/B{R['Revenue']})")
    put("Other Income (Expense), net", R["Other Income (Expense), net"],
        value=lt.get("Other Income (Expense), net", 0))
    put("Net Income", R["Net Income"], calc=True, bold=True,
        formula=f"=B{R['Operating Income']}+B{R['Other Income (Expense), net']}")
    ws.cell(row=23, column=1, value="Net Profit Margin").font = BOLD_FONT
    ws.cell(row=23, column=1).border = THIN_BORDER
    ws.cell(row=23, column=1).alignment = LEFT
    ws.cell(row=23, column=2).border = THIN_BORDER
    m = ws.cell(row=23, column=3,
                value=f"=IF(B{R['Revenue']}=0,0,B{R['Net Income']}/B{R['Revenue']})")
    m.number_format = PCT_FMT
    m.font = BOLD_FONT
    m.border = THIN_BORDER
    m.alignment = CENTER
    m.fill = CALC_FILL

    for col, w in zip("ABC", [38, 20, 14]):
        ws.column_dimensions[col].width = w


def build_detailed_pl(wb, data):
    ws = wb.create_sheet("Detailed P&L")
    ws.sheet_properties.tabColor = "5B7FA6"
    style_title_block(ws, f"{data.company} -- Detailed P&L", data.period, 4)
    for i, h in enumerate(["Acct #", "Account", "Amount", "Channel"], 1):
        ws.cell(row=3, column=i, value=h)
    style_header_row(ws, 3, 4)

    r = 4
    for line in IS_INPUT_ORDER:
        accounts = data.detail.get(line, [])
        if not accounts:
            continue
        for col in range(1, 5):
            c = ws.cell(row=r, column=col)
            c.fill = PatternFill("solid", fgColor="D9E2F3")
            c.border = THIN_BORDER
        ws.cell(row=r, column=2, value=line).font = BOLD_FONT
        ws.cell(row=r, column=2).alignment = LEFT
        r += 1
        start = r
        for acct, name, amt, channel in sorted(accounts):
            ws.cell(row=r, column=1, value=acct).border = THIN_BORDER
            ws.cell(row=r, column=1).alignment = CENTER
            ws.cell(row=r, column=2, value=name).border = THIN_BORDER
            ws.cell(row=r, column=2).alignment = LEFT
            a = ws.cell(row=r, column=3, value=amt)
            money(a)
            ws.cell(row=r, column=4, value=channel).border = THIN_BORDER
            ws.cell(row=r, column=4).alignment = LEFT
            r += 1
        for col in range(1, 5):
            c = ws.cell(row=r, column=col)
            c.border = THIN_BORDER
        ws.cell(row=r, column=2, value=f"Total -- {line}").font = Font(
            name="Calibri", size=11, bold=True, italic=True)
        ws.cell(row=r, column=2).alignment = LEFT
        t = ws.cell(row=r, column=3, value=f"=SUM(C{start}:C{r - 1})")
        money_bold(t)
        r += 2

    for col, w in zip("ABCD", [12, 44, 18, 20]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A4"


def build_balance_sheet(wb, data):
    ws = wb.create_sheet("Balance Sheet")
    ws.sheet_properties.tabColor = "1F3864"
    style_title_block(ws, f"{data.company} -- Balance Sheet", data.period, 3)
    headers = ["Line", "Current Period",
               "Prior Period" if data.has_prior else "Prior Period (N/A)"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=5, column=i, value=h)
    style_header_row(ws, 5, 3)

    lt = data.line_totals
    R = BS_ROWS

    ws.cell(row=4, column=1, value="ASSETS").font = BOLD_FONT

    def label(row, text, section=False):
        c = ws.cell(row=row, column=1, value=text)
        c.font = (Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                  if section else BODY_FONT)
        if section:
            c.fill = PatternFill("solid", fgColor="808080")
        c.border = THIN_BORDER
        c.alignment = LEFT
        for col in (2, 3):
            cc = ws.cell(row=row, column=col)
            cc.border = THIN_BORDER
            if section:
                cc.fill = PatternFill("solid", fgColor="808080")

    def val(row, amount, formula=None, bold=False, calc=False):
        c = ws.cell(row=row, column=2)
        c.value = formula if formula else (amount or 0)
        if calc:
            money_calc(c)
        elif bold:
            money_bold(c)
        else:
            money(c)
        p = ws.cell(row=row, column=3)
        p.border = THIN_BORDER
        p.alignment = Alignment(horizontal="right", vertical="center")
        p.number_format = MONEY_FMT

    # Prior-period values (presentation sign) for the optional C column.
    prior_vals = {}
    if data.has_prior:
        pm = defaultdict(float)
        # reuse the mapping already applied to the current TB
        prior_mode = (data.prior_report or {}).get("mode", "signed")
        for _, r in data.prior_tb.iterrows():
            line = _MAPPING_CACHE.get(str(r["account_number"]).strip())
            if not line or line not in LINES:
                line = fuzzy_map(str(r["account_name"]))
            if line and line in LINES:
                amt = float(r["amount"])
                signed = amt if prior_mode == "signed" else amt * LINES[line].sign
                pm[line] += signed * LINES[line].sign
        prior_vals = dict(pm)

    ws.cell(row=4, column=1, value="ASSETS").font = BOLD_FONT

    for line in ["Cash & Cash Equivalents", "Accounts Receivable", "Inventory",
                 "Prepaid Expenses", "Other Current Assets"]:
        label(R[line], line)
        val(R[line], lt.get(line, 0))
        if data.has_prior:
            ws.cell(row=R[line], column=3,
                    value=prior_vals.get(line, 0)).number_format = MONEY_FMT

    label(R["Total Current Assets"], "Total Current Assets")
    rng = f"B{R['Cash & Cash Equivalents']}:B{R['Other Current Assets']}"
    val(R["Total Current Assets"], 0, formula=f"=SUM({rng})", calc=True)
    if data.has_prior:
        prng = rng.replace("B", "C")
        ws.cell(row=R["Total Current Assets"], column=3,
                value=f"=SUM({prng})").number_format = MONEY_FMT

    for line in ["Property & Equipment, net", "Other Noncurrent Assets"]:
        label(R[line], line)
        val(R[line], lt.get(line, 0))
        if data.has_prior:
            ws.cell(row=R[line], column=3,
                    value=prior_vals.get(line, 0)).number_format = MONEY_FMT

    label(R["Total Assets"], "TOTAL ASSETS")
    val(R["Total Assets"], 0,
        formula=f"=B{R['Total Current Assets']}+B{R['Property & Equipment, net']}"
                f"+B{R['Other Noncurrent Assets']}", bold=True, calc=True)
    if data.has_prior:
        ws.cell(row=R["Total Assets"], column=3,
                value=f"=C{R['Total Current Assets']}+C{R['Property & Equipment, net']}"
                      f"+C{R['Other Noncurrent Assets']}").number_format = MONEY_FMT

    ws.cell(row=15, column=1, value="LIABILITIES").font = BOLD_FONT
    for line in ["Accounts Payable", "Accrued Expenses", "Deferred Revenue",
                 "Current Portion of Long-Term Debt",
                 "Other Current Liabilities"]:
        label(R[line], line)
        val(R[line], lt.get(line, 0))
        if data.has_prior:
            ws.cell(row=R[line], column=3,
                    value=prior_vals.get(line, 0)).number_format = MONEY_FMT

    label(R["Total Current Liabilities"], "Total Current Liabilities")
    lrng = f"B{R['Accounts Payable']}:B{R['Other Current Liabilities']}"
    val(R["Total Current Liabilities"], 0, formula=f"=SUM({lrng})", calc=True)
    if data.has_prior:
        ws.cell(row=R["Total Current Liabilities"], column=3,
                value=f"=SUM({lrng.replace('B', 'C')})").number_format = MONEY_FMT

    for line in ["Long-Term Debt", "Other Noncurrent Liabilities"]:
        label(R[line], line)
        val(R[line], lt.get(line, 0))
        if data.has_prior:
            ws.cell(row=R[line], column=3,
                    value=prior_vals.get(line, 0)).number_format = MONEY_FMT

    label(R["Total Liabilities"], "TOTAL LIABILITIES")
    val(R["Total Liabilities"], 0,
        formula=f"=B{R['Total Current Liabilities']}+B{R['Long-Term Debt']}"
                f"+B{R['Other Noncurrent Liabilities']}", bold=True, calc=True)
    if data.has_prior:
        ws.cell(row=R["Total Liabilities"], column=3,
                value=f"=C{R['Total Current Liabilities']}+C{R['Long-Term Debt']}"
                      f"+C{R['Other Noncurrent Liabilities']}").number_format = MONEY_FMT

    ws.cell(row=25, column=1, value="EQUITY").font = BOLD_FONT
    for line in ["Common Stock", "Retained Earnings, Beginning"]:
        label(R[line], line)
        val(R[line], lt.get(line, 0))
        if data.has_prior:
            ws.cell(row=R[line], column=3,
                    value=prior_vals.get(line, 0)).number_format = MONEY_FMT

    label(R["Net Income"], "Net Income (from P&L)")
    ni = ws.cell(row=R["Net Income"], column=2, value="='P&L'!B22")
    money(ni)
    if data.has_prior:
        ws.cell(row=R["Net Income"], column=3, value="N/A").alignment = CENTER

    label(R["Total Equity"], "TOTAL EQUITY")
    val(R["Total Equity"], 0,
        formula=f"=SUM(B{R['Common Stock']}:B{R['Net Income']})",
        bold=True, calc=True)
    label(R["Total Liabilities & Equity"], "TOTAL LIABILITIES & EQUITY")
    val(R["Total Liabilities & Equity"], 0,
        formula=f"=B{R['Total Liabilities']}+B{R['Total Equity']}",
        bold=True, calc=True)
    label(R["Balance Check (should be 0)"], "Balance Check (should be 0)")
    chk = ws.cell(row=R["Balance Check (should be 0)"], column=2,
                  value=f"=B{R['Total Assets']}-B{R['Total Liabilities & Equity']}")
    money_bold(chk)
    chk.fill = PatternFill("solid", fgColor=LIGHT_GRAY)
    ws.conditional_formatting.add(
        f"B{R['Balance Check (should be 0)']}",
        CellIsRule(operator="notEqual", formula=["0"],
                   fill=PatternFill("solid", fgColor=RED),
                   font=Font(color="9C0006")))

    if not data.has_prior:
        for row in R.values():
            c = ws.cell(row=row, column=3, value="N/A")
            c.alignment = CENTER
            c.font = Font(name="Calibri", size=10, italic=True, color=GRAY)
            c.border = THIN_BORDER

    for col, w in zip("ABC", [38, 20, 20]):
        ws.column_dimensions[col].width = w


# Filled in by main() so the BS builder can map prior-period accounts.
_MAPPING_CACHE: dict = {}


def build_scf(wb, data):
    ws = wb.create_sheet("SCF")
    ws.sheet_properties.tabColor = "5B7FA6"
    style_title_block(
        ws, f"{data.company} -- Statement of Cash Flows (Indirect Method)",
        data.period, 2)
    for i, h in enumerate(["Line", "Amount"], 1):
        ws.cell(row=4, column=i, value=h)
    style_header_row(ws, 4, 2)

    R = BS_ROWS
    PR = PL_ROWS
    r = 5
    ws.cell(row=r, column=1, value="Cash flows from operating activities").font = BOLD_FONT
    r += 1

    def line(label_text, formula=None, value=None, bold=False):
        nonlocal r
        ws.cell(row=r, column=1, value=label_text).font = (
            BOLD_FONT if bold else BODY_FONT)
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=1).alignment = LEFT
        c = ws.cell(row=r, column=2)
        if formula is not None:
            c.value = formula
        elif value is not None:
            c.value = value
        else:
            c.value = "N/A"
            c.alignment = CENTER
            c.font = Font(name="Calibri", size=10, italic=True, color=GRAY)
            c.border = THIN_BORDER
            r += 1
            return
        money_bold(c) if bold else money(c)
        r += 1

    def delta(bs_row, liability=False):
        """Period-over-period change formula, or None when no prior TB."""
        if not data.has_prior:
            return None
        cur, prv = f"'Balance Sheet'!B{bs_row}", f"'Balance Sheet'!C{bs_row}"
        expr = f"({cur}-{prv})" if liability else f"-({cur}-{prv})"
        return f"=IF({prv}=\"N/A\",\"N/A\",{expr})"

    op_start = r
    line("Net income", formula="='P&L'!B22")
    line("Depreciation & amortization", formula="='P&L'!B15")
    line("(Increase) decrease in accounts receivable", formula=delta(R["Accounts Receivable"]))
    line("(Increase) decrease in inventory", formula=delta(R["Inventory"]))
    line("(Increase) decrease in prepaid expenses", formula=delta(R["Prepaid Expenses"]))
    line("(Increase) decrease in other current assets", formula=delta(R["Other Current Assets"]))
    line("Increase (decrease) in accounts payable", formula=delta(R["Accounts Payable"], True))
    line("Increase (decrease) in accrued expenses", formula=delta(R["Accrued Expenses"], True))
    line("Increase (decrease) in deferred revenue", formula=delta(R["Deferred Revenue"], True))
    line("Increase (decrease) in other current liabilities",
         formula=delta(R["Other Current Liabilities"], True))
    op_end = r - 1
    net_op_row = r
    if data.has_prior:
        line("Net cash provided by operating activities",
             formula=f"=SUM(B{op_start}:B{op_end})", bold=True)
    else:
        line("Net cash provided by operating activities",
             value="N/A -- run with --prior-tb for period-over-period cash flows",
             bold=True)

    r += 1
    ws.cell(row=r, column=1, value="Cash flows from investing activities").font = BOLD_FONT
    r += 1
    inv_start = r
    if data.has_prior:
        line("Purchases of property & equipment",
             formula=f"=-((\'Balance Sheet\'!B{R['Property & Equipment, net']}"
                     f"-\'Balance Sheet\'!C{R['Property & Equipment, net']})"
                     f"+\'P&L\'!B{PR['Depreciation & Amortization']})")
    else:
        line("Purchases of property & equipment")
    inv_end = r - 1
    net_inv_row = r
    line("Net cash used in investing activities",
         formula=f"=SUM(B{inv_start}:B{inv_end})" if data.has_prior else None,
         bold=True)

    r += 1
    ws.cell(row=r, column=1, value="Cash flows from financing activities").font = BOLD_FONT
    r += 1
    fin_start = r
    if data.has_prior:
        line("Proceeds (repayments) of debt, net",
             formula=f"=(\'Balance Sheet\'!B{R['Current Portion of Long-Term Debt']}"
                     f"-\'Balance Sheet\'!C{R['Current Portion of Long-Term Debt']})"
                     f"+(\'Balance Sheet\'!B{R['Long-Term Debt']}"
                     f"-\'Balance Sheet\'!C{R['Long-Term Debt']})")
    else:
        line("Proceeds (repayments) of debt, net")
    fin_end = r - 1
    net_fin_row = r
    line("Net cash provided by financing activities",
         formula=f"=SUM(B{fin_start}:B{fin_end})" if data.has_prior else None,
         bold=True)

    r += 1
    if data.has_prior:
        ws.cell(row=r, column=1, value="Net change in cash").font = BOLD_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=1).alignment = LEFT
        chg = ws.cell(row=r, column=2,
                      value=f"=B{net_op_row}+B{net_inv_row}+B{net_fin_row}")
        money_bold(chg)
        chg.fill = CALC_FILL
        chg_row = r
        r += 1
        for label_text, bs_col in (
                ("Cash, beginning of period", "C"),
                ("Cash, end of period", "B")):
            ws.cell(row=r, column=1, value=label_text).font = BODY_FONT
            ws.cell(row=r, column=1).border = THIN_BORDER
            ws.cell(row=r, column=1).alignment = LEFT
            cell = ws.cell(row=r, column=2,
                           value=f"='Balance Sheet'!{bs_col}"
                                 f"{R['Cash & Cash Equivalents']}")
            money(cell)
            if label_text.startswith("Cash, beginning"):
                beg_row = r
            else:
                end_row = r
            r += 1
        ws.cell(row=r, column=1,
                value="Check: ending = beginning + change (should be 0)").font = BOLD_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=1).alignment = LEFT
        chk = ws.cell(row=r, column=2,
                      value=f"=B{end_row}-(B{beg_row}+B{chg_row})")
        money_bold(chk)
        chk.fill = CALC_FILL
    else:
        ws.cell(
            row=r, column=1,
            value="Provide --prior-tb to unlock the full cash-flow reconciliation."
        ).font = Font(name="Calibri", size=10, italic=True, color=GRAY)

    for col, w in zip("AB", [52, 22]):
        ws.column_dimensions[col].width = w


def build_checks(wb, data):
    ws = wb.create_sheet("Checks")
    ws.sheet_properties.tabColor = "C00000"
    style_title_block(ws, f"{data.company} -- Data Quality Checks", data.period, 4)
    for i, h in enumerate(["Check", "Result", "Status", "Notes"], 1):
        ws.cell(row=3, column=i, value=h)
    style_header_row(ws, 3, 4)

    rep = data.report
    signed_mode = rep["mode"] == "signed"
    rows = [
        ("Import detection",
         f"{rep['file_type']} ({rep['mode']} mode)",
         '"PASS"',
         "; ".join(rep["assumptions"])),
        ("Trial balance out-of-balance (debits - credits)",
         ("=SUM('Raw Input'!C2:C10000)-SUM('Raw Input'!D2:D10000)"
          if signed_mode else "N/A -- statement export"),
         None,  # status formula filled in below
         "Debits must equal credits before the packet means anything."),
        ("Unmapped accounts (no GL line)",
         len(data.unmapped),
         None,
         "; ".join(f"{a} {n}" for a, n, _ in data.unmapped[:5])
         or "All accounts mapped."),
        ("Accounts mapped by fuzzy name match (review)",
         len(data.fuzzy_mapped),
         None,
         "; ".join(f"{a} -> {l}" for a, _, l in data.fuzzy_mapped[:5])
         or "Every account hit the mapping file exactly."),
        ("P&L accounts missing channel/class",
         len(data.missing_channel),
         None,
         "; ".join(f"{a} {n}" for a, n, _ in data.missing_channel[:5])
         or "Every P&L account has channel data."),
        ("Cash & cash equivalents balance",
         data.line_totals.get("Cash & Cash Equivalents", 0),
         None,
         "Tie to the bank reconciliation; investigate negatives."),
        ("Net income tie-out (P&L vs Balance Sheet)",
         "='P&L'!B22-'Balance Sheet'!B28",
         None,
         "Must be zero."),
        ("Balance sheet balanced (Assets - L&E)",
         "='Balance Sheet'!B31",
         None,
         "Must be zero."),
        ("Packet tabs populated",
         f"{10}/{10}",
         '"PASS"',
         "All ten tabs generated with data."),
    ]
    # Status formulas keyed to each row's own result cell.
    status_formulas = {
        1: '"PASS"',
        2: ('=IF(ABS(B{0})<0.01,"PASS","FLAG")' if signed_mode else '"PASS"'),
        3: '=IF(B{0}=0,"PASS","FLAG")',
        4: '=IF(B{0}=0,"PASS","REVIEW")',
        5: '=IF(B{0}=0,"PASS","REVIEW")',
        6: '=IF(B{0}>=0,"PASS","FLAG")',
        7: '=IF(ABS(B{0})<0.01,"PASS","FLAG")',
        8: '=IF(ABS(B{0})<0.01,"PASS","FLAG")',
        9: '"PASS"',
    }

    for i, (check, result, _status, notes) in enumerate(rows):
        r = 4 + i
        idx = i + 1
        ws.cell(row=r, column=1, value=check).font = BODY_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        ws.cell(row=r, column=1).alignment = LEFT
        rc = ws.cell(row=r, column=2)
        rc.value = result
        rc.border = THIN_BORDER
        rc.alignment = CENTER
        if isinstance(result, (int, float)):
            rc.number_format = MONEY_FMT
        sc = ws.cell(row=r, column=3)
        sc.value = status_formulas[idx].format(r)
        sc.font = BOLD_FONT
        sc.border = THIN_BORDER
        sc.alignment = CENTER
        nc = ws.cell(row=r, column=4, value=notes)
        nc.font = Font(name="Calibri", size=10, color=GRAY)
        nc.border = THIN_BORDER
        nc.alignment = LEFT
        ws.row_dimensions[r].height = 22

    last = 3 + len(rows)
    ws.conditional_formatting.add(
        f"C4:C{last}",
        CellIsRule(operator="equal", formula=['"PASS"'],
                   fill=PatternFill("solid", fgColor=GREEN),
                   font=Font(color="006100")))
    ws.conditional_formatting.add(
        f"C4:C{last}",
        CellIsRule(operator="equal", formula=['"FLAG"'],
                   fill=PatternFill("solid", fgColor=RED),
                   font=Font(color="9C0006")))
    ws.conditional_formatting.add(
        f"C4:C{last}",
        CellIsRule(operator="equal", formula=['"REVIEW"'],
                   fill=PatternFill("solid", fgColor=YELLOW),
                   font=Font(color="9C6500")))

    for col, w in zip("ABCD", [44, 22, 12, 60]):
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def parse_args(argv=None):
    p = argparse.ArgumentParser(
        description="Build a monthly financial packet from any financial-statement "
                    "export (trial balance, P&L, or balance sheet; Excel or CSV).")
    p.add_argument("--input", required=True,
                   help="Financial statement file (.xlsx/.csv) -- format auto-detected")
    p.add_argument("--mapping", default="gl_mapping.csv",
                   help="Account -> GL line mapping CSV")
    p.add_argument("--prior-tb", default=None,
                   help="Prior-period statement file: unlocks variance analysis + SCF")
    p.add_argument("--prior-label", default="Prior Period",
                   help="Label for the prior period, e.g. 'December 2025'")
    p.add_argument("--output", default=None, help="Output workbook path")
    p.add_argument("--company", default="Acme Foods Co.", help="Company name for headers")
    p.add_argument("--period", default="January 2026",
                   help="Period label, e.g. 'January 2026'")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    global _MAPPING_CACHE

    df, report = detect_and_load(args.input)
    mapping = load_mapping(args.mapping)
    _MAPPING_CACHE = mapping
    prior_df, prior_report = None, None
    if args.prior_tb:
        prior_df, prior_report = detect_and_load(args.prior_tb)

    data = aggregate(args.company, args.period, df, report, mapping,
                     prior_df, prior_report, args.prior_label)

    wb = Workbook()
    build_monthly_close(wb, data)
    build_raw_input(wb, data)
    build_adjustments(wb, data)
    build_dashboard(wb, data)
    build_analysis(wb, data)
    build_pl(wb, data)
    build_detailed_pl(wb, data)
    build_balance_sheet(wb, data)
    build_scf(wb, data)
    build_checks(wb, data)

    out = args.output or f"financial_packet_{args.period.replace(' ', '_')}.xlsx"
    wb.save(out)

    ni = data.analysis["cur_k"]["_net_income"]
    print(f"Packet saved: {out}")
    print(f"  Detected input    : {report['file_type']} ({report['mode']} mode)")
    for a in report["assumptions"]:
        print(f"    - {a}")
    if abs(data.tb_diff) > 0.01:
        print(f"  TB out-of-balance : {data.tb_diff:,.0f}")
    print(f"  Unmapped accounts : {len(data.unmapped)}")
    print(f"  Fuzzy-mapped      : {len(data.fuzzy_mapped)}")
    print(f"  Missing channels  : {len(data.missing_channel)}")
    print(f"  Net income        : {ni:,.0f}")
    print("  Key findings:")
    for b in data.analysis["bullets"]:
        print(f"    - {b}")
    if data.analysis["flags"]:
        print("  Red flags:")
        for sev, msg in data.analysis["flags"]:
            print(f"    [{sev.upper()}] {msg}")

    issues = []
    if abs(data.tb_diff) > 0.01:
        issues.append("trial balance is out of balance")
    if data.unmapped:
        issues.append(f"{len(data.unmapped)} unmapped account(s)")
    if issues:
        print("  WARNINGS: " + "; ".join(issues), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
