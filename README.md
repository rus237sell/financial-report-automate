# Financial Statement Analyzer

Turn a month-end trial-balance Excel export into a board-ready financial packet —
with automatic GL mapping and an adjustable KPI grading scale.

I built this while doing month-end close as a Financial Analyst / Accounting
Consultant. The close ran on five journal-entry areas (broker commissions,
co-man payroll at hours × rates + 11% payroll tax load, production labor &
overhead reallocated by channel revenue %, loan interest tie-outs, and prepaid
amortization), and the packet had to come out in the exact same layout every
month for the board. This tool automates the second half: point it at the
post-close trial balance and it builds the packet.

> **Note:** the sample data in `sample_input/` is entirely fictional
> ("Acme Foods Co."). No client data is included in this repo.

## What it does

`generate_financials.py` reads a QBO-style trial balance
(`Account Number | Account Name | Debit | Credit | Channel`) and:

1. **Auto-maps every account to its GL statement line** via `gl_mapping.csv`
   (account → Revenue, COGS, Salaries & Wages, …, Cash, AR, AP, …).
2. **Builds a 9-tab Excel packet** with live formulas — subtotals, margins,
   and tie-outs calculate in Excel, not just on generation:
   - `Monthly Close` — close checklist (JE areas, owners, status dropdown)
   - `Raw Input` — the trial balance exactly as imported
   - `Adjustments` — **the adjustable grading scale** (see below)
   - `Dashboard` — KPI scorecard with live A–F grades
   - `P&L` — income statement
   - `Detailed P&L` — account-level detail grouped by statement line
   - `Balance Sheet` — classified, with a balance check
   - `SCF` — indirect-method cash flow statement
   - `Checks` — data-quality checks (TB balance, unmapped accounts, missing
     channel data, cash tie-out, net-income tie-out, blank tabs)
3. **Grades performance on an adjustable scale.** The `Adjustments` tab holds
   each KPI's target plus the B/C/D tolerance bands. Change any yellow cell
   and every grade on the `Dashboard` recalculates instantly — e.g. set the
   Net Profit Margin target to 15% with 2/5/8-point bands and a 14.6% actual
   grades a B.

Default KPIs: Net Profit Margin, Gross Margin, Operating Margin,
Operating Expense Ratio, Current Ratio, Debt to Equity.

## Quick start

```bash
pip install -r requirements.txt

# Run against the fictional sample data
python generate_financials.py \
    --input sample_input/trial_balance_sample.xlsx \
    --period "January 2026" \
    --company "Acme Foods Co."
```

This writes `financial_packet_January_2026.xlsx` in the working directory.

### Options

| Flag | Description |
|---|---|
| `--input` | Trial-balance `.xlsx` (required) |
| `--mapping` | Account → GL line CSV (default: `gl_mapping.csv`) |
| `--prior-tb` | Prior-period trial balance — unlocks the full SCF reconciliation and prior-period BS column |
| `--output` | Output workbook path |
| `--company` | Company name in headers |
| `--period` | Period label, e.g. `"January 2026"` |

With `--prior-tb`, the `SCF` tab reconciles operating/investing/financing cash
flows from balance-sheet changes and checks ending cash; without it, change
columns show as N/A placeholders.

## Inputs

- **Trial balance export** — any QBO-style export with account number, account
  name, and debit/credit (or a single signed balance column), plus an
  optional Channel/Class column used for the missing-channel-data check.
- **`gl_mapping.csv`** — `account_number,account_name,statement,line`.
  Unmapped accounts are flagged on the `Checks` tab instead of failing
  silently.
- **`journal_entry_template.csv`** — the five recurring close areas as a
  starter JE template (broker commissions, co-man payroll, labor/overhead
  reallocation, loan interest, prepaid amortization).

## The grading scale

On the `Adjustments` tab each KPI row has:

| Column | Meaning |
|---|---|
| Target | The goal (e.g. 15% net margin) |
| B / C / D band | How far off target still earns that grade |
| Direction | `higher` (bigger is better) or `lower` (smaller is better) |

Anything worse than the D band is an F. Bands are percentage points for
margins and ratio points for ratios. The `Dashboard` grades are plain Excel
formulas, so the scale is auditable and adjustable without touching code.

## Tech

Python, `pandas` (input parsing), `openpyxl` (workbook build, formulas,
conditional formatting, data validation). No macros.
