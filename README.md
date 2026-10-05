# Financial Statement Analyzer

Turn **any** financial-statement export into a board-ready financial packet —
with automatic GL mapping, an adjustable KPI grading scale, and an
auto-analyzer that writes the variance commentary for you.

I built this while doing month-end close as a Financial Analyst / Accounting
Consultant. The close ran on five journal-entry areas (broker commissions,
co-man payroll at hours × rates + 11% payroll tax load, production labor &
overhead reallocated by channel revenue %, loan interest tie-outs, and prepaid
amortization), and the packet had to come out in the exact same layout every
month for the board. v1 automated the packet build from a QBO trial balance;
**v2** accepts any statement format and adds the analysis a CFO would ask for
before the meeting.

> **Note:** the sample data in `sample_input/` is entirely fictional
> ("Acme Foods Co."). No client data is included in this repo.

## What's new in v2

- **Universal importer** (`ingest.py`) — drop in a trial balance, a P&L
  export, or a balance-sheet export, as Excel or CSV, with whatever the
  columns happen to be called (`Acct #`, `Account Title`, `Amount`, …).
  It sniffs the layout, detects the sign convention (debit/credit columns vs.
  signed balances vs. presentation-style positives), and tells you what it
  assumed on the `Checks` tab.
- **Fuzzy GL mapping** — exact account-number hits first, then keyword
  matching on the account name (`"Loan interest"` → Interest Expense even
  with no account number). Fuzzy matches are listed for review, never
  silent.
- **Auto-analyzer** (`analyze.py`) — with a prior period supplied, the new
  `Analysis` tab gives you variance by GL line ($ and %), plain-English
  key findings, and severity-ranked red flags (margin compression, opex
  outpacing revenue, receivables building, thin liquidity…).

## What it does

`generate_financials.py` normalizes the input and:

1. **Auto-maps every account to its GL statement line** via `gl_mapping.csv`
   (account → Revenue, COGS, Salaries & Wages, …, Cash, AR, AP, …),
   with fuzzy name matching as fallback.
2. **Builds a 10-tab Excel packet** with live formulas — subtotals, margins,
   and tie-outs calculate in Excel, not just on generation:
   - `Monthly Close` — close checklist (JE areas, owners, status dropdown)
   - `Raw Input` — the statement exactly as imported (layout adapts to input)
   - `Adjustments` — **the adjustable grading scale** (see below)
   - `Dashboard` — KPI scorecard with live A–F grades
   - `Analysis` — variance vs prior, narrative findings, red flags
   - `P&L` — income statement
   - `Detailed P&L` — account-level detail grouped by statement line
   - `Balance Sheet` — classified, with a balance check
   - `SCF` — indirect-method cash flow statement (full reconciliation when
     a prior period is supplied)
   - `Checks` — data-quality checks (import detection, TB balance, unmapped
     accounts, fuzzy matches, missing channel data, cash tie-out,
     net-income tie-out, blank tabs)
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

# v1-style: QBO trial balance
python generate_financials.py \
    --input sample_input/trial_balance_sample.xlsx \
    --period "January 2026" \
    --company "Acme Foods Co."

# v2: any format + prior period for variance analysis and SCF
python generate_financials.py \
    --input sample_input/trial_balance_sample.xlsx \
    --prior-tb sample_input/trial_balance_prior_sample.csv \
    --prior-label "December 2025" \
    --period "January 2026" \
    --company "Acme Foods Co."

# v2: messy presentation-style CSV, no account numbers (fuzzy mapping demo)
python generate_financials.py \
    --input sample_input/statement_messy_format.csv \
    --period "January 2026" \
    --company "Acme Foods Co."
```

This writes `financial_packet_January_2026.xlsx` in the working directory.

### Options

| Flag | Description |
|---|---|
| `--input` | Any financial-statement file, `.xlsx` or `.csv` — format auto-detected (required) |
| `--mapping` | Account → GL line CSV (default: `gl_mapping.csv`) |
| `--prior-tb` | Prior-period statement file — unlocks variance analysis + full SCF |
| `--prior-label` | Label for the prior period, e.g. `"December 2025"` |
| `--output` | Output workbook path |
| `--company` | Company name in headers |
| `--period` | Period label, e.g. `"January 2026"` |

## Inputs

- **Any statement export** — trial balance, P&L, or balance sheet. The
  importer detects debit/credit columns, signed balance columns, and
  presentation-style positive amounts, and reports its assumptions on the
  `Checks` tab. A `Channel`/`Class` column is used for the
  missing-channel-data check when present.
- **`gl_mapping.csv`** — `account_number,account_name,statement,line`.
  Unmapped accounts are flagged on the `Checks` tab instead of failing
  silently; name-based fuzzy matches are listed for review.
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
