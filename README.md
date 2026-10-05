# Financial Report Automate

**Goal:** automate the monotonous month-end close packet — the same
statements, same layout, every month — from any trial balance or statement
export, with an adjustable KPI grading scale and auto-written variance
analysis.

**Macro steps:**
1. Universal importer (`ingest.py`) sniffs the file layout — debit/credit
   columns, signed balances, or presentation-style positives, whatever the
   columns are called — and normalizes it.
2. Every account maps to a GL statement line via `gl_mapping.csv`
   (fuzzy name-matching fallback, flagged for review).
3. Each period (actuals, prior, budget) loads into a SQLite store (`db.py`);
   the Analysis tab's variance tables are produced by SQL
   (`sql/analysis_queries.sql`).
4. The builder emits a 10-tab Excel packet with live formulas: close
   checklist, raw input, adjustments/grading scale, dashboard, analysis,
   P&L, detailed P&L, balance sheet, cash flow, checks.
5. Optional prior period unlocks variance analysis + full SCF; optional
   budget unlocks budget-vs-actual.

**Versions:**
- **v1** — QBO trial balance → close packet with adjustable KPI grading scale.
- **v2** — universal importer (any statement format), fuzzy GL mapping,
  auto-analyzer (variance, narrative findings, red flags).
- **v3** — budget-vs-actual; SQLite analytical store with documented SQL
  queries, verified row-for-row against the Python engine.

**Short results** (fictional "Acme Foods Co." sample data): net income
$123,800 on $849,500 revenue (14.6% margin); beat budget by 9.6%; dashboard
grades A/B/B/C/A/C on default bands; SCF reconciles to the penny.

## Quick start

```bash
pip install -r requirements.txt
python generate_financials.py \
    --input sample_input/trial_balance_sample.xlsx \
    --prior-tb sample_input/trial_balance_prior_sample.csv \
    --prior-label "December 2025" \
    --budget sample_input/budget_sample.csv \
    --budget-label "2026 Budget" \
    --period "January 2026" --company "Acme Foods Co." \
    --db packet.db
```

Then query the store yourself: `sqlite3 packet.db < sql/analysis_queries.sql`

| Flag | Description |
|---|---|
| `--input` | Any statement file, `.xlsx`/`.csv` — format auto-detected (required) |
| `--mapping` | Account → GL line CSV (default: `gl_mapping.csv`) |
| `--prior-tb` / `--prior-label` | Prior-period file — unlocks variance + full SCF |
| `--budget` / `--budget-label` | Budget file — unlocks budget-vs-actual |
| `--db` | Persist the SQLite store for ad-hoc SQL |
| `--output` | Output workbook path |

**Files:** `generate_financials.py` (packet builder) · `ingest.py`
(universal importer) · `analyze.py` (variance/KPI/flags/narrative) ·
`db.py` (SQLite store) · `sql/` (schema + documented queries) ·
`gl_mapping.csv` · `journal_entry_template.csv` (five recurring close areas:
broker commissions, co-man payroll + 11% tax, labor/overhead reallocation,
loan interest tie-out, prepaid amortization) · `sample_input/` (fictional
data only — no client data in this repo).

**Grading scale** (Adjustments tab, all editable): net margin A ≥ 10% ·
gross margin A ≥ 45% · opex ratio A ≤ 30% · current ratio A ≥ 2.0 ·
debt/equity A ≤ 1.0 · interest coverage A ≥ 5x. Change any band and the
Dashboard re-scores live.
