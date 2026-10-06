"""
db.py -- SQLite store for multi-period financial data.

Every imported statement lands in one queryable database instead of living
only inside a generated workbook:

    periods    one row per imported statement (actuals, prior, budget, ...)
    accounts   every account, normalized: signed amount, presentation amount,
               GL line, channel, and whether the mapping was fuzzy
    gl_mapping the account -> GL line map

Views:
    v_line_totals  statement-line totals per period
    v_budget_var   actual vs budget by GL line (when a budget is loaded)

The Analysis tab's variance table is produced by SQL (see variance_rows),
and sql/analysis_queries.sql holds the documented analytical queries --
monthly P&L, variances, KPI ratios, top movers, red-flag checks -- so the
whole analysis is reproducible with plain SQL.
"""

from __future__ import annotations

import sqlite3
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS periods (
    id          INTEGER PRIMARY KEY,
    label       TEXT UNIQUE NOT NULL,
    source_file TEXT,
    mode        TEXT,
    kind        TEXT NOT NULL DEFAULT 'actual'   -- actual | prior | budget
);

CREATE TABLE IF NOT EXISTS gl_mapping (
    account_number TEXT PRIMARY KEY,
    line           TEXT NOT NULL,
    statement      TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS accounts (
    period_id      INTEGER NOT NULL REFERENCES periods(id),
    account_number TEXT,
    account_name   TEXT NOT NULL,
    amount         REAL NOT NULL,   -- signed: debits positive
    presented      REAL NOT NULL,   -- presentation amount per GL sign
    channel        TEXT,
    gl_line        TEXT,            -- NULL when unmapped
    fuzzy          INTEGER NOT NULL DEFAULT 0
);

CREATE VIEW IF NOT EXISTS v_line_totals AS
SELECT p.label            AS period,
       p.kind             AS kind,
       a.gl_line          AS line,
       SUM(a.presented)   AS total
FROM accounts a
JOIN periods p ON p.id = a.period_id
WHERE a.gl_line IS NOT NULL
GROUP BY p.label, p.kind, a.gl_line;
"""

VARIANCE_SQL = """
SELECT cur.line                                      AS line,
       cur.total                                     AS current,
       prior.total                                   AS prior,
       cur.total - COALESCE(prior.total, 0)          AS change,
       CASE WHEN COALESCE(prior.total, 0) != 0
            THEN (cur.total - prior.total) / ABS(prior.total)
       END                                          AS pct
FROM      (SELECT line, total FROM v_line_totals
           WHERE period = :cur AND kind = 'actual') cur
LEFT JOIN (SELECT line, total FROM v_line_totals
           WHERE period = :prior AND kind IN ('actual', 'prior')) prior
       ON prior.line = cur.line
WHERE NOT (cur.total = 0 AND COALESCE(prior.total, 0) = 0)
ORDER BY ABS(cur.total - COALESCE(prior.total, 0)) DESC, cur.line ASC;
"""

BUDGET_VARIANCE_SQL = """
SELECT cur.line                                      AS line,
       cur.total                                     AS actual,
       bud.total                                     AS budget,
       cur.total - COALESCE(bud.total, 0)            AS change,
       CASE WHEN COALESCE(bud.total, 0) != 0
            THEN (cur.total - bud.total) / ABS(bud.total)
       END                                          AS pct
FROM      (SELECT line, total FROM v_line_totals
           WHERE period = :cur AND kind = 'actual') cur
LEFT JOIN (SELECT line, total FROM v_line_totals
           WHERE kind = 'budget') bud
       ON bud.line = cur.line
WHERE NOT (cur.total = 0 AND COALESCE(bud.total, 0) = 0)
ORDER BY ABS(cur.total - COALESCE(bud.total, 0)) DESC, cur.line ASC;
"""

KPI_SQL = """
SELECT
  SUM(CASE WHEN line = 'Revenue' THEN total ELSE 0 END) AS revenue,
  SUM(CASE WHEN line = 'Cost of Goods Sold' THEN total ELSE 0 END) AS cogs,
  SUM(CASE WHEN line IN ('Salaries & Wages','Payroll Taxes & Benefits',
        'Rent & Facilities','Professional Fees','Sales & Marketing',
        'Depreciation & Amortization','Interest Expense',
        'Other Operating Expenses') THEN total ELSE 0 END) AS opex,
  SUM(CASE WHEN line = 'Other Income (Expense), net' THEN total ELSE 0 END)
      AS other_net,
  SUM(CASE WHEN line = 'Interest Expense' THEN total ELSE 0 END)
      AS interest_expense,
  SUM(CASE WHEN line IN ('Cash & Cash Equivalents','Accounts Receivable',
        'Inventory','Prepaid Expenses','Other Current Assets')
        THEN total ELSE 0 END) AS current_assets,
  SUM(CASE WHEN line IN ('Accounts Payable','Accrued Expenses',
        'Deferred Revenue','Current Portion of Long-Term Debt',
        'Other Current Liabilities') THEN total ELSE 0 END) AS current_liab,
  SUM(CASE WHEN line IN ('Accounts Payable','Accrued Expenses',
        'Deferred Revenue','Current Portion of Long-Term Debt',
        'Other Current Liabilities','Long-Term Debt',
        'Other Noncurrent Liabilities') THEN total ELSE 0 END) AS total_liab,
  SUM(CASE WHEN line = 'Common Stock' THEN total ELSE 0 END) AS common_stock,
  SUM(CASE WHEN line = 'Retained Earnings, Beginning' THEN total ELSE 0 END)
      AS retained_begin
FROM v_line_totals
WHERE period = :period AND kind = 'actual';
"""


def connect(db_path: str | Path | None = None) -> sqlite3.Connection:
    """Open the SQLite store (in-memory when no path is given)."""
    conn = sqlite3.connect(":memory:" if db_path is None else str(db_path))
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    return conn


def load_mapping(conn: sqlite3.Connection, mapping: dict[str, str],
                 lines: dict) -> None:
    for acct, line in mapping.items():
        conn.execute(
            "INSERT OR REPLACE INTO gl_mapping (account_number, line, statement)"
            " VALUES (?, ?, ?)",
            (acct, line, lines[line].statement if line in lines else ""),
        )
    conn.commit()


def load_period(conn: sqlite3.Connection, label: str, df, report: dict,
                resolve, lines: dict, kind: str = "actual") -> int:
    """Load one normalized statement into the store. Returns the period id.

    resolve(account_number, account_name) -> (gl_line | None, was_fuzzy).
    """
    cur = conn.execute(
        "INSERT OR REPLACE INTO periods (label, source_file, mode, kind)"
        " VALUES (?, ?, ?, ?)",
        (label, report.get("file"), report.get("mode"), kind),
    )
    period_id = cur.lastrowid
    mode = report.get("mode", "signed")

    rows = []
    for _, r in df.iterrows():
        acct, name = r["account_number"], r["account_name"]
        amount = float(r["amount"])
        line, fuzzy = resolve(acct, name)
        if line is None or line not in lines:
            signed, presented, gl_line, fuzzy = amount, amount, None, 0
        else:
            sign = lines[line].sign
            signed = amount if mode == "signed" else amount * sign
            presented = signed * sign
            gl_line = line
        rows.append((period_id, acct or None, name, signed, presented,
                     r["channel"] or None, gl_line, int(bool(fuzzy))))
    conn.executemany(
        "INSERT INTO accounts (period_id, account_number, account_name,"
        " amount, presented, channel, gl_line, fuzzy)"
        " VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        rows,
    )
    conn.commit()
    return period_id


def variance_rows(conn: sqlite3.Connection, cur_label: str,
                  prior_label: str) -> list[dict]:
    """Variance by GL line, computed in SQL. Mirrors analyze.variance_rows."""
    out = []
    for r in conn.execute(VARIANCE_SQL, {"cur": cur_label, "prior": prior_label}):
        out.append({"line": r["line"], "current": r["current"],
                    "prior": r["prior"], "change": r["change"],
                    "pct": r["pct"],
                    "flagged": (r["pct"] is not None
                                and abs(r["change"]) >= 1000
                                and abs(r["pct"]) >= 0.10)})
    return out


def budget_variance_rows(conn: sqlite3.Connection,
                         cur_label: str) -> list[dict]:
    """Actual vs budget by GL line, computed in SQL."""
    out = []
    for r in conn.execute(BUDGET_VARIANCE_SQL, {"cur": cur_label}):
        out.append({"line": r["line"], "actual": r["actual"],
                    "budget": r["budget"], "change": r["change"],
                    "pct": r["pct"],
                    "flagged": (r["pct"] is not None
                                and abs(r["change"]) >= 2000
                                and abs(r["pct"]) >= 0.15)})
    return out


def kpi_inputs(conn: sqlite3.Connection, period_label: str) -> dict:
    """Raw aggregates for KPI ratios, computed in SQL."""
    row = conn.execute(KPI_SQL, {"period": period_label}).fetchone()
    return dict(row) if row else {}
