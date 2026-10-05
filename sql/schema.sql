-- schema.sql -- SQLite schema for the financial-report-automate store.
-- Created by db.py; kept here for reference and for running the
-- analytical queries in sql/analysis_queries.sql by hand.

CREATE TABLE IF NOT EXISTS periods (
    id          INTEGER PRIMARY KEY,
    label       TEXT UNIQUE NOT NULL,   -- e.g. 'January 2026'
    source_file TEXT,                   -- the original export filename
    mode        TEXT,                   -- 'signed' or 'presentation'
    kind        TEXT NOT NULL DEFAULT 'actual'  -- actual | prior | budget
);

CREATE TABLE IF NOT EXISTS gl_mapping (
    account_number TEXT PRIMARY KEY,
    line           TEXT NOT NULL,        -- GL statement line
    statement      TEXT NOT NULL         -- 'IS' or 'BS'
);

CREATE TABLE IF NOT EXISTS accounts (
    period_id      INTEGER NOT NULL REFERENCES periods(id),
    account_number TEXT,
    account_name   TEXT NOT NULL,
    amount         REAL NOT NULL,        -- signed: debits positive
    presented      REAL NOT NULL,        -- presentation amount per GL sign
    channel        TEXT,
    gl_line        TEXT,                 -- NULL when unmapped
    fuzzy          INTEGER NOT NULL DEFAULT 0  -- 1 = matched by name, review it
);

-- Statement-line totals per period: the single source of truth every
-- packet tab and every query below builds on.
CREATE VIEW IF NOT EXISTS v_line_totals AS
SELECT p.label          AS period,
       p.kind           AS kind,
       a.gl_line        AS line,
       SUM(a.presented) AS total
FROM accounts a
JOIN periods p ON p.id = a.period_id
WHERE a.gl_line IS NOT NULL
GROUP BY p.label, p.kind, a.gl_line;
