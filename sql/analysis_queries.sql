-- analysis_queries.sql -- the analytical queries behind the packet.
-- Run these against the SQLite store built by db.py (or --db packet.db).
-- Replace the period labels with the months you want to compare.

-- 1. Monthly P&L by GL line ----------------------------------------------
SELECT line, total
FROM v_line_totals
WHERE period = 'January 2026' AND kind = 'actual'
ORDER BY total DESC;

-- 2. Variance vs prior period ($ and %) ----------------------------------
SELECT cur.line                               AS line,
       cur.total                               AS current,
       prior.total                             AS prior,
       cur.total - COALESCE(prior.total, 0)    AS change,
       CASE WHEN COALESCE(prior.total, 0) != 0
            THEN (cur.total - prior.total) / ABS(prior.total)
       END                                    AS pct_change
FROM      (SELECT line, total FROM v_line_totals
           WHERE period = 'January 2026' AND kind = 'actual') cur
LEFT JOIN (SELECT line, total FROM v_line_totals
           WHERE period = 'December 2025' AND kind IN ('actual','prior')) prior
       ON prior.line = cur.line
ORDER BY ABS(cur.total - COALESCE(prior.total, 0)) DESC;

-- 3. KPI ratios in one row -----------------------------------------------
WITH t AS (
  SELECT
    SUM(CASE WHEN line = 'Revenue' THEN total ELSE 0 END) AS revenue,
    SUM(CASE WHEN line = 'Cost of Goods Sold' THEN total ELSE 0 END) AS cogs,
    SUM(CASE WHEN line IN ('Salaries & Wages','Payroll Taxes & Benefits',
          'Rent & Facilities','Professional Fees','Sales & Marketing',
          'Depreciation & Amortization','Interest Expense',
          'Other Operating Expenses') THEN total ELSE 0 END) AS opex,
    SUM(CASE WHEN line = 'Other Income (Expense), net'
          THEN total ELSE 0 END) AS other_net,
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
    SUM(CASE WHEN line = 'Retained Earnings, Beginning'
          THEN total ELSE 0 END) AS retained_begin
  FROM v_line_totals
  WHERE period = 'January 2026' AND kind = 'actual'
)
SELECT
  (revenue - cogs - opex + other_net) / revenue        AS net_profit_margin,
  (revenue - cogs) / revenue                          AS gross_margin,
  (revenue - cogs - opex) / revenue                   AS operating_margin,
  opex / revenue                                      AS opex_ratio,
  current_assets / current_liab                       AS current_ratio,
  total_liab / (common_stock + retained_begin
                + (revenue - cogs - opex + other_net)) AS debt_to_equity
FROM t;

-- 4. Top movers vs prior (the story behind the variance) -----------------
SELECT line, change, pct_change
FROM (
  SELECT cur.line                            AS line,
         cur.total - COALESCE(prior.total, 0) AS change,
         CASE WHEN COALESCE(prior.total, 0) != 0
              THEN (cur.total - prior.total) / ABS(prior.total)
         END                                 AS pct_change
  FROM      (SELECT line, total FROM v_line_totals
             WHERE period = 'January 2026' AND kind = 'actual') cur
  LEFT JOIN (SELECT line, total FROM v_line_totals
             WHERE period = 'December 2025') prior
         ON prior.line = cur.line
)
WHERE ABS(change) >= 1000 AND ABS(pct_change) >= 0.10
ORDER BY ABS(change) DESC
LIMIT 5;

-- 5. Red-flag checks ------------------------------------------------------
-- Margin compression > 2pp vs prior
SELECT 'margin compression' AS flag,
       (cur_m - prior_m) * 100 AS pp_change
FROM (SELECT (SUM(CASE WHEN line='Revenue' THEN total ELSE 0 END)
             - SUM(CASE WHEN line='Cost of Goods Sold' THEN total ELSE 0 END)
             - SUM(CASE WHEN line IN ('Salaries & Wages','Payroll Taxes & Benefits',
                 'Rent & Facilities','Professional Fees','Sales & Marketing',
                 'Depreciation & Amortization','Interest Expense',
                 'Other Operating Expenses') THEN total ELSE 0 END)
             + SUM(CASE WHEN line='Other Income (Expense), net' THEN total ELSE 0 END))
             / SUM(CASE WHEN line='Revenue' THEN total ELSE 0 END) AS cur_m
      FROM v_line_totals WHERE period='January 2026' AND kind='actual'),
     (SELECT (SUM(CASE WHEN line='Revenue' THEN total ELSE 0 END)
             - SUM(CASE WHEN line='Cost of Goods Sold' THEN total ELSE 0 END)
             - SUM(CASE WHEN line IN ('Salaries & Wages','Payroll Taxes & Benefits',
                 'Rent & Facilities','Professional Fees','Sales & Marketing',
                 'Depreciation & Amortization','Interest Expense',
                 'Other Operating Expenses') THEN total ELSE 0 END)
             + SUM(CASE WHEN line='Other Income (Expense), net' THEN total ELSE 0 END))
             / SUM(CASE WHEN line='Revenue' THEN total ELSE 0 END) AS prior_m
      FROM v_line_totals WHERE period='December 2025')
WHERE (cur_m - prior_m) * 100 <= -2;

-- Receivables growing > 5pp faster than revenue
SELECT 'AR outpacing revenue' AS flag,
       ar_growth - rev_growth AS pp_gap
FROM (SELECT
        (SELECT (cur.total - prior.total) / ABS(prior.total)
         FROM (SELECT total FROM v_line_totals
               WHERE period='January 2026' AND line='Accounts Receivable') cur,
              (SELECT total FROM v_line_totals
               WHERE period='December 2025' AND line='Accounts Receivable') prior
        ) AS ar_growth,
        (SELECT (cur.total - prior.total) / ABS(prior.total)
         FROM (SELECT total FROM v_line_totals
               WHERE period='January 2026' AND line='Revenue') cur,
              (SELECT total FROM v_line_totals
               WHERE period='December 2025' AND line='Revenue') prior
        ) AS rev_growth)
WHERE ar_growth - rev_growth > 0.05;

-- 6. Budget vs actual -----------------------------------------------------
SELECT cur.line                            AS line,
       cur.total                            AS actual,
       bud.total                            AS budget,
       cur.total - COALESCE(bud.total, 0)   AS variance,
       CASE WHEN COALESCE(bud.total, 0) != 0
            THEN (cur.total - bud.total) / ABS(bud.total)
       END                                  AS pct_variance
FROM      (SELECT line, total FROM v_line_totals
           WHERE period = 'January 2026' AND kind = 'actual') cur
LEFT JOIN (SELECT line, total FROM v_line_totals
           WHERE kind = 'budget') bud
       ON bud.line = cur.line
ORDER BY ABS(cur.total - COALESCE(bud.total, 0)) DESC;
