"""
analyze.py -- automatic financial analysis (v2).

Given current- and prior-period statement-line totals, this module produces
what a CFO would ask for before the board meeting:

  * variance analysis  -- $ and % change by line, with movers flagged
  * red flags          -- rule-based warnings (margin compression, AR
                          outpacing revenue, thin liquidity, ...)
  * narrative          -- plain-English bullets summarizing the period

All thresholds are parameters, not hard-coded magic, so the analyzer can
be tuned without touching the rules.
"""

from __future__ import annotations


def _pct(cur: float, prior: float) -> float | None:
    if prior == 0:
        return None
    return (cur - prior) / abs(prior)


def variance_rows(cur: dict, prior: dict | None,
                  pct_threshold: float = 0.10,
                  abs_threshold: float = 1000.0) -> list[dict]:
    """One row per statement line: current, prior, $ change, % change, flag."""
    rows = []
    for line in cur:
        c = cur.get(line, 0.0)
        p = (prior or {}).get(line, 0.0)
        change = c - p
        pct = _pct(c, p)
        flagged = (
            prior is not None
            and abs(change) >= abs_threshold
            and pct is not None
            and abs(pct) >= pct_threshold
        )
        rows.append({
            "line": line, "current": c, "prior": p,
            "change": change, "pct": pct, "flagged": flagged,
        })
    # Biggest absolute movers first.
    rows.sort(key=lambda r: (-abs(r["change"]), r["line"]))
    return rows


def kpis(totals: dict) -> dict[str, float | None]:
    """The six dashboard KPIs from statement-line totals."""
    rev = totals.get("Revenue", 0) or 0
    gp = totals.get("Revenue", 0) - totals.get("Cost of Goods Sold", 0)
    opex = sum(totals.get(l, 0) for l in [
        "Salaries & Wages", "Payroll Taxes & Benefits", "Rent & Facilities",
        "Professional Fees", "Sales & Marketing",
        "Depreciation & Amortization", "Interest Expense",
        "Other Operating Expenses"])
    oi = gp - opex
    ni = oi + totals.get("Other Income (Expense), net", 0)
    ca = sum(totals.get(l, 0) for l in [
        "Cash & Cash Equivalents", "Accounts Receivable", "Inventory",
        "Prepaid Expenses", "Other Current Assets"])
    cl = sum(totals.get(l, 0) for l in [
        "Accounts Payable", "Accrued Expenses", "Deferred Revenue",
        "Current Portion of Long-Term Debt", "Other Current Liabilities"])
    tl = cl + totals.get("Long-Term Debt", 0) + totals.get(
        "Other Noncurrent Liabilities", 0)
    eq = totals.get("Common Stock", 0) + totals.get(
        "Retained Earnings, Beginning", 0) + ni
    return {
        "Net Profit Margin": (ni / rev) if rev else None,
        "Gross Margin": (gp / rev) if rev else None,
        "Operating Margin": (oi / rev) if rev else None,
        "Operating Expense Ratio": (opex / rev) if rev else None,
        "Current Ratio": (ca / cl) if cl else None,
        "Debt to Equity": (tl / eq) if eq else None,
        "_revenue": rev, "_net_income": ni, "_opex": opex, "_cash":
            totals.get("Cash & Cash Equivalents", 0),
    }


def red_flags(cur: dict, prior: dict | None,
              cur_k: dict, prior_k: dict | None) -> list[tuple[str, str]]:
    """Return [(severity, message)] -- 'high', 'medium', or 'low'."""
    flags: list[tuple[str, str]] = []

    def pp_move(key):
        if prior_k and cur_k.get(key) is not None and prior_k.get(key) is not None:
            return (cur_k[key] - prior_k[key]) * 100
        return None

    if prior is not None:
        npm = pp_move("Net Profit Margin")
        if npm is not None and npm <= -2:
            flags.append(("high",
                          f"Net margin compressed {abs(npm):.1f}pp vs prior period."))
        rev_pct = _pct(cur.get("Revenue", 0), prior.get("Revenue", 0))
        if rev_pct is not None and rev_pct <= -0.05:
            flags.append(("high", f"Revenue declined {abs(rev_pct):.1%} vs prior period."))
        opex_pct = _pct(cur_k["_opex"], (prior_k or {}).get("_opex", 0))
        if (rev_pct is not None and opex_pct is not None
                and opex_pct - rev_pct > 0.03):
            flags.append(("medium",
                          "Operating expenses growing faster than revenue."))
        ar_pct = _pct(cur.get("Accounts Receivable", 0),
                      prior.get("Accounts Receivable", 0))
        if (rev_pct is not None and ar_pct is not None
                and ar_pct - rev_pct > 0.05):
            flags.append(("medium",
                          "Receivables outpacing revenue -- collection risk building."))
        gm = pp_move("Gross Margin")
        if gm is not None and gm <= -2:
            flags.append(("medium",
                          f"Gross margin compressed {abs(gm):.1f}pp -- check pricing/COGS."))

    cash = cur_k["_cash"]
    if cash < 0:
        flags.append(("high", "Cash balance is negative -- immediate attention."))
    elif cur_k["_opex"] and cash < cur_k["_opex"] / 12:
        flags.append(("medium", "Cash covers less than one month of operating expenses."))

    cr, dte = cur_k.get("Current Ratio"), cur_k.get("Debt to Equity")
    if cr is not None and cr < 1.5:
        flags.append(("medium", f"Current ratio {cr:.2f} -- liquidity is thin."))
    if dte is not None and dte > 1.5:
        flags.append(("medium", f"Debt-to-equity {dte:.2f} -- leverage is elevated."))

    order = {"high": 0, "medium": 1, "low": 2}
    flags.sort(key=lambda f: order[f[0]])
    return flags


def _fmt_money(x: float) -> str:
    return f"${x:,.0f}"


def budget_variance_rows(cur: dict, budget: dict | None,
                         pct_threshold: float = 0.15,
                         abs_threshold: float = 2000.0) -> list[dict]:
    """Actual vs budget by GL line, mirroring variance_rows."""
    rows = []
    for line in cur:
        c = cur.get(line, 0.0)
        b = (budget or {}).get(line, 0.0)
        if c == 0 and b == 0:
            continue
        change = c - b
        pct = _pct(c, b)
        flagged = (
            budget is not None
            and abs(change) >= abs_threshold
            and pct is not None
            and abs(pct) >= pct_threshold
        )
        rows.append({"line": line, "actual": c, "budget": b,
                     "change": change, "pct": pct, "flagged": flagged})
    rows.sort(key=lambda r: (-abs(r["change"]), r["line"]))
    return rows


def budget_flags(cur: dict, budget: dict | None) -> list[tuple[str, str]]:
    """Flag material budget misses: (severity, message)."""
    flags: list[tuple[str, str]] = []
    if budget is None:
        return flags
    cur_k, bud_k = kpis(cur), kpis(budget)
    ni_pct = _pct(cur_k["_net_income"], bud_k["_net_income"])
    if ni_pct is not None and abs(ni_pct) >= 0.10:
        verb = "beat" if ni_pct > 0 else "missed"
        flags.append(("medium" if abs(ni_pct) < 0.25 else "high",
                      f"Net income {verb} budget by {abs(ni_pct):.1%}."))
    for r in budget_variance_rows(cur, budget):
        if r["flagged"] and r["line"] not in ("Revenue", "Net Income"):
            d = "over" if r["change"] > 0 else "under"
            pct = f"{abs(r['pct']):.1%}" if r["pct"] is not None else "n/a"
            flags.append(("low",
                          f"{r['line']} came in {d} budget by {pct} "
                          f"({_fmt_money(abs(r['change']))})."))
    return flags


def budget_bullet(cur: dict, budget: dict | None,
                  budget_label: str = "budget") -> str | None:
    """One-line budget verdict for the narrative, or None without a budget."""
    if budget is None:
        return None
    cur_k, bud_k = kpis(cur), kpis(budget)
    ni_pct = _pct(cur_k["_net_income"], bud_k["_net_income"])
    if ni_pct is None:
        return None
    verb = "beat" if ni_pct >= 0 else "missed"
    return (f"Against {budget_label}, net income of {_fmt_money(cur_k['_net_income'])} "
            f"{verb} by {abs(ni_pct):.1%} "
            f"(budgeted {_fmt_money(bud_k['_net_income'])}).")


def narrative_bullets(cur: dict, prior: dict | None,
                      cur_k: dict, prior_k: dict | None,
                      prior_label: str = "prior period") -> list[str]:
    """Plain-English summary bullets for the packet's Analysis tab."""
    bullets: list[str] = []
    rev, ni = cur_k["_revenue"], cur_k["_net_income"]
    npm = cur_k.get("Net Profit Margin")

    if prior is not None and prior_k is not None:
        rev_pct = _pct(rev, prior_k["_revenue"])
        ni_pct = _pct(ni, prior_k["_net_income"])
        rev_dir = ("up" if (rev_pct or 0) >= 0 else "down")
        ni_dir = ("up" if (ni_pct or 0) >= 0 else "down")
        rev_s = f"{abs(rev_pct):.1%}" if rev_pct is not None else "n/a"
        ni_s = f"{abs(ni_pct):.1%}" if ni_pct is not None else "n/a"
        npm_pp = ""
        if npm is not None and prior_k.get("Net Profit Margin") is not None:
            dpp = (npm - prior_k["Net Profit Margin"]) * 100
            npm_pp = f" ({dpp:+.1f}pp vs {prior_label})"
        bullets.append(
            f"Revenue of {_fmt_money(rev)} was {rev_dir} {rev_s} vs {prior_label}; "
            f"net income of {_fmt_money(ni)} was {ni_dir} {ni_s}, "
            f"a {npm:.1%} net margin{npm_pp}."
        )
        # Top movers.
        movers = [r for r in variance_rows(cur, prior) if r["flagged"]][:3]
        if movers:
            parts = []
            for m in movers:
                d = "up" if m["change"] >= 0 else "down"
                pct = f"{abs(m['pct']):.1%}" if m["pct"] is not None else "n/a"
                parts.append(f"{m['line']} {d} {pct} ({_fmt_money(abs(m['change']))})")
            bullets.append("Largest movers: " + "; ".join(parts) + ".")
        # Margin story.
        gm = cur_k.get("Gross Margin")
        if gm is not None and prior_k.get("Gross Margin") is not None:
            dpp = (gm - prior_k["Gross Margin"]) * 100
            verb = "expanded" if dpp >= 0 else "compressed"
            bullets.append(
                f"Gross margin {verb} to {gm:.1%} ({dpp:+.1f}pp) -- "
                f"{'COGS grew slower than revenue' if dpp >= 0 else 'watch pricing and input costs'}."
            )
    else:
        bullets.append(
            f"Revenue of {_fmt_money(rev)} with net income of {_fmt_money(ni)} "
            f"({npm:.1%} net margin)." if npm is not None else
            f"Revenue of {_fmt_money(rev)}; net income {_fmt_money(ni)}."
        )

    cr, dte = cur_k.get("Current Ratio"), cur_k.get("Debt to Equity")
    if cr is not None and dte is not None:
        bullets.append(
            f"Liquidity is {'comfortable' if cr >= 2 else 'adequate' if cr >= 1.5 else 'tight'} "
            f"(current ratio {cr:.2f}); leverage at {dte:.2f}x debt-to-equity."
        )
    return bullets
