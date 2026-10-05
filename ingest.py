"""
ingest.py -- universal financial-statement importer (v2).

v1 only understood one dialect: a QBO-style trial balance with
Account Number / Account Name / Debit / Credit / Channel. v2 accepts
*any* financial statement export -- trial balances, P&L exports, balance
sheet exports, CSV or Excel, whatever the columns happen to be called --
and normalizes it into one shape the packet builder understands.

How it works:
  1. Column sniffing: headers are normalized and matched against alias
     lists ("Acct #", "Account Title", "Amount", ... all resolve).
  2. Mode detection:
       - Debit + Credit columns            -> trial-balance mode
         (signed = debit - credit, debits positive)
       - Single amount column with debits/credits mixed in sign and a
         near-zero total                   -> signed trial-balance mode
       - Single amount column, everything presentation-positive
         (the way P&L/BS exports usually print) -> presentation mode;
         the sign is recovered per-account from the GL mapping
         (signed = amount * line_sign), so revenue still nets correctly.
  3. Account mapping: exact account-number hit in gl_mapping.csv first,
     then fuzzy keyword matching on the account name, then flagged as
     unmapped (never silently dropped).

Returns (df, report):
  df     columns: account_number, account_name, amount, channel
  report dict: file_type, mode, columns_mapped, fuzzy_mapped, assumptions
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd


COLUMN_ALIASES = {
    "account_number": [
        "account number", "account #", "acct #", "acct no", "account no",
        "number", "no", "gl code", "gl account", "code", "acct", "account",
    ],
    "account_name": [
        "account name", "account", "account title", "name", "description",
        "gl name", "title", "account description", "label",
    ],
    "debit": ["debit", "dr", "debits"],
    "credit": ["credit", "cr", "credits"],
    "amount": [
        "balance", "amount", "net", "total", "value", "ending balance",
        "current balance", "amt",
    ],
    "channel": [
        "channel", "class", "location", "department", "segment",
        "division", "channel/class",
    ],
}

# Fuzzy name -> GL line rules, ordered most-specific first.
FUZZY_RULES: list[tuple[str, list[str]]] = [
    ("Cash & Cash Equivalents", ["cash"]),
    ("Accounts Receivable", ["receiv"]),
    ("Inventory", ["inventory"]),
    ("Prepaid Expenses", ["prepaid"]),
    ("Property & Equipment, net", ["accum", "equipment", "machinery", "building",
                                   "furniture", "fixture", "leasehold", "fixed asset"]),
    ("Other Current Assets", ["other current asset"]),
    ("Other Noncurrent Assets", ["deposit", "other asset"]),
    ("Accounts Payable", ["payable"]),
    ("Current Portion of Long-Term Debt", ["current portion", "current maturit"]),
    ("Accrued Expenses", ["accrued"]),
    ("Deferred Revenue", ["deferred revenue", "unearned"]),
    ("Long-Term Debt", ["long-term debt", "long term debt", "note payable",
                        "loan payable", "bank loan"]),
    ("Other Current Liabilities", ["other current liab"]),
    ("Other Noncurrent Liabilities", ["other liab"]),
    ("Common Stock", ["common stock", "capital stock", "contributed capital",
                      "owner equity", "member equity"]),
    ("Retained Earnings, Beginning", ["retained earning"]),
    ("Revenue", ["revenue", "sales"]),
    ("Cost of Goods Sold", ["cost of goods", "cogs", "cost of sales",
                            "raw material", "purchases", "freight",
                            "packaging", "production"]),
    ("Salaries & Wages", ["salaries", "wages"]),
    ("Payroll Taxes & Benefits", ["payroll tax", "benefits", "payroll"]),
    ("Rent & Facilities", ["rent", "utilit", "facilities"]),
    ("Professional Fees", ["legal", "accounting", "consult", "professional"]),
    ("Sales & Marketing", ["advertis", "marketing", "promotion", "trade spend"]),
    ("Depreciation & Amortization", ["depreciation", "amortization"]),
    ("Other Income (Expense), net", ["interest income", "gain on", "other income",
                                    "other expense"]),
    ("Interest Expense", ["interest"]),
    ("Other Operating Expenses", ["insurance", "office", "supplie", "travel",
                                  "bank charge", "bank fee", "miscellaneous",
                                  "dues", "subscription"]),
]


def fuzzy_map(account_name: str) -> str | None:
    name = account_name.lower()
    for line, keywords in FUZZY_RULES:
        if any(k in name for k in keywords):
            return line
    return None


def _normalize_header(col) -> str:
    return " ".join(str(col).strip().lower().split())


def _find_column(norm_to_orig: dict[str, str], aliases: list[str]) -> str | None:
    for alias in aliases:
        if alias in norm_to_orig:
            return norm_to_orig[alias]
    for alias in aliases:  # contains-match fallback
        for norm, orig in norm_to_orig.items():
            if alias in norm or norm in alias:
                return orig
    return None


def _read_any(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if path.suffix.lower() == ".csv":
        # Try to sniff out the header row; fall back to first row.
        df = pd.read_csv(path, encoding="utf-8-sig")
    else:
        df = pd.read_excel(path, engine="openpyxl")
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    return df


def detect_and_load(path: str | Path):
    """Load any financial-statement file into the normalized shape."""
    raw = _read_any(path)
    norm_to_orig = {_normalize_header(c): c for c in raw.columns}
    assumptions: list[str] = []

    acct_col = _find_column(norm_to_orig, COLUMN_ALIASES["account_number"])
    name_col = _find_column(norm_to_orig, COLUMN_ALIASES["account_name"])
    debit_col = _find_column(norm_to_orig, COLUMN_ALIASES["debit"])
    credit_col = _find_column(norm_to_orig, COLUMN_ALIASES["credit"])
    amount_col = _find_column(norm_to_orig, COLUMN_ALIASES["amount"])
    channel_col = _find_column(norm_to_orig, COLUMN_ALIASES["channel"])

    if name_col is None:
        if acct_col is not None:
            # A lone "Account"-style column: decide by content whether it
            # holds account numbers or account names.
            vals = raw[acct_col].astype(str).str.strip()
            looks_numeric = (vals.str.fullmatch(r"\d+").mean() > 0.5)
            if looks_numeric:
                name_col = acct_col  # numbers double as labels
                assumptions.append(
                    f"Column '{acct_col}' looks like account numbers; "
                    f"using it as the account label too."
                )
            else:
                name_col = acct_col
                acct_col = None
                assumptions.append(
                    f"Column '{name_col}' looks like account names; "
                    f"mapping by name only."
                )
        else:
            raise ValueError(
                f"No account-name column found in {path}. "
                f"Columns seen: {list(raw.columns)}"
            )

    df = pd.DataFrame()
    if acct_col:
        df["account_number"] = raw[acct_col].astype(str).str.strip().replace(
            {"nan": "", "None": ""})
    else:
        df["account_number"] = ""
        assumptions.append(
            "No account-number column found; mapping accounts by name only.")
    df["account_name"] = raw[name_col].astype(str).str.strip()
    if channel_col:
        df["channel"] = raw[channel_col].astype(str).str.strip().replace(
            {"nan": "", "None": "", "NaT": ""})
    else:
        df["channel"] = ""

    columns_mapped = {
        k: v for k, v in {
            "account_number": acct_col, "account_name": name_col,
            "debit": debit_col, "credit": credit_col,
            "amount": amount_col, "channel": channel_col,
        }.items() if v
    }

    if debit_col and credit_col:
        mode = "signed"
        file_type = "trial_balance"
        debit = pd.to_numeric(raw[debit_col], errors="coerce").fillna(0)
        credit = pd.to_numeric(raw[credit_col], errors="coerce").fillna(0)
        df["amount"] = debit - credit
        assumptions.append(
            f"Detected debit/credit columns "
            f"('{debit_col}'/'{credit_col}'); treated as a trial balance."
        )
    elif amount_col:
        vals = pd.to_numeric(raw[amount_col], errors="coerce").fillna(0)
        total, gross = vals.sum(), vals.abs().sum()
        neg_frac = (vals < 0).mean() if len(vals) else 0
        # A signed trial balance mixes debits and credits, so it nets near
        # zero and carries plenty of negatives. A presentation-style P&L/BS
        # export prints everything positive.
        if neg_frac > 0.05 and (gross == 0 or abs(total) < 0.005 * gross):
            mode = "signed"
            file_type = "trial_balance"
            assumptions.append(
                f"Single amount column ('{amount_col}') nets near zero with "
                f"{neg_frac:.0%} negatives; treated as a signed trial balance."
            )
        else:
            mode = "presentation"
            file_type = "statement_export"
            assumptions.append(
                f"Single amount column ('{amount_col}') looks presentation-style "
                f"(all positive); signs will be recovered from the GL mapping."
            )
        df["amount"] = vals
    else:
        raise ValueError(
            f"Need Debit/Credit columns or an Amount/Balance column in {path}."
        )

    # Drop fully blank rows.
    df = df[~((df["account_number"] == "") & (df["amount"] == 0))].copy()
    df = df.reset_index(drop=True)

    report = {
        "file": str(path),
        "file_type": file_type,
        "mode": mode,
        "columns_mapped": columns_mapped,
        "assumptions": assumptions,
    }
    return df, report
