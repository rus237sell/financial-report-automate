"""Unit tests for analyze.py KPIs -- run: python -m unittest test_analyze."""

import unittest

from analyze import kpis, narrative_bullets, red_flags


def totals(**over):
    base = {
        "Revenue": 1200.0,
        "Cost of Goods Sold": 600.0,
        "Salaries & Wages": 100.0,
        "Interest Expense": 10.0,
        "Cash & Cash Equivalents": 120.0,
        "Accounts Receivable": 60.0,
        "Inventory": 60.0,
        "Accounts Payable": 30.0,
        "Current Portion of Long-Term Debt": 120.0,
    }
    base.update(over)
    return base


class DebtServiceCoverageTests(unittest.TestCase):
    def test_hand_math(self):
        # EBIT = (1200-600) - (100+10) + 10 = 500; debt service for the
        # month = interest 10 + CPLTD 120/12 = 20 -> 500/20 = 25.0x.
        k = kpis(totals())
        self.assertAlmostEqual(k["Debt Service Coverage"], 25.0)

    def test_interest_only_matches_interest_coverage(self):
        k = kpis(totals(**{"Current Portion of Long-Term Debt": 0.0}))
        self.assertAlmostEqual(k["Debt Service Coverage"],
                               k["Interest Coverage"])

    def test_principal_only(self):
        # No interest: EBIT = 1200-600-100 = 500, service = 240/12 = 20.
        k = kpis(totals(**{"Interest Expense": 0.0,
                           "Current Portion of Long-Term Debt": 240.0}))
        self.assertAlmostEqual(k["Debt Service Coverage"], 25.0)

    def test_no_debt_service_is_none(self):
        k = kpis(totals(**{"Interest Expense": 0.0,
                           "Current Portion of Long-Term Debt": 0.0}))
        self.assertIsNone(k["Debt Service Coverage"])
        self.assertIsNone(k["Interest Coverage"])

    def test_red_flag_high_below_1x(self):
        # EBIT = 1200-600-(480+100)+100 = 120; service = 100 + 600/12 = 150.
        t = totals(**{"Salaries & Wages": 480.0, "Interest Expense": 100.0,
                      "Current Portion of Long-Term Debt": 600.0})
        k = kpis(t)
        self.assertAlmostEqual(k["Debt Service Coverage"], 0.8)
        msgs = [m for _sev, m in red_flags(t, None, k, None)]
        self.assertTrue(any("Debt service coverage 0.80x" in m for m in msgs))

    def test_red_flag_medium_between_1_and_1_25(self):
        # EBIT = 120; service = 100 + 240/12 = 120 -> exactly 1.00x.
        t = totals(**{"Salaries & Wages": 480.0, "Interest Expense": 100.0,
                      "Current Portion of Long-Term Debt": 240.0})
        k = kpis(t)
        flags = red_flags(t, None, k, None)
        dscr_flags = [(s, m) for s, m in flags if "Debt service" in m]
        self.assertEqual(len(dscr_flags), 1)
        self.assertEqual(dscr_flags[0][0], "medium")

    def test_no_flag_at_or_above_1_25(self):
        t = totals()
        k = kpis(t)
        msgs = [m for _sev, m in red_flags(t, None, k, None)]
        self.assertFalse(any("Debt service" in m for m in msgs))

    def test_narrative_bullet(self):
        t = totals()
        k = kpis(t)
        bullets = narrative_bullets(t, None, k, None)
        self.assertTrue(any("Debt service coverage is 25.00x" in b
                            for b in bullets))


class ExistingKpiRegressionTests(unittest.TestCase):
    def test_quick_ratio_and_ccc_unchanged(self):
        k = kpis(totals())
        # Quick: (120+60) / (30+120) = 1.20 (AP + CPLTD are the only CLs).
        self.assertAlmostEqual(k["Quick Ratio"], 1.20)
        # DSO = 60/(1200*12/365); DIO = 60/(600*12/365); DPO = 30/(...)
        dso = 60 / (1200 * 12 / 365)
        dio = 60 / (600 * 12 / 365)
        dpo = 30 / (600 * 12 / 365)
        self.assertAlmostEqual(k["Cash Conversion Cycle"], dso + dio - dpo)


if __name__ == "__main__":
    unittest.main()
