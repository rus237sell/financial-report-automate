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


class EquityRatioTests(unittest.TestCase):
    def totals_er(self, **over):
        # Assets: CA = 120 cash + 60 AR + 60 inv = 240; + 300 PP&E
        # + 60 other noncurrent = 600 total assets.
        # Equity: 100 common + 50 retained begin + NI. NI = GP 600
        # - opex (100 salaries + 10 interest) + 0 other = 490.
        # Equity = 640 -- deliberately NOT 600 - liabilities, so the
        # identity test below has something real to check.
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
            "Property & Equipment, net": 300.0,
            "Other Noncurrent Assets": 60.0,
            "Common Stock": 100.0,
            "Retained Earnings, Beginning": 50.0,
        }
        base.update(over)
        return base

    def test_hand_math(self):
        k = kpis(self.totals_er())
        self.assertAlmostEqual(k["_total_assets"], 600.0)
        self.assertAlmostEqual(k["_equity"], 640.0)
        self.assertAlmostEqual(k["Equity Ratio"], 640.0 / 600.0)

    def test_no_assets_is_none(self):
        t = self.totals_er(**{
            "Cash & Cash Equivalents": 0.0, "Accounts Receivable": 0.0,
            "Inventory": 0.0, "Property & Equipment, net": 0.0,
            "Other Noncurrent Assets": 0.0})
        self.assertIsNone(kpis(t)["Equity Ratio"])

    def test_balanced_identity_with_debt_to_equity(self):
        # Build a balanced sheet: liabilities = AP 30 + CPLTD 120 +
        # LTD 200 = 350; assets 600, so equity must be 250 to balance.
        # NI is 490 (above), so common + retained begin = 250 - 490.
        t = self.totals_er(**{"Long-Term Debt": 200.0,
                              "Common Stock": -300.0,
                              "Retained Earnings, Beginning": 60.0})
        k = kpis(t)
        self.assertAlmostEqual(k["_equity"], 250.0)
        self.assertAlmostEqual(k["_total_liabilities"], 350.0)
        # Balanced: equity + liabilities == assets, and the two
        # solvency KPIs must satisfy ER = 1 / (1 + D/E).
        self.assertAlmostEqual(
            k["_equity"] + k["_total_liabilities"], k["_total_assets"])
        self.assertAlmostEqual(k["Equity Ratio"],
                               1 / (1 + k["Debt to Equity"]))

    def test_red_flag_medium_between_20_and_30_pct(self):
        # Equity 150 / assets 600 = 25.0%: NI 490, so common+retained
        # = -340.
        t = self.totals_er(**{"Common Stock": -340.0,
                              "Retained Earnings, Beginning": 0.0})
        k = kpis(t)
        self.assertAlmostEqual(k["Equity Ratio"], 0.25)
        er_flags = [(s, m) for s, m in red_flags(t, None, k, None)
                    if "Equity ratio" in m]
        self.assertEqual(len(er_flags), 1)
        self.assertEqual(er_flags[0][0], "medium")

    def test_red_flag_high_below_20_pct(self):
        t = self.totals_er(**{"Common Stock": -400.0,
                              "Retained Earnings, Beginning": 0.0})
        k = kpis(t)
        self.assertAlmostEqual(k["Equity Ratio"], 0.15)
        er_flags = [(s, m) for s, m in red_flags(t, None, k, None)
                    if "Equity ratio" in m]
        self.assertEqual(len(er_flags), 1)
        self.assertEqual(er_flags[0][0], "high")

    def test_no_flag_at_or_above_30_pct(self):
        t = self.totals_er()
        k = kpis(t)
        msgs = [m for _s, m in red_flags(t, None, k, None)]
        self.assertFalse(any("Equity ratio" in m for m in msgs))

    def test_narrative_bullet(self):
        t = self.totals_er()
        k = kpis(t)
        bullets = narrative_bullets(t, None, k, None)
        self.assertTrue(any("Equity ratio is 106.7%" in b for b in bullets))


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
