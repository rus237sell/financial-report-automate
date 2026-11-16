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


class ReturnOnAssetsTests(unittest.TestCase):
    def totals_roa(self, **over):
        # Same balance sheet as the equity-ratio tests: total assets
        # 600 (CA 240 + PP&E 300 + other noncurrent 60); NI = GP 600
        # - opex (100 salaries + 10 interest) = 490.
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
        k = kpis(self.totals_roa())
        self.assertAlmostEqual(k["_net_income"], 490.0)
        self.assertAlmostEqual(k["_total_assets"], 600.0)
        self.assertAlmostEqual(k["Return on Assets"], 490.0 / 600.0)

    def test_no_assets_is_none(self):
        t = self.totals_roa(**{
            "Cash & Cash Equivalents": 0.0, "Accounts Receivable": 0.0,
            "Inventory": 0.0, "Property & Equipment, net": 0.0,
            "Other Noncurrent Assets": 0.0})
        self.assertIsNone(kpis(t)["Return on Assets"])

    def test_dupont_identity(self):
        # ROA must equal net margin x asset turnover (revenue / assets).
        k = kpis(self.totals_roa())
        turnover = k["_revenue"] / k["_total_assets"]
        self.assertAlmostEqual(k["Return on Assets"],
                               k["Net Profit Margin"] * turnover)

    def test_roe_cross_check_via_equity_ratio(self):
        # ROA / equity ratio = net income / equity (return on equity).
        k = kpis(self.totals_roa())
        self.assertAlmostEqual(k["Return on Assets"] / k["Equity Ratio"],
                               k["_net_income"] / k["_equity"])

    def test_red_flag_high_when_negative(self):
        # Salaries 700 -> NI = 600 - 710 = -110; ROA = -110/600 < 0.
        t = self.totals_roa(**{"Salaries & Wages": 700.0})
        k = kpis(t)
        self.assertLess(k["Return on Assets"], 0)
        roa_flags = [(s, m) for s, m in red_flags(t, None, k, None)
                     if "Return on assets" in m]
        self.assertEqual(len(roa_flags), 1)
        self.assertEqual(roa_flags[0][0], "high")

    def test_red_flag_medium_between_0_and_5_pct(self):
        # Salaries 580 -> NI = 600 - 590 = 10; ROA = 10/600 = 1.7%.
        t = self.totals_roa(**{"Salaries & Wages": 580.0})
        k = kpis(t)
        self.assertAlmostEqual(k["Return on Assets"], 10.0 / 600.0)
        roa_flags = [(s, m) for s, m in red_flags(t, None, k, None)
                     if "Return on assets" in m]
        self.assertEqual(len(roa_flags), 1)
        self.assertEqual(roa_flags[0][0], "medium")

    def test_no_flag_at_or_above_5_pct(self):
        t = self.totals_roa()
        k = kpis(t)
        msgs = [m for _s, m in red_flags(t, None, k, None)]
        self.assertFalse(any("Return on assets" in m for m in msgs))

    def test_narrative_bullet(self):
        t = self.totals_roa()
        k = kpis(t)
        bullets = narrative_bullets(t, None, k, None)
        self.assertTrue(any("Return on assets is 81.7%" in b
                            for b in bullets))


class ReturnOnEquityTests(unittest.TestCase):
    def totals_roe(self, **over):
        # Same balance sheet as the ROA tests: total assets 600,
        # NI = GP 600 - opex (100 salaries + 10 interest) = 490,
        # equity = 100 common + 50 retained begin + 490 NI = 640.
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
        k = kpis(self.totals_roe())
        self.assertAlmostEqual(k["_net_income"], 490.0)
        self.assertAlmostEqual(k["_equity"], 640.0)
        self.assertAlmostEqual(k["Return on Equity"], 490.0 / 640.0)

    def test_no_equity_is_none(self):
        # Common -490 + retained 0 + NI 490 = 0 equity.
        t = self.totals_roe(**{"Common Stock": -490.0,
                               "Retained Earnings, Beginning": 0.0})
        self.assertAlmostEqual(kpis(t)["_equity"], 0.0)
        self.assertIsNone(kpis(t)["Return on Equity"])

    def test_roa_over_equity_ratio_identity(self):
        # ROE must equal ROA / equity ratio (the v7 cross-check, now
        # a KPI in its own right).
        k = kpis(self.totals_roe())
        self.assertAlmostEqual(k["Return on Equity"],
                               k["Return on Assets"] / k["Equity Ratio"])

    def test_dupont_three_factor_identity(self):
        # ROE = net margin x asset turnover x (assets / equity).
        k = kpis(self.totals_roe())
        turnover = k["_revenue"] / k["_total_assets"]
        leverage = k["_total_assets"] / k["_equity"]
        self.assertAlmostEqual(
            k["Return on Equity"],
            k["Net Profit Margin"] * turnover * leverage)

    def test_red_flag_high_when_negative(self):
        # Salaries 700 -> NI = -110; equity = 100 + 50 - 110 = 40;
        # ROE = -110/40 < 0.
        t = self.totals_roe(**{"Salaries & Wages": 700.0})
        k = kpis(t)
        self.assertLess(k["Return on Equity"], 0)
        roe_flags = [(s, m) for s, m in red_flags(t, None, k, None)
                     if "Return on equity" in m]
        self.assertEqual(len(roe_flags), 1)
        self.assertEqual(roe_flags[0][0], "high")

    def test_red_flag_medium_between_0_and_8_pct(self):
        # Salaries 580 -> NI = 10; equity = 160; ROE = 10/160 = 6.25%.
        t = self.totals_roe(**{"Salaries & Wages": 580.0})
        k = kpis(t)
        self.assertAlmostEqual(k["Return on Equity"], 10.0 / 160.0)
        roe_flags = [(s, m) for s, m in red_flags(t, None, k, None)
                     if "Return on equity" in m]
        self.assertEqual(len(roe_flags), 1)
        self.assertEqual(roe_flags[0][0], "medium")

    def test_no_flag_at_or_above_8_pct(self):
        t = self.totals_roe()
        k = kpis(t)
        msgs = [m for _s, m in red_flags(t, None, k, None)]
        self.assertFalse(any("Return on equity" in m for m in msgs))

    def test_narrative_bullet(self):
        t = self.totals_roe()
        k = kpis(t)
        bullets = narrative_bullets(t, None, k, None)
        self.assertTrue(any("Return on equity is 76.6%" in b
                            for b in bullets))


class AssetTurnoverTests(unittest.TestCase):
    def totals_at(self, **over):
        # Same balance sheet as the ROA/ROE tests: total assets 600
        # (CA 240 + PP&E 300 + other noncurrent 60); revenue 1200.
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
        k = kpis(self.totals_at())
        self.assertAlmostEqual(k["_total_assets"], 600.0)
        self.assertAlmostEqual(k["Asset Turnover"], 1200.0 / 600.0)

    def test_no_assets_is_none(self):
        t = self.totals_at(**{
            "Cash & Cash Equivalents": 0.0, "Accounts Receivable": 0.0,
            "Inventory": 0.0, "Property & Equipment, net": 0.0,
            "Other Noncurrent Assets": 0.0})
        self.assertIsNone(kpis(t)["Asset Turnover"])

    def test_dupont_identity_roa(self):
        # ROA must equal net margin x asset turnover.
        k = kpis(self.totals_at())
        self.assertAlmostEqual(k["Return on Assets"],
                               k["Net Profit Margin"] * k["Asset Turnover"])

    def test_dupont_three_factor_via_roe(self):
        # ROE = net margin x asset turnover x (assets / equity).
        k = kpis(self.totals_at())
        leverage = k["_total_assets"] / k["_equity"]
        self.assertAlmostEqual(
            k["Return on Equity"],
            k["Net Profit Margin"] * k["Asset Turnover"] * leverage)

    def test_red_flag_high_below_0_25(self):
        # Revenue 120 / assets 600 = 0.20x.
        t = self.totals_at(**{"Revenue": 120.0})
        k = kpis(t)
        self.assertAlmostEqual(k["Asset Turnover"], 0.20)
        at_flags = [(s, m) for s, m in red_flags(t, None, k, None)
                    if "Asset turnover" in m]
        self.assertEqual(len(at_flags), 1)
        self.assertEqual(at_flags[0][0], "high")

    def test_red_flag_medium_between_0_25_and_0_50(self):
        # Revenue 240 / assets 600 = 0.40x.
        t = self.totals_at(**{"Revenue": 240.0})
        k = kpis(t)
        self.assertAlmostEqual(k["Asset Turnover"], 0.40)
        at_flags = [(s, m) for s, m in red_flags(t, None, k, None)
                    if "Asset turnover" in m]
        self.assertEqual(len(at_flags), 1)
        self.assertEqual(at_flags[0][0], "medium")

    def test_no_flag_at_or_above_0_50(self):
        t = self.totals_at()
        k = kpis(t)
        msgs = [m for _s, m in red_flags(t, None, k, None)]
        self.assertFalse(any("Asset turnover" in m for m in msgs))

    def test_narrative_bullet(self):
        t = self.totals_at()
        k = kpis(t)
        bullets = narrative_bullets(t, None, k, None)
        self.assertTrue(any("Asset turnover is 2.00x" in b
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
