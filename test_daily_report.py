import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace

from openpyxl import load_workbook

import daily_report as dr
import deal_sheet as ds
from ibkr_market_scanner import ScanCandidate


def deal_row(**kwargs):
    deal = ds.Deal("123", "BHF", "Brighthouse Financial")
    for key, value in {"cash": 70.0, "price": 44.73, "close_date": date(2026, 12, 6), "close_basis": "guided",
                       "undisturbed": 40.0, "domicile": "United States", "exchanges": ["Nasdaq"],
                       "market_cap": 2.6e9, **kwargs}.items():
        setattr(deal, key, value)
    return ds.deal_row(deal, date(2026, 10, 28))


def lead(symbol, kind, provider="UK Takeover Panel", basis="Domicile United Kingdom"):
    contract = SimpleNamespace(conId=1, symbol=symbol, localSymbol=symbol, secType="STK", exchange=provider)
    return ScanCandidate(contract=contract, description=f"{symbol} plc", headlines=[{
        "provider": provider, "headline": f"{symbol} in offer period", "time": "2026-11-09", "url": "https://x"}],
        research={"tier": f"Hard catalyst · {kind}", "risks": f"{basis}. Price it in the deal calculator"})


class DailyReportTests(unittest.TestCase):
    def build(self):
        rows = [deal_row(), deal_row(domicile="Peru")]
        rows[1]["ticker"] = "CPAC"
        candidates = [lead("GB00B0KM9T71", "Merger / takeover"),
                      lead("SRL.AX", "Merger / takeover", "ASX announcements", "Listed in Australia; domicile not stated")]
        summary = {"screened_out": [{"ticker": "ULBI", "company": "Ultralife", "reason": "Keyword match in an SEC 8-K",
                                     "headline": "8-K", "source": "SEC EDGAR · 8-K", "url": "https://sec"}]}
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        path = dr.build_workbook(Path(folder.name) / "x.xlsx", datetime(2026, 11, 10, 7, 0).astimezone(),
                                 candidates, ["warn"], summary, rows, dr.run_config("public", 270))
        return load_workbook(path)

    def test_tabs_and_header_layout_match_the_event_monitor(self):
        wb = self.build()
        self.assertEqual(wb.sheetnames, ["Funnel", "Ideas", "Assessed", "Events", "Pipeline", "Rejected", "Config"])
        ideas = wb["Ideas"]
        self.assertEqual(ideas["B4"].value, "Price Related")
        self.assertEqual(ideas["AK4"].value, "Sources")
        self.assertEqual([ideas[f"{c}5"].value for c, _ in dr.IDEA_COLUMNS], [h for _, h in dr.IDEA_COLUMNS])
        self.assertEqual(ideas.freeze_panes, "D6")
        self.assertEqual(ideas["C6"].value, "BHF")   # a 56% spread clears, in the verify tier
        self.assertEqual(ideas["V6"].value, "VERIFY")
        self.assertEqual(ideas["R6"].value, '=IFERROR(L6/D6-1,"--")')
        self.assertTrue(ideas["U6"].value.startswith("=IFERROR("))
        self.assertIn("Termination", ideas["AE6"].value)

    def test_assessed_lists_priced_deals_then_unpriced_catalysts(self):
        assessed = self.build()["Assessed"]
        verdicts = [assessed[f"V{r}"].value for r in range(6, 9)]
        self.assertEqual(verdicts, ["VERIFY", "SCREENED", "INSUFFICIENT_DATA"])
        self.assertEqual(assessed["C8"].value, "SRL.AX")
        self.assertEqual(assessed["F8"].value, "AU")
        self.assertEqual(assessed.auto_filter.ref, "B5:AK8")

    def test_events_pipeline_rejected_and_funnel(self):
        wb = self.build()
        self.assertEqual(wb["Events"]["A2"].value, "SRL.AX")
        self.assertEqual(wb["Pipeline"]["A2"].value, "GB00B0KM9T71 plc")   # ISIN only: no ticker yet
        rejected = {wb["Rejected"][f"A{r}"].value: wb["Rejected"][f"D{r}"].value for r in range(2, 4)}
        self.assertEqual(rejected, {"ULBI": "Catalyst", "CPAC": "Deal sheet"})
        funnel = {wb["Funnel"][f"A{r}"].value: wb["Funnel"][f"B{r}"].value for r in range(2, 12)}
        self.assertEqual(funnel["Event leads detected (all sources)"], 3)
        self.assertEqual(funnel["Ideas clearing hurdle"], 1)
        config = {wb["Config"][f"A{r}"].value: wb["Config"][f"B{r}"].value for r in range(2, 6)}
        self.assertEqual(config["hurdle_risk_adjusted_annualized"], 0.25)


if __name__ == "__main__":
    unittest.main()
