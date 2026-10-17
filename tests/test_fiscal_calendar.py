"""財年週次自動換年測試。執行：venv/bin/python -m unittest discover -s tests"""
import sys
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import fill_weekly_excel as engine  # noqa: E402

D2_FY2027 = date(2026, 9, 27)   # 目前模板 D2
D2_FY2026 = date(2025, 9, 28)   # 沒換年的舊 D2


class FiscalYearStartRuleTest(unittest.TestCase):
    def test_rule_start_dates(self):
        self.assertEqual(engine.fiscal_year_start_for(date(2026, 10, 11)), date(2026, 9, 27))
        self.assertEqual(engine.fiscal_year_start_for(date(2026, 9, 26)), date(2025, 9, 28))
        self.assertEqual(engine.fiscal_year_start_for(date(2027, 9, 26)), date(2027, 9, 26))
        self.assertEqual(engine.fiscal_year_start_for(date(2028, 9, 24)), date(2027, 9, 26))
        self.assertEqual(engine.fiscal_year_start_for(date(2028, 10, 1)), date(2028, 10, 1))

    def test_d2_none_keeps_calendar_fallback(self):
        self.assertIsNone(engine.resolve_fiscal_year_start(date(2026, 10, 11), None))

    def test_d2_correct_is_kept(self):
        self.assertEqual(engine.resolve_fiscal_year_start(date(2026, 10, 11), D2_FY2027), D2_FY2027)

    def test_d2_stale_uses_rule(self):
        self.assertEqual(engine.resolve_fiscal_year_start(date(2026, 10, 11), D2_FY2026), date(2026, 9, 27))

    def test_d2_ahead_of_week_uses_rule(self):
        # 補做舊財年的週：D2 已是新財年，週次仍要算在舊財年
        self.assertEqual(engine.resolve_fiscal_year_start(date(2026, 9, 20), D2_FY2027), date(2025, 9, 28))


class FiscalWeekNumberTest(unittest.TestCase):
    CASES = [
        # (週日, D2, 期望 (Q, W), 期望季起始日)
        (date(2026, 10, 11), D2_FY2027, (1, 3), date(2026, 9, 27)),
        (date(2026, 10, 11), D2_FY2026, (1, 3), date(2026, 9, 27)),
        (date(2027, 9, 26), D2_FY2027, (1, 1), date(2027, 9, 26)),
        (date(2028, 9, 24), D2_FY2027, (4, 14), date(2028, 6, 25)),   # FY2028 第 53 週
        (date(2028, 10, 1), D2_FY2027, (1, 1), date(2028, 10, 1)),
        (date(2026, 9, 20), D2_FY2027, (4, 13), date(2026, 6, 28)),   # 近 8 週走勢跨年的前一週
    ]

    def test_week_numbers(self):
        for ws, d2, qw, _ in self.CASES:
            with self.subTest(week_start=ws, d2=d2):
                self.assertEqual(engine.compute_fiscal_week_number(ws, d2), qw)

    def test_quarter_start(self):
        for ws, d2, _, qs in self.CASES:
            with self.subTest(week_start=ws, d2=d2):
                self.assertEqual(engine.compute_quarter_start(ws, d2), qs)

    def test_quarter_start_calendar_fallback_unchanged(self):
        self.assertEqual(engine.compute_quarter_start(date(2026, 10, 11), None), date(2026, 10, 4))


if __name__ == '__main__':
    unittest.main()
