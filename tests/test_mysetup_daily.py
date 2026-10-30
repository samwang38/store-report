"""MySetup 每人每日每裝置計數測試。執行：venv/bin/python -m unittest discover -s tests"""
import sys
import tempfile
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import mysetup  # noqa: E402
import mysetup_daily as md  # noqa: E402
import mysetup_mail as mm  # noqa: E402
from test_mysetup_mail import make_xlsx  # noqa: E402

D1 = date.today() - timedelta(days=3)
D2 = D1 + timedelta(days=1)
at = lambda d, h: datetime(d.year, d.month, d.day, h)
OTHERS = [(f"9{i}", aid, "其他店", "iPhone", "某 人", at(D1, 9)) for i, aid in enumerate((444912, 1293891, 411723, 624351))]
ROWS_D1 = [("1", 337277, "士林", "iPhone", "吳 學諭", at(D1, 10)),
           ("2", 337277, "士林", "iPhone", "吳 學諭", at(D1, 11)),
           ("3", 337277, "士林", "Watch", "Zheng CSR NaNa", at(D1, 12)),
           ("4", 337277, "士林", "iPhone", "王 CSR 彥翔", at(D1, 13)),      # 規則指定
           ("5", 337277, "士林", "iPhone", "路人 甲", at(D1, 14)),          # 對不上
           ("6", 337277, "士林", "iPhone", "不計 入", at(D1, 15))] + OTHERS  # 規則 SKIP
CANDIDATES = {"004": {"SA1614": ("吳學諭", "Shiyu"), "SA3011": ("鄭娜娜", "NaNa")}}
RULES = {"004": [{"name": "王 CSR 彥翔", "code": "SA2765"}, {"name": "不計 入", "code": mysetup.SKIP}]}


class DailyCountsTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.orig = mm.ARCHIVE_PATH
        mm.ARCHIVE_PATH = Path(self.tmp.name) / "a.sqlite"
        mm.import_file(make_xlsx(ROWS_D1), f"Personal Setup_{D1}_{D1}.xlsx")
        others_d2 = [(f"8{i}", aid, "其他店", "iPhone", "某 人", at(D2, 9))
                     for i, aid in enumerate((444912, 1293891, 411723, 624351))]
        mm.import_file(make_xlsx([("7", 337277, "士林", "Watch", "吳 學諭", at(D2, 10))] + others_d2),
                       f"Personal Setup_{D2}_{D2}.xlsx")

    def tearDown(self):
        mm.ARCHIVE_PATH = self.orig
        self.tmp.cleanup()

    def test_counts_by_person_day_device(self):
        r = md.daily_counts(["004", "068"], D1 - timedelta(days=1), D2, RULES, candidates=CANDIDATES)
        c = r["counts"]["004"]
        self.assertEqual(sorted(c), [D1.isoformat(), D2.isoformat()])      # 沒收到信的前一天不出現
        self.assertEqual(c[D1.isoformat()]["SA1614"], {"Mac": 0, "iPad": 0, "iPhone": 2, "Watch": 0})
        self.assertEqual(c[D1.isoformat()]["SA3011"]["Watch"], 1)            # 英文名對上
        self.assertEqual(c[D1.isoformat()]["SA2765"]["iPhone"], 1)           # 店長規則優先
        self.assertNotIn("不計 入", str(r["unmatched"]))
        self.assertEqual(r["unmatched"]["004"][D1.isoformat()], {"路人 甲": {"Mac": 0, "iPad": 0, "iPhone": 1, "Watch": 0}})
        self.assertEqual(c[D2.isoformat()]["SA1614"]["Watch"], 1)
        self.assertEqual(r["counts"]["068"], {D1.isoformat(): {}, D2.isoformat(): {}})   # 有資料日但 0 筆

    def test_period_totals(self):
        c = md.daily_counts(["004"], D1, D2, RULES, candidates=CANDIDATES)["counts"]["004"]
        self.assertEqual(md.period_totals(c, D1, D2)["SA1614"], {"Mac": 0, "iPad": 0, "iPhone": 2, "Watch": 1})
        self.assertEqual(md.period_totals(c, D2, D2), {"SA1614": {"Mac": 0, "iPad": 0, "iPhone": 0, "Watch": 1}})


if __name__ == "__main__":
    unittest.main()
