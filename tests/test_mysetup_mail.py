"""MySetup 每日信件匯入測試。執行：venv/bin/python -m unittest discover -s tests"""
import sys
import tempfile
import unittest
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import mysetup  # noqa: E402
import mysetup_mail as mm  # noqa: E402

HEADER = ["Unique ID", "Channel", "HQ ID", "Reseller Name", "Apple ID", "POS Name", "Device", "Staff Name",
          "Apple Services", "Personal Setup Start Time", "Personal Setup Submission Time", "Customer Type"]


def make_xlsx(rows):
    """rows: [(uid, apple_id, pos, device, staff, submitted)]，依 Device 分頁（同 Apple 匯出檔）。"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet, dev in (("iPhone", "iPhone"), ("Apple Watch", "Watch")):
        ws = wb.create_sheet(sheet)
        ws.append(HEADER)
        for uid, aid, pos, d, staff, sub in rows:
            if d == dev:
                ws.append([uid, "Speciality", "264947", "晶實", aid, pos, d, staff, "N", sub, sub, "Walk-in"])
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


DAY = date.today() - timedelta(days=2)
T = lambda h: datetime(DAY.year, DAY.month, DAY.day, h, 0, 0)
DAILY = [("1001", 337277, "STUDIO A@士林文林", "iPhone", "王 CSR 彥翔", T(12)),
         ("1002", 337277, "STUDIO A@士林文林", "Watch", "盧 辰昕", T(15)),
         ("1003", 3842054, "STUDIOA@新店裕隆城", "iPhone", "陳 小明", T(16)),
         ("1004", 544641, "STUDIO A@台中彩虹3C店", "iPhone", "林 一", T(17)),
         ("1005", 444912, "STUDIO A＠高雄夢時代", "iPhone", "林 衿融", T(18)),
         ("1006", 1293891, "STUDIO A@台北微風", "Watch", "張 三", T(19))]
DAILY_NAME = f"Personal Setup_{DAY}_{DAY}.xlsx"


class ArchiveTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.db = Path(self.tmp.name) / "a.sqlite"

    def tearDown(self):
        self.tmp.cleanup()

    def test_parse_reads_apple_id(self):
        recs = mysetup.parse_personal_setup(make_xlsx(DAILY))
        self.assertEqual({r["apple_id"] for r in recs}, {"337277", "3842054", "544641", "444912", "1293891"})
        self.assertEqual(recs[0]["pos"], "STUDIO A@士林文林")

    def test_import_dedupes_and_splits_by_store(self):
        r = mm.import_file(make_xlsx(DAILY), DAILY_NAME, path=self.db)
        self.assertEqual((r["rows"], r["new"], r["from"], r["to"], r["company"]), (6, 5, DAY, DAY, True))   # 彩虹那筆不存
        self.assertEqual(mm.import_file(make_xlsx(DAILY), DAILY_NAME, path=self.db)["new"], 0)
        shilin = mm.records_for("004", DAY, DAY, path=self.db)
        self.assertEqual(sorted(x["uid"] for x in shilin), ["1001", "1002"])
        self.assertEqual([x["uid"] for x in mm.records_for("068", DAY, DAY, path=self.db)], ["1003"])
        self.assertEqual(mm.records_for("004", DAY + timedelta(days=1), DAY + timedelta(days=1), path=self.db), [])
        self.assertEqual(mm.records_for("057", DAY, DAY, path=self.db), [])   # 沒有對照的門市

    def test_status_and_missing_days(self):
        self.assertIsNone(mm.missing_days(DAY, DAY, path=self.db))
        mm.import_file(make_xlsx(DAILY), DAILY_NAME, path=self.db)
        st = mm.status(path=self.db)
        self.assertEqual((st["from"], st["through"]), (DAY.isoformat(), DAY.isoformat()))
        self.assertEqual(st["missing"], [(DAY + timedelta(days=1)).isoformat()])   # 昨天的信還沒到
        self.assertEqual(st["unknown"], [])                         # 台中彩虹 3C 店寫死排除，不算對不上
        start = DAY - timedelta(days=2)
        self.assertEqual(mm.missing_days(start, DAY + timedelta(days=5), path=self.db),
                         [start, DAY - timedelta(days=1), DAY + timedelta(days=1)])
        # 店長上傳檔涵蓋的日子不算缺
        self.assertEqual(mm.missing_days(start, DAY, path=self.db, also_covered=(start, DAY - timedelta(days=1))), [])

    def test_single_store_export_covers_only_that_store(self):
        """店長自己匯出寄到信箱的單店檔（10/01–10/25 只有士林）不能當成全公司那幾天都有資料。"""
        start, end = DAY - timedelta(days=5), DAY - timedelta(days=3)
        shilin_only = [("2001", 337277, "STUDIO A@士林文林", "iPhone", "蘇 悅", datetime(start.year, start.month, start.day, 10))]
        r = mm.import_file(make_xlsx(shilin_only), f"Personal Setup_{start}_{end}.xlsx", path=self.db)
        self.assertFalse(r["company"])
        self.assertEqual(mm.covered_days(start, end, path=self.db, shop_id="004"), [start, start + timedelta(days=1), end])
        self.assertEqual(mm.covered_days(start, end, path=self.db, shop_id="068"), [])
        self.assertEqual(mm.status(path=self.db)["through"], None)      # 還沒有全公司每日檔
        mm.import_file(make_xlsx(DAILY), DAILY_NAME, path=self.db)
        self.assertEqual(mm.missing_days(start, DAY, path=self.db, shop_id="068"),
                         [start, start + timedelta(days=1), end, end + timedelta(days=1), DAY - timedelta(days=1)])
        self.assertEqual(mm.missing_days(start, DAY, path=self.db, shop_id="004"),
                         [end + timedelta(days=1), DAY - timedelta(days=1)])
        self.assertEqual(mm.status(path=self.db)["from"], DAY.isoformat())

    def test_old_archive_without_scope_is_rebuilt(self):
        import sqlite3
        con = sqlite3.connect(self.db)
        con.executescript("""CREATE TABLE files (key TEXT PRIMARY KEY, filename TEXT, date_from TEXT, date_to TEXT, rows INTEGER, imported_at TEXT);
            INSERT INTO files VALUES ('mail:x','a.xlsx','2026-10-01','2026-10-25',1,'');
            CREATE TABLE meta (k TEXT PRIMARY KEY, v TEXT); INSERT INTO meta VALUES ('last_uid','99');""")
        con.close()
        self.assertFalse(mm.has_data(path=self.db))        # 檔案紀錄清掉，下一輪重抓
        con = sqlite3.connect(self.db)
        self.assertEqual(con.execute("SELECT COUNT(*) FROM meta").fetchone()[0], 0)
        con.close()

    def test_merge_with_upload(self):
        mm.import_file(make_xlsx(DAILY), DAILY_NAME, path=self.db)
        archived = mm.records_for("004", DAY, DAY, path=self.db)
        uploaded = mysetup.parse_personal_setup(make_xlsx([
            DAILY[0],                                                          # 重複（存檔已有）
            ("0999", 337277, "STUDIO A@士林文林", "iPhone", "蘇 悅", T(9)),   # 存檔沒有的
            DAILY[2],                                                          # 別家店的列
        ]))
        merged = mm.merge_records(uploaded, archived, "004")
        self.assertEqual(sorted(r["uid"] for r in merged), ["0999", "1001", "1002"])

    def test_fetch_mail_imports_attachment_once(self):
        msg = EmailMessage()
        msg["Subject"], msg["Message-ID"] = "Personal Setup Daily Report", "<a1@test>"
        msg.set_content("附件")
        msg.add_attachment(make_xlsx(DAILY), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=DAILY_NAME)
        other = EmailMessage()
        other["Subject"], other["Message-ID"] = "午餐", "<b2@test>"
        other.set_content("沒有附件")
        fake = FakeImap({11: msg.as_bytes(), 12: other.as_bytes()})
        orig = mm._connect
        mm._connect = lambda mc: fake
        try:
            r = mm.fetch_mail({}, log=lambda m: None, path=self.db)
            self.assertEqual((r["mails"], r["files"], r["new"], r["dates"]), (1, 1, 5, [DAY.isoformat()]))
            fake.searches.clear()
            r = mm.fetch_mail({}, log=lambda m: None, path=self.db)   # 第二次只看 UID 12 之後的新信
            self.assertEqual(r["files"], 0)
            self.assertEqual(fake.searches, [("UID", "13:*")])
            r = mm.fetch_mail({}, days=3, log=lambda m: None, path=self.db)   # 指定補抓：同一封不重複匯入
            self.assertEqual((r["mails"], r["files"]), (1, 0))
        finally:
            mm._connect = orig


class FakeImap:
    def __init__(self, msgs):
        self.msgs, self.searches = msgs, []

    def response(self, code):
        return code, [b"777"]

    def uid(self, cmd, *args):
        if cmd == "SEARCH":
            if args[1].startswith("UID "):   # 「UID n:*」沒有新信時，伺服器仍會回最後一封
                lo = int(args[1].split()[1].split(":")[0])
                self.searches.append(("UID", f"{lo}:*"))
                hits = [u for u in self.msgs if u >= lo] or [max(self.msgs)]
            else:
                self.searches.append(("SINCE",))
                hits = list(self.msgs)
            return "OK", [" ".join(map(str, hits)).encode()]
        if cmd == "FETCH":
            raw = self.msgs[int(args[0])]
            if "BODYSTRUCTURE" in args[1]:
                head = raw.split(b"\n\n", 1)[0]
                struct = b"(BODYSTRUCTURE (... \"xlsx\"))" if b"multipart" in head.lower() else b"(BODYSTRUCTURE (\"text\"))"
                return "OK", [(b"1 (BODY[HEADER] {n}", head + b"\r\n"), struct]
            return "OK", [(b"1 (BODY[] {n}", raw), b")"]
        raise AssertionError(cmd)

    def logout(self):
        pass


if __name__ == "__main__":
    unittest.main()
