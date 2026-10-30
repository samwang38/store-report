#!/usr/bin/env python3
"""
MySetup 每日信件匯入：Sam 的公司信箱每天收到前一天「全公司」的 Personal Setup 匯出檔
（附檔「Personal Setup_YYYY-MM-DD_YYYY-MM-DD.xlsx」），只有這個信箱收得到，所有門市共用。
常開 Mac（MBP-Hub）的 report_worker 用 IMAP 抓附檔，依每列的 Apple ID（門市在 Apple 的代號）
分店存進本機 SQLite；產週報時與店長上傳的檔案以 Unique ID 合併去重（server.build_report_excel）。
存檔只留統計要用的欄位，不存客人姓名／電話／信箱。

設定（local_config.json，帳密只放在這台 Mac）：
  "mysetupMail": {"username": "...@studioa.com.tw", "password": "...",
                  "host": "studioa.com.tw", "port": 143, "folder": "INBOX", "subject": ""}
  subject 留空時檢查每封新信的附檔名；有填就只看主旨含這段文字的信。
  沒有 mysetupMail 時沿用 "smtp" 的帳密（寄件帳號就是 Sam 的公司信箱，同一台郵件主機收信）。

手動：
  python mysetup_mail.py --file <xlsx> [...]   離線匯入檔案（測試、補資料）
  python mysetup_mail.py --once [--days N]     抓信一次（第一次會補抓近 90 天）
  python mysetup_mail.py --status              存檔涵蓋日期、缺的日子、對不上的 Apple ID
"""
import argparse
import email
import imaplib
import json
import re
import sqlite3
import ssl
import sys
from datetime import date, datetime, timedelta
from email import policy
from pathlib import Path

import mysetup

HERE = Path(__file__).resolve().parent
ARCHIVE_PATH = HERE / "mysetup_archive.sqlite"
CONFIG_PATH = HERE / "local_config.json"
KEEP_DAYS = 400              # 存檔保留天數（月報、年對年都夠用）
FIRST_BACKFILL_DAYS = 90     # 第一次抓信補抓的天數
FILE_RE = re.compile(r"personal\s*setup", re.I)
COMPANY_MIN_STORES = 5       # 檔裡出現幾家以上門市才算「全公司」每日檔；少於這個數是單店自己匯出的檔，只算那幾家有資料
RANGE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})")

# Apple ID → ERP 門市代碼（依 2026-10-29 全公司檔的 POS Name 對照）。
# 出現表上沒有的 Apple ID 時，status() 會列出來、report_worker 會通知 Sam 補上。
APPLE_ID_TO_SHOP = {
    "332944": "003",   # STUDIO A＠高雄大立
    "337277": "004",   # STUDIO A@士林文林
    "1293891": "005",  # STUDIO A@台北微風
    "380313": "012",   # STUDIO A＠台南大億
    "411723": "024",   # STUDIO A@台北美麗華
    "444912": "027",   # STUDIO A＠高雄夢時代
    "552148": "037",   # STUDIO A＠新竹光復
    "572037": "043",   # STUDIO A＠中壢大江
    "609084": "044",   # STUDIO A@高雄三多
    "624351": "046",   # STUDIO A@台北阿波羅
    "624352": "047",   # STUDIO A@台中金典
    "659937": "048",   # STUDIO A@新竹中正
    "321748": "050",   # STUDIO A@台北西門
    "819343": "053",   # STUDIO A@桃園台茂
    "851062": "054",   # STUDIO A@台北大葉高島屋
    "1139091": "058",  # STUDIO A@苗栗尚順
    "2041527": "061",  # STUDIO A@桃園統領
    "2247998": "062",  # STUDIO A@豐原太平洋
    "2407495": "063",  # STUDIO A@板橋遠百中山
    "2901726": "064",  # STUDIOA@新莊宏匯
    "3388951": "065",  # STUDIO A@竹北遠百
    "3522004": "066",  # STUDIOA@台中麗寶
    "3607000": "067",  # STUDIO A@SHOWTIME LIVE GANGSHAN STORE（岡山）
    "3842054": "068",  # STUDIOA@新店裕隆城
}

# 不納入的門市（Sam 2026-10-30 決定）：不存資料、不列為對不上的門市
EXCLUDED_APPLE_IDS = {
    "544641",   # STUDIO A@台中彩虹3C店
}


def apple_ids_of(shop_id):
    shop_id = str(shop_id or "").strip()
    return [aid for aid, sid in APPLE_ID_TO_SHOP.items() if sid == shop_id]


def shop_of(apple_id):
    return APPLE_ID_TO_SHOP.get(str(apple_id or "").strip())


# ─── 存檔（SQLite）────────────────────────────────────────────────────────
def _db(path=None):
    con = sqlite3.connect(str(path or ARCHIVE_PATH), timeout=30)
    con.executescript("""
        CREATE TABLE IF NOT EXISTS records (
          uid TEXT PRIMARY KEY, apple_id TEXT, pos TEXT, name TEXT NOT NULL, device TEXT, submitted TEXT NOT NULL);
        CREATE INDEX IF NOT EXISTS records_shop ON records (apple_id, submitted);
        CREATE TABLE IF NOT EXISTS files (          -- 匯入過的檔案；key 是「信的 Message-ID＋檔名」或離線匯入的檔名
          key TEXT PRIMARY KEY, filename TEXT, date_from TEXT, date_to TEXT, rows INTEGER, imported_at TEXT,
          scope TEXT);                              -- 涵蓋的門市：'*'＝全公司，否則是逗號分隔的 Apple ID
        CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT);
    """)
    if "scope" not in {r[1] for r in con.execute("PRAGMA table_info(files)")}:
        # 1.19.1 以前沒分全公司／單店檔：清掉檔案紀錄與抓信進度，下一輪重抓近 90 天重新判斷（筆數以 Unique ID 覆寫，不會重複）
        con.executescript("ALTER TABLE files ADD COLUMN scope TEXT; DELETE FROM files; DELETE FROM meta;")
        con.commit()
    return con


def file_range(filename, records=()):
    """檔名的日期範圍（Personal Setup_起_迄.xlsx）；檔名看不出來就用資料的最早／最晚提交日。"""
    m = RANGE_RE.search(filename or "")
    if m:
        return date.fromisoformat(m.group(1)), date.fromisoformat(m.group(2))
    days = [r["submitted"].date() for r in records]
    return (min(days), max(days)) if days else (None, None)


def import_file(data, filename, key=None, path=None):
    """匯入一份 Personal Setup 檔：回傳 {rows, new, from, to}。同一筆（Unique ID）重複匯入只會覆寫。"""
    records = mysetup.parse_personal_setup(data)
    d_from, d_to = file_range(filename, records)
    con = _db(path)
    try:
        if EXCLUDED_APPLE_IDS:
            con.execute(f"DELETE FROM records WHERE apple_id IN ({','.join('?' * len(EXCLUDED_APPLE_IDS))})",
                        tuple(EXCLUDED_APPLE_IDS))
        before = con.execute("SELECT COUNT(*) FROM records").fetchone()[0]
        for r in records:
            if r.get("apple_id") in EXCLUDED_APPLE_IDS:
                continue
            uid = r["uid"] or f"{r['apple_id']}|{r['name']}|{r['submitted'].isoformat()}"
            con.execute("INSERT OR REPLACE INTO records (uid,apple_id,pos,name,device,submitted) VALUES (?,?,?,?,?,?)",
                        (uid, r.get("apple_id"), r.get("pos"), r["name"], str(r.get("device") or ""),
                         r["submitted"].isoformat(sep=" ")))
        after = con.execute("SELECT COUNT(*) FROM records").fetchone()[0]
        aids = sorted({r["apple_id"] for r in records if r.get("apple_id")})
        scope = "*" if len(aids) >= COMPANY_MIN_STORES else ",".join(aids)
        con.execute("INSERT OR REPLACE INTO files (key,filename,date_from,date_to,rows,imported_at,scope) VALUES (?,?,?,?,?,?,?)",
                    (key or f"file:{filename}", filename, d_from and d_from.isoformat(), d_to and d_to.isoformat(),
                     len(records), datetime.now().isoformat(timespec="seconds"), scope))
        cutoff = (date.today() - timedelta(days=KEEP_DAYS)).isoformat()
        con.execute("DELETE FROM records WHERE submitted < ?", (cutoff,))
        con.execute("DELETE FROM files WHERE date_to < ?", (cutoff,))
        con.commit()
    finally:
        con.close()
    return {"rows": len(records), "new": after - before, "from": d_from, "to": d_to, "company": scope == "*"}


def has_data(path=None):
    if not Path(path or ARCHIVE_PATH).exists():
        return False
    con = _db(path)
    try:
        return con.execute("SELECT 1 FROM files LIMIT 1").fetchone() is not None
    finally:
        con.close()


def _covered(con, shop_id=None):
    """有資料的日子：全公司檔涵蓋所有門市；單店檔只算給檔裡那幾家（shop_id 沒給時只看全公司檔）。"""
    aids = set(apple_ids_of(shop_id)) if shop_id else set()
    days = set()
    for d_from, d_to, scope in con.execute("SELECT date_from, date_to, scope FROM files WHERE date_from IS NOT NULL"):
        if scope != "*" and not (aids & set((scope or "").split(","))):
            continue
        d, end = date.fromisoformat(d_from), date.fromisoformat(d_to)
        while d <= end:
            days.add(d)
            d += timedelta(days=1)
    return days


def covered_days(start, end, path=None, shop_id=None):
    """start~end 之間存檔有涵蓋（這家店）的日子（有收到信的日子，當天 0 筆也算）。"""
    if not has_data(path):
        return []
    con = _db(path)
    try:
        days = _covered(con, shop_id)
    finally:
        con.close()
    return sorted(d for d in days if start <= d <= end)


def missing_days(start, end, path=None, also_covered=None, shop_id=None):
    """start~end 之間（到昨天為止）存檔沒有涵蓋（這家店）的日子；also_covered=(起, 迄) 是店長上傳檔的日期範圍，
    那幾天不算缺。存檔是空的回傳 None。"""
    if not has_data(path):
        return None
    con = _db(path)
    try:
        days = _covered(con, shop_id)
    finally:
        con.close()
    if not days:
        return None
    if also_covered and also_covered[0] and also_covered[1]:
        d = also_covered[0]
        while d <= also_covered[1]:
            days.add(d)
            d += timedelta(days=1)
    out, d = [], start
    end = min(end, date.today() - timedelta(days=1))
    while d <= end:
        if d not in days:
            out.append(d)
        d += timedelta(days=1)
    return out


def records_for(shop_id, start, end, path=None):
    """本店 start~end（含迄日）的提交紀錄，格式同 mysetup.parse_personal_setup；存檔是空的回傳 []。"""
    aids = apple_ids_of(shop_id)
    if not aids or not Path(path or ARCHIVE_PATH).exists():
        return []
    con = _db(path)
    try:
        rows = con.execute(
            f"SELECT uid, name, device, submitted, apple_id, pos FROM records WHERE apple_id IN ({','.join('?' * len(aids))}) "
            "AND submitted >= ? AND submitted < ? ORDER BY submitted",
            (*aids, start.isoformat(), (end + timedelta(days=1)).isoformat())).fetchall()
    finally:
        con.close()
    return [{"uid": uid, "name": name, "device": device, "submitted": datetime.fromisoformat(sub),
             "apple_id": aid, "pos": pos} for uid, name, device, sub, aid, pos in rows]


def merge_records(uploaded, archived, shop_id):
    """店長上傳的檔＋每日存檔，以 Unique ID 去重。上傳檔裡屬於別家店（Apple ID 對到其他門市）的列不算。"""
    out, seen = [], set()
    for r in list(archived) + list(uploaded or []):
        owner = shop_of(r.get("apple_id"))
        if owner and owner != str(shop_id):
            continue
        if r.get("uid"):
            if r["uid"] in seen:
                continue
            seen.add(r["uid"])
        out.append(r)
    return out


def status(path=None):
    """全公司每日檔的 {from, through, missing[], unknown[{apple_id, pos, rows}], files}，給 report_worker 回報 Worker 與通知用。"""
    empty = {"from": None, "through": None, "missing": [], "unknown": [], "files": 0}
    if not has_data(path):
        return empty
    con = _db(path)
    try:
        days = _covered(con)
        nfiles = con.execute("SELECT COUNT(*) FROM files").fetchone()[0]
        unknown = [{"apple_id": aid, "pos": pos, "rows": n} for aid, pos, n in con.execute(
            "SELECT apple_id, MAX(pos), COUNT(*) FROM records GROUP BY apple_id ORDER BY apple_id")
            if aid and aid not in APPLE_ID_TO_SHOP and aid not in EXCLUDED_APPLE_IDS]
    finally:
        con.close()
    if not days:   # 只有單店匯出檔、還沒有全公司每日檔
        return {**empty, "unknown": unknown, "files": nfiles}
    first, last = min(days), max(days)
    missing, d = [], first
    while d < date.today():
        if d not in days:
            missing.append(d.isoformat())
        d += timedelta(days=1)
    return {"from": first.isoformat(), "through": last.isoformat(), "missing": missing[-31:],
            "unknown": unknown, "files": nfiles}


# ─── IMAP 抓信（寫法同自動日報 daily_job.imap_connect）──────────────────────
def _connect(mc):
    ctx = ssl.create_default_context()
    ctx.set_ciphers("DEFAULT:@SECLEVEL=1")  # 公司郵件主機用 1024 位元 DH 金鑰（同 smtp_test.connect）
    host, port = mc.get("host") or "studioa.com.tw", int(mc.get("port") or 143)
    if port == 993:
        m = imaplib.IMAP4_SSL(host, port, ssl_context=ctx, timeout=60)
    else:  # 公司主機沒開 993，走 143＋STARTTLS（不接受明文登入）
        m = imaplib.IMAP4(host, port, timeout=60)
        m.starttls(ssl_context=ctx)
    m.login(mc["username"], mc["password"])
    typ, data = m.select(mc.get("folder") or "INBOX", readonly=True)
    if typ != "OK":
        raise RuntimeError(f"IMAP 開啟信件匣失敗：{data}")
    return m


def _flatten(data):
    out = b""
    for x in data or []:
        if isinstance(x, tuple):
            out += b"".join(p for p in x if isinstance(p, bytes))
        elif isinstance(x, bytes):
            out += x
    return out


def _header(data, name):
    for x in data or []:
        if isinstance(x, tuple) and len(x) > 1:
            msg = email.message_from_bytes(x[1], policy=policy.default)
            return str(msg[name] or "").strip()
    return ""


def fetch_mail(mc, days=None, log=print, path=None):
    """抓新信裡的 Personal Setup 附檔並匯入：回傳 {mails, files, new, dates[匯入檔涵蓋的日子]}。
    記住信件匣的 UIDVALIDITY 與看過的最大 UID，之後每次只看新信；第一次（或 days 有指定）用 SINCE 日期補抓。"""
    con = _db(path)
    try:
        meta = dict(con.execute("SELECT k, v FROM meta").fetchall())
        done = {k for (k,) in con.execute("SELECT key FROM files WHERE key LIKE 'mail:%'")}
    finally:
        con.close()
    m = _connect(mc)
    summary = {"mails": 0, "files": 0, "new": 0, "dates": []}
    dates = set()
    try:
        validity = (m.response("UIDVALIDITY")[1] or [b""])[0]
        validity = (validity or b"").decode() if isinstance(validity, bytes) else str(validity or "")
        last_uid = int(meta.get("last_uid") or 0) if meta.get("uidvalidity") == validity else 0
        if days is None and last_uid:
            typ, data = m.uid("SEARCH", None, f"UID {last_uid + 1}:*")
        else:
            since = date.today() - timedelta(days=days or FIRST_BACKFILL_DAYS)
            typ, data = m.uid("SEARCH", None, "SINCE", since.strftime("%d-%b-%Y"))
        if typ != "OK":
            raise RuntimeError(f"IMAP 搜尋失敗：{data}")
        uids = sorted(int(u) for u in (data[0] or b"").split() if int(u) > last_uid or days is not None)
        want_subject = str(mc.get("subject") or "").strip().lower()
        for uid in uids:
            typ, head = m.uid("FETCH", str(uid), "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID SUBJECT)] BODYSTRUCTURE)")
            if typ != "OK":
                continue
            raw = _flatten(head).lower()
            if b"xlsx" not in raw and b"spreadsheetml" not in raw:
                continue
            subject = _header(head, "Subject")
            if want_subject and want_subject not in subject.lower():
                continue
            msg_id = _header(head, "Message-ID") or f"uid-{validity}-{uid}"
            typ, body = m.uid("FETCH", str(uid), "(BODY.PEEK[])")
            if typ != "OK":
                continue
            msg = email.message_from_bytes(_first_literal(body), policy=policy.default)
            hit = False
            for part in msg.iter_attachments():
                fn = Path(part.get_filename() or "").name
                if not (fn.lower().endswith(".xlsx") and FILE_RE.search(fn)):
                    continue
                hit = True
                key = f"mail:{msg_id}:{fn}"
                if key in done:
                    continue
                try:
                    r = import_file(part.get_payload(decode=True), fn, key=key, path=path)
                except Exception as e:  # 單一附檔壞掉不擋其他信
                    log(f"  ⚠ MySetup 附檔 {fn} 無法讀取：{e}")
                    continue
                summary["files"] += 1
                summary["new"] += r["new"]
                d = r["from"]
                while d and r["to"] and d <= r["to"]:
                    dates.add(d.isoformat())
                    d += timedelta(days=1)
                log(f"  匯入 {fn}（{subject[:40]}）：{r['rows']} 筆，新增 {r['new']} 筆")
            summary["mails"] += hit
        if uids:
            con = _db(path)
            try:
                con.executemany("INSERT OR REPLACE INTO meta (k, v) VALUES (?, ?)",
                                [("uidvalidity", validity), ("last_uid", str(max(uids[-1], last_uid)))])
                con.commit()
            finally:
                con.close()
    finally:
        try:
            m.logout()
        except Exception:
            pass
    summary["dates"] = sorted(dates)
    return summary


def _first_literal(data):
    for x in data or []:
        if isinstance(x, tuple) and len(x) > 1:
            return x[1]
    return b""


def mail_config(cfg):
    """收信設定：mysetupMail 有帳密就用它，沒有就沿用 smtp 的帳密與主機（IMAP 走 143）。"""
    mc = cfg.get("mysetupMail") or {}
    if mc.get("username") and mc.get("password"):
        return mc
    sm = cfg.get("smtp") or {}
    if sm.get("username") and sm.get("password"):
        return {**mc, "username": sm["username"], "password": sm["password"], "host": mc.get("host") or sm.get("host")}
    return None


def load_mail_config():
    try:
        return mail_config(json.loads(CONFIG_PATH.read_text("utf-8")))
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser(description="MySetup 每日信件匯入")
    ap.add_argument("--file", nargs="+", help="離線匯入 Personal Setup 檔")
    ap.add_argument("--once", action="store_true", help="抓信一次")
    ap.add_argument("--days", type=int, help="抓信時補抓近 N 天（預設只看新信，第一次 90 天）")
    ap.add_argument("--status", action="store_true", help="顯示存檔狀態")
    a = ap.parse_args()
    if a.file:
        for f in a.file:
            p = Path(f).expanduser()
            r = import_file(p.read_bytes(), p.name)
            print(f"{p.name}：{r['rows']} 筆，新增 {r['new']} 筆（{r['from']}～{r['to']}）")
    if a.once:
        mc = load_mail_config()
        if not mc:
            sys.exit("local_config.json 沒有 mysetupMail 或 smtp 帳密")
        print(fetch_mail(mc, days=a.days))
    if a.status or not (a.file or a.once):
        print(json.dumps(status(), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
