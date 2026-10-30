"""
MySetup 每人、每日、每種裝置的提交數（日報信進度表、排班系統「我的業績」共用）。

資料來自每日信件存檔（mysetup_mail）；名字對照同週報：店長在 LINE 設的規則優先，再比 EP_EMP 中文名／英文名。
候選員工＝各店近 90 天有交易的員工（ERP 一次查所有門市，SA999 總公司沖帳除外）。
只回存檔有該店資料的日子：當天沒人提交也會出現（空的 {}），沒收到信的日子不出現。
"""
from datetime import date, timedelta

import mysetup
import mysetup_mail

DEVICES = mysetup.MYSETUP_DEVICES   # ("Mac", "iPad", "iPhone", "Watch")


def store_candidates(shop_ids, days=90):
    """{門市: {員工代碼: (中文名, 英文名)}}：近 days 天在該店有交易的員工。"""
    import server
    shop_ids = sorted({str(s) for s in shop_ids if s})
    if not shop_ids:
        return {}
    since = (date.today() - timedelta(days=days)).isoformat()
    _, rows = server.run_remote(
        "select distinct shop_id, emp_id1 from poslinev_bi where org_id = '01' and shop_id in ("
        + ",".join(server.quote_sql(s) for s in shop_ids)
        + f") and doc_date >= to_date('{since}', 'yyyy-mm-dd')", timeout=180)
    by_shop = {}
    for row in rows:
        shop, code = server.text(row[0]), server.text(row[1])
        if shop and code and code != server.engine.VOUCHER_SETTLE_EMP:
            by_shop.setdefault(shop, set()).add(code)
    names = server.epb_emp_names({c for codes in by_shop.values() for c in codes})
    return {shop: {c: names.get(c, ("", "")) for c in codes} for shop, codes in by_shop.items()}


def _blank():
    return {d: 0 for d in DEVICES}


def daily_counts(shop_ids, start, end, rules_by_shop, candidates=None):
    """回傳 {"counts": {門市: {日期: {員工代碼: {Mac, iPad, iPhone, Watch}}}},
            "unmatched": {門市: {日期: {名字: {Mac, iPad, iPhone, Watch}}}}}，日期是 ISO 字串。
    rules_by_shop：{門市: [{"name", "code"}]}（Worker /api/mac/mysetup/rules）；code 為 SKIP 的名字不計入。"""
    shop_ids = [str(s) for s in shop_ids if mysetup_mail.apple_ids_of(s)]
    cands = candidates if candidates is not None else store_candidates(shop_ids)
    counts, unmatched = {}, {}
    for shop in shop_ids:
        covered = {d.isoformat() for d in mysetup_mail.covered_days(start, end, shop_id=shop)}
        if not covered:
            continue
        recs = mysetup_mail.records_for(shop, start, end)
        aliases = {mysetup.normalize(r.get("name")): str(r.get("code") or "").strip()
                   for r in rules_by_shop.get(shop) or [] if mysetup.normalize(r.get("name")) and r.get("code")}
        mapping = mysetup.match_names([r["name"] for r in recs], cands.get(shop, {}), aliases)
        per_day = {d: {} for d in sorted(covered)}
        miss = {}
        for rec in recs:
            day = rec["submitted"].date().isoformat()
            dev = mysetup.device_key(rec.get("device"))
            if day not in covered or not dev:
                continue
            code, _ = mapping.get(mysetup.normalize(rec["name"]), (None, ""))
            if code == mysetup.SKIP:
                continue
            if code:
                slot = per_day[day].setdefault(code, _blank())
            else:   # 同一正規化名的不同打法合併到第一次出現的寫法（同 mysetup.summarize）
                names = miss.setdefault(day, {})
                label = next((k for k in names if mysetup.normalize(k) == mysetup.normalize(rec["name"])), rec["name"])
                slot = names.setdefault(label, _blank())
            slot[dev] += 1
        counts[shop] = per_day
        if miss:
            unmatched[shop] = miss
    return {"counts": counts, "unmatched": unmatched}


def period_totals(per_day, start, end):
    """把 daily_counts 某店的 {日期: {鍵: {裝置: n}}} 加總成 start~end 的 {鍵: {裝置: n}}。"""
    out = {}
    for day, people in per_day.items():
        if start.isoformat() <= day <= end.isoformat():
            for key, devs in people.items():
                slot = out.setdefault(key, _blank())
                for d in DEVICES:
                    slot[d] += devs.get(d, 0)
    return out
