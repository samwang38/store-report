#!/usr/bin/env python3
"""
EPB → 預約工作台 銷售比對 同步腳本（Pattern B：店內 Mac 主動推送，零 inbound）

做什麼：
  1) 向 EPB 查近 N 天某門市的「已成交」銷售（trans_type A 銷售 / H 尾款，扣 E 銷退）。
  2) 整理成去識別化快照 {updatedAt, sold:[{v:會員碼, s:存貨碼, d:成交日}]}。
  3) POST 到 Cloudflare Worker 的 /epb/ingest（帶 secret）寫入 KV，供網頁讀取比對。

自我節流（搭配 launchd 每 60 秒呼叫）：
  - 距上次同步 ≥ 1 小時 → 同步
  - 或 Worker 端有「立即同步」旗標（網頁按鈕設的）→ 同步
  否則直接結束，不打 EPB。

多店：EPB 是中央 ERP，一台中央機器即可一次撈全部店、依「門市名稱」分別推送快照
      （門市名稱 = 兩系統的對應橋樑；前端用預約的 shopName 比對同店銷售）。

設定（擇一）：
  - 環境變數 EPB_WORKER_URL、EPB_INGEST_SECRET、EPB_SHOPS(逗號分隔 EPB shop_id)、EPB_WINDOW_DAYS(預設90)
  - 或 local_config.json 內：
      "epb": { "workerUrl": "https://studioa-reservation.<帳號>.workers.dev",
               "ingestSecret": "<和 Worker 一致的 secret>",
               "shops": ["004", "005", "010"],   // 多店；要全公司可留空/移除
               "windowDays": 90 }
    （相容舊設定：只設 "shopId":"004" 等同 "shops":["004"]）

需在公司 VPN/內網下執行（EPB 連線前提）。
"""
import json
import os
import sys
import time
import urllib.request
from datetime import datetime, timezone, timedelta
from pathlib import Path

import server  # 同層 gateway：run_remote / 引擎，import 不會啟動 web server

HERE = Path(__file__).resolve().parent
LAST_SYNC_FILE = HERE / ".epb_last_sync"
TPE = timezone(timedelta(hours=8))


def cfg(key, env, default=None):
    v = os.environ.get(env)
    if v:
        return v
    try:
        c = json.loads((HERE / "local_config.json").read_text("utf-8")).get("epb", {})
        if c.get(key) is not None:
            return c[key]
    except Exception:
        pass
    return default


WORKER_URL = (cfg("workerUrl", "EPB_WORKER_URL") or "").rstrip("/")
SECRET = cfg("ingestSecret", "EPB_INGEST_SECRET") or ""
WINDOW_DAYS = int(cfg("windowDays", "EPB_WINDOW_DAYS", 90))


def _resolve_shops():
    """要同步的 EPB shop_id 清單。
    優先 epb.shops（list 或逗號字串）；否則退回舊的單店 epb.shopId；都沒有 → None（全公司）。"""
    shops = cfg("shops", "EPB_SHOPS", None)
    if isinstance(shops, str):
        shops = [s.strip() for s in shops.split(",") if s.strip()]
    if not shops:
        single = cfg("shopId", "EPB_SHOP_ID", None)
        shops = [str(single).strip()] if single else None
    return [str(s).strip() for s in shops] if shops else None


SHOPS = _resolve_shops()


def _post(path, body, with_secret=False):
    req = urllib.request.Request(
        WORKER_URL + path,
        data=json.dumps(body).encode("utf-8"),
        method="POST",
        headers={"Content-Type": "application/json",
                 # 自訂 UA：Cloudflare 會以 error 1010 擋掉預設的 Python-urllib UA
                 "User-Agent": "studioa-epb-sync/1.0",
                 **({"Authorization": f"Bearer {SECRET}"} if with_secret else {})},
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8") or "{}")


def claim_sync_flag():
    """取走並清掉 Worker 端的手動同步旗標；回 True 代表使用者按了「立即同步」。"""
    try:
        return bool(_post("/epb/claim-sync", {}, with_secret=True).get("requested"))
    except Exception as e:
        print(f"[warn] claim-sync 失敗（忽略）：{e}", file=sys.stderr)
        return False


def due_by_schedule():
    try:
        last = float(LAST_SYNC_FILE.read_text().strip())
    except Exception:
        return True
    return (time.time() - last) >= 3600


def shop_id_to_name():
    """EPB shop_id → 門市名稱（與預約系統 shopName 對應的橋樑）。
    以 server.FALLBACK_STORES 硬編清單為底（可靠），pos_shop 查得到才覆蓋。
    注意：pos_shop 經 EPB gateway 常回 0 列，不能單靠它——否則名稱對照為空，
    run_sync 會退回拿「代碼」當店名推快照，導致前端（用預約 shopName 比對）永遠對不上。"""
    names = dict(getattr(server, "FALLBACK_STORES", {}))
    try:
        headers, rows = server.run_remote("select shop_id, name from pos_shop where org_id = '01'")
        idx = {h.upper(): i for i, h in enumerate(headers)}
        for r in rows:
            sid = str(r[idx["SHOP_ID"]]).strip()
            nm = (r[idx["NAME"]] or "").strip()
            if sid and nm:
                names[sid] = nm  # EPB 有回傳才以其為準
    except Exception as e:
        print(f"[warn] pos_shop 取名失敗，改用後備清單：{e}", file=sys.stderr)
    return names


def fetch_all_sold():
    """查 EPB 近 N 天已成交，依 shop_id 分組回 {shop_id: [{v,s,d}]}。
    同會員同品取最近成交日；扣銷退後淨量 > 0 才算。SHOPS=None 時撈全公司。"""
    shop_filter = ""
    if SHOPS:
        ids = ",".join("'" + s.replace("'", "") + "'" for s in SHOPS)
        shop_filter = f" and shop_id in ({ids})"
    sql = f"""
        select vip_id, stk_id, trans_type, stk_qty, shop_id,
               to_char(doc_date, 'YYYY-MM-DD') doc_date
        from poslinev_bi
        where org_id = '01'
          and doc_date >= trunc(sysdate) - {WINDOW_DAYS}
          and trans_type in ('A', 'H', 'E'){shop_filter}
    """
    headers, rows = server.run_remote(sql)
    idx = {h.upper(): i for i, h in enumerate(headers)}
    iv, isk, it, iq, ish, id_ = (idx["VIP_ID"], idx["STK_ID"], idx["TRANS_TYPE"],
                                 idx["STK_QTY"], idx["SHOP_ID"], idx["DOC_DATE"])

    per = {}  # shop_id -> {"net": {(v,s):qty}, "last": {(v,s):day}}
    for r in rows:
        vip = (r[iv] or "").strip()
        stk = (r[isk] or "").strip()
        if not vip or vip == "0" or not stk:
            continue
        shop = str(r[ish]).strip()
        try:
            qty = float(r[iq] or 0)
        except ValueError:
            qty = 0.0
        b = per.setdefault(shop, {"net": {}, "last": {}})
        key = (vip, stk)
        b["net"][key] = b["net"].get(key, 0.0) + qty
        if r[it] in ("A", "H"):
            d = r[id_] or ""
            if d and d > b["last"].get(key, ""):
                b["last"][key] = d

    out = {shop: [{"v": v, "s": s, "d": b["last"].get((v, s), "")}
                  for (v, s), q in b["net"].items() if q > 0]
           for shop, b in per.items()}
    # 設定清單裡的店即使無資料也回空（讓索引有同步時間）
    if SHOPS:
        for s in SHOPS:
            out.setdefault(s, [])
    return out


# ───────── 存貨代碼 → 型號 對照（供預約工作台「型號」欄） ─────────
# 前端把快取沒有的存貨代碼排進 Worker 待查清單；本腳本在排程同步時取走、
# 查 EPB 商品主檔 STKMAS 取型號（MODEL 欄，即 Apple 料號 如 MHRV4ZP/A）、回填 Worker KV。
# 已查過的不會再排入（前端快取命中）。表名/欄位已實測確認，可用設定覆蓋備用。
#   modelTable   EPB_MODEL_TABLE     商品主檔表名（預設 stkmas）
#   modelNameCol EPB_MODEL_NAME_COL  型號欄位名（預設 model；想存全名可改 name）
MODEL_TABLE = cfg("modelTable", "EPB_MODEL_TABLE", "stkmas")
MODEL_NAME_COL = cfg("modelNameCol", "EPB_MODEL_NAME_COL", "model")


def claim_pending_models():
    """取走 Worker 端待查型號的存貨代碼清單（需 secret）。"""
    try:
        return _post("/epb/models/claim-pending", {}, with_secret=True).get("stks") or []
    except Exception as e:
        print(f"[warn] claim-pending(models) 失敗（忽略）：{e}", file=sys.stderr)
        return []


def fetch_models(stks):
    """查 STKMAS 解出 stks 的型號 → {stk_id: 型號}。分批避開 Oracle IN 上限。"""
    found = {}
    for i in range(0, len(stks), 900):
        chunk = stks[i:i + 900]
        ids = ",".join("'" + s.replace("'", "") + "'" for s in chunk)
        sql = (f"select stk_id, {MODEL_NAME_COL} as model from {MODEL_TABLE} "
               f"where org_id = '01' and stk_id in ({ids})")
        headers, rows = server.run_remote(sql)
        idx = {h.upper(): j for j, h in enumerate(headers)}
        si, mi = idx["STK_ID"], idx["MODEL"]
        for r in rows:
            sid = str(r[si]).strip()
            mdl = (r[mi] or "").strip()
            if sid and mdl:
                found[sid] = mdl
    return found


def sync_models():
    """取待查存貨代碼 → 查型號 → 回填 Worker。回 (解出數, 待查數)。"""
    stks = claim_pending_models()
    norm = [str(s).strip() for s in stks if str(s).strip()]
    if not norm:
        return 0, 0
    found = fetch_models(norm)
    if found:
        _post("/epb/models/ingest", {"models": found}, with_secret=True)
    print(f"[ok] 型號回填 {len(found)}/{len(norm)} 筆")
    return len(found), len(norm)


def post_sync_log(entry):
    """回報一筆同步紀錄到 Worker（供系統狀態頁；無個資）；失敗忽略。"""
    try:
        _post("/epb/sync-log", entry, with_secret=True)
    except Exception as e:
        print(f"[warn] 回報 sync-log 失敗（忽略）：{e}", file=sys.stderr)


def run_sync():
    names = shop_id_to_name()
    by_shop = fetch_all_sold()
    ts = datetime.now(TPE).strftime("%Y-%m-%d %H:%M:%S")
    total = 0
    shops_summary = []
    for shop_id, sold in by_shop.items():
        name = names.get(shop_id) or shop_id  # 以門市名稱當比對橋樑
        _post("/epb/ingest", {"shop": name, "updatedAt": ts, "sold": sold}, with_secret=True)
        total += len(sold)
        shops_summary.append({"name": name, "count": len(sold)})
        print(f"  · {name}（{shop_id}）：{len(sold)} 筆")
    LAST_SYNC_FILE.write_text(str(time.time()))
    print(f"[ok] 已同步 {len(by_shop)} 店、共 {total} 筆已成交（近 {WINDOW_DAYS} 天）→ {ts}")

    # 同一趟順便回填預約工作台需要的型號（只查前端排入的待查存貨代碼）
    model_found, model_asked = 0, 0
    try:
        model_found, model_asked = sync_models()
    except Exception as e:
        print(f"[warn] 型號同步失敗（忽略，不影響銷售同步）：{e}", file=sys.stderr)

    post_sync_log({
        "at": ts, "ok": True, "shops": shops_summary, "sold": total,
        "models": {"found": model_found, "asked": model_asked},
    })


def main():
    if not WORKER_URL or not SECRET:
        sys.exit("缺少設定：請設 EPB_WORKER_URL 與 EPB_INGEST_SECRET（或 local_config.json 的 epb.*）")

    force = "--force" in sys.argv
    manual = False if force else claim_sync_flag()
    if force or manual or due_by_schedule():
        why = "強制" if force else ("手動立即同步" if manual else "排程到期")
        print(f"[run] 觸發同步（{why}）")
        try:
            run_sync()
        except Exception as e:
            ts = datetime.now(TPE).strftime("%Y-%m-%d %H:%M:%S")
            post_sync_log({"at": ts, "ok": False, "error": str(e)[:300]})
            raise
    else:
        print("[skip] 未到同步時點、也無手動請求")


if __name__ == "__main__":
    main()
