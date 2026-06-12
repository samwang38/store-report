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

設定（擇一）：
  - 環境變數 EPB_WORKER_URL、EPB_INGEST_SECRET、EPB_SHOP_ID(預設004)、EPB_WINDOW_DAYS(預設90)
  - 或 local_config.json 內：
      "epb": { "workerUrl": "https://studioa-reservation.<帳號>.workers.dev",
               "ingestSecret": "<和 Worker 一致的 secret>",
               "shopId": "004", "windowDays": 90 }

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
SHOP_ID = str(cfg("shopId", "EPB_SHOP_ID", "004"))
WINDOW_DAYS = int(cfg("windowDays", "EPB_WINDOW_DAYS", 90))


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


def fetch_sold():
    """查 EPB 近 N 天已成交，回 [{v,s,d}]（同會員同品取最近成交日；扣銷退後淨量 > 0 才算）。"""
    sql = f"""
        select vip_id, stk_id, trans_type, stk_qty,
               to_char(doc_date, 'YYYY-MM-DD') doc_date
        from poslinev_bi
        where org_id = '01'
          and shop_id = '{SHOP_ID}'
          and doc_date >= trunc(sysdate) - {WINDOW_DAYS}
          and trans_type in ('A', 'H', 'E')
    """
    headers, rows = server.run_remote(sql)
    idx = {h.upper(): i for i, h in enumerate(headers)}
    iv, isk, it, iq, id_ = (idx["VIP_ID"], idx["STK_ID"], idx["TRANS_TYPE"],
                            idx["STK_QTY"], idx["DOC_DATE"])

    net = {}          # (vip, stk) -> 淨數量（A/H 正、E 銷退本身為負）
    last_day = {}     # (vip, stk) -> 最近成交日（只看 A/H）
    for r in rows:
        vip = (r[iv] or "").strip()
        stk = (r[isk] or "").strip()
        if not vip or vip == "0" or not stk:
            continue
        try:
            qty = float(r[iq] or 0)
        except ValueError:
            qty = 0.0
        key = (vip, stk)
        net[key] = net.get(key, 0.0) + qty
        if r[it] in ("A", "H"):
            d = r[id_] or ""
            if d and d > last_day.get(key, ""):
                last_day[key] = d

    return [{"v": v, "s": s, "d": last_day.get((v, s), "")}
            for (v, s), q in net.items() if q > 0]


def run_sync():
    sold = fetch_sold()
    snapshot = {"updatedAt": datetime.now(TPE).strftime("%Y-%m-%d %H:%M:%S"), "sold": sold}
    _post("/epb/ingest", snapshot, with_secret=True)
    LAST_SYNC_FILE.write_text(str(time.time()))
    print(f"[ok] 已同步 {len(sold)} 筆已成交（shop {SHOP_ID}, 近 {WINDOW_DAYS} 天）→ {snapshot['updatedAt']}")


def main():
    if not WORKER_URL or not SECRET:
        sys.exit("缺少設定：請設 EPB_WORKER_URL 與 EPB_INGEST_SECRET（或 local_config.json 的 epb.*）")

    force = "--force" in sys.argv
    manual = False if force else claim_sync_flag()
    if force or manual or due_by_schedule():
        why = "強制" if force else ("手動立即同步" if manual else "排程到期")
        print(f"[run] 觸發同步（{why}）")
        run_sync()
    else:
        print("[skip] 未到同步時點、也無手動請求")


if __name__ == "__main__":
    main()
