#!/usr/bin/env python3
"""
週報小幫手 Mac 端工作程式（launchd 每 60 秒呼叫一次，用 install_report_worker.sh 安裝）。

做什麼（Pattern B：Mac 主動拉取，外部不連進公司網路）：
  0) 自動更新：最多每 10 分鐘檢查 GitHub，有新版就 fast-forward 後以新版重新啟動；
     本機有未 commit 的修改或分歧時不更新，改寄信通知 Sam（每天最多一次）。
  1) 向 Worker 拉待寄信件（綁定驗證碼等），用公司 SMTP 寄出並回報。
  2) EPB 連得上才領產報工作：呼叫 server.build_report_excel 產出週報，
     以附件寄給店長並回報結果；失敗時 Worker 會退回次數並推播店長，同時寄信通知 Sam。

設定（local_config.json）：
  "smtp":         由 smtp_test.py 測試成功後存入（host/port/username/password/from）
  "reportWeekly": {"workerUrl": "...", "macSecret": "..."}

手動測試（不經 Worker，直接產一份寄出）：
  python3 report_worker.py --local 004 2026-10-10 收件人@studioa.com.tw [編制人數]
"""
import json
import os
import socket
import subprocess
import sys
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, timedelta
from email.message import EmailMessage
from pathlib import Path

from smtp_test import connect

HERE = Path(__file__).resolve().parent
MAX_JOBS_PER_RUN = 3
UPDATE_INTERVAL_SEC = 600
UPDATE_STAMP = HERE / ".last_update_check"
UPDATE_BLOCKED_STAMP = HERE / ".update_blocked_notified"


def log(msg):
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}", flush=True)


def load_cfg():
    cfg = json.loads((HERE / "local_config.json").read_text("utf-8"))
    if not cfg.get("smtp"):
        raise SystemExit("local_config.json 沒有 smtp 設定，請先執行 smtp_test.py 並選擇存檔")
    return cfg


# ─── Worker API ───────────────────────────────────────────
def worker(cfg, method, path, body=None):
    rw = cfg["reportWeekly"]
    req = urllib.request.Request(
        rw["workerUrl"].rstrip("/") + path, method=method,
        data=json.dumps(body).encode() if body is not None else None,
        headers={"authorization": f"Bearer {rw['macSecret']}", "content-type": "application/json",
                 "user-agent": "report-weekly-mac/1.0"})
    with urllib.request.urlopen(req, timeout=30) as res:
        return json.loads(res.read().decode())


# ─── 自動更新 ─────────────────────────────────────────────
def git(*args, timeout=60):
    r = subprocess.run(["git", *args], cwd=HERE, capture_output=True, text=True, timeout=timeout)
    if r.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)}：{(r.stderr or r.stdout).strip()[:300]}")
    return r.stdout.strip()


def app_version():
    try:
        v = (HERE / "VERSION").read_text().strip()
    except OSError:
        v = "?"
    try:
        return f"{v} ({git('rev-parse', '--short', 'HEAD', timeout=10)})"
    except Exception:
        return v


def _notify_update_blocked(cfg, reason):
    today = time.strftime("%Y-%m-%d")
    try:
        if UPDATE_BLOCKED_STAMP.read_text().strip() == today:
            return
    except OSError:
        pass
    UPDATE_BLOCKED_STAMP.write_text(today)
    notify_sam(cfg, "自動更新暫停", f"常開 Mac 上的週報工具沒有自動更新：{reason}\n\n"
               f"目前版本：{app_version()}。請到那台 Mac 的工具資料夾處理後執行 git pull。")


def maybe_update(cfg):
    """有新版就更新並以新版重新啟動（os.execv），沒有就直接返回。"""
    if not (HERE / ".git").exists():
        return
    try:
        if time.time() - UPDATE_STAMP.stat().st_mtime < UPDATE_INTERVAL_SEC:
            return
    except OSError:
        pass
    UPDATE_STAMP.write_text(time.strftime("%Y-%m-%d %H:%M:%S"))
    try:
        if git("status", "--porcelain", "--untracked-files=no"):
            _notify_update_blocked(cfg, "這台 Mac 上的程式有未 commit 的修改。")
            return
        git("fetch", "--quiet", "origin", "main")
        head, remote = git("rev-parse", "HEAD"), git("rev-parse", "origin/main")
        if head == remote:
            return
        try:
            git("merge-base", "--is-ancestor", "HEAD", "origin/main")
        except RuntimeError:
            _notify_update_blocked(cfg, "這台 Mac 的版本和 GitHub 分歧，無法直接更新。")
            return
        before = app_version()
        git("merge", "--ff-only", "--quiet", "origin/main")
        log(f"已自動更新：{before} → {app_version()}，以新版重新啟動")
    except Exception as e:
        log(f"自動更新檢查失敗（沿用目前版本）：{e}")
        return
    os.execv(sys.executable, [sys.executable, *sys.argv])


# ─── 寄信 ─────────────────────────────────────────────────
def send_mail(cfg, to, subject, body, attachment=None):
    s_cfg = cfg["smtp"]
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = s_cfg["from"]
    msg["To"] = to
    msg.set_content(body)
    if attachment:
        name, data = attachment
        msg.add_attachment(data, maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=name)
    s, _ = connect(s_cfg["host"], int(s_cfg["port"]), timeout=60)
    try:
        if s_cfg.get("username"):
            s.login(s_cfg["username"], s_cfg["password"])
        s.send_message(msg)
    finally:
        try:
            s.quit()
        except Exception:
            pass


def notify_sam(cfg, subject, body):
    try:
        send_mail(cfg, cfg["smtp"]["from"], f"【週報小幫手・系統通知】{subject}", body)
    except Exception as e:
        log(f"通知 Sam 失敗：{e}")


# ─── 產報 ─────────────────────────────────────────────────
def epb_reachable(timeout=3):
    import server  # 只為了取得 EPB 位址；import 不會啟動 web server
    u = urllib.parse.urlparse(server.EPB_WSDL_URL)
    try:
        with socket.create_connection((u.hostname, u.port or 80), timeout=timeout):
            return True
    except OSError:
        return False


def build_and_send(cfg, shop_id, shop_name, week_end, yoy_end, employee_count, to, manager_name="", traffic=None):
    import server
    import shoppertrak_traffic
    wk_end = date.fromisoformat(week_end)
    wk_start = wk_end - timedelta(days=6)
    payload = {"shopId": shop_id, "weekEnd": week_end, "weekStart": wk_start.isoformat(),
               "employeeCount": employee_count}
    if yoy_end:
        payload["yoyEnd"] = yoy_end
    t0 = time.time()
    t = traffic or {}
    if t.get("username"):
        log("  使用門市自己的人流帳號")
    with shoppertrak_traffic.credentials_override(t.get("username"), t.get("password")):
        filename, body, _meta = server.build_report_excel(payload, log=lambda m: log(f"  {m}"))
    span = f"{wk_start:%m/%d}–{wk_end:%m/%d}"
    send_mail(cfg, to, f"【週報】{shop_name} {span}",
              f"{manager_name or '店長'} 你好：\n\n附件是 {shop_name} {wk_start:%Y/%m/%d}（日）到 {wk_end:%Y/%m/%d}（六）的週報。\n"
              "數字有疑問，直接在 LINE 告訴 Sam，確認後修正重寄不扣次。\n\n週報小幫手",
              attachment=(filename, body))
    log(f"  已寄出 {filename}（{len(body):,} bytes，{time.time() - t0:.0f} 秒）")


# ─── 主流程 ───────────────────────────────────────────────
def process_outbox(cfg):
    items = worker(cfg, "GET", "/api/mac/outbox")["items"]
    for it in items:
        try:
            send_mail(cfg, it["to_email"], it["subject"], it["body"])
            worker(cfg, "POST", "/api/mac/outbox/result", {"id": it["id"], "ok": True})
            log(f"已寄出 {it['kind']} 信 #{it['id']}")
        except Exception as e:
            worker(cfg, "POST", "/api/mac/outbox/result", {"id": it["id"], "ok": False, "error": str(e)})
            log(f"寄信失敗 #{it['id']}：{e}")


def process_traffic_verify(cfg):
    """店長在 LINE 填的人流帳號：每輪試登入一組，結果回報給 Worker。"""
    item = worker(cfg, "GET", "/api/mac/traffic/verify/next")["item"]
    if not item:
        return
    import shoppertrak_traffic
    try:
        shoppertrak_traffic.verify_login(item["username"], item["password"])
        worker(cfg, "POST", "/api/mac/traffic/verify/result", {"shop_id": item["shop_id"], "ok": True})
        log(f"門市 {item['shop_id']} 人流帳號驗證成功")
    except Exception as e:
        worker(cfg, "POST", "/api/mac/traffic/verify/result", {"shop_id": item["shop_id"], "ok": False, "error": str(e)[:200]})
        log(f"門市 {item['shop_id']} 人流帳號驗證失敗：{e}")


def process_jobs(cfg):
    if not epb_reachable():
        log("EPB 連不上（VPN／內網？），本輪不領產報工作")
        return
    for _ in range(MAX_JOBS_PER_RUN):
        job = worker(cfg, "POST", "/api/mac/jobs/claim", {})["job"]
        if not job:
            return
        log(f"領到工作 #{job['id']}：{job['shop_id']} {job['shop_name']} 週結束 {job['week_end']}")
        try:
            build_and_send(cfg, job["shop_id"], job["shop_name"], job["week_end"], job.get("yoy_end"),
                           job.get("employee_count"), job["to_email"], job.get("manager_name", ""), job.get("traffic"))
            worker(cfg, "POST", "/api/mac/jobs/result", {"id": job["id"], "ok": True, "version": app_version()})
        except Exception as e:
            tb = traceback.format_exc()
            log(f"工作 #{job['id']} 失敗：{e}")
            worker(cfg, "POST", "/api/mac/jobs/result",
                   {"id": job["id"], "ok": False, "error": f"{type(e).__name__}: {e}", "version": app_version()})
            notify_sam(cfg, f"產報失敗 {job['shop_name']} {job['week_end']}",
                       f"工作 #{job['id']}，收件人 {job['to_email']}，版本 {app_version()}。"
                       f"已自動退回次數並通知店長。\n\n{tb}")


def main():
    cfg = load_cfg()
    if len(sys.argv) >= 5 and sys.argv[1] == "--local":
        shop_id, week_end, to = sys.argv[2], sys.argv[3], sys.argv[4]
        emp = int(sys.argv[5]) if len(sys.argv) > 5 else None
        build_and_send(cfg, shop_id, shop_id, week_end, None, emp, to)
        return
    maybe_update(cfg)
    try:
        process_outbox(cfg)
        process_traffic_verify(cfg)
        process_jobs(cfg)
    except urllib.error.URLError as e:
        log(f"Worker 連線失敗：{e}")


if __name__ == "__main__":
    main()
