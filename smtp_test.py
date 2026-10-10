#!/usr/bin/env python3
"""
公司信箱 SMTP 寄信測試（給週報小幫手寄驗證碼與週報用）。

用法（在終端機執行，密碼不會顯示也不會印出）：
    python3 smtp_test.py

流程：
  1) 詢問 SMTP 主機／port／帳號／密碼／寄件人（預設值可直接按 Enter）。
  2) 依 port 選連線方式：465 用 SSL，其餘用 STARTTLS（若伺服器不支援則用明文連線並提示）。
  3) 寄一封含 xlsx 小附件的測試信給自己。
  4) 成功後詢問是否存進 local_config.json 的 "smtp"（存檔前自動備份）。
"""
import getpass
import io
import json
import shutil
import smtplib
import ssl
import sys
import time
from email.message import EmailMessage
from pathlib import Path

HERE = Path(__file__).resolve().parent
CONFIG = HERE / "local_config.json"
DEFAULT_HOST = "studioa.com.tw"   # 公司 MX 即此主機；465/587 皆有開


def ask(prompt, default=""):
    v = input(f"{prompt}" + (f"（預設 {default}）" if default else "") + "：").strip()
    return v or default


def tiny_xlsx():
    """產生一個很小的 xlsx 當附件，測試公司郵件系統會不會擋 Excel 附件。"""
    try:
        import openpyxl
    except ImportError:
        return None
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "測試"
    ws.append(["項目", "數值"])
    ws.append(["測試時間", time.strftime("%Y-%m-%d %H:%M:%S")])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def connect(host, port, timeout=20):
    ctx = ssl.create_default_context()
    # 公司郵件主機用 1024 位元 DH 金鑰，OpenSSL 預設等級（2048）會拒絕；降到等級 1，憑證照樣驗證
    ctx.set_ciphers("DEFAULT:@SECLEVEL=1")
    if port == 465:
        return smtplib.SMTP_SSL(host, port, timeout=timeout, context=ctx), "SSL"
    s = smtplib.SMTP(host, port, timeout=timeout)
    s.ehlo()
    if s.has_extn("starttls"):
        try:
            s.starttls(context=ctx)
        except ssl.SSLCertVerificationError:
            # 公司自簽憑證常見：改用不驗證憑證的連線，並提示
            print("⚠ 伺服器憑證無法驗證（可能是公司自簽憑證），改用不驗證憑證的加密連線。")
            s.close()
            s = smtplib.SMTP(host, port, timeout=timeout)
            s.ehlo()
            s.starttls(context=ssl._create_unverified_context())
            s.ehlo()
            return s, "STARTTLS（未驗證憑證）"
        s.ehlo()
        return s, "STARTTLS"
    print("⚠ 伺服器不支援加密（STARTTLS），密碼會以明文傳送，建議改用 465。")
    return s, "明文"


def main():
    print("== 公司信箱 SMTP 寄信測試 ==")
    cfg = {}
    try:
        cfg = json.loads(CONFIG.read_text("utf-8"))
    except Exception:
        pass
    old = cfg.get("smtp", {})

    host = ask("SMTP 主機", old.get("host", DEFAULT_HOST))
    port_in = ask("Port（465、587，不確定就按 Enter 自動測試）", str(old.get("port", "自動")))
    ports = [465, 587] if port_in in ("自動", "auto", "") else [int(port_in)]
    user = ask("登入帳號（通常是完整公司信箱）", old.get("username", ""))
    pw = getpass.getpass("密碼（輸入時不會顯示）：") or old.get("password", "")
    sender = ask("寄件人信箱", old.get("from", user))
    to = ask("測試收件人", sender)

    msg = EmailMessage()
    msg["Subject"] = "【週報小幫手】寄信測試"
    msg["From"] = sender
    msg["To"] = to
    msg.set_content("這是一封寄信測試。收得到、附件也打得開，代表週報小幫手可以用這個信箱寄出驗證碼和週報。")
    att = tiny_xlsx()
    if att:
        msg.add_attachment(att, maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename="寄信測試.xlsx")

    port = None
    for p in ports:
        try:
            s, mode = connect(host, p)
            print(f"已連線 {host}:{p}（{mode}）")
            if user:
                s.login(user, pw)
                print("登入成功")
            s.send_message(msg)
            s.quit()
            port = p
            break
        except smtplib.SMTPAuthenticationError as e:
            print(f"✗ 登入失敗：帳號或密碼錯誤，或公司不允許此帳號用 SMTP 登入。（{e.smtp_code}）")
            sys.exit(1)
        except Exception as e:
            print(f"✗ port {p} 寄信失敗：{type(e).__name__}: {e}")
    if port is None:
        print("  兩個 port 都失敗。請確認目前有連公司 VPN／內網，或把錯誤訊息給 Claude 看。")
        sys.exit(1)

    print(f"✓ 已寄出測試信到 {to}，請到信箱確認是否收到（也看看垃圾郵件匣），附件能否打開。")

    if ask("要把這組設定存進 local_config.json 嗎？(y/N)", "N").lower() == "y":
        if CONFIG.exists():
            bak = CONFIG.with_name(f"local_config.backup-{time.strftime('%Y%m%d-%H%M%S')}.json")
            shutil.copy2(CONFIG, bak)
            print(f"已備份原設定到 {bak.name}")
        cfg["smtp"] = {"host": host, "port": port, "username": user,
                       "password": pw, "from": sender}
        CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        print("✓ 已存檔（local_config.json 不會進 git）")


if __name__ == "__main__":
    main()
