#!/usr/bin/env python3
"""DSS 決策支援系統（dss.studioa.com.tw）— 登入與資料查詢

登入流程無法全自動（6 碼圖形驗證碼＋不定時 Email OTP），改採「前端互動式登入」：
  1. start_login()：開新 session，抓登入頁 _csrf 與驗證碼圖，前端顯示給使用者輸入。
  2. submit_captcha(code)：送出帳密＋驗證碼。依回應判定：成功 / 驗證碼錯 / 需 Email OTP。
  3. （視情況）submit_otp(code)：送出信箱收到的驗證碼。
登入成功後 session cookie 存入 local_config.json，server 重啟時嘗試還原（best-effort），
讓重新登入的頻率降到最低。

帳密與 ShopperTrak 同模式存在 local_config.json（"dss" 鍵）；讀寫沿用
shoppertrak_traffic 的 load_config / save_config，共用同一把 _CFG_LOCK 避免競寫。
"""
import html
import re
import threading
import time
from urllib.parse import urljoin

import shoppertrak_traffic as _cfgmod

BASE_URL = "https://dss.studioa.com.tw:8443"
LOGIN_PATH = "/mailCap"
CAPTCHA_PATH = "/cpa/getCaptchaImg"

_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

# ─── 帳密（local_config.json "dss" 鍵）─────────────────────────────────


def get_credentials():
    d = _cfgmod.load_config().get("dss") or {}
    return d.get("username"), d.get("password")


def has_credentials():
    u, p = get_credentials()
    return bool(u and p)


def set_credentials(username, password):
    cfg = _cfgmod.load_config()
    cfg.setdefault("dss", {})
    cfg["dss"]["username"] = (username or "").strip()
    cfg["dss"]["password"] = password or ""
    _cfgmod.save_config(cfg)


def clear_credentials():
    cfg = _cfgmod.load_config()
    cfg.pop("dss", None)
    _cfgmod.save_config(cfg)


def _save_cookies(session):
    cfg = _cfgmod.load_config()
    cfg.setdefault("dss", {})
    cfg["dss"]["cookies"] = session.cookies.get_dict()
    _cfgmod.save_config(cfg)


def _stored_cookies():
    return (_cfgmod.load_config().get("dss") or {}).get("cookies") or {}


# ─── 登入狀態機 ────────────────────────────────────────────────────────
# idle → need_captcha → (need_otp) → logged_in；任何錯誤回 need_captcha/error
_LOCK = threading.Lock()
_state = {
    "state": "idle",        # idle / need_captcha / need_otp / logged_in / error
    "error": "",
    "session": None,
    "csrf": None,
    "captcha_png": None,    # bytes
    "otp_form": None,       # {"action","method","fields","code_field"}
    "otp_sent": False,
    "report_csrf": None,    # /report 頁的 CSRF token（查詢資料用）
}


def _requests():
    import requests  # lazy：缺套件時不影響 server 啟動與帳密設定
    return requests


def _new_session():
    s = _requests().Session()
    s.headers.update({"User-Agent": _UA})
    return s


def _parse_csrf(html_text):
    m = re.search(r'name="_csrf"\s+value="([^"]+)"', html_text)
    return m.group(1) if m else None


def _looks_like_login_page(html_text):
    return 'name="inputCaptcha"' in html_text or 'id="loginForm"' in html_text


def _form_action(attrs, default=""):
    m = re.search(r'action\s*=\s*["\']([^"\']*)["\']', attrs, re.I)
    return html.unescape(m.group(1)) if m else default


def _form_method(attrs):
    m = re.search(r'method\s*=\s*["\']([^"\']*)["\']', attrs, re.I)
    return (m.group(1) if m else "post").lower()


def _input_attrs(tag):
    attrs = {}
    for k, _, v1, v2, v3 in re.findall(
        r'([:\w-]+)(\s*=\s*(?:"([^"]*)"|\'([^\']*)\'|([^\s>]+)))?', tag, re.I
    ):
        attrs[k.lower()] = html.unescape(v1 or v2 or v3 or "")
    return attrs


def _button_fields(body, send_like=False):
    """Return a likely clicked submit button name/value for simple HTML forms."""
    candidates = []
    for tag_m in re.finditer(r"<input\b[^>]*>", body, re.I):
        candidates.append((tag_m.group(0), ""))
    for tag_m in re.finditer(r"<button\b([^>]*)>(.*?)</button>", body, re.S | re.I):
        candidates.append((tag_m.group(0), tag_m.group(2) or ""))
    for tag, visible in candidates:
        attrs = _input_attrs(tag)
        typ = (attrs.get("type") or "submit").lower()
        if typ not in ("submit", "button"):
            continue
        visible = re.sub(r"<[^>]+>", "", visible)
        text = " ".join((attrs.get("value", "") + " " + visible + " " + tag).split())
        if send_like and not re.search(r"send|mail|email|otp|code|verify|寄|發|送|信|驗證", text, re.I):
            continue
        name = attrs.get("name")
        return {name: attrs.get("value", "")} if name else {}
    return {}


def _parse_otp_forms(html_text):
    """偵測 Email 驗證流程。

    有些 DSS 帳號 captcha 後會先出現「寄送 Email 驗證碼」按鈕，
    送出後才出現輸入 OTP 的欄位；舊版只找 OTP 輸入欄，因此不會觸發寄信。
    """
    verify_form, send_form = None, None
    for form_m in re.finditer(r"<form\b([^>]*)>(.*?)</form>", html_text, re.S | re.I):
        attrs, body = form_m.group(1), form_m.group(2)
        if "type=\"password\"" in body or "type='password'" in body:
            continue  # 還是登入頁，不是 OTP 頁
        inputs = re.findall(r"<input\b[^>]*>", body, re.I)
        fields, code_field = {}, None
        for tag in inputs:
            tag_attrs = _input_attrs(tag)
            name = tag_attrs.get("name")
            if not name:
                continue
            typ = (tag_attrs.get("type") or "text").lower()
            if typ in ("hidden", "submit"):
                fields[name] = tag_attrs.get("value", "")
            elif typ in ("text", "number", "tel") and re.search(
                r"otp|cap|code|verify|驗證", name + tag, re.I
            ):
                code_field = name
        form = {
            "action": _form_action(attrs),
            "method": _form_method(attrs),
            "fields": fields,
        }
        if code_field and verify_form is None:
            verify_form = {
                **form,
                "fields": fields,
                "code_field": code_field,
            }
            continue
        text = " ".join(re.sub(r"<[^>]+>", " ", body).split())
        if send_form is None and re.search(r"email|mail|otp|驗證碼|寄送|發送|信箱|郵件", text + " " + body, re.I):
            send_form = {
                **form,
                "fields": {**fields, **_button_fields(body, send_like=True)},
            }
    return verify_form, send_form


def _parse_otp_form(html_text):
    verify_form, _ = _parse_otp_forms(html_text)
    return verify_form


def _submit_parsed_form(session, form, extra=None):
    data = dict(form.get("fields") or {})
    if extra:
        data.update(extra)
    action = form.get("action") or LOGIN_PATH
    url = action if action.startswith("http") else urljoin(BASE_URL + "/", action)
    method = (form.get("method") or "post").lower()
    if method == "get":
        return session.get(url, params=data, allow_redirects=False, timeout=30)
    return session.post(url, data=data, allow_redirects=False, timeout=30)


def _set_error(msg):
    _state["state"] = "error"
    _state["error"] = msg


def status():
    u, _ = get_credentials()
    with _LOCK:
        return {
            "available": True,
            "hasCredentials": has_credentials(),
            "username": u or "",
            "state": _state["state"],
            "error": _state["error"],
            "otpSent": _state.get("otp_sent", False),
        }


def get_captcha_png():
    with _LOCK:
        return _state["captcha_png"]


def _fetch_captcha(session):
    res = session.get(BASE_URL + CAPTCHA_PATH, timeout=30)
    res.raise_for_status()
    return res.content


def start_login():
    """開新 session 抓登入頁與驗證碼。回傳 status()。"""
    if not has_credentials():
        with _LOCK:
            _set_error("尚未設定 DSS 帳密")
        return status()
    try:
        session = _new_session()
        res = session.get(BASE_URL + "/", timeout=30)
        res.raise_for_status()
        csrf = _parse_csrf(res.text)
        if not csrf:
            raise RuntimeError("登入頁解析失敗：找不到 _csrf（網站改版？）")
        png = _fetch_captcha(session)
        with _LOCK:
            _state.update(state="need_captcha", error="", session=session,
                          csrf=csrf, captcha_png=png, otp_form=None, otp_sent=False)
    except Exception as exc:
        with _LOCK:
            _set_error(f"無法連線 DSS：{exc}")
    return status()


def refresh_captcha():
    """重新取一張驗證碼圖（同 session）。"""
    with _LOCK:
        session = _state["session"]
        if _state["state"] != "need_captcha" or session is None:
            return status()
    try:
        png = _fetch_captcha(session)
        with _LOCK:
            _state["captcha_png"] = png
    except Exception as exc:
        with _LOCK:
            _set_error(f"驗證碼更新失敗：{exc}")
    return status()


def _after_auth_response(session, res, allow_otp_send=True):
    """登入/OTP 送出後的共同判讀：成功 / 仍在登入頁 / OTP 頁。"""
    # 跟隨 redirect 取最終頁
    page = res
    if res.is_redirect or res.is_permanent_redirect:
        loc = res.headers.get("Location", "")
        url = loc if loc.startswith("http") else BASE_URL + loc
        page = session.get(url, timeout=30)
    text = page.text or ""

    if _looks_like_login_page(text):
        # 回到登入頁＝失敗（驗證碼錯或帳密錯）。重抓 csrf+captcha 讓使用者重試。
        csrf = _parse_csrf(text)
        png = _fetch_captcha(session)
        with _LOCK:
            _state.update(state="need_captcha", session=session,
                          csrf=csrf or _state["csrf"], captcha_png=png, otp_form=None)
            _state["error"] = "登入失敗：請確認帳密與驗證碼後重試"
        return

    otp, otp_send = _parse_otp_forms(text)
    if otp_send and allow_otp_send:
        try:
            sent_res = _submit_parsed_form(session, otp_send)
            ctype = sent_res.headers.get("Content-Type", "")
            if "json" in ctype or not (sent_res.text or "").strip():
                with _LOCK:
                    if otp:
                        _state.update(state="need_otp", error="", session=session,
                                      otp_form=otp, otp_sent=True)
                    else:
                        _set_error("Email 驗證碼可能已送出，但無法解析輸入驗證碼的表單")
                return
            return _after_auth_response(session, sent_res, allow_otp_send=False)
        except Exception as exc:
            with _LOCK:
                _state.update(state="need_otp", error=f"Email 驗證碼寄送失敗：{exc}",
                              session=session, otp_form=otp, otp_sent=False)
            return

    if otp:
        csrf = _parse_csrf(text)
        if csrf:
            otp["fields"].setdefault("_csrf", csrf)
        with _LOCK:
            _state.update(state="need_otp", error="", session=session,
                          otp_form=otp, otp_sent=True)
        return

    with _LOCK:
        _state.update(state="logged_in", error="", session=session,
                      captcha_png=None, otp_form=None, otp_sent=False, report_csrf=None)
    _save_cookies(session)


def submit_captcha(code):
    with _LOCK:
        session, csrf = _state["session"], _state["csrf"]
        if _state["state"] != "need_captcha" or session is None:
            return status()
    username, password = get_credentials()
    try:
        res = session.post(
            BASE_URL + LOGIN_PATH,
            data={"_csrf": csrf, "username": username,
                  "password": password, "inputCaptcha": (code or "").strip()},
            allow_redirects=False, timeout=30,
        )
        _after_auth_response(session, res)
    except Exception as exc:
        with _LOCK:
            _set_error(f"登入送出失敗：{exc}")
    return status()


def submit_otp(code):
    with _LOCK:
        session, otp = _state["session"], _state["otp_form"]
        if _state["state"] != "need_otp" or session is None or not otp:
            return status()
    try:
        res = _submit_parsed_form(session, otp, {otp["code_field"]: (code or "").strip()})
        _after_auth_response(session, res)
    except Exception as exc:
        with _LOCK:
            _set_error(f"驗證碼送出失敗：{exc}")
    return status()


def restore_session():
    """server 啟動時呼叫：還原上次 cookie 並驗證是否仍有效（best-effort）。"""
    cookies = _stored_cookies()
    if not cookies:
        return False
    try:
        session = _new_session()
        for k, v in cookies.items():
            session.cookies.set(k, v)
        res = session.get(BASE_URL + "/", timeout=30)
        if res.ok and not _looks_like_login_page(res.text or ""):
            with _LOCK:
                _state.update(state="logged_in", error="", session=session)
            return True
    except Exception:
        pass
    return False


def is_logged_in(probe=False):
    with _LOCK:
        if _state["state"] != "logged_in" or _state["session"] is None:
            return False
        session = _state["session"]
    if not probe:
        return True
    try:
        res = session.get(BASE_URL + "/", timeout=30)
        ok = res.ok and not _looks_like_login_page(res.text or "")
    except Exception:
        ok = False
    if not ok:
        with _LOCK:
            if _state["session"] is session:
                _state.update(state="idle", error="DSS 登入已過期，請重新登入")
    return ok


def get_session():
    """取得已登入的 requests.Session；未登入丟 RuntimeError（供資料查詢使用）。"""
    with _LOCK:
        if _state["state"] != "logged_in" or _state["session"] is None:
            raise RuntimeError("DSS 尚未登入")
        return _state["session"]


# ─── 搭售統計查詢（3PP搭售率報表(人)，rptId=3PP_aggregate_E）────────────
# 端點為 jqGrid 後端：POST /rpt/search（JSON body＋X-CSRF-TOKEN header），
# whereColumn 帶單據日期區間；回應 {totalRecords, totalPages, data:[...]}。
# 每列＝一位員工，五組機種欄位：{k}_m=搭售台數、{k}_s=配件數、{k}_ms=零搭售台數
# （k ∈ cpu/iphone/ipad/watch/airpods；cpu_m 實際欄名為 cpu_m，比例欄忽略、自行計算）。

REPORT_ID = "3PP_aggregate_E"
BUNDLE_GROUPS = ("cpu", "iphone", "ipad", "watch", "airpods")


def _report_csrf(session):
    """取得報表頁的 CSRF token（每 session 抓一次即可）。"""
    with _LOCK:
        tok = _state.get("report_csrf")
    if tok:
        return tok
    res = session.get(BASE_URL + f"/report?id={REPORT_ID}", timeout=30)
    res.raise_for_status()
    if _looks_like_login_page(res.text or ""):
        _mark_expired(session)
        raise RuntimeError("DSS 登入已過期，請重新登入")
    tok = _parse_csrf(res.text)
    if not tok:
        raise RuntimeError("無法取得 DSS 報表 CSRF token（網站改版？）")
    with _LOCK:
        _state["report_csrf"] = tok
    return tok


def _mark_expired(session):
    with _LOCK:
        if _state["session"] is session:
            _state.update(state="idle", error="DSS 登入已過期，請重新登入")
            _state["report_csrf"] = None


def _rpt_search(session, csrf, start_date, end_date, page):
    payload = {
        "_search": False,
        "nd": int(time.time() * 1000),
        "rows": 500,
        "page": page,
        "sidx": "",
        "sord": "asc",
        "rptId": REPORT_ID,
        "whereColumn": [
            {"code": "start_doc_date", "value": start_date.isoformat()},
            {"code": "end_doc_date", "value": end_date.isoformat()},
            {"code": "org_id", "value": ""},
            {"code": "dept_id", "value": ""},
        ],
    }
    res = session.post(
        BASE_URL + "/rpt/search",
        json=payload,
        headers={"X-CSRF-TOKEN": csrf},
        timeout=60,
    )
    if res.status_code in (401, 403):
        _mark_expired(session)
        raise RuntimeError(f"DSS 搭售查詢被拒（{res.status_code}），請重新登入")
    res.raise_for_status()
    ctype = res.headers.get("Content-Type", "")
    if "json" not in ctype:
        _mark_expired(session)
        raise RuntimeError("DSS 回應非 JSON（登入已過期？），請重新登入")
    return res.json() or {}


def fetch_bundle_stats(shop_id, start_date, end_date, log=lambda m: None):
    """查詢區間內指定門市「每位員工 × 機種」的搭售統計。

    回傳 list[{empId, empName, cpu:{m,s,ms}, iphone:{...}, ...}]。
    未登入或 session 過期丟 RuntimeError，由呼叫端 fail-graceful 略過。
    """
    session = get_session()
    csrf = _report_csrf(session)
    shop = str(shop_id).strip()
    rows, page = [], 1
    total_pages_seen = 1
    while True:
        data = _rpt_search(session, csrf, start_date, end_date, page)
        rows.extend(data.get("data") or [])
        total_pages = int(data.get("totalPages") or data.get("total") or 1)
        total_pages_seen = max(total_pages_seen, total_pages)
        if page >= total_pages:
            break
        page += 1
    out = []
    for r in rows:
        if str(r.get("SHOP_ID", "")).strip() != shop:
            continue
        item = {"empId": (r.get("emp_id1") or "").strip(),
                "empName": (r.get("emp_name") or "").strip()}
        for k in BUNDLE_GROUPS:
            item[k] = {
                "m": int(float(r.get(f"{k}_m") or 0)),
                "s": int(float(r.get(f"{k}_s") or 0)),
                "ms": int(float(r.get(f"{k}_ms") or 0)),
            }
        out.append(item)
    if not out:
        log(f"    DSS {shop} {start_date}~{end_date}：查 {total_pages_seen} 頁、{len(rows)} 筆，無符合門市資料")
    return out
