#!/usr/bin/env python3
"""MyCoach（我的私教管理平台）— 登入與充值紀錄查詢，供第 13/14 頁「MyCoach 應充值數／實際充值數／儲值率」。

登入要圖形驗證碼，且登入狀態撐不過一天 → 每次產報前由人輸入驗證碼登入一次：
  本機網頁：start_login() 開新 session 抓驗證碼圖 → 前端顯示 → submit_code(code) 登入，session 留在記憶體。
  週報小幫手：report_worker 依 Worker 的要求呼叫 new_session / fetch_captcha / login，
              cookie 以門市存在 .mycoach_sessions.json，產報時用 session_override 帶入。
充值紀錄「操作人」欄位就是員工代碼（SA1893…），不需要名字對應；實際充值數＝課時數加總。

帳密與 DSS 同模式存在 local_config.json（"mycoach" 鍵）。
"""
import json
import random
import re
import threading
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path

import shoppertrak_traffic as _cfgmod

BASE_URL = "https://api-tw.china-dt.com/amcm/"
SESSIONS_PATH = Path(__file__).resolve().parent / ".mycoach_sessions.json"
_UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")


class LoginError(RuntimeError):
    pass


# ─── 連線與查詢（本機、小幫手共用）──────────────────────────────────────
def new_session():
    import requests  # lazy：缺套件時不影響 server 啟動
    s = requests.Session()
    s.headers["User-Agent"] = _UA
    s.get(BASE_URL, timeout=30).raise_for_status()   # 先拿 JSESSIONID，驗證碼綁在這個 session
    return s


def fetch_captcha(session):
    """回傳 (圖片 bytes, content-type)。"""
    res = session.get(f"{BASE_URL}service/verifyCode/code?{random.random()}", timeout=30)
    res.raise_for_status()
    return res.content, (res.headers.get("content-type") or "image/jpeg").split(";")[0]


def login(session, username, password, code):
    """同網頁 doLogin（jQuery 預設 GET）；回應是 JS 物件字面值 {code:1,...}，code≠1 丟 LoginError(訊息)。"""
    res = session.get(BASE_URL + "service/loginController/login", timeout=30, params={
        "managerName": username, "password": password, "verifyCode": (code or "").strip()})
    res.raise_for_status()
    text = res.text or ""
    try:
        data = json.loads(text)
    except ValueError:
        m_code = re.search(r"""["']?code["']?\s*:\s*["']?(-?\d+)""", text)
        m_msg = re.search(r"""["']?message["']?\s*:\s*["']([^"']*)""", text)
        data = {"code": m_code.group(1) if m_code else None, "message": m_msg.group(1) if m_msg else text[:100]}
    if str(data.get("code")) != "1":
        raise LoginError(str(data.get("message") or "登入失敗"))


def _check_page(session):
    res = session.get(BASE_URL + "service/amcClassHourUseLog/checkIndex", timeout=30, allow_redirects=False)
    if res.status_code != 200 or 'id="dealerId"' not in (res.text or ""):
        return None
    return res.text


def session_valid(session):
    try:
        return _check_page(session) is not None
    except Exception:
        return False


def _store_ids(session, page):
    """經銷商 ID 在頁面隱藏欄位；門市清單由 findAllbyRole 取得。storeId 留空會查到全經銷商所有門市，所以一定要帶。"""
    m = re.search(r'id="dealerId"[^>]*value="([^"]*)"', page)
    dealer_id = m.group(1) if m else ""
    res = session.get(BASE_URL + "service/amcStoreController/findAllbyRole", timeout=30,
                      params={"dealerId": dealer_id, "storeId": ""})
    res.raise_for_status()
    stores = [s for s in json.loads(res.text or "[]") if s.get("storeId")]
    if len(stores) != 1:
        raise RuntimeError(f"MyCoach 帳號對應到 {len(stores)} 間門市，請改用門市帳號登入")
    return dealer_id, stores[0]["storeId"]


def fetch_recharge(session, start, end):
    """充值紀錄（網頁「充值記錄」，type=6）：[{operator, hours, date}]，含迄日。"""
    page = _check_page(session)
    if page is None:
        raise LoginError("MyCoach 登入已過期，請重新登入")
    dealer_id, store_id = _store_ids(session, page)
    rows, offset, page_size = [], 0, 500
    while True:
        res = session.get(BASE_URL + "service/amcClassHourUseLog/findUseLogList", timeout=60, params={
            "type": 6, "startDate": start.isoformat(), "endDate": end.isoformat(), "mobile": "",
            "dealerId": dealer_id, "storeId": store_id, "channelId": "",
            "draw": 1, "start": offset, "length": page_size})
        res.raise_for_status()
        if not res.text.strip():
            raise LoginError("MyCoach 登入已過期，請重新登入")
        data = res.json()
        batch = data.get("data") or []
        for r in batch:   # 只取員工代碼、課時數、日期；顧客資料不留
            created = str(r.get("createDateTime") or "")[:10]
            try:
                d = date.fromisoformat(created)
            except ValueError:
                continue
            rows.append({"operator": str(r.get("operator") or "").strip().upper(),
                         "hours": float(r.get("hourCount") or 0), "date": d})
        offset += len(batch)
        if not batch or offset >= int(data.get("recordsFiltered") or data.get("recordsTotal") or 0):
            break
    return rows


def hours_by_operator(rows, start, end):
    out = {}
    for r in rows:
        if start <= r["date"] <= end and r["operator"]:
            out[r["operator"]] = out.get(r["operator"], 0) + r["hours"]
    return {k: (int(v) if v == int(v) else v) for k, v in out.items()}


# ─── 產報時用哪個 session ───────────────────────────────────────────────
_OVERRIDE = None   # 週報小幫手工作：report_worker 帶入該門市的 session（一次只跑一份）


@contextmanager
def session_override(session):
    global _OVERRIDE
    prev = _OVERRIDE
    _OVERRIDE = session
    try:
        yield
    finally:
        _OVERRIDE = prev


def current_session():
    """產報用：小幫手工作的 session 優先，否則用本機網頁登入的；都沒有回 None。"""
    if _OVERRIDE is not None:
        return _OVERRIDE
    with _LOCK:
        return _state["session"] if _state["state"] == "logged_in" else None


# ─── 小幫手：各門市 session cookie 存檔（report_worker 每輪是新程序）────────
def _load_sessions():
    try:
        return json.loads(SESSIONS_PATH.read_text("utf-8"))
    except Exception:
        return {}


def save_shop_session(shop_id, session, logged_in):
    data = _load_sessions()
    data[shop_id] = {"cookies": session.cookies.get_dict(), "loggedIn": bool(logged_in),
                     "savedAt": datetime.now().isoformat(timespec="seconds")}
    SESSIONS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")


def load_shop_session(shop_id, require_login=True):
    item = _load_sessions().get(shop_id)
    if not item or (require_login and not item.get("loggedIn")):
        return None
    import requests
    s = requests.Session()
    s.headers["User-Agent"] = _UA
    for k, v in (item.get("cookies") or {}).items():
        s.cookies.set(k, v)
    return s


# ─── 本機網頁登入（帳密存 local_config.json "mycoach"）──────────────────
_LOCK = threading.Lock()
_state = {"state": "idle", "error": "", "session": None, "captcha": None, "captcha_type": "image/jpeg", "logged_in_at": None}


def get_credentials():
    d = _cfgmod.load_config().get("mycoach") or {}
    return d.get("username"), d.get("password")


def has_credentials():
    u, p = get_credentials()
    return bool(u and p)


def set_credentials(username, password):
    cfg = _cfgmod.load_config()
    cfg["mycoach"] = {"username": (username or "").strip(), "password": password or ""}
    _cfgmod.save_config(cfg)


def clear_credentials():
    cfg = _cfgmod.load_config()
    cfg.pop("mycoach", None)
    _cfgmod.save_config(cfg)


def status(probe=False):
    if probe:
        with _LOCK:
            session = _state["session"] if _state["state"] == "logged_in" else None
        if session is not None and not session_valid(session):
            with _LOCK:
                if _state["session"] is session:
                    _state.update(state="idle", session=None, error="MyCoach 登入已過期，請重新登入")
    u, _ = get_credentials()
    with _LOCK:
        return {"available": True, "hasCredentials": has_credentials(), "username": u or "",
                "state": _state["state"], "error": _state["error"], "loggedInAt": _state["logged_in_at"]}


def get_captcha():
    with _LOCK:
        return _state["captcha"], _state["captcha_type"]


def start_login():
    """開新 session 抓驗證碼（也用於「換一張」）。"""
    if not has_credentials():
        with _LOCK:
            _state.update(state="error", error="尚未設定 MyCoach 帳密")
        return status()
    try:
        session = new_session()
        img, ctype = fetch_captcha(session)
        with _LOCK:
            _state.update(state="need_captcha", error="", session=session, captcha=img, captcha_type=ctype)
    except Exception as exc:
        with _LOCK:
            _state.update(state="error", error=f"無法連線 MyCoach：{exc}", session=None)
    return status()


def submit_code(code):
    with _LOCK:
        session = _state["session"]
        if _state["state"] != "need_captcha" or session is None:
            return status()
    username, password = get_credentials()
    try:
        login(session, username, password, code)
        with _LOCK:
            _state.update(state="logged_in", error="", captcha=None,
                          logged_in_at=datetime.now().isoformat(timespec="seconds"))
    except LoginError as exc:
        # 驗證碼用過就失效：自動換一張讓使用者重輸
        start_login()
        with _LOCK:
            _state["error"] = f"登入失敗：{exc}"
    except Exception as exc:
        with _LOCK:
            _state.update(state="error", error=f"登入送出失敗：{exc}", session=None)
    return status()
