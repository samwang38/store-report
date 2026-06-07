#!/usr/bin/env python3
"""ShopperTrak 來客數（人流）查詢 — 後端整合

登入與查詢邏輯比照「人流插件」(content.js)：
  1. Playwright 無頭 Chromium 開啟 analytics.shoppertrak.com，
     自動填入帳號／密碼（沿用插件的「找帳號欄→送出→找密碼欄→送出」啟發式），
     登入完成後從 sessionStorage 取得 authToken + tenantId。
  2. 用該 token 呼叫 rdc-api.shoppertrak.com 的 kpis 端點，取回區間來客總數。

帳密與每店編制人數記錄在本機 local_config.json（已加入 .gitignore）。
Playwright 採 lazy import：未安裝時 get_traffic_total() 會丟出 RuntimeError，
由 server.py 捕捉後「略過來客數、其餘照常產生」。
"""
import json
import threading
import time
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "local_config.json"

# ShopperTrak REST API（同插件）
ORG_ID = 5536
API_BASE = "https://rdc-api.shoppertrak.com"
LOGIN_URL = "https://analytics.shoppertrak.com/"

# EPB 門市代碼（storeCode）→ ShopperTrak siteId（取自插件 content.js 的 STORES）
SITE_BY_SHOP = {
    "004": 82751,     # 士林
    "005": 80028316,  # 微風
    "024": 10094800,  # 美麗華
    "046": 80009128,  # 阿波羅
    "054": 80031194,  # 大葉高島屋
    "050": 82750,     # 西門
    "063": 80172833,  # 板橋遠百
    "064": 80174282,  # 新莊宏匯
    "068": 80215488,  # 新店裕隆城
    "037": 80007164,  # 新竹光復
    "043": 80028206,  # 中壢大江
    "048": 80019208,  # 新竹中正
    "053": 80239830,  # 桃園台茂
    "058": 80052238,  # 苗栗尚順
    "061": 80144772,  # 桃園統領
    "065": 80190022,  # 竹北遠百
    "047": 80009129,  # 台中金典
    "062": 80209110,  # 台中豐原
    "066": 80194639,  # 台中麗寶
    "003": 82752,     # 高雄大立
    "012": 80019436,  # 大億西門
    "027": 10096210,  # 夢時代
    "044": 80007483,  # 三多
    "067": 80195854,  # 高雄岡山
}


# ─── 本機設定檔（帳密 / 編制人數）──────────────────────────────────────
_CFG_LOCK = threading.Lock()


def load_config():
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(cfg):
    with _CFG_LOCK:
        CONFIG_PATH.write_text(
            json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8"
        )


def get_credentials():
    """回傳 (username, password)；未設定則回 (None, None)。"""
    st = load_config().get("shoppertrak") or {}
    return st.get("username"), st.get("password")


def has_credentials():
    u, p = get_credentials()
    return bool(u and p)


def set_credentials(username, password):
    cfg = load_config()
    cfg.setdefault("shoppertrak", {})
    cfg["shoppertrak"]["username"] = (username or "").strip()
    cfg["shoppertrak"]["password"] = password or ""
    save_config(cfg)


def clear_credentials():
    cfg = load_config()
    cfg.pop("shoppertrak", None)
    save_config(cfg)


def get_employee_count(shop_id):
    counts = load_config().get("employeeCount") or {}
    val = counts.get(str(shop_id))
    try:
        return int(val) if val not in (None, "") else None
    except (TypeError, ValueError):
        return None


def set_employee_count(shop_id, count):
    if count in (None, ""):
        return
    try:
        n = int(count)
    except (TypeError, ValueError):
        return
    cfg = load_config()
    cfg.setdefault("employeeCount", {})
    cfg["employeeCount"][str(shop_id)] = n
    save_config(cfg)


def site_id_for_shop(shop_id):
    return SITE_BY_SHOP.get(str(shop_id).strip().upper())


# ─── 登入（Playwright 無頭 Chromium，比照插件流程）──────────────────────
_AUTH_LOCK = threading.Lock()
_auth_cache = {"token": None, "tenant": None}

# 帳號／密碼欄位偵測（對應插件 findUsernameField / findPasswordField）
_USER_HINTS = ("email", "user", "account", "login", "帳號", "使用者")
_SUBMIT_HINTS = (
    "login", "sign in", "next", "continue", "submit",
    "下一步", "繼續", "登入", "登录",
)


def _find_password_selector(page):
    return page.query_selector(
        "input[type='password']:not([disabled]):not([readonly])"
    )


def _find_username_field(page):
    inputs = page.query_selector_all("input")
    for inp in inputs:
        try:
            if inp.get_attribute("type") in ("password",):
                continue
            if inp.get_attribute("disabled") is not None or inp.get_attribute("readonly") is not None:
                continue
            typ = (inp.get_attribute("type") or "text").lower()
            if typ not in ("text", "email", "tel", "search"):
                continue
            hay = " ".join(
                (inp.get_attribute(a) or "")
                for a in ("name", "id", "placeholder", "autocomplete", "aria-label")
            ).lower()
            if any(h in hay for h in _USER_HINTS):
                return inp
        except Exception:
            continue
    # 後備：第一個可見 text/email 欄位
    for inp in inputs:
        try:
            typ = (inp.get_attribute("type") or "text").lower()
            if typ in ("text", "email") and inp.is_visible():
                return inp
        except Exception:
            continue
    return None


def _submit_from(page, field):
    """送出登入表單：優先送出所在 form，否則點擊登入按鈕。"""
    if field is not None:
        try:
            field.press("Enter")
            return
        except Exception:
            pass
    buttons = page.query_selector_all(
        "button, input[type='submit'], [role='button']"
    )
    for btn in buttons:
        try:
            if btn.get_attribute("disabled") is not None:
                continue
            if (btn.get_attribute("aria-disabled") or "") == "true":
                continue
            txt = (btn.inner_text() or btn.get_attribute("value") or "").strip().lower()
            typ = (btn.get_attribute("type") or "").lower()
            if typ == "submit" or any(h in txt for h in _SUBMIT_HINTS):
                btn.click()
                return
        except Exception:
            continue
    raise RuntimeError("找不到登入送出按鈕")


def _read_auth(page):
    return page.evaluate(
        "() => ({ token: sessionStorage.getItem('authToken'),"
        " tenant: sessionStorage.getItem('tenantId') })"
    )


def _login(username, password, log=lambda m: None, timeout_ms=60000):
    """以 Playwright 登入並回傳 (authToken, tenantId)。"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "未安裝 Playwright，無法查詢來客數。請執行："
            "pip3 install playwright && python3 -m playwright install chromium"
        ) from exc

    log("登入 ShopperTrak…")
    deadline = time.time() + timeout_ms / 1000
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.goto(LOGIN_URL, wait_until="domcontentloaded", timeout=timeout_ms)

            # 已是登入狀態？
            auth = _read_auth(page)
            if auth.get("token") and auth.get("tenant"):
                return auth["token"], auth["tenant"]

            # 等待帳號／密碼欄位出現
            user_field = None
            while time.time() < deadline:
                user_field = _find_username_field(page)
                if user_field or _find_password_selector(page):
                    break
                page.wait_for_timeout(400)
            if not user_field and not _find_password_selector(page):
                raise RuntimeError("找不到 ShopperTrak 登入表單")

            pwd_field = _find_password_selector(page)
            if user_field:
                user_field.fill(username)
            if not pwd_field:
                # 兩步式登入：先送出帳號，等待密碼欄出現
                _submit_from(page, user_field)
                while time.time() < deadline:
                    pwd_field = _find_password_selector(page)
                    if pwd_field:
                        break
                    page.wait_for_timeout(400)
            if not pwd_field:
                raise RuntimeError("找不到密碼欄位（請確認帳號正確、未啟用兩步驟驗證）")

            pwd_field.fill(password)
            _submit_from(page, pwd_field)

            # 等待 token 寫入 sessionStorage
            while time.time() < deadline:
                auth = _read_auth(page)
                if auth.get("token") and auth.get("tenant"):
                    log("登入成功")
                    return auth["token"], auth["tenant"]
                page.wait_for_timeout(500)
            raise RuntimeError("登入逾時：未取得 ShopperTrak token（帳密是否正確？）")
        finally:
            browser.close()


def _get_auth(log=lambda m: None, force=False):
    """取得快取或重新登入後的 (token, tenant)。"""
    with _AUTH_LOCK:
        if not force and _auth_cache["token"] and _auth_cache["tenant"]:
            return _auth_cache["token"], _auth_cache["tenant"]
        username, password = get_credentials()
        if not (username and password):
            raise RuntimeError("尚未設定 ShopperTrak 帳密")
        token, tenant = _login(username, password, log=log)
        _auth_cache["token"] = token
        _auth_cache["tenant"] = tenant
        return token, tenant


def _auth_headers(token, tenant):
    return {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token,
        "tenant": tenant,
    }


def _iso_day(d: date):
    return d.isoformat() + "T00:00:00.000Z"


def _fetch_traffic_rows(site_id, start_date, end_date, token, tenant):
    import requests  # lazy：缺套件時不影響 server 與帳密設定
    url = f"{API_BASE}/api/v1/kpis/organizations/{ORG_ID}/sites/{site_id}"
    body = {
        "groupBy": "day",
        "operatingHours": True,
        "reportStartDate": _iso_day(start_date),
        "reportEndDate": _iso_day(end_date),
        "add_aggregated_data": True,
        "kpi": ["traffic"],
    }
    res = requests.post(url, headers=_auth_headers(token, tenant), json=body, timeout=60)
    return res


def get_traffic_total(shop_id, start_date, end_date, log=lambda m: None):
    """回傳 [start_date, end_date]（含端點）區間的來客總數（int）。

    失敗時丟出 RuntimeError；由呼叫端決定是否略過。
    """
    site_id = site_id_for_shop(shop_id)
    if not site_id:
        raise RuntimeError(f"門市 {shop_id} 無對應 ShopperTrak siteId")
    if end_date < start_date:
        return 0

    token, tenant = _get_auth(log=log)
    res = _fetch_traffic_rows(site_id, start_date, end_date, token, tenant)
    if res.status_code in (401, 419):
        # token 失效 → 重登一次再試
        token, tenant = _get_auth(log=log, force=True)
        res = _fetch_traffic_rows(site_id, start_date, end_date, token, tenant)
    if not res.ok:
        raise RuntimeError(f"來客數 API {res.status_code}：{res.text[:160]}")

    rows = (((res.json() or {}).get("result") or [{}])[0].get("currentPeriod") or {}).get("data")
    if not isinstance(rows, list):
        raise RuntimeError("來客數回應格式異常")
    return int(sum(float(r.get("traffic") or 0) for r in rows))
