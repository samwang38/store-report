"""
MySetup（Apple Personal Setup）提交數量：解析匯出的 Excel、把 Staff Name 對到員工代碼。

Excel 只有員工自己打的 Staff Name（如「吳 學諭」「王 CSR 彥翔」「Zheng CSR NaNa」），
週報以 SA 員工代碼為主鍵，對應順序：
  1) 人工對照規則（local_config.json "mysetupAliases"，店長在網頁設定卡事先設好，永遠優先）
  2) 中文全名 == EP_EMP.NAME
  3) 英文詞 == EP_EMP.NAME_LANG，且候選中只有一人符合
都對不上的名字記到 "mysetupUnmatched"，網頁設定卡列出來讓店長指定，下次產報即套用。
規則 key 為正規化後的名字（去空白、去 CSR、英文轉小寫），值為員工代碼或 SKIP（不計入）。
"""
import json
import re
import shutil
import threading
import warnings
from contextlib import contextmanager
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import openpyxl

CONFIG_PATH = Path(__file__).resolve().parent / "local_config.json"
SKIP = "__skip__"
_CFG_LOCK = threading.Lock()
# 週報小幫手的工作：對照規則由 Worker 隨工作帶來（店長在 LINE 設定），不讀寫本機 local_config。
# 產報時一次只跑一份，用模組層級變數即可（同 shoppertrak_traffic.credentials_override）。
_ALIAS_OVERRIDE = None

COL_UID = "Unique ID"
COL_NAME = "Staff Name"
COL_DEVICE = "Device"
COL_SUBMIT = "Personal Setup Submission Time"


# ─── 名字正規化 ──────────────────────────────────────────────────────────
def _strip_title(name):
    return re.sub(r"(?i)(?<![a-z])csr(?![a-z])", " ", str(name or ""))


def normalize(name):
    """去空白、去 CSR 職稱字樣、英文轉小寫：「王 CSR 彥翔」→「王彥翔」、「Zheng CSR NaNa」→「zhengnana」。"""
    return re.sub(r"\s+", "", _strip_title(name)).lower()


def english_tokens(name):
    return {t.lower() for t in re.findall(r"[A-Za-z]+", _strip_title(name))}


# ─── 解析 Excel ──────────────────────────────────────────────────────────
def _to_datetime(value):
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def parse_personal_setup(data):
    """讀所有分頁（iPhone/iPad/Mac/Apple Watch），回傳 [{uid, name, device, submitted}]，以 Unique ID 去重。"""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # Apple 匯出檔沒有預設樣式，openpyxl 會警告
        wb = openpyxl.load_workbook(BytesIO(data), read_only=True, data_only=True)
    records, seen = [], set()
    found_sheet = False
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            continue
        header = [str(h or "").strip() for h in header]
        if COL_NAME not in header or COL_SUBMIT not in header:
            continue
        found_sheet = True
        i_name, i_submit = header.index(COL_NAME), header.index(COL_SUBMIT)
        i_uid = header.index(COL_UID) if COL_UID in header else None
        i_dev = header.index(COL_DEVICE) if COL_DEVICE in header else None
        for row in rows:
            name = str(row[i_name] or "").strip() if i_name < len(row) else ""
            submitted = _to_datetime(row[i_submit]) if i_submit < len(row) else None
            if not name or submitted is None:
                continue
            uid = str(row[i_uid]).strip() if i_uid is not None and row[i_uid] is not None else None
            if uid:
                if uid in seen:
                    continue
                seen.add(uid)
            records.append({
                "uid": uid,
                "name": name,
                "device": (row[i_dev] if i_dev is not None else None) or ws.title,
                "submitted": submitted,
            })
    wb.close()
    if not found_sheet:
        raise ValueError(f"不是 Personal Setup 匯出檔：找不到「{COL_NAME}」「{COL_SUBMIT}」欄位")
    return records


# ─── 對應 ────────────────────────────────────────────────────────────────
def match_names(names, emp_names, aliases):
    """names: Excel 原始名字；emp_names: {code: (中文名, 英文名)}（本店候選員工）；aliases: {正規化名: code/SKIP}。
    回傳 {正規化名: (code / SKIP / None, 依據)}。"""
    by_cn = {}
    by_en = {}
    for code, (cn, en) in emp_names.items():
        if cn:
            by_cn.setdefault(normalize(cn), set()).add(code)
        if en:
            by_en.setdefault(en.strip().lower(), set()).add(code)
    result = {}
    for raw in names:
        key = normalize(raw)
        if key in result:
            continue
        if key in aliases:
            result[key] = (aliases[key], "人工對照")
            continue
        hits = by_cn.get(key, set())
        if len(hits) == 1:
            result[key] = (next(iter(hits)), "中文名")
            continue
        hits = set()
        for tok in english_tokens(raw):
            hits |= by_en.get(tok, set())
        if len(hits) == 1:
            result[key] = (next(iter(hits)), "英文名")
            continue
        result[key] = (None, "同名多人" if hits else "未對上")
    return result


def summarize(records, mapping, start, end):
    """依 Submission Time 落在 start~end（含迄日）統計：回傳 (每員工筆數, 未對上 {顯示名: 筆數}, 不計入 {顯示名: 筆數})。"""
    counts, unmatched, skipped = {}, {}, {}
    for rec in records:
        d = rec["submitted"].date()
        if d < start or d > end:
            continue
        code, _ = mapping.get(normalize(rec["name"]), (None, ""))
        if code == SKIP:
            skipped[rec["name"]] = skipped.get(rec["name"], 0) + 1
        elif code:
            counts[code] = counts.get(code, 0) + 1
        else:
            # 同一正規化名的不同打法合併到第一次出現的寫法
            label = next((k for k in unmatched if normalize(k) == normalize(rec["name"])), rec["name"])
            unmatched[label] = unmatched.get(label, 0) + 1
    return counts, unmatched, skipped


# ─── 設定檔（local_config.json）──────────────────────────────────────────
MYSETUP_DEVICES = ("Mac", "iPad", "iPhone", "Watch")


def device_key(device):
    """Device 欄（或分頁名）正規化成 Mac／iPad／iPhone／Watch；其他回 None。"""
    d = str(device or "").strip().lower()
    if d.startswith("mac"):
        return "Mac"
    if d.startswith("ipad"):
        return "iPad"
    if d.startswith("iphone"):
        return "iPhone"
    if d.startswith("watch") or d.startswith("apple watch"):
        return "Watch"
    return None


def summarize_by_device(records, mapping, start, end):
    """同 summarize，但再依機種拆：回傳 ({員工代碼: {機種: 筆數}}, {未對上顯示名: {機種: 筆數}})；SKIP 不計入。
    機種不是 Mac／iPad／iPhone／Watch 的記在「其他」。"""
    counts, unmatched = {}, {}
    for rec in records:
        d = rec["submitted"].date()
        if d < start or d > end:
            continue
        dev = device_key(rec.get("device")) or "其他"
        code, _ = mapping.get(normalize(rec["name"]), (None, ""))
        if code == SKIP:
            continue
        if code:
            slot = counts.setdefault(code, {})
        else:
            label = next((k for k in unmatched if normalize(k) == normalize(rec["name"])), rec["name"])
            slot = unmatched.setdefault(label, {})
        slot[dev] = slot.get(dev, 0) + 1
    return counts, unmatched


def _load():
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save(cfg):
    if CONFIG_PATH.exists():
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        shutil.copy2(CONFIG_PATH, CONFIG_PATH.with_name(f"local_config.backup-{stamp}.json"))
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def get_aliases(shop_id):
    """{正規化名: {"name": 原始寫法, "code": 員工代碼或 SKIP}}"""
    return dict(_load().get("mysetupAliases", {}).get(shop_id, {}))


def alias_codes(shop_id):
    if _ALIAS_OVERRIDE is not None:
        return dict(_ALIAS_OVERRIDE)
    return {k: v.get("code") for k, v in get_aliases(shop_id).items() if v.get("code")}


@contextmanager
def aliases_override(rules):
    """rules: [{"name", "code"}]（週報小幫手的門市規則）；期間內不讀本機規則、也不寫未對上清單。"""
    global _ALIAS_OVERRIDE
    prev = _ALIAS_OVERRIDE
    _ALIAS_OVERRIDE = {normalize(r.get("name")): str(r.get("code") or "").strip()
                       for r in (rules or []) if normalize(r.get("name")) and r.get("code")}
    try:
        yield
    finally:
        _ALIAS_OVERRIDE = prev


def save_aliases(shop_id, rules):
    """rules: [{"name": Excel 名字, "code": 員工代碼 / SKIP}]，整批覆寫該店規則；已設定的名字從未對上清單移除。"""
    clean = {}
    for r in rules:
        name = str(r.get("name") or "").strip()
        code = str(r.get("code") or "").strip()
        key = normalize(name)
        if key and code:
            clean[key] = {"name": name, "code": code}
    with _CFG_LOCK:
        cfg = _load()
        cfg.setdefault("mysetupAliases", {})[shop_id] = clean
        pending = cfg.get("mysetupUnmatched", {}).get(shop_id, [])
        if pending:
            cfg["mysetupUnmatched"][shop_id] = [u for u in pending if normalize(u.get("name")) not in clean]
        _save(cfg)
    return clean


def get_unmatched(shop_id):
    return list(_load().get("mysetupUnmatched", {}).get(shop_id, []))


def save_unmatched(shop_id, unmatched, last_seen):
    """unmatched: {顯示名: 筆數}；每次產報覆寫，沒有未對上就清空。"""
    if _ALIAS_OVERRIDE is not None:
        return
    items = [{"name": n, "count": c, "lastSeen": last_seen} for n, c in sorted(unmatched.items(), key=lambda x: -x[1])]
    with _CFG_LOCK:
        cfg = _load()
        if cfg.get("mysetupUnmatched", {}).get(shop_id, []) == items:
            return
        cfg.setdefault("mysetupUnmatched", {})[shop_id] = items
        _save(cfg)
