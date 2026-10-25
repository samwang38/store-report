#!/usr/bin/env python3
"""
週報 Excel 自動填充腳本
Usage:
  python fill_weekly_excel.py \
    --week-start 2026-04-05 \
    --week-end   2026-04-11 \
    --data-dir   /path/to/銷售資料/ \
    --template   /path/to/週報模板.xlsx \
    --output     /path/to/週報_2026-04-05~04-11.xlsx
"""
from __future__ import annotations

VERSION = "1.4.1"

import argparse, glob, re, shutil, sys
from datetime import date, timedelta
from calendar import monthrange
from pathlib import Path

import pandas as pd
import openpyxl

# ─── Employee list ─────────────────────────────────────────────────────────────
# Default list (used as fallback if template has no 設定 sheet).
# At runtime, main() overrides these from the template's 設定 sheet,
# so other branches only need to edit that sheet — no Python changes required.
EMPLOYEES = [
    ('SA1092', 'Sam'), ('SA1614', 'Dean'), ('SA1893', 'Jun'),
    ('SA2765', 'Nana'), ('SA3011', 'Zoe'), ('SA3154', 'Winter'), ('SA3221', 'Tanya'),
]
EMP_CODES = [e[0] for e in EMPLOYEES]

def load_employees(wb):
    """Read employee list from 設定 sheet.
    Columns: A = 員工代碼, B = 顯示名稱, C = 排除毛利 (已廢棄，忽略), D = 年度起始日.
    Returns list of (code, name) tuples, or None if sheet is absent / empty."""
    if '設定' not in wb.sheetnames:
        return None
    ws = wb['設定']
    result = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        code = str(row[0]).strip() if row[0] is not None else ''
        name = str(row[1]).strip() if row[1] is not None else ''
        if code and name:
            result.append((code, name))
    return result or None

def load_fiscal_year_start(wb) -> 'date | None':
    """Read fiscal year start date from 設定 sheet col D row 2.
    Returns a date object, or None if not set (will fall back to calendar-quarter logic)."""
    if '設定' not in wb.sheetnames:
        return None
    ws = wb['設定']
    val = ws.cell(row=2, column=4).value   # D2
    if val is None:
        return None
    if hasattr(val, 'date'):               # datetime → date
        return val.date()
    if isinstance(val, date):
        return val
    try:
        return date.fromisoformat(str(val).strip())
    except ValueError:
        return None

# ─── C6 device codes (Mac only — C6 still needed to split mini/iMac/MacBook) ───
C6_CPU     = {6001.0, 6002.0, 6007.0, 6008.0, 6340.0, 6341.0, 6342.0, 6343.0, 6344.0}
#  6001 = iMac, 6002 = Mac mini
#  6340 = MBA 13 M5, 6341 = MBA 15 M5, 6342 = MBP 14 M4/M5, 6343 = MBP 16 M5, 6344 = MacBook Neo
#  ↑ 新 Mac 世代上市時只需更新此處，其餘計算自動跟進
C6_MACBOOK = C6_CPU - {6001.0, 6002.0}   # MacBook 系列（排除 mini / iMac）
C6_AIRPODS = {6258.0, 6312.0, 6330.0, 6348.0}   # 6348 = AirPods 5（2026-09 上市）
C6_SA      = {'cpu': {6533.0}, 'ipad': {6534.0}, 'iphone': {6535.0},
              'watch': {6536.0}, 'airpods': {6537.0}}

# ─── C4 device codes (stable across model years — use for iPhone/Watch/iPad) ───
# C4 represents product category and does NOT change when new models are released.
# Confirmed consistent between 2025 (iPhone 16 era) and 2026 (iPhone 17 era).
C4_IPHONE  = {4004.0}                    # all iPhone models, all years
C4_IPAD    = {4005.0, 4006.0, 4041.0}   # iPad std/Air(4005) + mini(4006) + Pro(4041)
C4_WATCH   = {4038.0}                    # all Watch models, all years
# Mac: C4=4002→MacBook, C4=4001→Mac mini+iMac (split still done via C6)

# C4 categories for 3PP accessory sheet
C4_ROWS = {
    4007.0: 2, 4009.0: 3, 4010.0: 4, 4012.0: 5, 4013.0: 6, 4014.0: 7,
    4017.0: 8, 4021.0: 9, 4022.0: 10, 4026.0: 11, 4039.0: 12,
    4050.0: 13, 4053.0: 14, 4069.0: 15,
}
C4_CPU_SET     = {4007.0}
C4_IPHONE_SET  = {4009.0}
C4_IPAD_SET    = {4010.0}
C4_WATCH_SET   = {4039.0}
C4_AIRPODS_SET = {4069.0}
VAP_BRANDS     = {59.0, 224.0, 277.0, 106.0}

# ─── Sheet 3/4（每月重點 / Speakers）───────────────────────────────────────────
# 每月重點：本期統計對象為品牌代碼 496（Starter Kit / Hello Nature 等入門組）。
# 此頁每月會換統計對象、且未必以品牌區分，換檔時直接改這裡。
MONTHLY_FOCUS_BRANDS = {496.0}
MONTHLY_FOCUS_LABEL  = '品牌 496 Woongjin（ARpedia 繪本：Hello Nature／Starter Kit 入門組、Disney AR 互動繪本組等）'
SPEAKERS_C4          = {4013.0}   # 3PP 藍牙喇叭（同第 7 頁 3PP配件比較的 Speakers 列）
SPEAKERS_EXCL_BRANDS = {453.0}    # 第 4 頁「排除 Momax」區塊：Momax 喇叭皆掛品牌 453

# ─── Sheet 13/14 Y 欄「服務費」──────────────────────────────────────────────
# 99900947 Mac資料轉移/重置系統、99900948 iOS資料轉移/重置系統，統計含稅金額。
SERVICE_FEE_SKUS = {'99900947', '99900948'}

# ─── D 欄「原廠商品營業額」排除條件 ──────────────────────────────────────────
C1_EXCLUDED_FROM_APPLE = {1002.0, 1004.0, 1008.0}
C3_EXCLUDED_FROM_APPLE = {3047.0, 3003.0, 3004.0, 3018.0, 3019.0, 3012.0}
C6_EXCLUDED_FROM_APPLE = {6888.0, 6889.0}
SKU_EXCLUDED_FROM_APPLE = {
    '99200168', '99500006', '99900946', '99900947', '99900948', '99900949',
    '99900950', '99901684', '99901685', '99902607', '99902608', '99902609',
    '99902610',
    '888',       # 2點折1元（點數折抵）
    '99200202',  # 教育價活動（教育優惠內部代碼）
    '99901780',  # SA Care 檢測新機活動代碼
}

# Slide 1 product row definitions
PRODUCT_ROWS = {
    1:  (3001.0, {6001.0}, None, None),
    2:  (3001.0, {6002.0}, None, None),
    3:  (3001.0, None, 'MBP 14', None),   # 名稱辨別，自動涵蓋各世代 MacBook Pro 14
    4:  (3001.0, None, 'MBP 16', None),   # MacBook Pro 16
    5:  (3001.0, None, 'MBA 13', None),   # MacBook Air 13
    6:  (3001.0, None, 'MBA 15', None),   # MacBook Air 15
    7:  (3001.0, None, 'MBN',    None),   # MacBook Neo
    # ── row 8 = CPU 小計（由 SUBTOTAL_ROWS 寫入）──
    9:  (3002.0, {6330.0}, None, None),           # AirPods Pro 3
    10: (3002.0, {6312.0}, '主動式降噪', None),    # AirPods 4 主動降噪
    11: (3002.0, {6312.0}, None, '主動式降噪'),    # AirPods 4
    12: (3002.0, {6348.0}, None, None),           # AirPods 5
    13: (3002.0, {6258.0}, None, None),           # AirPods Max
    14: (3002.0, {6073.0}, 'MINI', None),         # HomePod mini
    15: (3002.0, {6073.0}, None, 'MINI'),         # HomePod
    17: (3001.0, {6327.0}, None, None),           # Watch S11
    18: (3001.0, {6349.0}, None, None),           # Watch S12
    19: (3001.0, {6328.0}, None, None),           # Watch Ultra 3
    20: (3001.0, {6350.0}, None, None),           # Watch Ultra 4
    21: (3001.0, {6329.0}, None, None),           # Watch SE3
    23: (3001.0, {6323.0}, None, None),           # iPhone 17
    24: (3001.0, {6324.0}, None, None),           # iPhone Air
    25: (3001.0, {6325.0}, None, None),           # iPhone 17 Pro
    26: (3001.0, {6326.0}, None, None),           # iPhone 17 Pro Max
    27: (3001.0, {6335.0}, None, None),           # iPhone 17e
    28: (3001.0, {6345.0}, None, None),           # iPhone 18 Pro
    29: (3001.0, {6346.0}, None, None),           # iPhone 18 Pro Max
    30: (3001.0, {6347.0}, None, None),           # iPhone Duo（2026-10 上市）
    32: (3001.0, {6321.0, 6322.0}, None, None),
    33: (3001.0, {6317.0, 6318.0, 6336.0}, None, None),
    34: (3001.0, {6319.0, 6338.0}, None, None),
    35: (3001.0, {6331.0, 6332.0}, None, None),
    36: (3001.0, {6333.0}, None, None),
    37: (3001.0, {6313.0}, None, None),
    39: (3002.0, {6100.0}, '第一代', None),
    40: (3002.0, {6100.0}, '2nd', None),
    41: (3002.0, {6100.0}, 'USB-C', None),
    42: (3002.0, {6100.0}, 'PRO', 'keyboard'),  # 排除 Magic Keyboard for iPad Pro
}
# 新機型上市：在模板第 1 頁插列、這裡補一列並重編號（row = Excel 列 − 1）
SUBTOTAL_ROWS = {
    8:  list(range(1, 8)),   # CPU 小計（含 MacBook Neo row 7）
    16: list(range(9, 16)),  # 音頻小計
    22: list(range(17, 22)), # Watch 小計
    31: list(range(23, 31)), # iPhone 小計
    38: list(range(32, 38)), # iPad 小計
    43: list(range(39, 43)), # Apple Pencil 小計
}

# ─── CLI ───────────────────────────────────────────────────────────────────────
def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument('--version',    action='store_true', help='顯示版本號後結束')
    p.add_argument('--week-start', required=False)
    p.add_argument('--week-end',   required=False)
    p.add_argument('--yoy-end',    required=False, help='Sheet 16/17 年對年截止日（YYYY-MM-DD）；預設＝週結束日')
    p.add_argument('--data-dir',   required=False)
    p.add_argument('--template',   required=False)
    p.add_argument('--output',     required=False)
    p.add_argument('--output-dir', required=False)
    return p.parse_args()

# ─── File discovery ────────────────────────────────────────────────────────────
def find_800ab(data_dir: Path, year: int) -> str | None:
    # Try year-prefixed patterns first
    for pat in [f'800AB_{year}*.xlsx', f'800AB{year}*.xlsx']:
        found = sorted(glob.glob(str(data_dir / pat)))
        if found:
            if year < date.today().year:
                annual = [f for f in found if '整年' in f]
                return annual[0] if annual else found[-1]
            else:
                non_annual = [f for f in found if '整年' not in f]
                return non_annual[-1] if non_annual else found[-1]
    # Fallback: any 800AB*.xlsx without 整年 for current year,
    # or with 整年 for previous year
    all_files = sorted(glob.glob(str(data_dir / '800AB*.xlsx')))
    if year < date.today().year:
        annual = [f for f in all_files if '整年' in f and str(year) in f]
        return annual[0] if annual else None
    else:
        non_annual = [f for f in all_files if '整年' not in f]
        return non_annual[-1] if non_annual else None

def find_sacare(data_dir: Path) -> Path | None:
    for name in ['SAcare對應價目表.xlsx', 'SAcare對應價目表.xlsx']:
        p = data_dir / name
        if p.exists():
            return p
    return None

# ─── Data loading ──────────────────────────────────────────────────────────────
def find_header_row(filepath):
    for h in range(0, 20):
        try:
            df = pd.read_excel(filepath, header=h, nrows=1)
            # 去除 BOM 及前後空白再比對，避免欄位名稱有隱藏字元導致比對失敗
            cols = {str(c).strip().lstrip('\ufeff') for c in df.columns.tolist()}
            # 不限定欄位順序，只要 '單據日期' 與 '單據代碼' 同時出現即為正確 header 列
            if '單據日期' in cols and '單據代碼' in cols:
                return h
        except Exception:
            pass
    return 8  # 自店 004 預設 header 在第 9 行（0-indexed = 8）

# 所有報表計算所需的必要欄位，及其用途說明（用於錯誤訊息）
REQUIRED_COLUMNS = {
    '單據日期':     '期間篩選（本週 / 本月 / 去年同期）',
    '單據代碼':     '成交筆數計算',
    '交易類型':     '銷售=一般／訂金=定金計營業額／尾款=完成計台數／銷退=退貨扣除',
    '存貨代碼':     '辨識 SAcare 品項',
    '名稱':         'MacBook 機型區分、ACPP、Office、3PP 排名',
    '數量':         '台數與數量計算',
    '銷售金額(含稅)':'營業額計算',
    '銷退金額':     '營業額計算（退貨扣除）',
    '單位成本':     '毛利計算',
    '品牌代碼':     '認證機台數（881/885/886/888）、VAP 業績、毛利排除（297）',
    '類別1代碼':    '毛利排除條件（C1=21）',
    '類別3代碼':    '主要品類（3001 主機 / 3002 原廠配件 / 3003 三方 / 3032 ACPP）',
    '類別4代碼':    'iPhone / iPad / Watch 台數（C4 穩定代碼）',
    '類別6代碼':    'Mac 機型區分、SAcare 品類、Sheet 1 產品列',
    '員工代碼':         '個人業績計算（Sheet 12～15）',
    '折扣':             '排除贈品交易（Sheet 8/9 銷售排名）',
    '等級代碼':         'Sheet 12 排除等級 05（非全職／兼職員工）',
    '淨銷售金額(未稅)': 'Sheet 12 H/I 欄毛利計算主要依據（缺少時毛利會變負數）',
}

def validate_columns(df, filepath):
    """Check all required columns exist; exit with clear message if any are missing."""
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        print(f'\n❌ 錯誤：800AB 檔案缺少必要欄位')
        print(f'   檔案：{Path(filepath).name}')
        print(f'\n   缺少的欄位（共 {len(missing)} 個）：')
        for col in missing:
            print(f'     • {col}　→　{REQUIRED_COLUMNS[col]}')
        print(f'\n   請重新匯出 800AB，確認以上欄位已勾選後再執行。')
        sys.exit(1)

def load_800(filepath):
    print(f'  Loading {Path(filepath).name}...', flush=True)
    hdr = find_header_row(filepath)
    try:
        df = pd.read_excel(filepath, sheet_name='Sheet', header=hdr)
    except Exception:
        df = pd.read_excel(filepath, header=hdr)
    # 統一去除欄位名稱的 BOM 及前後空白，防止 KeyError
    df.columns = [str(c).strip().lstrip('\ufeff') for c in df.columns]
    df['單據日期'] = pd.to_datetime(df['單據日期'])
    for col in ['品牌代碼', '類別3代碼', '類別4代碼', '類別6代碼']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    if '員工代碼' in df.columns:
        df['員工代碼'] = df['員工代碼'].astype(str).str.strip()
    if '等級代碼' in df.columns:
        df['等級代碼'] = df['等級代碼'].astype(str).str.strip()
    df['NET'] = df.get('銷售金額(含稅)', 0).fillna(0) + df.get('銷退金額', 0).fillna(0)
    validate_columns(df, filepath)
    return df

def load_sacare(path: Path) -> dict:
    print(f'  Loading SAcare prices...', flush=True)
    dfs = pd.read_excel(path, header=0)
    prices = {}
    for _, row in dfs.iterrows():
        if pd.notna(row.iloc[0]) and pd.notna(row.iloc[2]):
            prices[str(row.iloc[0]).strip()] = float(row.iloc[2])
    print(f'    {len(prices)} SAcare codes', flush=True)
    return prices

# ─── Date utilities ────────────────────────────────────────────────────────────
def period(df, start: date, end: date) -> pd.DataFrame:
    s, e = pd.Timestamp(start), pd.Timestamp(f'{end} 23:59:59')
    return df[(df['單據日期'] >= s) & (df['單據日期'] <= e)]

def fiscal_year_start_for(d: date) -> date:
    """Apple 財年起始日（規則）：財年在九月最後一個週六結束，新財年從隔天週日開始。
    回傳包含 d 的那個財年的起始日。財年多為 52 週，偶爾 53 週（例：2027-09-26～2028-09-30）。"""
    def start_of(year):
        sep30 = date(year, 9, 30)
        last_sat = sep30 - timedelta(days=(sep30.weekday() - 5) % 7)
        return last_sat + timedelta(days=1)
    s = start_of(d.year)
    return s if d >= s else start_of(d.year - 1)


def resolve_fiscal_year_start(week_start: date, configured: 'date | None') -> 'date | None':
    """決定 week_start 該用的財年起始日。
    configured＝設定頁 D2；沒設定（None）就回 None（維持曆法季度 fallback）。
    D2 落在「規則起始日 ～ week_start」之間（正常就是等於規則起始日）就沿用 D2；
    D2 已落後（week_start 進入之後的財年）或晚於 week_start（補做舊財年的週）就改用規則，不改寫 Excel。"""
    if configured is None:
        return None
    rule = fiscal_year_start_for(week_start)
    return configured if rule <= configured <= week_start else rule


def compute_quarter_start(week_start: date, fiscal_year_start: 'date | None' = None) -> date:
    """Return the start date of the fiscal quarter containing week_start.

    If fiscal_year_start is provided (read from 設定 sheet col D2):
      - D2 先經 resolve_fiscal_year_start 校正（D2 落後時自動換年）
      - Quarter start = fiscal_year_start + (quarter_index × 13 weeks)
      - 53 週年度的第 53 週歸 Q4（Q4 起始日不變）

    Falls back to calendar-quarter logic if fiscal_year_start is None.
    """
    if fiscal_year_start is not None:
        fiscal_year_start = resolve_fiscal_year_start(week_start, fiscal_year_start)
        delta_days = (week_start - fiscal_year_start).days
        week_in_year = delta_days // 7          # 0-based week index from year start
        quarter_idx  = min(week_in_year // 13, 3)   # 0-based quarter index (0=Q1,1=Q2…)；第 53 週歸 Q4
        return fiscal_year_start + timedelta(weeks=quarter_idx * 13)

    # Fallback: calendar-quarter (original logic)
    q_first_month = ((week_start.month - 1) // 3) * 3 + 1
    q_first_day   = date(week_start.year, q_first_month, 1)
    days_ahead    = (week_start.weekday() - q_first_day.weekday()) % 7
    return q_first_day + timedelta(days=days_ahead)


def compute_fiscal_week_number(week_start: date, fiscal_year_start: date) -> tuple:
    """Returns (quarter_num, week_in_quarter) — both 1-based.
    fiscal_year_start（D2）先經 resolve_fiscal_year_start 校正；53 週年度的第 53 週標 Q4W14。"""
    fiscal_year_start = resolve_fiscal_year_start(week_start, fiscal_year_start)
    delta_days      = (week_start - fiscal_year_start).days
    week_in_year    = delta_days // 7
    quarter_idx     = min(week_in_year // 13, 3)
    quarter_num     = quarter_idx + 1
    week_in_quarter = week_in_year - quarter_idx * 13 + 1
    return quarter_num, week_in_quarter

def safe_rate(num, den):
    if not den or num is None:
        return 0
    return round(num / den, 4)

# 台數在尾款當天計入（銷售完成日）；訂金只計營業額，不計台數
# 一般現金交易用「銷售」；分期付款完成用「尾款」
SALE_TYPES = {'銷售', '尾款'}

# ─── Core metric calculator ────────────────────────────────────────────────────
def calc_metrics(d: pd.DataFrame, sacare_prices: dict) -> dict:
    """Calculate all KPIs for a filtered DataFrame (already period-filtered)."""
    sa_codes = set(sacare_prices.keys())
    non_sa = d[~d['存貨代碼'].astype(str).str.strip().isin(sa_codes)]

    # Revenue by category
    rev_3001 = non_sa.loc[non_sa['類別3代碼'] == 3001.0, 'NET'].sum()
    rev_3002 = non_sa.loc[non_sa['類別3代碼'] == 3002.0, 'NET'].sum()
    rev_3003 = non_sa.loc[non_sa['類別3代碼'] == 3003.0, 'NET'].sum()

    # SAcare：只計 SALE_TYPES（銷售/尾款），排除訂金避免雙重計算
    sa_rows = d[d['存貨代碼'].astype(str).str.strip().isin(sa_codes)].copy()
    sa_sold = sa_rows[sa_rows['交易類型'].isin(SALE_TYPES)]
    sa_ret  = sa_rows[sa_rows['交易類型'] == '銷退']
    sa_sold_net = (sa_sold['存貨代碼'].astype(str).str.strip().map(sacare_prices) * sa_sold['數量'].fillna(0)).sum()
    sa_ret_net  = (sa_ret['存貨代碼'].astype(str).str.strip().map(sacare_prices) * sa_ret['數量'].abs()).sum()
    sa_rev = sa_sold_net - sa_ret_net
    sa_gross = sa_rev / 2

    # 日報口徑：a 零售營業額＝非 SA 的 銷售+銷退；c 零售訂金營業額＝訂金+退訂；
    #           b 訂單出貨金額＝尾款列折後金額（尾款 NET=0，實際出貨金額在 LINE_TOTAL_AFTDISC）
    # 總營業額 = a + c + e(SA)（d 專案＝0；尾款 NET 不計）
    retail_rev  = non_sa.loc[non_sa['交易類型'].isin(['銷售', '銷退']), 'NET'].sum()
    deposit_rev = non_sa.loc[non_sa['交易類型'].isin(['訂金', '退訂']), 'NET'].sum()
    order_ship  = d.loc[d['交易類型'] == '尾款', '折後金額'].sum() if '折後金額' in d.columns else 0
    total_rev = retail_rev + deposit_rev + sa_rev

    # ACPP（C3=3032 AppleCare+ 代收保費）
    acpp_rev = non_sa.loc[non_sa['類別3代碼'] == 3032.0, 'NET'].sum()

    # Gross profit (Apple C3 ∈ {3001,3002,3032,3033}, excl SAcare, brand≠297, C6≠31, C1≠21)
    excl = (d['存貨代碼'].astype(str).str.strip().isin(sa_codes) |
            (d.get('品牌代碼', pd.Series(dtype=float)) == 297.0) |
            (d.get('類別6代碼', pd.Series(dtype=float)) == 31.0) |
            (d.get('類別1代碼', pd.Series(dtype=float)) == 21.0))
    apl_mask = d['類別3代碼'].isin([3001.0, 3002.0, 3033.0]) & ~excl
    tpp_mask = (d['類別3代碼'] == 3003.0) & ~excl
    acpp_mask = (d['類別3代碼'] == 3032.0) & ~excl

    def gross(mask):
        # 排除尾款：尾款 NET=0 但 ERP 仍記錄完整成本，會造成毛利負數
        sub = d[mask & (d['交易類型'] != '尾款')]
        return (sub['NET'] - (sub.get('單位成本', 0).fillna(0) * 1.05).round() * sub['數量'].fillna(0)).sum()

    apl_gross  = gross(apl_mask)
    tpp_gross  = gross(tpp_mask)
    acpp_gross = gross(acpp_mask)

    # Unit counts
    # Mac: C6-based (need to distinguish MacBook / mini / iMac within same C4=4001/4002)
    cert_brands = {881.0, 885.0, 886.0, 888.0}
    unit_mask = (d['類別3代碼'] == 3001.0) | (d['品牌代碼'].isin(cert_brands))

    def net_units_c6(c6_set):
        m = unit_mask & d['類別6代碼'].isin(c6_set)
        sale = d.loc[m & d['交易類型'].isin(SALE_TYPES), '數量'].sum()
        ret  = d.loc[m & (d['交易類型'] == '銷退'),  '數量'].abs().sum()
        return int(sale - ret)

    # iPhone / Watch / iPad: C4-based + 認證機品牌 bypass
    def net_units_c4(c4_set):
        m = ((d['類別3代碼'] == 3001.0) | d['品牌代碼'].isin(cert_brands)) & \
            d['類別4代碼'].isin(c4_set)
        sale = d.loc[m & d['交易類型'].isin(SALE_TYPES), '數量'].sum()
        ret  = d.loc[m & (d['交易類型'] == '銷退'),  '數量'].abs().sum()
        return int(sale - ret)

    cpu_non_mini = net_units_c6(C6_MACBOOK)   # MacBook（引用 C6_MACBOOK 常數，新世代只需更新 C6_CPU）
    cpu_mini     = net_units_c6({6002.0})
    cpu_imac     = net_units_c6({6001.0})
    cpu_total    = cpu_non_mini + cpu_mini + cpu_imac
    iphone_units = net_units_c4(C4_IPHONE)
    ipad_units   = net_units_c4(C4_IPAD)
    watch_units  = net_units_c4(C4_WATCH)
    # AirPods：C3=3002（或認證機）且為 AirPods 主機；C6_AIRPODS 只含現行世代，
    # 去年的舊款（如 AirPods Pro 2）由 _is_airpods 以品名補判，否則年對年台數/搭售率失真
    ap_mask = ((d['類別3代碼'] == 3002.0) | d['品牌代碼'].isin(cert_brands)) & _is_airpods(d)
    airpods_units = int(d.loc[ap_mask & d['交易類型'].isin(SALE_TYPES), '數量'].sum() -
                        d.loc[ap_mask & (d['交易類型'] == '銷退'), '數量'].abs().sum())

    # ACPP 套數：C3=3032，依名稱關鍵字分裝置，銷售/尾款正計、銷退扣回（同 calc_employee）
    ac_name = d['名稱'].astype(str).str.lower()
    def acpp_cnt(keyword):
        m = (d['類別3代碼'] == 3032.0) & ac_name.str.contains(keyword, na=False)
        return int(d.loc[m & d['交易類型'].isin(SALE_TYPES), '數量'].sum() -
                   d.loc[m & (d['交易類型'] == '銷退'), '數量'].abs().sum())

    # SAcare counts by device
    def sa_units(c6_set):
        m = d['類別6代碼'].isin(c6_set) & d['存貨代碼'].astype(str).str.strip().isin(sa_codes)
        return int(d.loc[m & d['交易類型'].isin(SALE_TYPES), '數量'].sum() -
                   d.loc[m & (d['交易類型'] == '銷退'), '數量'].abs().sum())

    sa_cpu     = sa_units(C6_SA['cpu'])
    sa_ipad    = sa_units(C6_SA['ipad'])
    sa_iphone  = sa_units(C6_SA['iphone'])
    sa_watch   = sa_units(C6_SA['watch'])
    sa_airpods = sa_units(C6_SA['airpods'])

    # 舊換新台數：回收舊機以 C3=3020「舊換新-機型」銷退列入帳（同單另有正數折抵列），數銷退列台數
    tradein = int(d.loc[(d['類別3代碼'] == 3020.0) & (d['交易類型'] == '銷退') &
                        d['名稱'].astype(str).str.startswith('舊換新'), '數量'].abs().sum())

    # 成交筆數 (excluding 正負沖銷)
    sale_codes = set(d.loc[d['交易類型'] == '銷售', '單據代碼'].dropna().astype(str))
    ret_codes  = set(d.loc[d['交易類型'] == '銷退',  '單據代碼'].dropna().astype(str))
    txn_count  = len(sale_codes - ret_codes)

    m = dict(
        rev_3001=int(rev_3001), rev_3002=int(rev_3002), rev_3003=int(rev_3003),
        sa_rev=int(sa_rev), sa_gross=int(sa_gross),
        total_rev=int(total_rev),
        retail_rev=int(retail_rev), deposit_rev=int(deposit_rev), order_ship=int(order_ship),
        retail_income=int(retail_rev + order_ship),
        acpp_rev=int(acpp_rev),
        apl_gross=int(apl_gross), tpp_gross=int(tpp_gross), acpp_gross=int(acpp_gross),
        total_gross=int(apl_gross + tpp_gross + acpp_gross + sa_gross),
        cpu_non_mini=cpu_non_mini, cpu_mini=cpu_mini, cpu_total=cpu_total,
        iphone=iphone_units, ipad=ipad_units, watch=watch_units, airpods=airpods_units,
        acpp_mac=acpp_cnt('mac'), acpp_ipad=acpp_cnt('ipad'), acpp_iphone=acpp_cnt('iphone'),
        acpp_watch=acpp_cnt('watch'), acpp_airpods=acpp_cnt('airpods'),
        sa_cpu=sa_cpu, sa_ipad=sa_ipad, sa_iphone=sa_iphone,
        sa_watch=sa_watch, sa_airpods=sa_airpods,
        txn_count=txn_count, tradein=tradein,
    )
    # 保固（AC+SA）＝ ACPP + SA Care：營業額與各機種套數
    m['warranty_rev'] = m['acpp_rev'] + m['sa_rev']
    for dev, ac, sa_k in (('cpu', 'acpp_mac', 'sa_cpu'), ('ipad', 'acpp_ipad', 'sa_ipad'),
                          ('iphone', 'acpp_iphone', 'sa_iphone'), ('watch', 'acpp_watch', 'sa_watch'),
                          ('airpods', 'acpp_airpods', 'sa_airpods')):
        m[f'warranty_{dev}'] = m[ac] + m[sa_k]
    m['warranty_gross'] = m['acpp_gross'] + m['sa_gross']
    return m

# ─── Sheet 1: 主機銷售台數 ─────────────────────────────────────────────────────
def fill_sheet1(ws, df_cur: pd.DataFrame, quarter_start: date, week_end: date):
    print('  Sheet 1: 主機銷售台數', flush=True)
    d = df_cur.copy()   # 尾款計台數，訂金不計，透過 SALE_TYPES 控制
    _cert = {881.0, 885.0, 886.0, 888.0}
    _brand = d.get('品牌代碼', pd.Series(dtype=float, index=d.index)).fillna(0)

    def calc_units(start, end, c3, c6_set, must_inc, must_exc):
        s, e = pd.Timestamp(start), pd.Timestamp(f'{end} 23:59:59')
        sub = d[(d['單據日期'] >= s) & (d['單據日期'] <= e)]
        # 認證機（品牌 881/885/886/888）無論 C3 為何，只要符合條件就計入
        brand_sub = _brand.reindex(sub.index).fillna(0)
        c3_mask = (sub['類別3代碼'] == c3) | brand_sub.isin(_cert)
        if c6_set is None:
            # 名稱辨別模式：只用 C3 + must_inc/must_exc（不限 C6），適用 MacBook 各型號
            sub = sub[c3_mask]
        else:
            sub = sub[c3_mask & sub['類別6代碼'].isin(c6_set)]
        if must_inc: sub = sub[sub['名稱'].str.contains(must_inc, case=False, na=False)]
        if must_exc: sub = sub[~sub['名稱'].str.contains(must_exc, case=False, na=False)]
        # 台數邏輯：尾款當天計入（銷售完成），訂金不計，銷退扣除
        sale = sub.loc[sub['交易類型'].isin(SALE_TYPES), '數量'].sum()
        ret  = sub.loc[sub['交易類型'] == '銷退', '數量'].abs().sum()
        return int(sale - ret)

    prev_end   = quarter_start - timedelta(days=1)
    prev_start = prev_end - timedelta(days=6)

    weeks = [(quarter_start + timedelta(weeks=i),
              quarter_start + timedelta(weeks=i, days=6)) for i in range(13)]

    table = {}
    for row_idx, (c3, c6_set, inc, exc) in PRODUCT_ROWS.items():
        prev = calc_units(prev_start, prev_end, c3, c6_set, inc, exc)
        weekly = []
        for ws_, we_ in weeks:
            if ws_ > week_end:
                weekly.append(0)
            else:
                weekly.append(calc_units(ws_, min(we_, week_end), c3, c6_set, inc, exc))
        table[row_idx] = [prev] + weekly

    for st_row, members in SUBTOTAL_ROWS.items():
        combined = [0] * 14
        for r in members:
            if r in table:
                for i, v in enumerate(table[r]):
                    combined[i] += v
        table[st_row] = combined

    # Excel row = table row + 1 (header is row 1)
    # Excel cols: B=2(prev W13), C=3(W01)...O=15(W13), P=16(Total)
    for row_idx, vals in table.items():
        excel_row = row_idx + 1
        total = sum(vals[1:])
        ws.cell(row=excel_row, column=2).value = vals[0] or None   # prev W13
        for wi, v in enumerate(vals[1:], start=3):
            ws.cell(row=excel_row, column=wi).value = v or None    # W01-W13
        ws.cell(row=excel_row, column=16).value = total or None    # Total

# ─── Sheet 2: 教育價 ─────────────────────────────────────────────────────────
# 教育價交易＝下列兩條件的聯集（兩種登錄方式都會漏，所以取聯集）：
#   1. 同一張單據裡有存貨代碼 99500274「教育證號」（銷退單也會帶，故退貨能正常扣回）
#   2. ERP「促銷活動名稱」(POSLINEV_BI.MC_ID → POS_MC_CODE.MC_NAME) 含「教育價」
#      涵蓋 教育價活動／電腦單機教育價／iPad單機教育價／教育價ARpedia活動／ACC 教育價／教育價DG活動
# Keyboard 與 Pencil 只算 Apple 原廠（C3=3002 且 C4=4010）；Pencil 排除「筆尖」耗材。
EDU_KEYWORD = '教育價'
EDU_CERT_SKU = '99500274'   # 教育證號（單據上的存代，非商品）
EDU_ROWS = ['CPU', 'iPad', 'Watch', 'Pencil', 'Keyboard']


def _edu_row_mask(d: pd.DataFrame, row: str):
    c3, c4, c6 = d['類別3代碼'], d['類別4代碼'], d['類別6代碼']
    name = d['名稱'].fillna('')
    if row == 'CPU':
        return (c3 == 3001.0) & c6.isin(C6_CPU)
    if row == 'iPad':
        return c4.isin(C4_IPAD)
    if row == 'Watch':
        return c4.isin(C4_WATCH)
    apple_ipad_acc = (c3 == 3002.0) & c4.isin(C4_IPAD_SET)   # 原廠 iPad 週邊配件
    if row == 'Keyboard':
        return apple_ipad_acc & (name.str.contains('keyboard', case=False, na=False) |
                                 name.str.contains('巧控鍵盤', na=False))
    # Pencil：Apple Pencil 本體，排除「筆尖」耗材
    return apple_ipad_acc & name.str.contains('pencil', case=False, na=False) & \
        ~name.str.contains('筆尖', na=False)


def fill_sheet_edu(ws, df_cur: pd.DataFrame, quarter_start: date, week_end: date):
    """第 2 頁 教育價：當季 13 週逐週的總銷售台數 / 教育價台數 / 教育價佔比。
    區塊：列 1-6 總銷售台數、列 7-12 教育價台數、列 13-18 教育價佔比；
    欄：B~N＝W1…W13、O＝總計。台數口徑同第 1 頁（銷售＋尾款正計、銷退扣回）。"""
    print('  Sheet 2: 教育價', flush=True)
    d = df_cur
    none_mask = pd.Series(False, index=d.index)
    # ① 同單有「教育證號」；② 促銷活動名稱含「教育價」
    # （CLI 讀 Excel 的路徑沒有單據代碼／促銷活動名稱這兩欄 → 該條件視為不成立）
    doc = d.get('單據代碼')
    if doc is not None:
        sku = d['存貨代碼'].astype(str).str.strip()
        by_cert = doc.isin(set(doc[sku == EDU_CERT_SKU]))
    else:
        by_cert = none_mask
    promo = d.get('促銷活動名稱')
    by_promo = promo.fillna('').str.contains(EDU_KEYWORD, na=False) if promo is not None else none_mask
    edu_all = by_cert | by_promo

    weeks = [(quarter_start + timedelta(weeks=i),
              quarter_start + timedelta(weeks=i, days=6)) for i in range(13)]

    def units(mask) -> int:
        sale = d.loc[mask & d['交易類型'].isin(SALE_TYPES), '數量'].sum()
        ret  = d.loc[mask & (d['交易類型'] == '銷退'), '數量'].abs().sum()
        return int(sale - ret)

    for j, row_name in enumerate(EDU_ROWS):
        base = _edu_row_mask(d, row_name)
        totals, edus = [], []
        for w_start, w_end in weeks:
            if w_start > week_end:
                totals.append(None)
                edus.append(None)
                continue
            in_week = ((d['單據日期'] >= pd.Timestamp(w_start)) &
                       (d['單據日期'] <= pd.Timestamp(f'{min(w_end, week_end)} 23:59:59')))
            totals.append(units(base & in_week))
            edus.append(units(base & in_week & edu_all))
        for col, (tot, edu) in enumerate(zip(totals, edus), start=2):
            ws.cell(row=2 + j,  column=col).value = tot
            ws.cell(row=8 + j,  column=col).value = edu
            ws.cell(row=14 + j, column=col).value = (edu / tot) if tot else None
        tot_sum = sum(v for v in totals if v)
        edu_sum = sum(v for v in edus if v)
        ws.cell(row=2 + j,  column=15).value = tot_sum
        ws.cell(row=8 + j,  column=15).value = edu_sum
        ws.cell(row=14 + j, column=15).value = (edu_sum / tot_sum) if tot_sum else None


# ─── Sheet 6 / 16 共用列定義（2026-10 公司日報格式）──────────────────────────────
# row -> (列名, 值)；值：str＝指標 key、(num, den)＝比率（num 可為 key 的 tuple＝加總）、None＝留空
# per_capita（人均產值）/ visitors（來客數）不在 calc_metrics，由填表時補進指標 dict
# 保固（AC+SA）＝ ACPP + SA Care
STORE_KPI_ROWS = {
    2:  ('總營業額(a + c + e)',      'total_rev'),
    3:  ('零售營業額(a)',            'retail_rev'),
    4:  ('訂單出貨金額(b)',          'order_ship'),
    5:  ('當日零售營業收入(a + b)',  'retail_income'),
    6:  ('零售訂金營業額(c)',        'deposit_rev'),
    7:  ('零售專案營業額(d)',        None),
    8:  ('Apple 主機營業額',         'rev_3001'),
    9:  ('Apple 配件營業額',         'rev_3002'),
    10: ('3PP配件營業額',            'rev_3003'),
    11: ('3PP搭售率',                ('rev_3003', 'total_rev')),
    12: ('保固營業額（AC+SA）',      'warranty_rev'),
    13: ('保固搭售率（AC+SA）',      ('warranty_rev', 'total_rev')),
    14: ('總毛利額',                 'total_gross'),
    15: ('Apple 毛利額',             'apl_gross'),
    16: ('3PP 毛利額',               'tpp_gross'),
    17: ('保固毛利額（AC+SA）',      'warranty_gross'),
    18: ('CPU 台數',                 'cpu_total'),
    19: ('保固 for CPU 套數（AC+SA）', 'warranty_cpu'),
    20: ('保固 for CPU 搭售率（AC+SA）', ('warranty_cpu', 'cpu_total')),
    21: ('iPad 台數',                'ipad'),
    22: ('保固 for iPad 套數（AC+SA）', 'warranty_ipad'),
    23: ('保固加總 iPad 搭售率（AC+SA）', ('warranty_ipad', 'ipad')),
    24: ('iPhone 台數',              'iphone'),
    25: ('保固 for iPhone 套數（AC+SA）', 'warranty_iphone'),
    26: ('保固加總 iPhone 搭售率（AC+SA）', ('warranty_iphone', 'iphone')),
    27: ('Watch 台數',               'watch'),
    28: ('保固 for Watch 套數（AC+SA）', 'warranty_watch'),
    29: ('保固加總 Watch 搭售率（AC+SA）', ('warranty_watch', 'watch')),
    30: ('AirPods 台數',             'airpods'),
    31: ('保固 for AirPods 套數（AC+SA）', 'warranty_airpods'),
    32: ('保固加總 AirPods 搭售率（AC+SA）', ('warranty_airpods', 'airpods')),
    33: ('人均產值',                 'per_capita'),
    34: ('來客數',                   'visitors'),
    35: ('成交筆數',                 'txn_count'),
    36: ('提袋率',                   ('txn_count', 'visitors')),
    37: ('舊換新台數:',              'tradein'),
}


def _kpi_value(m: dict, spec):
    """依 STORE_KPI_ROWS 的值規格取數；指標不存在（如沒有來客數）回 None。"""
    if spec is None:
        return None
    if isinstance(spec, str):
        return m.get(spec)
    num_keys, den_key = spec
    num_keys = num_keys if isinstance(num_keys, tuple) else (num_keys,)
    if den_key not in m or any(k not in m for k in num_keys):
        return None
    return safe_rate(sum(m[k] for k in num_keys), m[den_key])


def _apply_gross_bi(m: dict, bi: 'dict | None') -> dict:
    """用 ERP BI（02/03-日報 毛利額 含稅）覆寫 Apple/3PP/ACPP 毛利並重算總毛利；bi 為 None 時沿用引擎值。"""
    if bi:
        m.update({k: bi[k] for k in ('apl_gross', 'tpp_gross', 'acpp_gross') if k in bi})
        m['total_gross'] = m['apl_gross'] + m['tpp_gross'] + m['acpp_gross'] + m['sa_gross']
        m['warranty_gross'] = m['acpp_gross'] + m['sa_gross']
    return m


# ─── Sheet 6: 門市週報（fill_sheet2）────────────────────────────────────────────────────────
def fill_sheet2(ws, df_cur, df_prev, sacare_prices, dates: dict, traffic=None, emp_count=None,
                gross_bi=None):
    """gross_bi：{'上週'|'本週'|'本月'|'上月'|'去年': {apl_gross, tpp_gross, acpp_gross}}（server 查 ERP BI）"""
    print('  Sheet 6: 門市週報', flush=True)

    def get(df, start, end):
        return calc_metrics(period(df, start, end), sacare_prices)

    # Column: B=2(上週), C=3(本週), F=6(本月), I=9(上月同期), L=12(去年同期)
    cols = {'上週': 2, '本週': 3, '本月': 6, '上月': 9, '去年': 12}
    data = {
        '上週': get(df_cur,  dates['prev_wk_start'], dates['prev_wk_end']),
        '本週': get(df_cur,  dates['wk_start'],      dates['wk_end']),
        '本月': get(df_cur,  dates['mo_start'],      dates['mo_end']),
        '上月': get(df_cur,  dates['lm_start'],      dates['lm_end']),
        '去年': get(df_prev, dates['ly_start'],      dates['ly_end']),
    }
    for col, m in data.items():
        _apply_gross_bi(m, (gross_bi or {}).get(col))
        # 人均產值 = 零售營業額(a) / 總員工數（編制人數，前端輸入）；來客數＝ShopperTrak 人流（查不到則留空）
        if emp_count:
            m['per_capita'] = int(m.get('retail_rev', 0) / emp_count)
        visitors = (traffic or {}).get(col)
        if visitors is not None:
            m['visitors'] = int(visitors)

    has_traffic = bool(traffic) and any(v is not None for v in traffic.values())
    active_rows = []
    for r, (_, spec) in STORE_KPI_ROWS.items():
        if spec is None:                                   # 零售專案營業額(d) 留空
            continue
        if spec == 'per_capita' and not emp_count:
            continue
        if not has_traffic and 'visitors' in (spec if isinstance(spec, tuple) else (spec,)):
            continue                                       # 來客數、提袋率需有人流
        active_rows.append(r)
        is_pct = isinstance(spec, tuple)
        for col_name, col in cols.items():
            v = _kpi_value(data[col_name], spec)
            if is_pct:
                ws.cell(row=r, column=col).value = v if data[col_name].get(spec[1]) else None
            else:
                ws.cell(row=r, column=col).value = int(v) if v else None

    # ── 計算差異欄位 D/E/H/J/K/M/N (全部用公式重算) ──────────────────────────
    for r in active_rows:
        B = ws.cell(row=r, column=2).value
        C = ws.cell(row=r, column=3).value
        F = ws.cell(row=r, column=6).value
        G = ws.cell(row=r, column=7).value   # 本月目標 (user-filled)
        I = ws.cell(row=r, column=9).value
        L = ws.cell(row=r, column=12).value
        # D = C - B (週差異)
        if C is not None and B is not None:
            ws.cell(row=r, column=4).value = round(C - B, 6)
        else:
            ws.cell(row=r, column=4).value = None
        # E = D/B (週差異%)
        ws.cell(row=r, column=5).value = safe_rate((C or 0) - (B or 0), B)
        # H = F/G (達成率) — G might be a string ('-')
        G_num = G if isinstance(G, (int, float)) else None
        ws.cell(row=r, column=8).value = safe_rate(F, G_num)
        # J = F - I (上月同期差異)
        if F is not None and I is not None:
            ws.cell(row=r, column=10).value = round(F - I, 6)
        else:
            ws.cell(row=r, column=10).value = None
        # K = J/I (上月同期差異%)
        ws.cell(row=r, column=11).value = safe_rate((F or 0) - (I or 0), I)
        # M = F - L (去年同期差異)
        if F is not None and L is not None:
            ws.cell(row=r, column=13).value = round(F - L, 6)
        else:
            ws.cell(row=r, column=13).value = None
        # N = M/L (去年同期差異%)
        ws.cell(row=r, column=14).value = safe_rate((F or 0) - (L or 0), L)

# ─── Sheet 7: 3PP配件比較（fill_sheet3）──────────────────────────────────────────────────────
def fill_sheet3(ws, df_cur, df_prev, sacare_prices, dates: dict):
    print('  Sheet 7: 3PP配件比較', flush=True)
    sa_codes = set(sacare_prices.keys())

    def c4_rev(df, start, end, c4_code):
        d = period(df, start, end)
        d = d[~d['存貨代碼'].astype(str).str.strip().isin(sa_codes)]
        d = d[(d['類別3代碼'] == 3003.0) & (d['類別4代碼'] == c4_code)]
        return int(d['NET'].sum())

    def sa_rev_total(df, start, end):
        d = period(df, start, end)
        sa = d[d['存貨代碼'].astype(str).str.strip().isin(sa_codes)].copy()
        sa['SA_NET'] = sa['存貨代碼'].astype(str).str.strip().map(sacare_prices) * sa['數量'].fillna(0)
        return int(sa['SA_NET'].sum())

    # Columns: C=3(上週), D=4(本週), G=7(本月), H=8(上月同期), K=11(去年同期)
    #          N=14(2025 YTD), O=15(2026 YTD)
    periods = {
        'prev_wk': (df_cur,  dates['prev_wk_start'], dates['prev_wk_end']),
        'wk':      (df_cur,  dates['wk_start'],       dates['wk_end']),
        'mo':      (df_cur,  dates['mo_start'],        dates['mo_end']),
        'lm':      (df_cur,  dates['lm_start'],        dates['lm_end']),
        'ly':      (df_prev, dates['ly_start'],        dates['ly_end']),
        'ytd_cur': (df_cur,  dates['ytd_cur_start'],   dates['mo_end']),
        'ytd_prv': (df_prev, dates['ytd_prv_start'],   dates['ly_end']),
    }

    col_map = {'prev_wk': 3, 'wk': 4, 'mo': 7, 'lm': 8,
               'ly': 11, 'ytd_prv': 14, 'ytd_cur': 15}

    def write_diff_cols(ws, row):
        """Compute and write difference/percentage columns E/F/I/J/L/M/P/Q."""
        C = ws.cell(row=row, column=3).value or 0
        D = ws.cell(row=row, column=4).value or 0
        G = ws.cell(row=row, column=7).value or 0
        H = ws.cell(row=row, column=8).value or 0
        K = ws.cell(row=row, column=11).value or 0
        N = ws.cell(row=row, column=14).value or 0
        O = ws.cell(row=row, column=15).value or 0
        # E/F: 本週 vs 上週
        ws.cell(row=row, column=5).value  = (D - C) or None
        ws.cell(row=row, column=6).value  = safe_rate(D - C, C)
        # I/J: 本月 vs 上月同期
        ws.cell(row=row, column=9).value  = (G - H) or None
        ws.cell(row=row, column=10).value = safe_rate(G - H, H)
        # L/M: 本月 vs 去年同期
        ws.cell(row=row, column=12).value = (G - K) or None
        ws.cell(row=row, column=13).value = safe_rate(G - K, K)
        # P/Q: 今年YTD vs 去年YTD
        ws.cell(row=row, column=16).value = (O - N) or None
        ws.cell(row=row, column=17).value = safe_rate(O - N, N)

    for c4, excel_row in C4_ROWS.items():
        for pname, (df, s, e) in periods.items():
            v = c4_rev(df, s, e, c4)
            ws.cell(row=excel_row, column=col_map[pname]).value = v or None
        write_diff_cols(ws, excel_row)

    # SA CARE row (row 16)
    sa_row = 16
    for pname, (df, s, e) in periods.items():
        v = sa_rev_total(df, s, e)
        ws.cell(row=sa_row, column=col_map[pname]).value = v or None
    write_diff_cols(ws, sa_row)

    # Row 17 (加總): sum rows 2-16 (all C4 + SA Care)
    # Row 18 (純配件): sum rows 2-15 (C4 only, exclude SA Care row 16)
    for value_col in [3, 4, 7, 8, 11, 14, 15]:
        total_all  = sum(ws.cell(row=r, column=value_col).value or 0 for r in range(2, 17))
        total_pure = sum(ws.cell(row=r, column=value_col).value or 0 for r in range(2, 16))
        ws.cell(row=17, column=value_col).value = total_all  or None
        ws.cell(row=18, column=value_col).value = total_pure or None
    write_diff_cols(ws, 17)
    write_diff_cols(ws, 18)

# ─── Sheet 8/9: 銷售排名（fill_sheet45）──────────────────────────────────────────────────────
def fill_sheet45(ws4, ws5, df_cur, sacare_prices, dates: dict):
    print('  Sheet 8/9: 銷售排名', flush=True)
    sa_codes = set(sacare_prices.keys())

    def get_ranking(df, start, end, vap_only=False):
        d = period(df, start, end)
        d = d[~d['存貨代碼'].astype(str).str.strip().isin(sa_codes)]
        d = d[d['類別3代碼'] == 3003.0]
        # Exclude gift (折扣 >= 80 or 金額 == 0)
        is_return = d['交易類型'] == '銷退'
        is_gift = (d.get('折扣', pd.Series(0, index=d.index)).fillna(0) >= 80) | \
                  ((d['NET'].abs() == 0) & ~is_return)
        d = d[~is_gift]
        if vap_only:
            d = d[d['品牌代碼'].isin(VAP_BRANDS)]
        else:
            d = d[~d['品牌代碼'].isin(VAP_BRANDS)]

        grouped = d.groupby('存貨代碼').agg(
            品名=('名稱', 'first'),
            qty=('數量', 'sum'),
            rev=('NET', 'sum')
        ).reset_index()
        return grouped.sort_values('qty', ascending=False).head(10)

    def write_ranking(ws, wk_rank, mo_rank):
        for i in range(10):
            row = i + 2
            # 週銷 (cols B=2, C=3, D=4)
            if i < len(wk_rank):
                r = wk_rank.iloc[i]
                ws.cell(row=row, column=2).value = str(r['存貨代碼'])
                ws.cell(row=row, column=3).value = str(r['品名'])
                ws.cell(row=row, column=4).value = int(r['qty'])
            else:
                for c in [2, 3, 4]:
                    ws.cell(row=row, column=c).value = None
            # 月銷 (cols E=5, F=6, G=7, H=8)
            if i < len(mo_rank):
                r = mo_rank.iloc[i]
                ws.cell(row=row, column=5).value = str(r['存貨代碼'])
                ws.cell(row=row, column=6).value = str(r['品名'])
                ws.cell(row=row, column=7).value = int(r['qty'])
                ws.cell(row=row, column=8).value = int(r['rev'])
            else:
                for c in [5, 6, 7, 8]:
                    ws.cell(row=row, column=c).value = None

    wk_3pp = get_ranking(df_cur, dates['wk_start'], dates['wk_end'], vap_only=False)
    mo_3pp = get_ranking(df_cur, dates['mo_start'], dates['mo_end'], vap_only=False)
    write_ranking(ws4, wk_3pp, mo_3pp)

    wk_vap = get_ranking(df_cur, dates['wk_start'], dates['wk_end'], vap_only=True)
    mo_vap = get_ranking(df_cur, dates['mo_start'], dates['mo_end'], vap_only=True)
    write_ranking(ws5, wk_vap, mo_vap)

# ─── Per-employee calculations ─────────────────────────────────────────────────
def calc_employee(df: pd.DataFrame, emp_code: str, sacare_prices: dict) -> dict:
    """Calculate all per-employee metrics for a filtered period."""
    d = df[df['員工代碼'] == emp_code].copy()
    sa_codes = set(sacare_prices.keys())
    non_sa = d[~d['存貨代碼'].astype(str).str.strip().isin(sa_codes)]

    # Revenue (excl SA)
    rev_excl_sa = int(non_sa['NET'].sum())
    rev_3003    = int(non_sa.loc[non_sa['類別3代碼'] == 3003.0, 'NET'].sum())
    rev_3002    = int(non_sa.loc[non_sa['類別3代碼'] == 3002.0, 'NET'].sum())  # Apple 原廠配件
    # D 欄：原廠商品營業額（不含 SAcare，尾款才紀錄，訂金不計）
    # 排除：C1∈{1002,1004,1008} / C3∈{3047,3003,3004,3018,3019,3012}
    #        C6∈{6888,6889} / 指定存貨代碼
    non_sa_d = non_sa[non_sa['交易類型'].isin(SALE_TYPES | {'銷退'})]
    _c1  = non_sa_d.get('類別1代碼', pd.Series(dtype=float, index=non_sa_d.index)).fillna(0)
    _c6  = non_sa_d.get('類別6代碼', pd.Series(dtype=float, index=non_sa_d.index)).fillna(0)
    _sku = non_sa_d['存貨代碼'].astype(str).str.strip()
    apple_mask = (
        ~_c1.isin(C1_EXCLUDED_FROM_APPLE) &
        ~non_sa_d['類別3代碼'].isin(C3_EXCLUDED_FROM_APPLE) &
        ~_c6.isin(C6_EXCLUDED_FROM_APPLE) &
        ~_sku.isin(SKU_EXCLUDED_FROM_APPLE)
    )
    rev_apple = int(non_sa_d.loc[apple_mask, 'NET'].sum())

    # SAcare：只計 SALE_TYPES（銷售/尾款），排除訂金避免雙重計算
    sa_rows = d[d['存貨代碼'].astype(str).str.strip().isin(sa_codes)].copy()
    sa_sold = sa_rows[sa_rows['交易類型'].isin(SALE_TYPES)]
    sa_ret  = sa_rows[sa_rows['交易類型'] == '銷退']
    sa_sold_net = (sa_sold['存貨代碼'].astype(str).str.strip().map(sacare_prices) * sa_sold['數量'].fillna(0)).sum()
    sa_ret_net  = (sa_ret['存貨代碼'].astype(str).str.strip().map(sacare_prices) * sa_ret['數量'].abs()).sum()
    sa_rev = int(sa_sold_net - sa_ret_net)

    # Gross profit (Sheet 12 H/I 未稅毛利)
    # H 原廠毛利：依 ERP「13-門市獎金Apple毛利額未稅-員工」
    #   C3 白名單 {3001,3002,3032,3033,3046} + 99901689（抵用券兌換, C3=NaN）
    #   交易類型：包含 銷售/訂金/銷退，排除 尾款/退訂
    #   原因：800AB 訂金列記錄原廠商品全額未稅收入；尾款列收入=0
    # I 3PP毛利：依 ERP「14-門市獎金3PP毛利額未稅-員工」
    #   C3 白名單 {3003,3006,3004,3018,3019,3012}
    #   交易類型：排除 尾款（訂金/退訂/銷退 皆納入）
    #   原因：ERP 以訂金發生時點認列 3PP 收入與成本；尾款列為 3PP 贈品成本補登列
    #         (NET=0，cost>0)，ERP 未計入；退訂會自動與原訂金相抵。
    #   v1.1.15 修正：改為只排除 尾款（原先錯誤地排除 訂金/退訂）
    # 公式：淨銷售金額(未稅) - 單位成本 × 數量（不加銷退金額(未稅)，不乘/除 1.05）
    SA_CARE_GROSS_EXCL = {
        '99903303','99903302','99200168','99500006','99900946','99900947','99900948',
        '99900949','99900950','99901684','99901685','99902607','99902608','99902609',
        '99902610','99903343','99903339',
        '99200202',  # 教育價活動（內部代碼，非實際商品，已排除於 D 欄）
        '99200201',  # 促銷組合（內部代碼）
        '99500203',  # v1.1.16: 預收訂金（會計科目，非實際商品，未來由商品銷售/尾款列認列）
    }
    # Apple-specific additional exclusions (800AB format, leading zeros removed)
    SA_APPLE_GROSS_EXCL = SA_CARE_GROSS_EXCL | {
        '7307154','7309136','7309137','7310037','7310042',
        '7310053','7310093','7311242',
        '88600895','88601027','90501795','90501799',
    }

    sku_str = d['存貨代碼'].astype(str).str.strip()
    c6_excl = d.get('類別6代碼', pd.Series(dtype=float, index=d.index)).isin([6888.0, 6889.0])
    c1_excl = d.get('類別1代碼', pd.Series(dtype=float, index=d.index)).isin([1002.0, 1004.0, 1008.0])
    excl_apl = sku_str.isin(SA_APPLE_GROSS_EXCL) | c6_excl | c1_excl
    excl_tpp = sku_str.isin(SA_CARE_GROSS_EXCL)  | c6_excl | c1_excl

    apl_mask = d['類別3代碼'].isin([3001.0, 3002.0, 3032.0, 3033.0, 3046.0]) | (sku_str == '99901689')

    def gross_apl():
        sub = d[apl_mask & ~excl_apl & ~d['交易類型'].isin(['尾款', '退訂'])]
        net_ux = sub.get('淨銷售金額(未稅)', pd.Series(0, index=sub.index)).fillna(0)
        # C3=3032 (ACPP 代收保費): 800AB 將含稅金額存入 淨銷售金額(未稅) 欄位，
        # 而 ERP 銷貨收入(未稅) 使用含稅÷1.05 的真正未稅金額，需修正。
        is_acpp = sub['類別3代碼'] == 3032.0
        net_ux = net_ux.where(~is_acpp, net_ux / 1.05)
        cost   = sub.get('單位成本', pd.Series(0, index=sub.index)).fillna(0) * sub['數量'].fillna(0)
        return int((net_ux - cost).sum())

    tpp_mask = d['類別3代碼'].isin([3003.0, 3006.0, 3004.0, 3018.0, 3019.0, 3012.0])

    def gross_tpp():
        sub = d[tpp_mask & ~excl_tpp & ~d['交易類型'].isin(['尾款'])]
        net_ux = sub.get('淨銷售金額(未稅)', pd.Series(0, index=sub.index)).fillna(0)
        cost   = sub.get('單位成本', pd.Series(0, index=sub.index)).fillna(0) * sub['數量'].fillna(0)
        return int((net_ux - cost).sum())

    apl_gross = gross_apl()
    tpp_gross = gross_tpp()

    # 認證機品牌（無論 C3 為何，只要 C6/C4 符合就計入）
    _cert  = {881.0, 885.0, 886.0, 888.0}
    _brand = d.get('品牌代碼', pd.Series(dtype=float, index=d.index)).fillna(0)

    # Unit counts
    # Mac (C6): need to cover all Mac C6 codes across model years
    def net_units_c6(c6_set, c3=None):
        m = d['類別6代碼'].isin(c6_set)
        if c3:
            m &= (d['類別3代碼'] == c3) | _brand.isin(_cert)
        sale = d.loc[m & d['交易類型'].isin(SALE_TYPES), '數量'].sum()
        ret  = d.loc[m & (d['交易類型'] == '銷退'),  '數量'].abs().sum()
        return int(sale - ret)

    # iPhone / Watch / iPad (C4): stable across model years + 認證機品牌 bypass
    def net_units_c4(c4_set, c3=3001.0):
        m = ((d['類別3代碼'] == c3) | _brand.isin(_cert)) & d['類別4代碼'].isin(c4_set)
        sale = d.loc[m & d['交易類型'].isin(SALE_TYPES), '數量'].sum()
        ret  = d.loc[m & (d['交易類型'] == '銷退'),  '數量'].abs().sum()
        return int(sale - ret)

    cpu     = net_units_c6(C6_CPU, c3=3001.0)   # 引用 C6_CPU 常數，新世代只需更新 C6_CPU
    iphone  = net_units_c4(C4_IPHONE)
    ipad    = net_units_c4(C4_IPAD)
    watch   = net_units_c4(C4_WATCH)
    airpods = net_units_c6(C6_AIRPODS, c3=3002.0)

    # ACPP+ = 類別3=3032 (AppleCare+ 代收保費), grouped by device type via product name
    ac_rows = d[(d['類別3代碼'] == 3032.0) & d['交易類型'].isin(SALE_TYPES)].copy()
    ac_ret  = d[(d['類別3代碼'] == 3032.0) & (d['交易類型'] == '銷退')].copy()
    def acpp_name(keyword):
        kw = keyword.lower()
        sold = ac_rows[ac_rows['名稱'].astype(str).str.lower().str.contains(kw, na=False)]['數量'].sum()
        retn = ac_ret[ ac_ret['名稱'].astype(str).str.lower().str.contains(kw, na=False)]['數量'].abs().sum()
        return int(sold - retn)

    mac_acpp     = acpp_name('mac')
    ipad_acpp    = acpp_name('ipad')
    iphone_acpp  = acpp_name('iphone')
    watch_acpp   = acpp_name('watch')
    airpods_acpp = acpp_name('airpods')

    # SAcare counts by device
    def sa_cnt(c6_set):
        m = d['類別6代碼'].isin(c6_set) & d['存貨代碼'].astype(str).str.strip().isin(sa_codes)
        return int(d.loc[m & d['交易類型'].isin(SALE_TYPES), '數量'].sum() -
                   d.loc[m & (d['交易類型'] == '銷退'),  '數量'].abs().sum())

    mac_sa    = sa_cnt(C6_SA['cpu'])
    ipad_sa   = sa_cnt(C6_SA['ipad'])
    iphone_sa = sa_cnt(C6_SA['iphone'])
    watch_sa  = sa_cnt(C6_SA['watch'])
    airpods_sa = sa_cnt(C6_SA['airpods'])

    # C4 3PP accessory revenues
    def acc_rev(c4_set):
        m = d['類別4代碼'].isin(c4_set) & (d['類別3代碼'] == 3003.0)
        return int(d.loc[m, 'NET'].sum())

    vap_rev    = int(d.loc[d['品牌代碼'].isin(VAP_BRANDS), 'NET'].sum())
    # 服務費（資料轉移/重置系統）含稅金額，銷退列的 NET 為負值故直接加總即為淨額
    service_fee = int(d.loc[d['存貨代碼'].astype(str).str.strip().isin(SERVICE_FEE_SKUS), 'NET'].sum())
    _office_mask = d['名稱'].str.lower().str.contains('office|microsoft', na=False)
    office_qty = int(d.loc[_office_mask & d['交易類型'].isin(SALE_TYPES), '數量'].sum()) - \
                 int(d.loc[_office_mask & (d['交易類型'] == '銷退'), '數量'].abs().sum())

    return dict(
        rev_excl_sa=rev_excl_sa, rev_3003=rev_3003, rev_3002=rev_3002, rev_apple=rev_apple,
        sa_rev=sa_rev, apl_gross=apl_gross, tpp_gross=tpp_gross,
        cpu=cpu, iphone=iphone, ipad=ipad, watch=watch, airpods=airpods,
        mac_acpp=mac_acpp, ipad_acpp=ipad_acpp,
        iphone_acpp=iphone_acpp, watch_acpp=watch_acpp, airpods_acpp=airpods_acpp,
        mac_sa=mac_sa, ipad_sa=ipad_sa, iphone_sa=iphone_sa,
        watch_sa=watch_sa, airpods_sa=airpods_sa,
        cpu_acc=acc_rev(C4_CPU_SET), iphone_acc=acc_rev(C4_IPHONE_SET),
        ipad_acc=acc_rev(C4_IPAD_SET), watch_acc=acc_rev(C4_WATCH_SET),
        airpods_acc=acc_rev(C4_AIRPODS_SET),
        service_fee=service_fee,
        cpu_ios_acc=acc_rev({4012.0}),   # CPU/iOS通用週邊
        ios_acc=acc_rev({4022.0}),        # iOS通用週邊配件
        vap_rev=vap_rev, office_qty=office_qty,
    )

# ─── Sheet 12: 個人新制獎金（fill_sheet6）────────────────────────────────────────────────────
def fill_sheet6(ws, df_cur, sacare_prices, dates: dict):
    print('  Sheet 12: 個人新制獎金', flush=True)
    d_mo = period(df_cur, dates['mo_start'], dates['mo_end'])
    # 排除等級代碼 05（非全職／兼職員工，不納入獎金計算）
    if '等級代碼' in d_mo.columns:
        d_mo = d_mo[d_mo['等級代碼'] != '05']

    for i, (code, _) in enumerate(EMPLOYEES):
        row = i + 2
        m = calc_employee(d_mo, code, sacare_prices)
        ws.cell(row=row, column=2).value  = m['rev_excl_sa'] or None   # B 營業額
        ws.cell(row=row, column=3).value  = m['rev_3003'] or None       # C 3PP 營業額
        ws.cell(row=row, column=4).value  = m['rev_apple'] or None       # D 原廠商品營業額（排除3003/3004/3012/3018/3019/3047）
        ws.cell(row=row, column=5).value  = m['sa_rev'] or None         # E SA Care
        # 未稅毛利由 calc_employee 直接依 ERP 公式計算（不需 /1.05）
        apl_gross_ex = m['apl_gross'] or None
        tpp_gross_ex = m['tpp_gross'] or None
        ws.cell(row=row, column=8).value  = apl_gross_ex                # H 原廠毛利(未稅)
        ws.cell(row=row, column=9).value  = tpp_gross_ex                # I 3PP 毛利(未稅)
        # F, G, J, K = Excel formulas — skip

    # 加總列：動態掃描 A 欄找「加總」文字，避免因模板結構差異（如有無總公司列）而寫錯列
    emp_end    = 1 + len(EMPLOYEES)          # 最後一位員工的列號
    emp_rows_6 = range(2, emp_end + 1)
    # 從員工結束列之後向下找「加總」
    total_row = None
    for r in range(emp_end + 1, emp_end + 10):
        cell_val = ws.cell(row=r, column=1).value
        if cell_val is not None and '加總' in str(cell_val):
            total_row = r
            break
    if total_row is None:
        total_row = emp_end + 1              # fallback：緊接在最後員工後面
    for col in [2, 3, 4, 5, 8, 9]:
        total = sum(ws.cell(row=r, column=col).value or 0 for r in emp_rows_6)
        ws.cell(row=total_row, column=col).value = total or None

# ─── Sheet 13/14: 個人週/月主機（fill_sheet78）─────────────────────────────────────────────────
def fill_sheet78(ws7, ws8, df_cur, sacare_prices, dates: dict):
    print('  Sheet 13/14: 個人週/月主機', flush=True)
    d_wk = period(df_cur, dates['wk_start'], dates['wk_end'])
    d_mo = period(df_cur, dates['mo_start'], dates['mo_end'])

    def write_emp_row(ws, row, m):
        cpu, ipad, iphone, watch, airpods = m['cpu'], m['ipad'], m['iphone'], m['watch'], m['airpods']
        mac_acpp, mac_sa = m['mac_acpp'], m['mac_sa']
        ipad_acpp, ipad_sa = m['ipad_acpp'], m['ipad_sa']
        iphone_acpp, iphone_sa = m['iphone_acpp'], m['iphone_sa']
        watch_acpp, watch_sa = m['watch_acpp'], m['watch_sa']
        airpods_acpp, airpods_sa = m['airpods_acpp'], m['airpods_sa']

        mac_rate    = safe_rate(mac_acpp + mac_sa, cpu)
        ipad_rate   = safe_rate(ipad_acpp + ipad_sa, ipad)
        iphone_rate = safe_rate(iphone_acpp + iphone_sa, iphone)
        ios_rate    = safe_rate(iphone_acpp + iphone_sa + ipad_acpp + ipad_sa, iphone + ipad)
        watch_rate  = safe_rate(watch_acpp + watch_sa, watch)
        ap_rate     = safe_rate(airpods_acpp + airpods_sa, airpods)

        vals = [
            cpu or None, mac_acpp or None, mac_sa or None, mac_rate,
            ipad or None, ipad_acpp or None, ipad_sa or None, ipad_rate,
            iphone or None, iphone_acpp or None, iphone_sa or None, iphone_rate,
            ios_rate,
            watch or None, watch_acpp or None, watch_sa or None, watch_rate,
            airpods or None, airpods_acpp or None, airpods_sa or None, ap_rate,
        ]
        for col_offset, v in enumerate(vals):
            ws.cell(row=row, column=col_offset + 2).value = v
        # Y(25) 服務費金額（含稅，銷退已由 NET 帶負號扣回）；W/X 為 MySetup，由 server 依上傳的 Personal Setup Excel 填入
        ws.cell(row=row, column=25).value = m['service_fee'] or None

    for i, (code, _) in enumerate(EMPLOYEES):
        row = i + 3  # data starts at row 3
        write_emp_row(ws7, row, calc_employee(d_wk, code, sacare_prices))
        write_emp_row(ws8, row, calc_employee(d_mo, code, sacare_prices))

    # ── 加總列動態：起始 row 3，加總列 = 3 + 員工人數 ────────────────────────
    def write_total_row(ws):
        total_row = 3 + len(EMPLOYEES)
        emp_rows  = range(3, total_row)
        # Count columns: B(2),C(3),D(4), F(6),G(7),H(8), J(10),K(11),L(12), O(15),P(16),Q(17), S(19),T(20),U(21)
        count_cols = [2,3,4, 6,7,8, 10,11,12, 15,16,17, 19,20,21, 25]
        totals = {}
        for col in count_cols:
            totals[col] = sum(ws.cell(row=r, column=col).value or 0 for r in emp_rows)
            ws.cell(row=total_row, column=col).value = totals[col] or None
        # Rate columns recomputed from totals
        # E(5): Mac rate = (C+D)/B
        ws.cell(row=total_row, column=5).value  = safe_rate(totals[3]+totals[4], totals[2])
        # I(9): iPad rate = (G+H)/F
        ws.cell(row=total_row, column=9).value  = safe_rate(totals[7]+totals[8], totals[6])
        # M(13): iPhone rate = (K+L)/J
        ws.cell(row=total_row, column=13).value = safe_rate(totals[11]+totals[12], totals[10])
        # N(14): iOS rate = (K+L+G+H)/(J+F)
        ws.cell(row=total_row, column=14).value = safe_rate(totals[11]+totals[12]+totals[7]+totals[8], totals[10]+totals[6])
        # R(18): Watch rate = (P+Q)/O
        ws.cell(row=total_row, column=18).value = safe_rate(totals[16]+totals[17], totals[15])
        # V(22): AirPods rate = (T+U)/S
        ws.cell(row=total_row, column=22).value = safe_rate(totals[20]+totals[21], totals[19])

    write_total_row(ws7)
    write_total_row(ws8)

# ─── Sheet 15: 個人月3PP（fill_sheet9）───────────────────────────────────────────────────────
def fill_sheet9(ws, df_cur, sacare_prices, dates: dict):
    print('  Sheet 15: 個人月3PP', flush=True)
    d_mo = period(df_cur, dates['mo_start'], dates['mo_end'])

    for i, (code, _) in enumerate(EMPLOYEES):
        row = i + 2
        m = calc_employee(d_mo, code, sacare_prices)
        ws.cell(row=row, column=2).value  = m['cpu_acc']     or None  # B CPU配件
        ws.cell(row=row, column=3).value  = m['iphone_acc']  or None  # C iPhone配件
        ws.cell(row=row, column=4).value  = m['ipad_acc']    or None  # D iPad配件
        ws.cell(row=row, column=5).value  = m['cpu_ios_acc'] or None  # E CPU/iOS通用週邊
        ws.cell(row=row, column=6).value  = m['ios_acc']     or None  # F iOS通用週邊配件
        ws.cell(row=row, column=7).value  = m['watch_acc']   or None  # G Watch配件
        ws.cell(row=row, column=8).value  = m['airpods_acc'] or None  # H AirPods配件
        ws.cell(row=row, column=9).value  = m['vap_rev']     or None  # I VAP
        ws.cell(row=row, column=10).value = m['office_qty']  or None  # J Office數量

    # 加總列動態：起始 row 2，加總列 = 2 + 員工人數
    total_row  = 2 + len(EMPLOYEES)
    emp_rows_9 = range(2, total_row)
    for col in range(2, 11):   # B～J
        total = sum(ws.cell(row=r, column=col).value or 0 for r in emp_rows_9)
        ws.cell(row=total_row, column=col).value = total or None

# ─── Sheet 3/4: 每月重點 / Speakers ───────────────────────────────────────────
def _net_qty(d: pd.DataFrame, mask) -> int:
    """淨數量：銷售/尾款 正計、銷退 取絕對值扣回（同 calc_employee 的 net_units_* 慣例）。"""
    sale = d.loc[mask & d['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = d.loc[mask & (d['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)

def _focus_mask(d: pd.DataFrame):
    return d['品牌代碼'].isin(MONTHLY_FOCUS_BRANDS)

def _speakers_mask(d: pd.DataFrame):
    return (d['類別3代碼'] == 3003.0) & d['類別4代碼'].isin(SPEAKERS_C4)

def _label_row(ws, label: str, start: int = 2):
    """在 A 欄由 start 列往下找標籤所在列（員工列數為動態，不能寫死列號）。"""
    for r in range(start, ws.max_row + 1):
        v = ws.cell(row=r, column=1).value
        if v is not None and str(v).strip() == label:
            return r
    return None

def _copy_row_style(ws, src_row: int, dst_row: int, ncol: int):
    from copy import copy
    for c in range(1, ncol + 1):
        ws.cell(row=dst_row, column=c)._style = copy(ws.cell(row=src_row, column=c)._style)


def _clear_below(ws, row: int):
    """刪掉 row 以下的列（第 3/4 頁表格下方的附加區塊每次重寫；員工列重建時會把它們當尾端列搬動）。"""
    if ws.max_row > row:
        ws.delete_rows(row + 1, ws.max_row - row)


def fill_sheet_focus(ws, df_cur, dates: dict):
    """第 3 頁 每月重點：C=週銷售數、D=月銷售數（品牌 MONTHLY_FOCUS_BRANDS 的淨數量）。
    表格下方標示統計項目（MONTHLY_FOCUS_LABEL）並列出本月實際計入的商品。"""
    print('  Sheet 3: 每月重點', flush=True)
    d_wk = period(df_cur, dates['wk_start'], dates['wk_end'])
    d_mo = period(df_cur, dates['mo_start'], dates['mo_end'])
    total_row = _label_row(ws, 'Total') or (2 + len(EMPLOYEES))
    _clear_below(ws, total_row)

    for i, (code, _) in enumerate(EMPLOYEES):
        row = i + 2
        for col, d in ((3, d_wk), (4, d_mo)):
            e = d[d['員工代碼'] == code]
            ws.cell(row=row, column=col).value = _net_qty(e, _focus_mask(e)) or None

    for col in (3, 4):
        total = sum(ws.cell(row=r, column=col).value or 0 for r in range(2, total_row))
        ws.cell(row=total_row, column=col).value = total or None

    # 統計項目說明＋本月計入商品（全店淨數量，依數量排序）
    from openpyxl.styles import Font
    bold = Font(name='Arial', size=12, bold=True)
    plain = Font(name='Arial', size=12)
    r = total_row + 2
    ws.cell(row=r, column=1).value = f'統計項目：{MONTHLY_FOCUS_LABEL}'
    ws.cell(row=r, column=1).font = bold
    r += 1
    ws.cell(row=r, column=1).value = f"本月計入商品（{dates['mo_start']:%m/%d}~{dates['mo_end']:%m/%d}）"
    ws.cell(row=r, column=4).value = '月銷售數'
    for c in (1, 4):
        ws.cell(row=r, column=c).font = bold
    items = d_mo[_focus_mask(d_mo)]
    qty = {name: _net_qty(g, _focus_mask(g)) for name, g in items.groupby('名稱')}
    for name, q in sorted(qty.items(), key=lambda kv: -kv[1]):
        if not q:
            continue
        r += 1
        ws.cell(row=r, column=1).value = str(name)
        ws.cell(row=r, column=4).value = q
        ws.cell(row=r, column=1).font = ws.cell(row=r, column=4).font = plain

def fill_sheet_speakers(ws, df_cur, dates: dict):
    """第 4 頁 Speakers（3PP 藍牙喇叭 C3=3003 且 C4∈SPEAKERS_C4）：
    C/D=本週銷售數/金額、E/F=月銷售數/金額，另加一格全店年累積金額。"""
    print('  Sheet 4: Speakers', flush=True)
    d_wk = period(df_cur, dates['wk_start'], dates['wk_end'])
    d_mo = period(df_cur, dates['mo_start'], dates['mo_end'])
    total_row = _label_row(ws, 'Total') or (2 + len(EMPLOYEES))

    for i, (code, _) in enumerate(EMPLOYEES):
        row = i + 2
        for (qty_col, amt_col), d in (((3, 4), d_wk), ((5, 6), d_mo)):
            e = d[d['員工代碼'] == code]
            m = _speakers_mask(e)
            ws.cell(row=row, column=qty_col).value = _net_qty(e, m) or None
            ws.cell(row=row, column=amt_col).value = int(e.loc[m, 'NET'].sum()) or None

    for col in range(3, 7):
        total = sum(ws.cell(row=r, column=col).value or 0 for r in range(2, total_row))
        ws.cell(row=total_row, column=col).value = total or None

    # 年累積金額（全店，不限報表列出的員工）：今年 1/1 ～ 截止日。
    # 截止日＝前端「年對年截止日」有填就用它（與第 16/17 頁同步），留空則用本月結束日。
    ytd_end = dates.get('ytd_end') or dates['mo_end']
    d_ytd = period(df_cur, date(ytd_end.year, 1, 1), ytd_end)
    ytd_row = _label_row(ws, '年累積金額', start=total_row + 1)
    if ytd_row:
        amt = int(d_ytd.loc[_speakers_mask(d_ytd), 'NET'].sum())
        ws.cell(row=ytd_row, column=2).value = amt or None

    # 排除 Momax 區塊：A=員工編號、B=姓名、C=週銷數量、D=月數量、E=週銷金額、F=月金額
    base = ytd_row or total_row
    _clear_below(ws, base)
    title_row = base + 2
    ws.cell(row=title_row, column=1).value = 'Speakers（排除 Momax）'
    from openpyxl.styles import Font
    ws.cell(row=title_row, column=1).font = Font(name='Arial', size=12, bold=True)
    head_row = title_row + 1
    _copy_row_style(ws, 1, head_row, 6)
    for c, h in enumerate(('員工編號', '姓名', '週銷數量', '月數量', '週銷金額', '月金額'), start=1):
        ws.cell(row=head_row, column=c).value = h

    def excl_mask(d):
        return _speakers_mask(d) & ~d['品牌代碼'].isin(SPEAKERS_EXCL_BRANDS)

    first = head_row + 1
    for i, (code, display) in enumerate(EMPLOYEES):
        row = first + i
        _copy_row_style(ws, 2, row, 6)
        # 顯示名為 "SA1092 Sam" 時拆出姓名（同 server.rebuild_two_col_employee_sheet）
        name = display[len(code):].strip() if display.startswith(code) else display
        ws.cell(row=row, column=1).value = code
        ws.cell(row=row, column=2).value = name or None
        ew, em = d_wk[d_wk['員工代碼'] == code], d_mo[d_mo['員工代碼'] == code]
        ws.cell(row=row, column=3).value = _net_qty(ew, excl_mask(ew)) or None
        ws.cell(row=row, column=4).value = _net_qty(em, excl_mask(em)) or None
        ws.cell(row=row, column=5).value = int(ew.loc[excl_mask(ew), 'NET'].sum()) or None
        ws.cell(row=row, column=6).value = int(em.loc[excl_mask(em), 'NET'].sum()) or None
    sum_row = first + len(EMPLOYEES)
    _copy_row_style(ws, total_row, sum_row, 6)
    ws.cell(row=sum_row, column=1).value = 'Total'
    for col in range(3, 7):
        total = sum(ws.cell(row=r, column=col).value or 0 for r in range(first, sum_row))
        ws.cell(row=sum_row, column=col).value = total or None

# ─── Sheet 5: 門市人流（ShopperTrak 逐日來客數）─────────────────────────────
def _traffic_sum(daily: dict, start: date, end: date) -> int:
    if not daily or end < start:
        return 0
    return int(sum(v for d, v in daily.items() if start <= d <= end))


def _traffic_week_row(daily: dict, start: date, end: date) -> list:
    """回傳 [週日, 週一, …, 週六] 七格值（該區間沒涵蓋到的星期為 None）。"""
    vals = [None] * 7
    d = start
    while d <= end:
        idx = (d.weekday() + 1) % 7          # 週日 = 0
        vals[idx] = (vals[idx] or 0) + int(daily.get(d, 0))
        d += timedelta(days=1)
    return vals


def fill_sheet_traffic(ws, daily: dict, dates: dict):
    """第 5 頁 門市人流：
    2~5 列＝上週／本週逐日（B=週日…H=週六）＋ TOTAL 與 WoW 差異人數／比例；
    7~8 列＝當年各月來客數與年度總計；
    11~12 列＝上月 vs 本月、去年同月 vs 本月、去年 vs 今年年度累積。
    daily＝{date: 來客數}；查不到（未設帳密／查詢失敗）時只寫標題、數值全部留空。"""
    print('  Sheet 5: 門市人流', flush=True)
    wk_s, wk_e = dates['wk_start'], dates['wk_end']
    year = wk_s.year

    # 標題會隨報表年月變動，沒有人流資料時也要寫
    ws.cell(row=7, column=1).value = year
    mo_month = dates['mo_start'].month
    lm_month = dates['lm_start'].month
    for i, label in enumerate([
        f"{dates['lm_start'].year}\n{lm_month}月累積", f'{year}\n{mo_month}月累積', '差異',
        f"{dates['ly_start'].year}\n{mo_month}月同期", f'{year}\n{mo_month}月累積', '差異',
        f'{year - 1}\n年度累積', f'{year}\n年度累積', '差異',
    ]):
        ws.cell(row=11, column=i + 1).value = label

    if not daily:
        return

    def rate(cur, base):
        return (cur - base) / base if base else None

    # ── 2~5 列：上週／本週逐日、WoW 差異人數與比例（比例一律除以上週）──
    prev = _traffic_week_row(daily, dates['prev_wk_start'], dates['prev_wk_end'])
    curr = _traffic_week_row(daily, wk_s, wk_e)
    prev_total = _traffic_sum(daily, dates['prev_wk_start'], dates['prev_wk_end'])
    curr_total = _traffic_sum(daily, wk_s, wk_e)
    for i in range(7):
        col = i + 2
        ws.cell(row=2, column=col).value = prev[i]
        ws.cell(row=3, column=col).value = curr[i]
        if prev[i] is None and curr[i] is None:
            continue
        ws.cell(row=4, column=col).value = (curr[i] or 0) - (prev[i] or 0)
        ws.cell(row=5, column=col).value = rate(curr[i] or 0, prev[i] or 0)
    ws.cell(row=2, column=9).value = prev_total or None
    ws.cell(row=3, column=9).value = curr_total or None
    ws.cell(row=4, column=9).value = curr_total - prev_total
    ws.cell(row=5, column=9).value = rate(curr_total, prev_total)

    # ── 7~8 列：當年各月來客數（未到的月份留空）與年度總計 ──
    for m in range(1, 13):
        m_start = date(year, m, 1)
        if m_start > wk_e:
            continue
        m_end = date(year, m, monthrange(year, m)[1])
        ws.cell(row=8, column=m + 1).value = _traffic_sum(daily, m_start, m_end) or None
    ws.cell(row=8, column=14).value = _traffic_sum(daily, date(year, 1, 1), wk_e) or None

    # ── 11~12 列：上月／本月／去年同月／年度累積（年度累積固定算到本週結束日）──
    _, ytd_cur_end, ytd_prv_start, ytd_prv_end = _yoy_periods({'wk_end': wk_e})
    lm = _traffic_sum(daily, dates['lm_start'], dates['lm_end'])
    mo = _traffic_sum(daily, dates['mo_start'], dates['mo_end'])
    ly = _traffic_sum(daily, dates['ly_start'], dates['ly_end'])
    ytd_cur = _traffic_sum(daily, date(year, 1, 1), ytd_cur_end)
    ytd_prv = _traffic_sum(daily, ytd_prv_start, ytd_prv_end)
    for col, value in ((1, lm), (2, mo), (3, rate(mo, lm)),
                       (4, ly), (5, mo), (6, rate(mo, ly)),
                       (7, ytd_prv), (8, ytd_cur), (9, rate(ytd_cur, ytd_prv))):
        ws.cell(row=12, column=col).value = value if value else None


# ─── Sheet 16/17 共用：YOY 累積區間 ───────────────────────────────────────────
def _yoy_periods(dates: dict) -> 'tuple[date, date, date, date]':
    """回傳 (今年起始, 今年截止, 去年起始, 去年截止)。
    截止日預設為本週末日；若 dates 內有 'yoy_end'（前端自訂年對年截止日）則優先採用。
    今年 = 截止日當年 1/1 ～ 截止日；去年 = 前一年 1/1 ～ 去年同月同日
    （2/29 等不存在時取當月最後一天）。"""
    end = dates.get('yoy_end') or dates['wk_end']
    cur_year = end.year
    prev_year = cur_year - 1
    last = monthrange(prev_year, end.month)[1]
    prv_e = date(prev_year, end.month, min(end.day, last))
    return date(cur_year, 1, 1), end, date(prev_year, 1, 1), prv_e


# ─── Sheet 16: 月報YOY（fill_sheet10）（年對年累積比較）────────────────────────────────────────
def fill_sheet10(ws, df_cur, df_prev, sacare_prices, dates: dict, traffic=None, emp_count=None,
                 gross_bi=None):
    """列項同第 6 頁（STORE_KPI_ROWS）。gross_bi：{'cur'|'prv': {apl_gross, tpp_gross, acpp_gross}}"""
    print('  Sheet 16: 月報YOY', flush=True)
    # 累積區間：今年 1/1～截止日、去年 1/1～去年同日（截止日預設週末，可由前端自訂）
    cur_s, cur_e, prv_s, prv_e = _yoy_periods(dates)
    cur = _apply_gross_bi(calc_metrics(period(df_cur,  cur_s, cur_e), sacare_prices), (gross_bi or {}).get('cur'))
    prv = _apply_gross_bi(calc_metrics(period(df_prev, prv_s, prv_e), sacare_prices), (gross_bi or {}).get('prv'))

    # 人均產值（累積 = 累積零售營業額(a) / 編制人數）、來客數（ShopperTrak）
    if emp_count:
        cur['per_capita'] = int(cur.get('retail_rev', 0) / emp_count)
        prv['per_capita'] = int(prv.get('retail_rev', 0) / emp_count)
    t = traffic or {}
    if t.get('prv') is not None or t.get('cur') is not None:
        cur['visitors'] = int(t.get('cur') or 0)
        prv['visitors'] = int(t.get('prv') or 0)

    for r, (_, spec) in STORE_KPI_ROWS.items():
        b = _kpi_value(prv, spec)
        c = _kpi_value(cur, spec)
        if b is None and c is None:        # 專案營業額(d)、無人均產值/來客數時留空
            continue
        is_pct = isinstance(spec, tuple)
        b, c = b or 0, c or 0
        ws.cell(row=r, column=2).value = (round(b, 4) if is_pct else int(b)) if (is_pct or b) else None  # B 去年同期累積
        ws.cell(row=r, column=3).value = (round(c, 4) if is_pct else int(c)) if (is_pct or c) else None  # C 今年累積
        # D 年差異 = 今年 − 去年
        diff = c - b
        ws.cell(row=r, column=4).value = round(diff, 6) if is_pct else (int(diff) if diff else None)
        # E 差異比例 = 年差異 / 去年
        ws.cell(row=r, column=5).value = safe_rate(diff, b)


# ─── Sheet 17: 3PP YOY（fill_sheet11）（各 3PP 類別年對年累積比較）─────────────────────────────
def fill_sheet11(ws, df_cur, df_prev, sacare_prices, dates: dict):
    print('  Sheet 17: 3PP YOY', flush=True)
    sa_codes = set(sacare_prices.keys())

    def c4_rev(df, start, end, c4_code):
        d = period(df, start, end)
        d = d[~d['存貨代碼'].astype(str).str.strip().isin(sa_codes)]
        d = d[(d['類別3代碼'] == 3003.0) & (d['類別4代碼'] == c4_code)]
        return int(d['NET'].sum())

    def sa_rev_total(df, start, end):
        d = period(df, start, end)
        sa = d[d['存貨代碼'].astype(str).str.strip().isin(sa_codes)].copy()
        sa['SA_NET'] = sa['存貨代碼'].astype(str).str.strip().map(sacare_prices) * sa['數量'].fillna(0)
        return int(sa['SA_NET'].sum())

    cur_s, cur_e, prv_s, prv_e = _yoy_periods(dates)

    def write_row(row, b, c):
        # C 去年度累積、D 今年度累積、E 差異金額、F 差異比例
        ws.cell(row=row, column=3).value = b or None
        ws.cell(row=row, column=4).value = c or None
        ws.cell(row=row, column=5).value = (c - b) or None
        ws.cell(row=row, column=6).value = safe_rate(c - b, b)

    # 依模板既有的「類別4代碼」(A 欄) 逐列對應，動態讀取避免硬編列號
    sa_row = None
    pure_b = pure_c = 0
    for r in range(2, ws.max_row + 1):
        code = ws.cell(row=r, column=1).value
        if code is None:
            continue
        code_f = float(code)
        if code_f == 4040.0:        # SA CARE
            b, c = sa_rev_total(df_prev, prv_s, prv_e), sa_rev_total(df_cur, cur_s, cur_e)
            sa_row = r
        else:                        # 一般 3PP 類別
            b, c = c4_rev(df_prev, prv_s, prv_e, code_f), c4_rev(df_cur, cur_s, cur_e, code_f)
            pure_b += b
            pure_c += c
        write_row(r, b, c)

    sa_b = sa_rev_total(df_prev, prv_s, prv_e)
    sa_c = sa_rev_total(df_cur, cur_s, cur_e)
    # 加總（純配件 + SA CARE）與 純配件：依 B 欄文字找列
    for r in range(2, ws.max_row + 1):
        label = ws.cell(row=r, column=2).value
        if label == '加總':
            write_row(r, pure_b + sa_b, pure_c + sa_c)
        elif label == '純配件':
            write_row(r, pure_b, pure_c)


# ─── Sheet 18/20：差異明細比較（週對週／本月對去年同月同期／年累積）────────────────
# 每列並排三組比較：上週|本週、去年同月同期|本月、去年累積|今年累積（同第 7、16/17 頁的區間）。
# 兩頁由程式建立（不在模板內），Top 10 等列數可變動。
DETAIL_SHEET_PRODUCTS = '商品明細比較'
DETAIL_SHEET_OPS      = '經營結構比較'
TPP_DETAIL_SHEET      = '3PP 類別細分'
DETAIL_GROUPS = [('週對週', '上週', '本週', '週差', '週差%'),
                 ('本月 vs 去年同月同期', '去年同期', '本月', '月差', '月差%'),
                 ('年累積 vs 去年累積', '去年累積', '今年累積', '年差', '年差%')]

# 主機品類：(列名, 判定)；台數口徑同 calc_metrics（認證機品牌一併計入），營業額＝同 mask 的 NET
CERT_BRANDS = {881.0, 885.0, 886.0, 888.0}
C6_HOMEPOD  = {6073.0}

def _is_airpods(d):
    """AirPods 主機：C6_AIRPODS 只含現行世代，舊款（如去年的 AirPods Pro 2）以 C3=3002＋C4=4014＋品名判定。"""
    by_name = (d['類別3代碼'] == 3002.0) & (d['類別4代碼'] == 4014.0) & \
              d['名稱'].astype(str).str.contains('airpods', case=False, na=False)
    return d['類別6代碼'].isin(C6_AIRPODS) | by_name

def _is_audio_device(d):
    return _is_airpods(d) | d['類別6代碼'].isin(C6_HOMEPOD)

DEVICE_LINES = [
    ('Mac 筆電', lambda d: d['類別6代碼'].isin(C6_MACBOOK), 3001.0),
    ('Mac 桌機', lambda d: d['類別6代碼'].isin({6001.0, 6002.0}), 3001.0),
    ('iPhone',   lambda d: d['類別4代碼'].isin(C4_IPHONE), 3001.0),
    ('iPad',     lambda d: d['類別4代碼'].isin(C4_IPAD), 3001.0),
    ('Watch',    lambda d: d['類別4代碼'].isin(C4_WATCH), 3001.0),
    ('AirPods',  _is_airpods, 3002.0),
    ('HomePod',  lambda d: d['類別6代碼'].isin(C6_HOMEPOD), 3002.0),
]

# Apple 原廠配件（C3=3002，AirPods/HomePod 已列在主機區）依 C4 分列；未列出的 C4 併入「其他」
APPLE_ACC_C4 = [
    (4007.0, 'CPU 配件（滑鼠／鍵盤／觸控板）'),
    (4009.0, 'iPhone 配件'),
    (4010.0, 'iPad 配件（含 Pencil）'),
    (4012.0, '通用線材／充電'),
    (4014.0, '耳機（不含 AirPods）'),
    (4039.0, 'Watch 配件'),
    (4067.0, 'Apple TV'),
]

# 禮券（只統計門市折抵；門市 POS 沒有賣券紀錄）
#   每月最後一天總部以員工 SA999 開單，用正數沖回當月全部折抵（月結沖帳），統計時整張排除。
#   折抵金額＝−NET（折抵列為負數；退貨／退訂的反向列為正數，自動扣回）。
#   即享券 C3=3024（銷售列負數）、享樂券 C3=3047（銷退列負數）、
#   STUDIO A 週邊商品抵用券 C3=3023 且品名含「抵用券」（排除折抵差額／餘額碼）。
#   STUDIO A 商品抵用券 200/100 元：贈送碼只計張數，兌換碼計張數＋折抵金額。
#   ↑ 新券種上市時只需調整此處
VOUCHER_SETTLE_EMP = 'SA999'
VOUCHER_RULES = [
    # (券名, 篩選)
    ('即享券', lambda d: d['類別3代碼'] == 3024.0),
    ('享樂券', lambda d: d['類別3代碼'] == 3047.0),
    ('STUDIO A 週邊商品抵用券',
     lambda d: (d['類別3代碼'] == 3023.0) & d['名稱'].astype(str).str.contains('抵用券')
               & ~d['存貨代碼'].isin({'99902213', '99500141'})),
]
# 類別 6 名稱（第 19 頁 3PP 細分用）：server 執行時由 ERP STKCAT6 載入；CLI 無來源時顯示代碼
C6_NAMES: dict = {}

def _c6_label(code) -> str:
    if pd.isna(code):
        return '未分類'
    name = C6_NAMES.get(float(code))
    return f'{int(code)} {name}' if name else f'類別6 {int(code)}'

VOUCHER_GIFT_SKUS   = {'99901687', '99901688'}   # STUDIO A 商品抵用券 200/100 元／贈送
VOUCHER_REDEEM_SKUS = {'99901689', '99901690'}   # STUDIO A 商品抵用券 200/100 元－兌換


def _detail_frames(df_cur, df_prev, dates):
    """回傳 6 個期間資料：上週、本週、去年同月同期、本月、去年累積、今年累積。"""
    cur_s, cur_e, prv_s, prv_e = _yoy_periods(dates)
    return [
        period(df_cur,  dates['prev_wk_start'], dates['prev_wk_end']),
        period(df_cur,  dates['wk_start'],      dates['wk_end']),
        period(df_prev, dates['ly_start'],      dates['ly_end']),
        period(df_cur,  dates['mo_start'],      dates['mo_end']),
        period(df_prev, prv_s, prv_e),
        period(df_cur,  cur_s, cur_e),
    ]


def _detail_styles():
    """分析頁共用樣式：深藍標題、淡色組別子標題、淺藍灰分區列、斑馬紋資料列、組別間中粗直線。"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    hair = Side(style='thin', color='D9D9D9')
    return dict(
        font=Font(name='Arial', size=12),
        bold=Font(name='Arial', size=12, bold=True),
        cur=Font(name='Arial', size=12, bold=True, color='1F3864'),   # 本期欄
        head=Font(name='Arial', size=12, bold=True, color='FFFFFF'),
        head_fill=PatternFill('solid', fgColor='1F3864'),
        group_fills=[PatternFill('solid', fgColor=c) for c in ('DDEBF7', 'E2EFDA', 'FFF2CC')],
        sec_fill=PatternFill('solid', fgColor='D6DCE4'),
        zebra_fill=PatternFill('solid', fgColor='F2F2F2'),
        total_fill=PatternFill('solid', fgColor='DDEBF7'),
        no_fill=PatternFill(fill_type=None),
        hair=hair,
        mid=Side(style='medium', color='7F7F7F'),
        border=Border(left=hair, right=hair, top=hair, bottom=hair),
        center=Alignment(horizontal='center', vertical='center', wrap_text=True),
        left=Alignment(horizontal='left', vertical='center'),
        left_wrap=Alignment(horizontal='left', vertical='center', wrap_text=True),
        indent=Alignment(horizontal='left', vertical='center', indent=1),
        zebra=0,
    )


def _paint_row(ws, row, st, role='data', cols=None, group_starts=(), fill=None, height=20):
    """依角色上底色與框線：data＝斑馬紋、section＝分區列、total＝小計/合計、head＝標題。
    cols：要塗的欄（預設 A～M）；group_starts：組別第一欄（左側畫中粗線）；fill：指定底色（覆蓋斑馬紋）。"""
    from openpyxl.styles import Border
    cols = cols or range(1, 14)
    if role == 'data':
        st['zebra'] += 1
        fill = fill or (st['zebra_fill'] if st['zebra'] % 2 == 0 else st['no_fill'])
    else:
        st['zebra'] = 0
        fill = fill or {'section': st['sec_fill'], 'total': st['total_fill']}.get(role, st['head_fill'])
    top = st['mid'] if role in ('section', 'total') else None
    for c in cols:
        cell = ws.cell(row=row, column=c)
        cell.border = Border(left=st['mid'] if c in group_starts else None, top=top, bottom=st['hair'])
        cell.fill = fill
    if height:   # 有自動換行的列傳 None，讓 Excel 依內容撐高
        ws.row_dimensions[row].height = height


def _reset_detail_sheet(wb, name, freeze='B3'):
    """清掉舊的同名頁，在最後重新建立（模板沒有這幾頁）。"""
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.column_dimensions['A'].width = 40
    for col in range(2, 14):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 13
    ws.freeze_panes = freeze   # 只能設一次：重設會留下重複的 selection，Excel 開檔會要求修復
    ws.sheet_view.showGridLines = False
    return ws


DETAIL_GROUP_STARTS = (2, 6, 10)


def _write_detail_header(ws, st):
    from openpyxl.styles import Border
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws.cell(row=1, column=1).value = '項目'
    for gi, (gname, a, b, diff, pct) in enumerate(DETAIL_GROUPS):
        c0 = 2 + gi * 4
        ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 3)
        ws.cell(row=1, column=c0).value = gname
        for j, lab in enumerate((a, b, diff, pct)):
            ws.cell(row=2, column=c0 + j).value = lab
    for r in (1, 2):
        for c in range(1, 14):
            cell = ws.cell(row=r, column=c)
            cell.font, cell.alignment = st['head'], st['center']
            cell.border = Border(left=st['mid'] if c in DETAIL_GROUP_STARTS else None, bottom=st['hair'])
            cell.fill = st['head_fill'] if c == 1 or r == 1 else st['group_fills'][(c - 2) // 4]
            if r == 2 and c > 1:
                cell.font = st['cur'] if (c - 2) % 4 == 1 else st['bold']
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 22


def _write_section(ws, row, title, st, ncols=13, group_starts=DETAIL_GROUP_STARTS):
    _paint_row(ws, row, st, 'section', cols=range(1, ncols + 1), group_starts=group_starts)
    ws.cell(row=row, column=1).value = title
    ws.cell(row=row, column=1).font = st['bold']
    ws.cell(row=row, column=1).alignment = st['left']
    return row + 1


_FMT = {
    'n': ('#,##0', '[Color10]+#,##0;[Red]-#,##0;0'),
    'f': ('0.00', '[Color10]+0.00;[Red]-0.00;0.00'),
    'r': ('0.0%', '[Color10]+0.0"pp";[Red]-0.0"pp";0.0"pp"'),
}
_PCT_FMT = '[Color10]+0.0%;[Red]-0.0%;0.0%'


def _is_total_label(label):
    return str(label).startswith('加總') or '合計' in str(label)


def _diff_pair(prev, cur, kind='n'):
    """(差異, 差異%)；比率類差異以百分點表示、% 留空。"""
    if prev is None or cur is None:
        return None, None
    if kind == 'r':
        return round((cur - prev) * 100, 2), None
    diff = cur - prev
    return (round(diff, 2) if kind == 'f' else diff), (safe_rate(diff, prev) if prev else None)


def _write_compare_row(ws, row, label, vals, st, kind='n', role=None, indent=True):
    """vals＝6 個值（上週, 本週, 去年同期, 本月, 去年累積, 今年累積）。
    kind：n＝整數金額/件數、f＝兩位小數、r＝比率（差異以百分點 pp 表示，% 欄留空）。
    role：data／total（預設依列名判斷：「合計」「加總」為 total）。"""
    role = role or ('total' if _is_total_label(label) else 'data')
    _paint_row(ws, row, st, role, group_starts=DETAIL_GROUP_STARTS)
    val_fmt, diff_fmt = _FMT[kind]
    a = ws.cell(row=row, column=1)
    a.value = label
    a.font = st['bold'] if role == 'total' else st['font']
    a.alignment = st['indent'] if (indent and role == 'data') else st['left']
    for gi in range(3):
        prev, cur = vals[gi * 2], vals[gi * 2 + 1]
        c0 = 2 + gi * 4
        diff, pct = _diff_pair(prev, cur, kind)
        if kind == 'f':
            prev = None if prev is None else round(prev, 2)
            cur = None if cur is None else round(cur, 2)
        for j, (v, fmt) in enumerate(((prev, val_fmt), (cur, val_fmt), (diff, diff_fmt), (pct, _PCT_FMT))):
            cell = ws.cell(row=row, column=c0 + j)
            cell.value = v
            cell.number_format = fmt
            cell.font = st['cur'] if j == 1 else (st['bold'] if role == 'total' else st['font'])
    return row + 1


def _net_units(d, mask):
    sale = d.loc[mask & d['交易類型'].isin(SALE_TYPES), '數量'].sum()
    ret  = d.loc[mask & (d['交易類型'] == '銷退'), '數量'].abs().sum()
    return int(sale - ret)


def _non_sa(d, sa_codes):
    return d[~d['存貨代碼'].astype(str).str.strip().isin(sa_codes)]


def _device_stats(d, sa_codes):
    """各主機品類 (台數, 營業額)。"""
    d = _non_sa(d, sa_codes)
    brand = d['品牌代碼'].isin(CERT_BRANDS)
    out = {}
    for name, sel, c3 in DEVICE_LINES:
        mask = ((d['類別3代碼'] == c3) | brand) & sel(d)
        out[name] = (_net_units(d, mask), int(d.loc[mask, 'NET'].sum()))
    return out


def _model_units(d, c3, c6_set, must_inc, must_exc):
    """機型台數，判定同 fill_sheet1 的 calc_units。"""
    mask = (d['類別3代碼'] == c3) | d['品牌代碼'].isin(CERT_BRANDS)
    if c6_set is not None:
        mask &= d['類別6代碼'].isin(c6_set)
    name = d['名稱'].astype(str)
    if must_inc: mask &= name.str.contains(must_inc, case=False, na=False)
    if must_exc: mask &= ~name.str.contains(must_exc, case=False, na=False)
    return _net_units(d, mask)


def fill_sheet_detail_products(wb, df_cur, df_prev, sacare_prices, dates: dict):
    print(f'  分析頁：{DETAIL_SHEET_PRODUCTS}', flush=True)
    sa_codes = set(sacare_prices.keys())
    frames = _detail_frames(df_cur, df_prev, dates)
    ws = _reset_detail_sheet(wb, DETAIL_SHEET_PRODUCTS)
    st = _detail_styles()
    _write_detail_header(ws, st)
    row = 3

    # 1. 主機（依品類）
    row = _write_section(ws, row, '主機（依品類）', st)
    dev = [_device_stats(f, sa_codes) for f in frames]
    for name, _, _ in DEVICE_LINES:
        units = [x[name][0] for x in dev]
        revs = [x[name][1] for x in dev]
        asp = [int(round(r / u)) if u > 0 else None for u, r in zip(units, revs)]
        row = _write_compare_row(ws, row, f'{name} 台數', units, st)
        row = _write_compare_row(ws, row, f'{name} 營業額', revs, st)
        row = _write_compare_row(ws, row, f'{name} 平均單價（營業額÷台數）', asp, st)

    # 2. 主機機型明細（列名取第 1 頁 A 欄；六期全為 0 的機型略過）
    row = _write_section(ws, row, '主機機型明細（台數；新機型去年同期為 0）', st)
    ws1 = wb['1.主機銷售台數'] if '1.主機銷售台數' in wb.sheetnames else None
    for row_idx, (c3, c6_set, inc, exc) in PRODUCT_ROWS.items():
        vals = [_model_units(f, c3, c6_set, inc, exc) for f in frames]
        if not any(vals):
            continue
        label = ws1.cell(row=row_idx + 1, column=1).value if ws1 else f'機型 {row_idx}'
        row = _write_compare_row(ws, row, str(label).strip(), vals, st)

    # 3. Apple 原廠配件（C3=3002）
    row = _write_section(ws, row, 'Apple 原廠配件（營業額）', st)
    acc = [_non_sa(f, sa_codes) for f in frames]
    acc = [a[a['類別3代碼'] == 3002.0] for a in acc]
    is_dev = [_is_audio_device(a) for a in acc]
    known = {c4 for c4, _ in APPLE_ACC_C4}
    for c4, label in APPLE_ACC_C4:
        vals = [int(a.loc[~m & (a['類別4代碼'] == c4), 'NET'].sum()) for a, m in zip(acc, is_dev)]
        row = _write_compare_row(ws, row, label, vals, st)
    vals = [int(a.loc[~m & ~a['類別4代碼'].isin(known), 'NET'].sum()) for a, m in zip(acc, is_dev)]
    row = _write_compare_row(ws, row, '其他原廠配件', vals, st)
    vals = [int(a.loc[m, 'NET'].sum()) for a, m in zip(acc, is_dev)]
    row = _write_compare_row(ws, row, 'AirPods／HomePod（同主機區）', vals, st)
    vals = [int(a['NET'].sum()) for a in acc]
    row = _write_compare_row(ws, row, 'Apple 配件營業額合計（同第 6 頁）', vals, st)

    mets = [calc_metrics(f, sacare_prices) for f in frames]

    # 5. 保固搭售明細（ACPP／SA Care 分開看）
    row = _write_section(ws, row, '保固搭售明細', st)
    for dev_name, unit_k, ac_k, sa_k, w_k in (
            ('CPU', 'cpu_total', 'acpp_mac', 'sa_cpu', 'warranty_cpu'),
            ('iPad', 'ipad', 'acpp_ipad', 'sa_ipad', 'warranty_ipad'),
            ('iPhone', 'iphone', 'acpp_iphone', 'sa_iphone', 'warranty_iphone'),
            ('Watch', 'watch', 'acpp_watch', 'sa_watch', 'warranty_watch'),
            ('AirPods', 'airpods', 'acpp_airpods', 'sa_airpods', 'warranty_airpods')):
        row = _write_compare_row(ws, row, f'{dev_name} 台數', [m[unit_k] for m in mets], st)
        row = _write_compare_row(ws, row, f'{dev_name} ACPP 套數', [m[ac_k] for m in mets], st)
        row = _write_compare_row(ws, row, f'{dev_name} SA Care 套數', [m[sa_k] for m in mets], st)
        row = _write_compare_row(ws, row, f'{dev_name} 搭售率（AC+SA）',
                                 [safe_rate(m[w_k], m[unit_k]) if m[unit_k] else None for m in mets],
                                 st, kind='r')
    row = _write_compare_row(ws, row, 'ACPP 營業額', [m['acpp_rev'] for m in mets], st)
    row = _write_compare_row(ws, row, 'SA Care 營業額', [m['sa_rev'] for m in mets], st)
    row = _write_compare_row(ws, row, '保固營業額（AC+SA）', [m['warranty_rev'] for m in mets], st)



TPP_GROUP_STARTS = (2, 10, 18)


def _write_tpp_header(ws, st):
    """3PP 類別細分頁表頭：每組 8 欄＝前期數量、前期營業額、本期數量、本期營業額、數量差異、數量差異%、金額差異、金額差異%。"""
    from openpyxl.styles import Border
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws.cell(row=1, column=1).value = '項目'
    for gi, (gname, a, b, _, _) in enumerate(DETAIL_GROUPS):
        c0 = 2 + gi * 8
        ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 7)
        ws.cell(row=1, column=c0).value = gname
        for j, lab in enumerate((f'{a}\n數量', f'{a}\n營業額', f'{b}\n數量', f'{b}\n營業額',
                                 '數量差異', '數量差異%', '金額差異', '金額差異%')):
            ws.cell(row=2, column=c0 + j).value = lab
    for r in (1, 2):
        for c in range(1, 26):
            cell = ws.cell(row=r, column=c)
            cell.font, cell.alignment = st['head'], st['center']
            cell.border = Border(left=st['mid'] if c in TPP_GROUP_STARTS else None, bottom=st['hair'])
            cell.fill = st['head_fill'] if c == 1 or r == 1 else st['group_fills'][(c - 2) // 8]
            if r == 2 and c > 1:
                cell.font = st['cur'] if (c - 2) % 8 in (2, 3) else st['bold']
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 36


def _write_qty_rev_row(ws, row, label, qtys, revs, st, role=None):
    """qtys／revs 各 6 個值（上週, 本週, 去年同期, 本月, 去年累積, 今年累積）。"""
    role = role or ('total' if _is_total_label(label) else 'data')
    _paint_row(ws, row, st, role, cols=range(1, 26), group_starts=TPP_GROUP_STARTS)
    a = ws.cell(row=row, column=1)
    a.value = label
    a.font = st['bold'] if role == 'total' else st['font']
    a.alignment = st['indent'] if role == 'data' else st['left']
    diff_fmt = _FMT['n'][1]
    for gi in range(3):
        pq, cq, pr, cr = qtys[gi * 2], qtys[gi * 2 + 1], revs[gi * 2], revs[gi * 2 + 1]
        dq, dqp = _diff_pair(pq, cq)
        dr, drp = _diff_pair(pr, cr)
        c0 = 2 + gi * 8
        cells = ((pq, '#,##0'), (pr, '#,##0'), (cq, '#,##0'), (cr, '#,##0'),
                 (dq, diff_fmt), (dqp, _PCT_FMT), (dr, diff_fmt), (drp, _PCT_FMT))
        for j, (v, fmt) in enumerate(cells):
            cell = ws.cell(row=row, column=c0 + j)
            cell.value, cell.number_format = v, fmt
            cell.font = st['cur'] if j in (2, 3) else (st['bold'] if role == 'total' else st['font'])
    return row + 1


def fill_sheet_tpp_detail(wb, df_cur, df_prev, sacare_prices, dates: dict):
    """3PP 類別細分頁：3PP 類別 4 → 類別 6 細分，數量與營業額並列（類別 4 小計營業額口徑同第 7/17 頁；列名取第 7 頁 B 欄）。
    類別 6 依今年累積營業額由大到小排，六期營業額全為 0 的略過。數量＝銷售＋尾款－銷退（含贈品）。"""
    print(f'  分析頁：{TPP_DETAIL_SHEET}', flush=True)
    sa_codes = set(sacare_prices.keys())
    frames = _detail_frames(df_cur, df_prev, dates)
    ws = _reset_detail_sheet(wb, TPP_DETAIL_SHEET)
    for col in range(2, 26):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 11
    st = _detail_styles()
    _write_tpp_header(ws, st)
    row = _write_section(ws, 3, '3PP 類別細分（粗體＝類別 4 小計，營業額同第 7、17 頁）', st,
                         ncols=25, group_starts=TPP_GROUP_STARTS)
    ws7 = wb['7.3PP配件比較'] if '7.3PP配件比較' in wb.sheetnames else None
    tpp = [_non_sa(f, sa_codes) for f in frames]
    tpp = [t[t['類別3代碼'] == 3003.0] for t in tpp]
    units = lambda t: _net_units(t, pd.Series(True, index=t.index))
    pure_q, pure_r = [0] * 6, [0] * 6
    for c4, excel_row in C4_ROWS.items():
        sub = [t[t['類別4代碼'] == c4] for t in tpp]
        revs = [int(t['NET'].sum()) for t in sub]
        qtys = [units(t) for t in sub]
        pure_r = [p + v for p, v in zip(pure_r, revs)]
        pure_q = [p + v for p, v in zip(pure_q, qtys)]
        if not any(revs) and not any(len(t) for t in sub):
            continue
        label = ws7.cell(row=excel_row, column=2).value if ws7 else f'類別4 {int(c4)}'
        row = _write_qty_rev_row(ws, row, f'{int(c4)} {label}', qtys, revs, st, role='total')
        keyed = [t.assign(_c6=t['類別6代碼'].fillna(-1)) for t in sub]
        by_c6 = [k.groupby('_c6')['NET'].sum() for k in keyed]
        codes = sorted(set().union(*[set(b.index) for b in by_c6]), key=lambda c: -by_c6[5].get(c, 0))
        for c6 in codes:
            c6_r = [int(b.get(c6, 0)) for b in by_c6]
            if not any(c6_r):
                continue
            c6_q = [units(k[k['_c6'] == c6]) for k in keyed]
            row = _write_qty_rev_row(ws, row, _c6_label(None if c6 == -1 else c6), c6_q, c6_r, st)
    mets = [calc_metrics(f, sacare_prices) for f in frames]
    sa_r = [m['sa_rev'] for m in mets]
    sa_q = [sum(m[k] for k in ('sa_cpu', 'sa_ipad', 'sa_iphone', 'sa_watch', 'sa_airpods')) for m in mets]
    row = _write_qty_rev_row(ws, row, 'SA CARE', sa_q, sa_r, st, role='total')
    row = _write_qty_rev_row(ws, row, '純配件', pure_q, pure_r, st, role='total')
    row = _write_qty_rev_row(ws, row, '加總（純配件＋SA CARE）', [p + q for p, q in zip(pure_q, sa_q)],
                             [p + r for p, r in zip(pure_r, sa_r)], st)
    ws.cell(row=row + 1, column=1).value = ('數量＝銷售＋尾款－銷退的淨件數（含贈品）；SA CARE 數量＝各機種 SA Care 套數合計。'
                                            '差異%＝差異÷前期。')
    ws.cell(row=row + 1, column=1).font = st['font']


def _doc_set(d, mask):
    return set(d.loc[mask, '單據代碼'].dropna().astype(str))


def _voucher_rows(d):
    """回傳 [(列名, 值)]；折抵金額以正數表示，單數＝折抵淨額 > 0 的單據數。"""
    rows = []
    d = d[d['員工代碼'].astype(str).str.strip() != VOUCHER_SETTLE_EMP]
    for name, sel in VOUCHER_RULES:
        v = d[sel(d)]
        amt = -v['NET']
        per_doc = amt.groupby(v['單據代碼']).sum()
        rows.append((f'{name} 折抵單數', int((per_doc > 0).sum())))
        rows.append((f'{name} 折抵金額', int(round(amt.sum()))))
    codes = d['存貨代碼'].astype(str).str.strip()
    qty_types = d['交易類型'].isin(['銷售', '尾款', '銷退'])
    gift = d[codes.isin(VOUCHER_GIFT_SKUS) & qty_types]
    redeem = d[codes.isin(VOUCHER_REDEEM_SKUS)]
    rows.append(('STUDIO A 商品抵用券 贈送張數', int(gift['數量'].sum())))
    rows.append(('STUDIO A 商品抵用券 兌換張數',
                 int(redeem.loc[redeem['交易類型'].isin(['銷售', '銷退']), '數量'].sum())))
    rows.append(('STUDIO A 商品抵用券 兌換金額', int(round(-redeem['NET'].sum()))))
    return rows


def fill_sheet_detail_ops(wb, df_cur, df_prev, sacare_prices, dates: dict):
    print(f'  分析頁：{DETAIL_SHEET_OPS}', flush=True)
    sa_codes = set(sacare_prices.keys())
    frames = _detail_frames(df_cur, df_prev, dates)
    ws = _reset_detail_sheet(wb, DETAIL_SHEET_OPS)
    st = _detail_styles()
    _write_detail_header(ws, st)
    row = 3
    mets = [calc_metrics(f, sacare_prices) for f in frames]

    # 1. 營收結構
    row = _write_section(ws, row, '營收結構', st)
    for label, key in (('總營業額', 'total_rev'), ('零售營業額', 'retail_rev'),
                       ('Apple 主機營業額', 'rev_3001'), ('Apple 配件營業額', 'rev_3002'),
                       ('3PP 配件營業額', 'rev_3003'), ('保固營業額（AC+SA）', 'warranty_rev')):
        row = _write_compare_row(ws, row, label, [m[key] for m in mets], st)
    for label, key in (('主機佔比', 'rev_3001'), ('Apple 配件佔比', 'rev_3002'),
                       ('3PP 配件佔比', 'rev_3003'), ('保固佔比', 'warranty_rev')):
        row = _write_compare_row(ws, row, label,
                                 [safe_rate(m[key], m['total_rev']) if m['total_rev'] else None for m in mets],
                                 st, kind='r')

    # 2. 客單結構（成交單＝有銷售列且非整單銷退，同 calc_metrics 的成交筆數）
    row = _write_section(ws, row, '客單結構', st)
    txn, aov, ipt, dev_share, tpp_share, tpp_aov = [], [], [], [], [], []
    for f, m in zip(frames, mets):
        nf = _non_sa(f, sa_codes)
        deals = _doc_set(f, f['交易類型'] == '銷售') - _doc_set(f, f['交易類型'] == '銷退')
        n = len(deals)
        in_deal = nf['單據代碼'].astype(str).isin(deals) & (nf['交易類型'] == '銷售')
        goods = in_deal & nf['類別3代碼'].isin([3001.0, 3002.0, 3003.0])
        dev_docs = _doc_set(nf, in_deal & (nf['類別3代碼'] == 3001.0))
        tpp_mask = in_deal & (nf['類別3代碼'] == 3003.0)
        tpp_docs = _doc_set(nf, tpp_mask)
        txn.append(m['txn_count'])
        aov.append(int(round(m['retail_rev'] / n)) if n else None)
        ipt.append(nf.loc[goods, '數量'].sum() / n if n else None)
        dev_share.append(safe_rate(len(dev_docs), n) if n else None)
        tpp_share.append(safe_rate(len(tpp_docs), n) if n else None)
        tpp_aov.append(int(round(nf.loc[tpp_mask, 'NET'].sum() / len(tpp_docs))) if tpp_docs else None)
    row = _write_compare_row(ws, row, '成交筆數', txn, st)
    row = _write_compare_row(ws, row, '客單價（零售營業額÷成交筆數）', aov, st)
    row = _write_compare_row(ws, row, '每筆件數（主機＋配件＋3PP）', ipt, st, kind='f')
    row = _write_compare_row(ws, row, '含主機單佔比', dev_share, st, kind='r')
    row = _write_compare_row(ws, row, '含 3PP 單佔比', tpp_share, st, kind='r')
    row = _write_compare_row(ws, row, '3PP 單均金額（含 3PP 單）', tpp_aov, st)

    # 3. 禮券
    row = _write_section(ws, row, '禮券折抵（不含總部月結沖帳；退貨／退訂已扣回）', st)
    per_period = [_voucher_rows(f) for f in frames]
    for i, (label, _) in enumerate(per_period[0]):
        row = _write_compare_row(ws, row, label, [p[i][1] for p in per_period], st)


# ─── Sheet 21：3PP 熱門商品（本週／本月，同款不同色合併）─────────────────────────
# 取該期 3PP 營業額前 10 大的「類別 4＋類別 6」，每組依數量列前 5 個系列。
# ERP 商品主檔沒有系列／顏色欄位，系列由品名去掉顏色字判定（同類別 4＋類別 6 內品名相同＝同系列）。
HOT_SERIES_SHEET  = '3PP 熱門商品'
HOT_SERIES_GROUPS = 10   # 每期展開的類別數
HOT_SERIES_TOP    = 5    # 每類列出的系列數
SERIES_COLOR_RE = re.compile(
    '黑|白|紅|藍|綠|黃|紫|粉|灰|銀|金|棕|橙|橘|咖啡|透明|米|杏|卡其|奶茶|奶油|午夜|星光|太空|石墨|沙|鈦|'
    '玫瑰|薰衣草|天空|森林|酒紅|勃根地|莎瓦納|朔月|炭|碳|曜石|珍珠|香檳|迷彩|褐|'
    'white|black|blue|red|green|pink|grey|gray|silver|gold|purple|clear|navy|midnight|starlight|'
    'beige|brown|orange|yellow|graphite|titanium', re.I)
# 含這些字的不是顏色（例：「抗藍光」「晶透抗黃殻」「鋁合金支架」「(藍色吊挂繩)」）
SERIES_NOT_COLOR_RE = re.compile(r'藍光|貼|殼|殻|膜|線|套|充|器|版|款|組|入|件|支架|扣|按鍵|繩|\d')


def _series_name(name) -> str:
    """去掉品名中的顏色（括號內顏色、結尾「-顏色」／「 顏色」），作為同款系列名。"""
    n = re.sub(r'\s+', ' ', str(name).replace('\xa0', ' ')).strip().lstrip('@').strip()

    def _paren(m):
        t = m.group(1)
        is_color = len(t) <= 6 and SERIES_COLOR_RE.search(t) and not SERIES_NOT_COLOR_RE.search(t)
        return '' if is_color else m.group(0)

    n = re.sub(r'[\(（]([^()（）]{1,8})[\)）]', _paren, n).strip()
    base = n   # 結尾沒有顏色可去時，回傳這個（不留下補的空白）
    # 只處理品名最後一段：「)黑色」→「) 黑色」、「1.2m銀」→「1.2m 銀」，好讓下面切出顏色
    n = re.sub(r'([)）])(?=[^\s)）]{1,6}$)', r'\1 ', n)
    n = re.sub(r'(\d(?:m|cm|mm|W))(?=[^\x00-\x7f][^\s]{0,5}$)', r'\1 ', n, flags=re.I)
    stripped = False
    for _ in range(3):   # 雙色「-黑/鈦」要連續去掉
        m = re.match(r'^(.*?)(?:\s*[-－/]\s*|\s+)([^\s\-－/]{1,6})$', n)
        if not (m and SERIES_COLOR_RE.search(m.group(2)) and not SERIES_NOT_COLOR_RE.search(m.group(2))):
            break
        n, stripped = m.group(1).strip(), True
    return n if stripped else base


def _c6_name(code) -> str:
    if pd.isna(code):
        return '未分類'
    return C6_NAMES.get(float(code)) or f'類別6 {int(code)}'


def _hot_series_block(ws, col0, title, d, c4_labels, st):
    """在 col0 起的 5 欄寫入一期的熱門系列；回傳寫到的最後一列。"""
    from openpyxl.styles import Alignment
    heads = ['排名', '系列名稱', '色數', '數量', '營業額']
    row = 1
    ws.merge_cells(start_row=row, start_column=col0, end_row=row, end_column=col0 + 4)
    ws.cell(row=row, column=col0).value = title
    cols = range(col0, col0 + 5)
    _paint_row(ws, row, st, 'head', cols=cols, height=26)
    for c in cols:
        ws.cell(row=row, column=c).font, ws.cell(row=row, column=c).alignment = st['head'], st['center']
    row += 1
    _paint_row(ws, row, st, 'head', cols=cols, fill=st['group_fills'][0], height=22)
    for j, h in enumerate(heads):
        cell = ws.cell(row=row, column=col0 + j)
        cell.value, cell.font, cell.alignment = h, st['bold'], st['center']
    row += 1

    d = d.assign(_c6=d['類別6代碼'].fillna(-1))
    group_rev = d.groupby(['類別4代碼', '_c6'])['NET'].sum().sort_values(ascending=False)
    group_rev = group_rev[group_rev > 0].head(HOT_SERIES_GROUPS)
    is_return = d['交易類型'] == '銷退'
    is_gift = (d['折扣'].fillna(0) >= 80) | ((d['NET'].abs() == 0) & ~is_return)
    items = d[~is_gift].assign(系列=lambda x: x['名稱'].map(_series_name))
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)

    for (c4, c6), rev in group_rev.items():
        g = items[(items['類別4代碼'] == c4) & (items['_c6'] == c6)]
        label = f"{c4_labels.get(c4, f'類別4 {int(c4)}')}｜{_c6_name(None if c6 == -1 else c6)}"
        qty_all = int(d.loc[(d['類別4代碼'] == c4) & (d['_c6'] == c6), '數量'].sum())
        vals = ['', label, None, qty_all, int(rev)]
        _paint_row(ws, row, st, 'section', cols=cols, height=None)
        for j, v in enumerate(vals):
            cell = ws.cell(row=row, column=col0 + j)
            cell.value, cell.font = v, st['bold']
            cell.alignment = left if j == 1 else st['center']
            if j >= 3:
                cell.number_format = '#,##0'
        row += 1
        ser = g.groupby('系列').agg(色數=('存貨代碼', 'nunique'), qty=('數量', 'sum'), rev=('NET', 'sum'))
        ser = ser[ser['qty'] > 0].sort_values(['qty', 'rev'], ascending=False).head(HOT_SERIES_TOP)
        for rank, (name, r) in enumerate(ser.iterrows(), start=1):
            vals = [rank, name, int(r['色數']), int(r['qty']), int(r['rev'])]
            _paint_row(ws, row, st, 'data', cols=cols, height=None)
            for j, v in enumerate(vals):
                cell = ws.cell(row=row, column=col0 + j)
                cell.value = v
                cell.font = st['cur'] if j == 3 else st['font']
                cell.alignment = left if j == 1 else st['center']
                if j >= 3:
                    cell.number_format = '#,##0'
            row += 1
    return row


def fill_sheet_hot_series(wb, df_cur, sacare_prices, dates: dict):
    print(f'  分析頁：{HOT_SERIES_SHEET}', flush=True)
    sa_codes = set(sacare_prices.keys())
    ws = _reset_detail_sheet(wb, HOT_SERIES_SHEET, freeze='A3')
    for col, width in zip('ABCDEFGHIJK', (6, 48, 6, 8, 11, 3, 6, 48, 6, 8, 11)):
        ws.column_dimensions[col].width = width
    st = _detail_styles()
    ws7 = wb['7.3PP配件比較'] if '7.3PP配件比較' in wb.sheetnames else None
    c4_labels = {c4: str(ws7.cell(row=r, column=2).value).strip() for c4, r in C4_ROWS.items()} if ws7 else {}

    def tpp(start, end):
        d = _non_sa(period(df_cur, start, end), sa_codes)
        return d[d['類別3代碼'] == 3003.0]

    wk_title = f"本週 {dates['wk_start']:%m/%d}~{dates['wk_end']:%m/%d}（依數量，同款不同色合併）"
    mo_title = f"本月 {dates['mo_start']:%m/%d}~{dates['mo_end']:%m/%d}（依數量，同款不同色合併）"
    _hot_series_block(ws, 1, wk_title, tpp(dates['wk_start'], dates['wk_end']), c4_labels, st)
    _hot_series_block(ws, 7, mo_title, tpp(dates['mo_start'], dates['mo_end']), c4_labels, st)


# ─── Sheet 22–27：主機活動／新機上市追蹤／員工分析（分析檔）────────────────────
PROMO_SHEET       = '主機活動'
LAUNCH_SHEET      = '新機上市追蹤'
EMP_AOV_SHEET     = '個人客單價'
EMP_WARRANTY_SHEET = '個人保固搭售率'
EMP_MIX_SHEET     = '個人業績組成'
EMP_TREND_SHEET   = '個人近8週走勢'
MYSETUP_DETAIL_SHEET = 'MySetup 提交細項'
ANALYSIS_SHEETS = [DETAIL_SHEET_PRODUCTS, TPP_DETAIL_SHEET, DETAIL_SHEET_OPS, HOT_SERIES_SHEET, PROMO_SHEET, LAUNCH_SHEET,
                   EMP_AOV_SHEET, EMP_WARRANTY_SHEET, EMP_MIX_SHEET, EMP_TREND_SHEET, MYSETUP_DETAIL_SHEET]

# 分析用主機品類（台數口徑同 _device_stats：C3 或認證機品牌，退貨扣回）
ANALYSIS_LINES = [
    ('iPhone',  lambda d: d['類別4代碼'].isin(C4_IPHONE), 3001.0),
    ('iPad',    lambda d: d['類別4代碼'].isin(C4_IPAD), 3001.0),
    ('Mac',     lambda d: d['類別6代碼'].isin(C6_CPU), 3001.0),
    ('Watch',   lambda d: d['類別4代碼'].isin(C4_WATCH), 3001.0),
    ('AirPods', _is_airpods, 3002.0),
]

# 新機上市追蹤：(新機, C6, C3, 對照機種, C6)。上市日＝第一筆銷售日；新機上市時在這裡加一組
LAUNCH_PAIRS = [
    ('iPhone 18 Pro',     {6345.0}, 3001.0, 'iPhone 17 Pro',     {6325.0}),
    ('iPhone 18 Pro Max', {6346.0}, 3001.0, 'iPhone 17 Pro Max', {6326.0}),
    ('Watch S12',         {6349.0}, 3001.0, 'Watch S11',         {6327.0}),
    ('Watch Ultra 4',     {6350.0}, 3001.0, 'Watch Ultra 3',     {6328.0}),
    ('AirPods 5',         {6348.0}, 3002.0, 'AirPods 4',         {6312.0}),
]
# 沒有對照機種、改追 AppleCare+ 搭售率的新機：(機種, C6, C3, ACPP 品名關鍵字)。只算 AppleCare+，不含 SAcare
LAUNCH_ACPP = [
    ('iPhone Duo', {6347.0}, 3001.0, 'iphone duo'),
]
LAUNCH_WEEKS = 12
LAUNCH_C6 = sorted({int(c) for p in LAUNCH_PAIRS for c in (p[1] | p[4])} |
                   {int(c) for p in LAUNCH_ACPP for c in p[1]})

C4_NAMES: dict = {}   # 類別 4 名稱（server 從 ERP 帶入；查不到時顯示代碼）
WARRANTY_DEVICES = [('Mac', 'cpu_total', 'acpp_mac', 'sa_cpu'), ('iPad', 'ipad', 'acpp_ipad', 'sa_ipad'),
                    ('iPhone', 'iphone', 'acpp_iphone', 'sa_iphone'), ('Watch', 'watch', 'acpp_watch', 'sa_watch'),
                    ('AirPods', 'airpods', 'acpp_airpods', 'sa_airpods')]


def _line_mask(d, sel, c3):
    return ((d['類別3代碼'] == c3) | d['品牌代碼'].isin(CERT_BRANDS)) & sel(d)


def _c4_name(code) -> str:
    return C4_NAMES.get(float(code)) or f'類別4 {int(code)}'


def _gift_mask(d):
    """贈品列（同 3PP 熱門商品頁）：折扣 ≥ 80 或金額 0 的非退貨列。"""
    return (d['折扣'].fillna(0) >= 80) | ((d['NET'].abs() == 0) & (d['交易類型'] != '銷退'))


def _avg_price(d):
    """平均單價＝(銷售＋銷退) 金額 ÷ 淨件數，排除贈品；訂金／尾款不計。"""
    x = d[d['交易類型'].isin(['銷售', '銷退']) & ~_gift_mask(d)]
    qty = x['數量'].sum()
    return int(round(x['NET'].sum() / qty)) if qty > 0 else None


def _emp_label(code, name):
    name = str(name or '').strip()
    return name if name.startswith(code) else f'{code} {name}'.strip()


def _emp_frames(d):
    """[(顯示名稱, 該員工的列)…, ('門市合計', 全部)]，員工順序同個人頁。"""
    rows = [(_emp_label(code, name), d[d['員工代碼'] == code]) for code, name in EMPLOYEES]
    return rows + [('門市合計', d)]


def _head_cells(ws, row, col0, labels, st, fill=None):
    """標題格：深藍底白字（fill 指定時用淡色底黑字）；st['groups'] 為組別第一欄，左側畫中粗線。"""
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import Border
    st['zebra'] = 0
    for j, lab in enumerate(labels):
        col = col0 + j
        cell = ws.cell(row=row, column=col)
        if not isinstance(cell, MergedCell):   # 合併儲存格只有左上角能寫值
            cell.value = lab
        cell.font = st['bold'] if fill else st['head']
        cell.alignment = st['center']
        cell.border = Border(left=st['mid'] if col in st.get('groups', ()) else None, bottom=st['hair'])
        cell.fill = fill or st['head_fill']


def _title_row(ws, row, text, ncols, st):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    _paint_row(ws, row, st, 'section', cols=range(1, ncols + 1), height=22)
    ws.cell(row=row, column=1).value = text
    ws.cell(row=row, column=1).font = st['bold']
    ws.cell(row=row, column=1).alignment = st['left']
    return row + 1


def _value_row(ws, row, label, vals, fmts, st, bold=False, colors=None):
    """寫一列：A 欄標籤，B 起數值；fmts 對應每個值的 number_format；colors 對應字色（None＝預設）。
    bold＝合計列（淡藍底、上方中粗線），其餘為斑馬紋資料列。"""
    from openpyxl.styles import Font
    _paint_row(ws, row, st, 'total' if bold else 'data', cols=range(1, len(vals) + 2),
               group_starts=st.get('groups', ()))
    c = ws.cell(row=row, column=1)
    c.value, c.font = label, st['bold'] if bold else st['font']
    c.alignment = st['left'] if bold else st['indent']
    for j, v in enumerate(vals):
        cell = ws.cell(row=row, column=2 + j)
        cell.value, cell.number_format = v, fmts[j]
        color = colors[j] if colors else None
        cell.font = Font(name='Arial', size=12, bold=True, color=color) if color else (st['bold'] if bold else st['font'])
    return row + 1


def _vs_color(v, avg):
    """高於門市平均綠色、低於紅色。"""
    if v is None or avg is None:
        return None
    return '2E7D32' if v > avg else ('C00000' if v < avg else None)


def _period_label(name, s, e):
    return f'{name} {s:%m/%d}~{e:%m/%d}'


# ── 22. 主機活動 ──
def fill_sheet_promo(wb, df_cur, sacare_prices, dates: dict):
    print(f'  分析頁：{PROMO_SHEET}', flush=True)
    ws = _reset_detail_sheet(wb, PROMO_SHEET, freeze='B3')
    ws.column_dimensions['A'].width = 64
    st = _detail_styles()
    st['groups'] = (2, 4)
    sa_codes = set(sacare_prices.keys())
    wk = _non_sa(period(df_cur, dates['wk_start'], dates['wk_end']), sa_codes)
    mo = _non_sa(period(df_cur, dates['mo_start'], dates['mo_end']), sa_codes)

    def label(d):
        sub = d['活動細項'] if '活動細項' in d.columns else pd.Series('', index=d.index)
        main = d['促銷活動名稱'] if '促銷活動名稱' in d.columns else pd.Series('', index=d.index)
        sub = sub.fillna('').astype(str).str.strip()
        main = main.fillna('').astype(str).str.strip()
        return sub.where(sub != '', main).where(lambda s: s != '', '無活動')

    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    _head_cells(ws, 1, 1, ['活動細項'], st)
    for gi, (name, s, e) in enumerate((('本週', dates['wk_start'], dates['wk_end']),
                                       ('本月', dates['mo_start'], dates['mo_end']))):
        c0 = 2 + gi * 2
        ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 1)
        _head_cells(ws, 1, c0, [_period_label(name, s, e), ''], st)
        _head_cells(ws, 2, c0, ['台數', '占比'], st, fill=st['group_fills'][gi])
    ws.row_dimensions[1].height = 26
    row = 3
    for name, sel, c3 in ANALYSIS_LINES:
        mw, mm = _line_mask(wk, sel, c3), _line_mask(mo, sel, c3)
        tw, tm = _net_units(wk, mw), _net_units(mo, mm)
        lw, lm = label(wk), label(mo)
        acts = sorted(set(lw[mw]) | set(lm[mm]))
        stats = [(a, _net_units(wk, mw & (lw == a)), _net_units(mo, mm & (lm == a))) for a in acts]
        stats = [s for s in stats if s[1] or s[2]]
        stats.sort(key=lambda s: (s[0] == '無活動', -s[2], -s[1], s[0]))
        _paint_row(ws, row, st, 'section', cols=range(1, 6), group_starts=st['groups'], height=22)
        for col, v in ((1, f'{name}（品類合計）'), (2, tw), (3, 1 if tw else None), (4, tm), (5, 1 if tm else None)):
            cell = ws.cell(row=row, column=col)
            cell.value, cell.font = v, st['bold']
            cell.number_format = '0.0%' if col in (3, 5) else '#,##0'
        row += 1
        for a, uw, um in stats:
            row = _value_row(ws, row, a, [uw, safe_rate(uw, tw) if tw else None, um, safe_rate(um, tm) if tm else None],
                             ['#,##0', '0.0%', '#,##0', '0.0%'], st)
    ws.cell(row=row + 1, column=1).value = '活動細項取 ERP 子活動說明；台數含尾款、扣除退貨，同第 1 頁與商品明細比較頁口徑。'
    ws.cell(row=row + 1, column=1).font = st['font']


# ── 23. 新機上市追蹤 ──
def fill_sheet_launch(wb, df_track, dates: dict, fiscal_start=None):
    print(f'  分析頁：{LAUNCH_SHEET}', flush=True)
    ws = _reset_detail_sheet(wb, LAUNCH_SHEET, freeze='A1')
    for col, width in zip('ABCDEFGH', (10, 16, 11, 11, 16, 11, 11, 11)):
        ws.column_dimensions[col].width = width
    st = _detail_styles()
    st['groups'] = (2, 5, 8)
    report_end = pd.Timestamp(dates['wk_end'])
    d = df_track if df_track is not None else pd.DataFrame(columns=['單據日期'])

    def first_sale(mask):
        x = d.loc[mask & d['交易類型'].isin(SALE_TYPES) & (d['數量'] > 0), '單據日期']
        return x.min().normalize() if len(x) else None

    def weekly(mask, start, upto):
        out = []
        for k in range(LAUNCH_WEEKS):
            s = start + pd.Timedelta(days=7 * k)
            e = s + pd.Timedelta(days=6)
            if e > upto:
                out.append((s, e, None))
                continue
            dm = mask & (d['單據日期'] >= s) & (d['單據日期'] < e + pd.Timedelta(days=1))
            out.append((s, e, _net_units(d, dm)))
        return out

    row = 1
    for new_name, new_c6, c3, old_name, old_c6 in LAUNCH_PAIRS:
        base = ((d['類別3代碼'] == c3) | d['品牌代碼'].isin(CERT_BRANDS)) if len(d) else None
        new_mask = base & d['類別6代碼'].isin(new_c6) if len(d) else None
        old_mask = base & d['類別6代碼'].isin(old_c6) if len(d) else None
        n0 = first_sale(new_mask) if len(d) else None
        o0 = first_sale(old_mask) if len(d) else None
        title = (f"{new_name}（首賣 {n0:%Y/%m/%d}）" if n0 is not None else f'{new_name}（尚未開賣）') + \
                (f"  vs  {old_name}（首賣 {o0:%Y/%m/%d}）" if o0 is not None else f'  vs  {old_name}（查無銷售）')
        row = _title_row(ws, row, title, 8, st)
        _head_cells(ws, row, 1, ['上市週', '今年期間', '今年台數', '今年累計', '對照期間', '對照台數', '對照累計', '累計比'], st)
        row += 1
        if n0 is None:
            row += 1
            continue
        new_w = weekly(new_mask, n0, report_end)
        old_w = weekly(old_mask, o0, report_end) if o0 is not None else [(None, None, None)] * LAUNCH_WEEKS
        cum_n = cum_o = 0
        for k in range(LAUNCH_WEEKS):
            s, e, u = new_w[k]
            os_, oe, ou = old_w[k]
            if u is not None:
                cum_n += u
            if ou is not None:
                cum_o += ou
            ratio = safe_rate(cum_n, cum_o) if (u is not None and ou is not None and cum_o) else None
            vals = [f'{s:%m/%d}~{e:%m/%d}', u, cum_n if u is not None else None,
                    f'{os_:%Y/%m/%d}' if os_ is not None else None, ou, cum_o if ou is not None else None, ratio]
            row = _value_row(ws, row, f'第 {k + 1} 週', vals,
                             ['@', '#,##0', '#,##0', '@', '#,##0', '#,##0', '0%'], st)
        row += 1

    # 無對照機種的新機：追 AppleCare+ 搭售率（套數 ÷ 台數，ACPP 依品名關鍵字對機種）
    st['groups'] = (3, 6)
    fmts = ['@', '#,##0', '#,##0', '0.0%', '#,##0', '#,##0', '0.0%']
    for name, c6, c3, kw in LAUNCH_ACPP:
        dev_mask = ((d['類別3代碼'] == c3) | d['品牌代碼'].isin(CERT_BRANDS)) & d['類別6代碼'].isin(c6) if len(d) else None
        ac_mask = (d['類別3代碼'] == 3032.0) & d['名稱'].astype(str).str.lower().str.contains(kw, na=False) if len(d) else None
        n0 = first_sale(dev_mask) if len(d) else None
        title = (f'{name} AppleCare+ 搭售率（首賣 {n0:%Y/%m/%d}；只算 AppleCare+，不含 SAcare）' if n0 is not None
                 else f'{name} AppleCare+ 搭售率（尚未開賣）')
        row = _title_row(ws, row, title, 8, st)
        _head_cells(ws, row, 1, ['週次', '期間', '台數', 'AppleCare+ 套數', '搭售率', '累計台數', '累計套數', '累計搭售率'], st)
        row += 1
        if n0 is None:
            row += 1
            continue
        # 期間跟週報週數走（週日～週六）：從首賣日所在的週報週起，到本週為止，最多 LAUNCH_WEEKS 週
        s = n0 - pd.Timedelta(days=(n0 - pd.Timestamp(dates['wk_start'])).days % 7)
        cum_u = cum_a = 0
        for k in range(LAUNCH_WEEKS):
            e = s + pd.Timedelta(days=6)
            if e > report_end:
                break
            wm = (d['單據日期'] >= s) & (d['單據日期'] < e + pd.Timedelta(days=1))
            u, a = _net_units(d, dev_mask & wm), _net_units(d, ac_mask & wm)
            cum_u += u
            cum_a += a
            if fiscal_start:
                q, w = compute_fiscal_week_number(s.date(), fiscal_start)
                label = f'Q{q}W{w:02d}'
            else:
                label = f'第 {k + 1} 週'
            vals = [f'{s:%m/%d}~{e:%m/%d}', u, a, safe_rate(a, u) if u else None,
                    cum_u, cum_a, safe_rate(cum_a, cum_u) if cum_u else None]
            row = _value_row(ws, row, label, vals, fmts, st)
            s += pd.Timedelta(days=7)
        tm = (d['單據日期'] >= n0) & (d['單據日期'] < report_end + pd.Timedelta(days=1))
        u, a = _net_units(d, dev_mask & tm), _net_units(d, ac_mask & tm)
        row = _value_row(ws, row, '上市至今', [f'{n0:%m/%d}~{report_end:%m/%d}', u, a, safe_rate(a, u) if u else None,
                                               None, None, None], fmts, st, bold=True)
        row += 1
    ws.cell(row=row, column=1).value = ('上市週從首賣日起每 7 天一週；只列已過完的週。對照期間為對照機種首賣後同一週的起始日。'
                                        '累計比＝今年累計 ÷ 對照累計。AppleCare+ 搭售率區塊的期間跟週報週數走（週日～週六），'
                                        '從首賣日所在的週起算到本週。')
    ws.cell(row=row, column=1).font = st['font']


# ── 24. 個人客單價 ──
def fill_sheet_emp_aov(wb, df_cur, sacare_prices, dates: dict):
    print(f'  分析頁：{EMP_AOV_SHEET}', flush=True)
    ws = _reset_detail_sheet(wb, EMP_AOV_SHEET, freeze='B3')
    ws.column_dimensions['A'].width = 22
    st = _detail_styles()
    st['groups'] = (2, 5, 8, 11)
    sa_codes = set(sacare_prices.keys())
    groups = [('原廠（主機＋原廠配件）', [3001.0, 3002.0]), ('3PP', [3003.0])]

    def stats(f_all, f_emp, c3s):
        deals = _doc_set(f_all, f_all['交易類型'] == '銷售') - _doc_set(f_all, f_all['交易類型'] == '銷退')
        nf = _non_sa(f_emp, sa_codes)
        m = nf['單據代碼'].astype(str).isin(deals) & (nf['交易類型'] == '銷售') & nf['類別3代碼'].isin(c3s)
        docs = _doc_set(nf, m)
        rev = int(nf.loc[m, 'NET'].sum())
        return len(docs), rev, (int(round(rev / len(docs))) if docs else None)

    periods = [('本週', dates['wk_start'], dates['wk_end']), ('本月', dates['mo_start'], dates['mo_end'])]
    # 表頭：員工 | 本週（原廠 筆數/業績/客單、3PP 筆數/業績/客單）| 本月（同）
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    _head_cells(ws, 1, 1, ['員工'], st)
    for pi, (pname, s, e) in enumerate(periods):
        for gi, (gname, _) in enumerate(groups):
            c0 = 2 + pi * 6 + gi * 3
            ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + 2)
            _head_cells(ws, 1, c0, [f'{_period_label(pname, s, e)}｜{gname}', '', ''], st)
            _head_cells(ws, 2, c0, ['筆數', '業績', '客單價'], st, fill=st['group_fills'][pi])
    for c in range(2, 14):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 12
    ws.row_dimensions[1].height = 36
    frames = [period(df_cur, s, e) for _, s, e in periods]
    store = [stats(f, f, c3s) for f in frames for _, c3s in groups]
    row = 3
    people = [(_emp_label(code, name), code) for code, name in EMPLOYEES] + [('門市合計', None)]
    for label, code in people:
        vals, colors = [], []
        for pi, f in enumerate(frames):
            f_emp = f if code is None else f[f['員工代碼'] == code]
            for gi, (_, c3s) in enumerate(groups):
                n, rev, aov = stats(f, f_emp, c3s)
                vals += [n, rev, aov]
                colors += [None, None, None if label == '門市合計' else _vs_color(aov, store[pi * 2 + gi][2])]
        row = _value_row(ws, row, label, vals, ['#,##0'] * 12, st, bold=label == '門市合計', colors=colors)

    # 類別 4 細項平均單價（本月）：原廠、3PP 各取門市本月營業額前幾名的類別 4
    mo = _non_sa(frames[1], sa_codes)
    for gname, c3s, top in (('原廠', [3001.0, 3002.0], 8), ('3PP', [3003.0], 12)):
        g = mo[mo['類別3代碼'].isin(c3s) & mo['交易類型'].isin(['銷售', '銷退'])]
        rev = g.groupby('類別4代碼')['NET'].sum().sort_values(ascending=False)
        c4s = [c for c, v in rev.items() if v > 0][:top]
        row += 1
        st['groups'] = ()
        row = _title_row(ws, row, f'本月 類別 4 細項平均單價｜{gname}（金額÷件數，不含贈品／訂金／尾款；綠＝高於門市、紅＝低於門市）',
                         max(len(c4s) + 1, 13), st)
        _head_cells(ws, row, 1, ['員工'] + [_c4_name(c) for c in c4s], st)
        ws.row_dimensions[row].height = 36
        row += 1
        store_avg = [_avg_price(g[g['類別4代碼'] == c]) for c in c4s]
        for label, f_emp in _emp_frames(g):
            vals = [_avg_price(f_emp[f_emp['類別4代碼'] == c]) for c in c4s]
            colors = None if label == '門市合計' else [_vs_color(v, a) for v, a in zip(vals, store_avg)]
            row = _value_row(ws, row, label, vals, ['#,##0'] * len(vals), st, bold=label == '門市合計', colors=colors)
        qty = [int(g.loc[(g['類別4代碼'] == c) & ~_gift_mask(g), '數量'].sum()) for c in c4s]
        row = _value_row(ws, row, '門市件數', qty, ['#,##0'] * len(qty), st)
    ws.cell(row=row + 1, column=1).value = ('筆數＝成交單（有銷售列且非整單退貨，同經營結構比較頁）中含該類商品的單數；業績＝這些單裡該類商品的金額；'
                                            '客單價＝業績÷筆數。綠＝高於門市、紅＝低於門市。')
    ws.cell(row=row + 1, column=1).font = st['font']


# ── 25. 個人保固搭售率 ──
def fill_sheet_emp_warranty(wb, df_cur, sacare_prices, dates: dict):
    print(f'  分析頁：{EMP_WARRANTY_SHEET}', flush=True)
    ws = _reset_detail_sheet(wb, EMP_WARRANTY_SHEET, freeze='B1')
    ws.column_dimensions['A'].width = 22
    for c in range(2, 2 + 4 * len(WARRANTY_DEVICES)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 9
    st = _detail_styles()
    st['groups'] = tuple(2 + 4 * i for i in range(len(WARRANTY_DEVICES)))
    row = 1
    for pname, s, e in (('本週', dates['wk_start'], dates['wk_end']), ('本月', dates['mo_start'], dates['mo_end'])):
        f = period(df_cur, s, e)
        row = _title_row(ws, row, f'{_period_label(pname, s, e)}（搭售率＝(AppleCare＋SA Care) 套數 ÷ 主機台數；綠＝高於門市、紅＝低於門市）',
                         1 + 4 * len(WARRANTY_DEVICES), st)
        _head_cells(ws, row, 1, ['員工'], st)
        for i, (dev, *_rest) in enumerate(WARRANTY_DEVICES):
            ws.merge_cells(start_row=row, start_column=2 + i * 4, end_row=row, end_column=5 + i * 4)
            _head_cells(ws, row, 2 + i * 4, [dev, '', '', ''], st)
        row += 1
        _head_cells(ws, row, 1, [''], st)
        for i in range(len(WARRANTY_DEVICES)):
            _head_cells(ws, row, 2 + i * 4, ['台數', 'AC', 'SA', '搭售率'], st, fill=st['group_fills'][i % 3])
        row += 1
        store = calc_metrics(f, sacare_prices)
        store_rate = [safe_rate(store[ac] + store[sa], store[u]) if store[u] else None
                      for _, u, ac, sa in WARRANTY_DEVICES]
        for label, f_emp in _emp_frames(f):
            m = store if label == '門市合計' else calc_metrics(f_emp, sacare_prices)
            vals, fmts, colors = [], [], []
            for i, (_, u, ac, sa) in enumerate(WARRANTY_DEVICES):
                rate = safe_rate(m[ac] + m[sa], m[u]) if m[u] else None
                vals += [m[u], m[ac], m[sa], rate]
                fmts += ['#,##0', '#,##0', '#,##0', '0%']
                colors += [None, None, None, None if label == '門市合計' else _vs_color(rate, store_rate[i])]
            row = _value_row(ws, row, label, vals, fmts, st, bold=label == '門市合計', colors=colors)
        row += 1
    ws.cell(row=row, column=1).value = '台數、AppleCare（ACPP）與 SA Care 套數同商品明細比較頁「保固搭售明細」口徑，依結帳員工歸屬，退貨扣回。'
    ws.cell(row=row, column=1).font = st['font']


# ── 26. 個人業績組成 ──
def fill_sheet_emp_mix(wb, df_cur, sacare_prices, dates: dict):
    print(f'  分析頁：{EMP_MIX_SHEET}', flush=True)
    ws = _reset_detail_sheet(wb, EMP_MIX_SHEET, freeze='B1')
    ws.column_dimensions['A'].width = 22
    for c in range(2, 11):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 13
    st = _detail_styles()
    st['groups'] = (2, 7)
    cats = [('主機', 'rev_3001'), ('原廠配件（含 AirPods）', 'rev_3002'), ('3PP', 'rev_3003'), ('保固（AC+SA）', 'warranty_rev')]
    row = 1
    for pname, s, e in (('本週', dates['wk_start'], dates['wk_end']), ('本月', dates['mo_start'], dates['mo_end'])):
        f = period(df_cur, s, e)
        row = _title_row(ws, row, f'{_period_label(pname, s, e)}（占比：綠＝高於門市、紅＝低於門市）', 10, st)
        _head_cells(ws, row, 1, ['員工'] + [c for c, _ in cats] + ['合計', '主機 占比', '原廠配件 占比', '3PP 占比', '保固 占比'], st)
        ws.row_dimensions[row].height = 36
        row += 1
        store = calc_metrics(f, sacare_prices)
        store_tot = sum(store[k] for _, k in cats)
        store_share = [safe_rate(store[k], store_tot) if store_tot else None for _, k in cats]
        for label, f_emp in _emp_frames(f):
            m = store if label == '門市合計' else calc_metrics(f_emp, sacare_prices)
            amts = [m[k] for _, k in cats]
            tot = sum(amts)
            shares = [safe_rate(a, tot) if tot else None for a in amts]
            colors = [None] * 5 + ([None] * 4 if label == '門市合計' else [_vs_color(v, a) for v, a in zip(shares, store_share)])
            row = _value_row(ws, row, label, amts + [tot] + shares, ['#,##0'] * 5 + ['0.0%'] * 4, st,
                             bold=label == '門市合計', colors=colors)
        row += 1
    ws.cell(row=row, column=1).value = '金額同第 6 頁口徑（含訂金、扣除退貨；SA Care 以價目表計），依結帳員工歸屬。'
    ws.cell(row=row, column=1).font = st['font']


# ── 27. 個人近 8 週走勢 ──
def fill_sheet_emp_trend(wb, df_cur, sacare_prices, dates: dict, bundle_weeks=None, fiscal_start=None):
    """bundle_weeks：[(週起日, 搭售統計 rows)…] 由舊到新 8 週；None＝搭售率留空。"""
    print(f'  分析頁：{EMP_TREND_SHEET}', flush=True)
    ws = _reset_detail_sheet(wb, EMP_TREND_SHEET, freeze='C2')
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 14
    for c in range(3, 13):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 13
    st = _detail_styles()
    starts = [dates['wk_start'] - timedelta(days=7 * i) for i in range(7, -1, -1)]
    labels = []
    for s in starts:
        wk = f'W{compute_fiscal_week_number(s, fiscal_start)[1]}\n' if fiscal_start else ''
        labels.append(f'{wk}{s:%m/%d}~{s + timedelta(days=6):%m/%d}')
    st['groups'] = (3, 11)
    _head_cells(ws, 1, 1, ['員工', '項目'] + labels + ['8 週平均', '本週 vs 平均'], st)
    ws.row_dimensions[1].height = 36
    frames = [period(df_cur, s, s + timedelta(days=6)) for s in starts]
    sa_codes = set(sacare_prices.keys())
    bundle_by_week = None
    if bundle_weeks:
        bundle_by_week = [{r['empId']: r for r in rows} for _, rows in bundle_weeks]

    def bundle_rate(i, code):
        if bundle_by_week is None:
            return None
        rows = list(bundle_by_week[i].values()) if code is None else [bundle_by_week[i].get(code)]
        rows = [r for r in rows if r]
        sm = sum(r[k]['m'] for r in rows for k in ('cpu', 'iphone', 'ipad', 'watch', 'airpods'))
        ss = sum(r[k]['s'] for r in rows for k in ('cpu', 'iphone', 'ipad', 'watch', 'airpods'))
        return round(ss / sm, 2) if sm else None

    def warranty_rate(m):
        units = sum(m[u] for _, u, _, _ in WARRANTY_DEVICES)
        sets = sum(m[ac] + m[sa] for _, _, ac, sa in WARRANTY_DEVICES)
        return safe_rate(sets, units) if units else None

    # (項目, 數值格式, 差值格式, 取值)；m＝該人該週的 calc_metrics（每人每週只算一次）
    items = [('業績', '#,##0', '[Color10]+#,##0;[Red]-#,##0;0', lambda i, f, m, code: m['total_rev']),
             ('3PP 搭售率', '0.00', '[Color10]+0.00;[Red]-0.00;0.00', lambda i, f, m, code: bundle_rate(i, code)),
             ('保固搭售率', '0.0%', '[Color10]+0.0"pp";[Red]-0.0"pp";0.0"pp"', lambda i, f, m, code: warranty_rate(m)),
             ('3PP 平均單價', '#,##0', '[Color10]+#,##0;[Red]-#,##0;0',
              lambda i, f, m, code: _avg_price(_non_sa(f, sa_codes)[lambda x: x['類別3代碼'] == 3003.0]))]
    from openpyxl.styles import Border
    row = 2
    people = [(_emp_label(code, name), code) for code, name in EMPLOYEES] + [('門市合計', None)]
    for pi, (label, code) in enumerate(people):
        fs = [f if code is None else f[f['員工代碼'] == code] for f in frames]
        ms = [calc_metrics(f, sacare_prices) for f in fs]
        # 每人一組 4 列同底色、員工之間交替；門市合計用淡藍；組與組之間中粗線
        fill = st['total_fill'] if code is None else (st['zebra_fill'] if pi % 2 else st['no_fill'])
        for k, (item, fmt, diff_fmt, fn) in enumerate(items):
            vals = [fn(i, f, m, code) for i, (f, m) in enumerate(zip(fs, ms))]
            got = [v for v in vals if v is not None]
            avg = sum(got) / len(got) if got else None
            cur = vals[-1]
            diff = (cur - avg) if (cur is not None and avg is not None) else None
            if diff is not None:   # 比率差以百分點顯示；其餘四捨五入，避免非 Excel 的預覽顯示一長串小數
                diff = round(diff * 100, 1) if fmt == '0.0%' else (round(diff, 2) if fmt == '0.00' else int(round(diff)))
            cells = [label if k == 0 else '', item] + vals + [avg, diff]
            for j, v in enumerate(cells):
                cell = ws.cell(row=row, column=1 + j)
                cell.value, cell.fill = v, fill
                cell.border = Border(top=st['mid'] if k == 0 else None, bottom=st['hair'],
                                     left=st['mid'] if j in (2, 10) else None)
                cell.font = st['cur'] if j == 9 else (st['bold'] if (code is None or j <= 1) else st['font'])
                if j <= 1:
                    cell.alignment = st['left']
                else:
                    cell.number_format = diff_fmt if j == 11 else fmt
            ws.row_dimensions[row].height = 20
            row += 1
        # 不合併 A 欄：直向合併在 Mac 預覽（QuickLook／Numbers）會把後面幾列的格子往右錯位
    ws.cell(row=row + 1, column=1).value = ('業績同第 6 頁總營業額口徑；3PP 搭售率同第 10 頁（配件數÷搭售台數）；'
                                            '保固搭售率＝(AppleCare＋SA Care) 套數÷主機台數（同個人保固搭售率頁）；'
                                            '3PP 平均單價＝3PP 金額÷件數（不含贈品／訂金／尾款）。最右欄＝本週減 8 週平均。')
    ws.cell(row=row + 1, column=1).font = st['font']


# ── MySetup 提交細項 ──
MYSETUP_DETAIL_DEVICES = ('Mac', 'iPad', 'iPhone', 'Watch')


def fill_sheet_mysetup_detail(wb, blocks=None):
    """blocks：[(期間標題, rows)…]；row＝{label, units: {機種: 台數 或 None}, subs: {機種: 筆數}, total}，
    最後一列為門市合計。blocks 為 None＝本次沒有 Personal Setup 匯出檔。"""
    print(f'  分析頁：{MYSETUP_DETAIL_SHEET}', flush=True)
    ws = _reset_detail_sheet(wb, MYSETUP_DETAIL_SHEET, freeze='B1')
    ws.column_dimensions['A'].width = 24
    for c in range(2, 17):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 9
    st = _detail_styles()
    if not blocks:
        ws.cell(row=1, column=1).value = '本次未提供 Personal Setup 匯出檔，沒有 MySetup 資料。'
        ws.cell(row=1, column=1).font = st['bold']
        return
    groups = list(MYSETUP_DETAIL_DEVICES) + ['合計']
    st['groups'] = tuple(2 + 3 * i for i in range(len(groups)))

    def split(r):
        """每組 (台數, 提交, 提交率)；台數未知（對不上的名字）時台數與提交率留空。"""
        out = []
        for g in groups:
            if g == '合計':
                known = [r['units'].get(d) for d in MYSETUP_DETAIL_DEVICES]
                u = sum(x for x in known if x) if any(x is not None for x in known) else None
                n = sum(r['subs'].values())
            else:
                u, n = r['units'].get(g), r['subs'].get(g, 0)
            out.append((u, n, safe_rate(n, u) if u else None))
        return out

    row = 1
    for title, rows in blocks:
        row = _title_row(ws, row, f'{title}（提交率＝MySetup 提交數÷主機台數，台數同週報第 13、14 頁；綠＝高於門市、紅＝低於門市）',
                         1 + 3 * len(groups), st)
        _head_cells(ws, row, 1, ['同仁'], st)
        for i, g in enumerate(groups):
            ws.merge_cells(start_row=row, start_column=2 + i * 3, end_row=row, end_column=4 + i * 3)
            _head_cells(ws, row, 2 + i * 3, [g, '', ''], st)
        row += 1
        _head_cells(ws, row, 1, [''], st)
        for i in range(len(groups)):
            _head_cells(ws, row, 2 + i * 3, ['台數', '提交', '提交率'], st, fill=st['group_fills'][i % 3])
        row += 1
        store = split(rows[-1])
        for r in rows:
            vals, fmts, colors = [], [], []
            for i, (u, n, rate) in enumerate(split(r)):
                vals += [u, n or None, rate]
                fmts += ['#,##0', '#,##0', '0%']
                colors += [None, None, None if r.get('total') else _vs_color(rate, store[i][2])]
            row = _value_row(ws, row, r['label'], vals, fmts, st, bold=bool(r.get('total')), colors=colors)
        row += 1
    ws.cell(row=row, column=1).value = ('提交數依 Personal Setup 的 Submission Time 與 Device 欄分週／月、分機種；'
                                        '名字對應同週報第 13、14 頁（含人工對照），對不上的名字列在合計上方。')
    ws.cell(row=row, column=1).font = st['font']


# ─── Main ─────────────────────────────────────────────────────────────────────
def main():
    args = parse_args()

    # --version 模式：只印版本號後離開
    if args.version:
        print(VERSION)
        sys.exit(0)

    # 確認必要參數都有傳入
    for attr in ['week_start', 'week_end', 'data_dir', 'template']:
        if not getattr(args, attr):
            sys.exit(f'ERROR: 缺少必要參數 --{attr.replace("_","-")}')
    if not args.output and not args.output_dir:
        sys.exit('ERROR: 缺少必要參數 --output 或 --output-dir')

    wk_start = date.fromisoformat(args.week_start)
    wk_end   = date.fromisoformat(args.week_end)
    data_dir = Path(args.data_dir).resolve()
    template = Path(args.template).resolve()

    # 解析輸出路徑（若提供 --output-dir，自動產生帶季週次的檔名）
    if args.output:
        output = Path(args.output).resolve()
    else:
        out_dir = Path(args.output_dir).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        _wb_tmp = openpyxl.load_workbook(template, read_only=True, data_only=True)
        _fys = load_fiscal_year_start(_wb_tmp)
        _wb_tmp.close()
        if _fys:
            _q, _w = compute_fiscal_week_number(wk_start, _fys)
            fname = f'Q{_q}W{_w:02d}_週報_{wk_start}~{wk_end.strftime("%m-%d")}.xlsx'
        else:
            fname = f'週報_{wk_start}~{wk_end.strftime("%m-%d")}.xlsx'
        output = out_dir / fname

    print(f'=== 週報 Excel 填充  v{VERSION} ===')
    print(f'  本週: {wk_start} ~ {wk_end}')

    # ── 期間計算 ─────────────────────────────────────────────────────────────
    year, prev_year = wk_start.year, wk_start.year - 1
    cross_month = wk_end.month != wk_start.month

    report_month = wk_start.month
    prev_month   = report_month - 1 if report_month > 1 else 12
    prev_mo_year = year if report_month > 1 else year - 1

    if cross_month:
        mo_start = date(year, report_month, 1)
        mo_end   = date(year, report_month, monthrange(year, report_month)[1])
        lm_start = date(prev_mo_year, prev_month, 1)
        lm_end   = date(prev_mo_year, prev_month, monthrange(prev_mo_year, prev_month)[1])
        ly_start = date(prev_year, report_month, 1)
        ly_end   = date(prev_year, report_month, monthrange(prev_year, report_month)[1])
    else:
        mo_start = date(year, report_month, 1)
        mo_end   = wk_end
        lm_last  = monthrange(prev_mo_year, prev_month)[1]
        lm_start = date(prev_mo_year, prev_month, 1)
        lm_end   = date(prev_mo_year, prev_month, min(wk_end.day, lm_last))
        ly_last  = monthrange(prev_year, report_month)[1]
        ly_start = date(prev_year, report_month, 1)
        ly_end   = date(prev_year, report_month, min(wk_end.day, ly_last))

    prev_wk_end   = wk_start - timedelta(days=1)
    prev_wk_start = prev_wk_end - timedelta(days=6)
    # quarter_start is computed after loading template (needs fiscal_year_start from 設定 sheet)

    dates = dict(
        wk_start=wk_start, wk_end=wk_end,
        prev_wk_start=prev_wk_start, prev_wk_end=prev_wk_end,
        mo_start=mo_start, mo_end=mo_end,
        lm_start=lm_start, lm_end=lm_end,
        ly_start=ly_start, ly_end=ly_end,
        ytd_cur_start=date(year, 1, 1),
        ytd_prv_start=date(prev_year, 1, 1),
    )
    if args.yoy_end:
        dates['yoy_end'] = date.fromisoformat(args.yoy_end)
        print(f'  年對年截止日(自訂): {dates["yoy_end"]}')

    print(f'  本月: {mo_start} ~ {mo_end}')
    print(f'  上月: {lm_start} ~ {lm_end}')
    print(f'  去年: {ly_start} ~ {ly_end}')

    # ── 找資料檔 ─────────────────────────────────────────────────────────────
    file_cur  = find_800ab(data_dir, year)
    file_prev = find_800ab(data_dir, prev_year)
    if not file_cur:  sys.exit(f'ERROR: 找不到 800AB_{year} 檔案')
    if not file_prev: sys.exit(f'ERROR: 找不到 800AB_{prev_year}整年 檔案')

    sacare_path = find_sacare(data_dir)
    if not sacare_path: sys.exit('ERROR: 找不到 SAcare對應價目表.xlsx')

    print(f'\n  本年: {Path(file_cur).name}')
    print(f'  去年: {Path(file_prev).name}')

    # ── 載入資料 ─────────────────────────────────────────────────────────────
    print('\nLoading data...', flush=True)
    sacare_prices = load_sacare(sacare_path)
    df_cur  = load_800(file_cur)
    df_prev = load_800(file_prev)

    # ── 開啟模板 ─────────────────────────────────────────────────────────────
    print('\nFilling sheets...', flush=True)
    shutil.copy(template, output)
    wb = openpyxl.load_workbook(output)

    # 從模板的「設定」工作表動態讀取員工清單 & 年度起始日
    global EMPLOYEES, EMP_CODES
    loaded = load_employees(wb)
    if loaded:
        EMPLOYEES = loaded
        EMP_CODES  = [e[0] for e in EMPLOYEES]
        print(f'  員工設定: 從模板「設定」工作表讀取，共 {len(EMPLOYEES)} 人')
    else:
        print(f'  員工設定: 未找到「設定」工作表，使用預設清單（{len(EMPLOYEES)} 人）')

    fiscal_year_start = resolve_fiscal_year_start(wk_start, load_fiscal_year_start(wb))
    if fiscal_year_start:
        q_num, w_num = compute_fiscal_week_number(wk_start, fiscal_year_start)
        print(f'  財務週次: 年度起始 {fiscal_year_start}，本週為 Q{q_num} W{w_num:02d}')
    else:
        print(f'  財務週次: 未設定年度起始日，使用曆法季度')

    quarter_start = compute_quarter_start(wk_start, fiscal_year_start)

    fill_sheet1(wb['1.主機銷售台數'], df_cur, quarter_start, wk_end)
    fill_sheet_edu(wb['2.教育價'], df_cur, quarter_start, wk_end)
    fill_sheet2(wb['6.門市週報'],   df_cur, df_prev, sacare_prices, dates)
    fill_sheet_focus(wb['3.每月重點'], df_cur, dates)
    fill_sheet_speakers(wb['4.Speakers'], df_cur, dates)
    fill_sheet_traffic(wb['5.門市人流'], {}, dates)   # CLI 無 ShopperTrak 來源 → 只寫標題
    fill_sheet3(wb['7.3PP配件比較'], df_cur, df_prev, sacare_prices, dates)
    fill_sheet45(wb['8.3PP 銷售排名'], wb['9.VAP銷售排名'], df_cur, sacare_prices, dates)
    fill_sheet6(wb['12.個人新制獎金'], df_cur, sacare_prices, dates)
    fill_sheet78(wb['13.個人週主機'], wb['14.個人月主機'], df_cur, sacare_prices, dates)
    fill_sheet9(wb['15.個人月3PP'], df_cur, sacare_prices, dates)
    fill_sheet10(wb['16.月報YOY'], df_cur, df_prev, sacare_prices, dates)
    fill_sheet11(wb['17.3PP YOY'], df_cur, df_prev, sacare_prices, dates)
    fill_sheet_detail_products(wb, df_cur, df_prev, sacare_prices, dates)
    fill_sheet_tpp_detail(wb, df_cur, df_prev, sacare_prices, dates)
    fill_sheet_detail_ops(wb, df_cur, df_prev, sacare_prices, dates)
    fill_sheet_hot_series(wb, df_cur, sacare_prices, dates)

    wb.save(output)
    print(f'\n✓ 完成: {output}')

if __name__ == '__main__':
    main()
