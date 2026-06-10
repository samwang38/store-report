# EPB 業績計算邏輯總整理

本文件整理 live-store-report-app 從 800AB ERP（EPB）抓取業績資料的完整邏輯：連線方式、資料表、欄位標準化、KPI 計算公式、業務慣例。供其他專案調用或對帳時參考。

程式碼對照：查詢層在 `server.py`，計算層在 `fill_weekly_excel.py`（下稱「引擎」）。

---

## 1. 連線與查詢層

### 架構
```
Python (server.py run_remote)
  └─ subprocess 啟動 JVM → EPBReportQuery.java
       └─ SOAP pullRowSetStream → EPB AP 服務（內網 192.168.1.177:8080）
            回傳 zlib 壓縮的 Java 序列化 RowSet → Java 端反序列化 → tab 分隔文字 stdout
```
- **不能純 Python 直連**：回應是 Java 序列化物件，必須由 JVM 解。
- 免帳密（系統帳號執行）、需公司內網/VPN。
- 可覆寫設定：`EPB_WSDL_URL`、`EPB_JAVA_HOME` 環境變數，或 `local_config.json` 的 `epb.wsdlUrl`。
- SQL 為 **Oracle 語法**；字串一律經 `quote_sql()` 跳脫。
- 效能慣例：多個日期區間用 OR 合併成單一 SQL（一次 JVM/SOAP 約 20 秒，大頭在伺服器端）。

### 資料表
| 資料表 | 用途 | 關鍵欄位 |
|---|---|---|
| `poslinev_bi` | POS 銷售明細（行層級） | trans_type, doc_date, doc_id, emp_id1, stk_id, name, stk_qty, line_total_net, line_tax, cost_price, trn_cost_price, brand_id, cat1_id, cat3_id, cat4_id, cat6_id, disc_num, line_no |
| `ep_user` | 員工名稱 | `e.user_id = l.emp_id1` |
| `pos_shop` | 門市主檔 | shop_id, name |

固定條件：`org_id = '01'`；士林 = `shop_id = '004'`。
日期過濾：`doc_date >= to_date(起) and doc_date < to_date(迄+1天)`（迄日含當天整天）。

### 欄位標準化（`standardize_remote_records`）
EPB 原始欄位 → 標準中文欄位 DataFrame（與 800AB 匯出格式一致）：
- `銷售金額(含稅)` = line_total_net + line_tax
- `淨銷售金額(未稅)` = line_total_net
- `單位成本` = trn_cost_price，為 0 時退回 cost_price（**未稅**成本）
- `NET` = 銷售金額(含稅) + 銷退金額（計算用主金額欄；銷退列數量/金額本身為負）
- 交易類型代碼：A=銷售、E=銷退、G=訂金、H=尾款、J=退訂

---

## 2. 核心業務慣例（所有計算共用）

### 交易類型處理（最重要的慣例）
- `SALE_TYPES = {'銷售', '尾款'}`：**台數**在銷售完成日計入 — 現金交易看「銷售」、分期完成看「尾款」；「訂金」只計營業額、不計台數。
- **SAcare 營業額**只計 銷售/尾款，排除訂金（避免雙重計算）；銷退另扣。
- **成交筆數** = 「銷售」單據代碼集合 − 「銷退」單據代碼集合（排除正負沖銷單）。

### 商品分類體系
- **C3（類別3）大分類**：3001=Apple 主機、3002=Apple 原廠配件、3003=3PP 第三方配件、3032=ACPP（AppleCare+ 代收保費）、3033/3046=其他原廠類
- **C4（類別4）產品線**（跨年代穩定，新機上市不變）：iPhone=4004；iPad=4005/4006/4041；Watch=4038；MacBook=4002、mini+iMac=4001
- **C6（類別6）細分型號**（每代更新）：Mac 台數必須用 C6 區分 MacBook/mini/iMac — `C6_CPU` 常數（6001=iMac、6002=mini、634x=各代 MacBook），**新 Mac 世代上市只需更新引擎的 `C6_CPU`**
- **SAcare 各機種 C6**：cpu=6533、ipad=6534、iphone=6535、watch=6536、airpods=6537
- **認證機（整新機）品牌** {881,885,886,888}：計台數時 bypass C3 條件
- **VAP 品牌** {59, 224, 277}

### SAcare 特殊處理
- SAcare 品項不在 EPB 內判定，以 `銷售資料/SAcare對應價目表.xlsx` 的存貨代碼→價格對照表為準。
- SAcare 營業額 = 對照表價格 × 數量（不用 ERP 的 NET），毛利 = 營業額 ÷ 2。
- 所有「非 SA」營收計算都先剔除 SAcare 存貨代碼。

---

## 3. 門市層級 KPI（`calc_metrics`，Sheet 2/10 使用）

輸入：已用 `period(df, start, end)` 過濾日期的標準 DataFrame + SAcare 價目。

| KPI | 公式 |
|---|---|
| 總營業額 total_rev | 全部非 SA 列的 NET 合計（含禮券/雜項折抵）＋ SAcare 營業額 |
| 分類營收 rev_3001/3002/3003 | 非 SA、該 C3 的 NET 合計 |
| Apple 毛利 apl_gross | C3∈{3001,3002,3032,3033}、排除（SA 品項、品牌297、C6=31、C1=21、**尾款列**）：Σ NET − round(單位成本×1.05)×數量。排除尾款是因尾款列 NET=0 但記錄完整成本，會造成假性負毛利 |
| 3PP 毛利 tpp_gross | 同上公式，C3=3003 |
| 總毛利 | apl_gross + tpp_gross + sa_gross（SA 毛利=SA 營收÷2） |
| 台數（Mac） | C6 判定（C6_MACBOOK / mini=6002 / iMac=6001），銷售+尾款 − 銷退 |
| 台數（iPhone/iPad/Watch） | C4 判定 +（C3=3001 或認證機品牌），銷售+尾款 − 銷退 |
| ACPP-MAC | C3=3032 且名稱含 "mac"，銷售+尾款數量 |
| SAcare 各機種件數 | C6_SA 判定 + SA 存貨代碼，銷售+尾款 − 銷退 |
| 成交筆數 txn_count | 銷售單據集合 − 銷退單據集合 |
| 來客數 | ShopperTrak（外部 API，非 EPB） |
| 人均產值 | 總營業額 ÷ 編制人數（前端輸入，存 local_config.json） |

附加率慣例（Sheet 2/7/8）：SAcare 附加率 = SA 件數 ÷ 該機種台數；ACPP 附加率同理。

## 4. 個人層級 KPI（`calc_employee`，Sheet 6-9 使用）

以 `員工代碼`（emp_id1）過濾後計算。與門市層級不同的重點：

### D 欄「原廠商品營業額」（獎金基準）
非 SA、交易類型∈{銷售,尾款,銷退}，**排除**：
- C1∈{1002,1004,1008}、C3∈{3047,3003,3004,3018,3019,3012}、C6∈{6888,6889}
- 指定存貨代碼（點數折抵 888、教育價 99200202、SA Care 檢測 99901780 等內部代碼，完整清單見引擎 `SKU_EXCLUDED_FROM_APPLE`）

### H 欄「Apple 毛利（未稅）」— 對齊 ERP 報表「13-門市獎金Apple毛利額未稅-員工」
- C3 白名單 {3001,3002,3032,3033,3046} ＋ 存貨代碼 99901689（抵用券兌換）
- 交易類型：**含 銷售/訂金/銷退，排除 尾款/退訂**（800AB 訂金列記原廠全額未稅收入；尾款列收入=0）
- 公式：`淨銷售金額(未稅) − 單位成本×數量`（不乘 1.05）
- C3=3032（ACPP）特例：800AB 把含稅金額存進未稅欄位，需 ÷1.05 修正
- 額外排除碼：`SA_APPLE_GROSS_EXCL`（引擎內，含 7307154 等門號/特殊代碼）

### I 欄「3PP 毛利（未稅）」— 對齊 ERP「14-門市獎金3PP毛利額未稅-員工」
- C3 白名單 {3003,3006,3004,3018,3019,3012}
- 交易類型：**只排除尾款**（訂金/退訂/銷退皆納入 — ERP 以訂金時點認列 3PP；尾款列是贈品成本補登，ERP 不計；退訂自動與訂金相抵）
- 排除碼：`SA_CARE_GROSS_EXCL`（教育價 99200202、促銷組合 99200201、預收訂金 99500203 等會計/內部代碼）

### 其他個人指標
- 個人台數/ACPP/SAcare 件數：與門市層級同邏輯，ACPP 依品名關鍵字分機種（mac/ipad/iphone/watch/airpods）
- 3PP 配件分類營收：C3=3003 依 C4 分（4007=CPU週邊、4009=iPhone週邊、4012=CPU/iOS通用、4022=iOS通用、4039=Watch、4069=AirPods…完整對照見引擎 `C4_ROWS`）
- 獎金計算（Sheet 6 公式）：F=SA營收/(SA+D欄)、G=3PP/原廠比、J=SA毛利=E÷2÷1.05、K=H+I+J

## 5. 期間與會計週期慣例

- **週**：日曜起算、週六結束（週日~週六）；預設「上個完整週六」為週結束日。
- **會計年度**：52 週（4 季 × 13 週），起始日讀模板「設定」sheet D2；季起始 = 年起 + 季序×13 週。Sheet 1 顯示當季 13 週逐週台數。
- **本月**：若週跨月則取整個報表月；否則 1 日~週結束日。
- **上月/去年同期**：同日期範圍對齊，月底日數不足取 `min(day, 當月天數)`（自動處理 2/29）。
- **YOY（Sheet 10/11）**：1/1 累積至截止日（可自訂 `yoyEnd`），對比前一年同日。

---

## 6. 其他專案如何調用（最省力做法）

### 方式 A：直接 import（Python 專案，推薦）
查詢層與計算層都是可直接 import 的模組，零改動即可重用：

```python
import sys
sys.path.insert(0, "/Users/sa/Claude/週報製作/報表製作_士林/live-store-report-app")
import server   # import 不會啟動 web server

from datetime import date
start, end = date(2026, 6, 1), date(2026, 6, 7)

# 1) 抓標準化銷售明細（任意門市、多區間合併單次查詢）
df = server.remote_sales_df("004", [(start, end)])

# 2) 門市 KPI（營業額/毛利/台數/附加率原料）
prices  = server.load_sacare_cached()                 # SAcare 價目表
d       = server.engine.period(df, start, end)        # 期間過濾
metrics = server.engine.calc_metrics(d, prices)       # → dict（total_rev, apl_gross, cpu_total, iphone, ...）

# 3) 個人 KPI
emp = server.engine.calc_employee(d, "員工代碼", prices)

# 4) 任意自訂 SQL（Oracle 語法）
headers, rows = server.run_remote("select shop_id, name from pos_shop where org_id = '01'")
```

注意事項：
- 跑在公司內網/VPN 下、本機需有 JDK 1.8 與 `/Library/EPBrowser` classpath（同一台 Mac 都已具備）。
- `server.engine` 就是 fill_weekly_excel 引擎；所有分類常數（C6_CPU、C4_ROWS、排除碼）都在引擎模組層級，可直接引用，新機種只要改引擎一處。
- SAcare 價目表路徑固定在本專案 `銷售資料/` 下，import 方式會自動沿用。

### 方式 B：HTTP API（非 Python 或不想 import）
啟動 `python3 server.py`（埠 8783）後：
- `POST /api/generate` `{"shopId":"004","weekEnd":"YYYY-MM-DD"}` → 輪詢 `/api/status` → `/api/download` 拿 11-sheet Excel
- 適合只要成品報表、或從 Shortcuts/排程腳本觸發的場景。

### 不建議
- 在新專案複製貼上 SQL 與公式 — 排除碼與機種常數每季可能更新，會出現兩份不同步的邏輯。一律以本專案為唯一邏輯來源（single source of truth）。
