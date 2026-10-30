# 門市報表製作（live-store-report-app）

可下拉選任一門市，產生該門市的士林式 17 張週報與 11 張分析（分成兩個 Excel，含個人項目動態填入）。
EPB 即時查詢免登入，採「選日期 → 背景產生 → 下載」的非同步流程。

## 啟動

雙擊 `啟動門市報表製作.command`，或：

```bash
cd live-store-report-app
python3 server.py            # 預設 http://127.0.0.1:8783
PORT=8899 python3 server.py  # 自訂 port
```

需先連上公司 VPN（EPB 查詢免帳密，但需在內網）。

## 操作

1. 選門市（預設 004 士林門市）。
2. 選週結束日期（週六，預設最近一個週六）。
3. 按「產生報表」，進度跑完後下載 Excel。

## 設計

報表核心（EPB 查詢、20 張填表、個人 12–15 sheet 依當週實際有交易員工動態增刪列）
移植自同層 `live-report-app`，邏輯零改動；本工具去掉登入閘、改非同步任務佇列、
加免登入門市下拉。Web 殼與前端流程仿 `北一區/北一區週報-app`。

本機資源（本資料夾內）：
- 範本：`週報模板.xlsx`
- 引擎：`fill_weekly_excel.py`
- SAcare 價目表：`銷售資料/SAcare對應價目表.xlsx`

## API

| 路由 | 說明 |
|------|------|
| `GET /api/stores` | 門市清單（EPB `pos_shop`，失敗時用後備清單） |
| `GET /api/default-date` | 最近一個週六 |
| `POST /api/generate` | `{shopId, weekEnd}` → `{jobId}` |
| `GET /api/status?jobId=` | 任務狀態與進度訊息 |
| `GET /api/download?jobId=` | 下載產生的 xlsx |
| `GET /api/dss/status` | DSS 登入狀態（idle / need_captcha / need_otp / logged_in） |
| `GET /api/dss/captcha` | 目前登入流程的圖形驗證碼（PNG） |
| `POST /api/dss/credentials` | 儲存／清除 DSS 帳密（本機 local_config.json） |
| `POST /api/dss/login/start` | 開始登入：建立 session、取得驗證碼 |
| `POST /api/dss/login/refresh-captcha` | 換一張驗證碼 |
| `POST /api/dss/login/captcha` | `{code}` 送出帳密＋驗證碼 |
| `POST /api/dss/login/otp` | `{code}` 送出 Email 驗證碼（二次認證時） |

DSS（搭售統計，Sheet 10/11）：登入需圖形驗證碼、不定時 Email 二次認證，
故採前端互動式登入；登入後 session 存於 local_config.json，重啟可續用。

## MySetup 每日信件自動匯入（第 13/14 頁）

Sam 的公司信箱每天收到前一天「全公司」的 Personal Setup 匯出檔（附檔 `Personal Setup_起_迄.xlsx`），只有這個信箱收得到，所有門市共用。
常開的 MBP-Hub 由 `report_worker.py` 每小時用 IMAP 抓一次（`mysetup_mail.py`），依每列的 Apple ID 分店存進本機 `mysetup_archive.sqlite`
（只存 Unique ID、Apple ID、POS Name、Staff Name、Device、提交時間，不存客人資料；保留約 400 天）。

- 產週報：本店存檔資料與店長上傳的檔以 Unique ID 合併去重；期間內缺信的日子寫在 log，並在週報信裡提醒店長。
- 日報信：自動日報程式（daily-report-mac）用 `mysetup_daily.py` 算本週／本月每人每裝置的提交率，附在日報信內文。
- 排班系統「我的業績」：有設定 `scheduleSystem` 時，匯入新資料後把每人每日每裝置數推到 `/api/sales/mysetup-sync`（第一次設定補推近 45 天；店長在 LINE 改名稱對照時重推該店近 45 天）。名字對照都沿用店長在 LINE 設的規則。
- 閉店網站的「MySetup 填寫／週進度」2026-10-30 停用，不再回填。
- 到中午還沒收到前一天的信、或出現 `APPLE_ID_TO_SHOP` 沒有的 Apple ID，會寄信通知 Sam。新門市要把 Apple ID 補進對照表。

收信預設沿用 `local_config.json` 的 `smtp` 帳密（寄件帳號就是 Sam 的公司信箱），不用另外設定。要改用別的信箱或指定信件匣、主旨時，再在 MBP-Hub 的 `local_config.json` 加（帳密不進版控）：

```json
"mysetupMail": {"username": "...@studioa.com.tw", "password": "...", "host": "studioa.com.tw", "port": 143, "folder": "INBOX", "subject": ""}
```

`subject` 留空時檢查每封新信的附檔名；有填就只看主旨含這段文字的信。

推排班系統要在 MBP-Hub 的 `local_config.json` 加（token 在排班系統後台「業績同步代理」新增一台取得，只顯示一次）：

```json
"scheduleSystem": {"url": "https://排班系統網址", "agentToken": "<agent_id>.<secret>"}
```

手動指令：

```bash
venv/bin/python mysetup_mail.py --once        # 抓信一次（第一次補抓近 90 天）
venv/bin/python mysetup_mail.py --status      # 存檔涵蓋日期、缺的日子、對不上的 Apple ID
venv/bin/python mysetup_mail.py --file 檔案.xlsx   # 離線匯入
```
