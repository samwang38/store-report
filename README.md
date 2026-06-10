# 門市報表製作（live-store-report-app）

可下拉選任一門市，產生該門市的士林式 13 張週報（含個人項目動態填入）。
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

報表核心（EPB 查詢、13 張填表、個人 8–11 sheet 依當週實際有交易員工動態增刪列）
移植自同層 `live-report-app`，邏輯零改動；本工具去掉登入閘、改非同步任務佇列、
加免登入門市下拉。Web 殼與前端流程仿 `北一區/北一區週報-app`。

共用上層資源（不複製）：
- 範本：`../週報模板.xlsx`
- 引擎：`../fill_weekly_excel.py`
- SAcare 價目表：`../銷售資料/SAcare對應價目表.xlsx`

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

DSS（搭售統計，Sheet 6/7）：登入需圖形驗證碼、不定時 Email 二次認證，
故採前端互動式登入；登入後 session 存於 local_config.json，重啟可續用。
