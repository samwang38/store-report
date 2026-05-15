#!/bin/bash
# 門市報表製作 — 一鍵啟動（自動安裝 / 自動更新 / 啟動）
# 同事只要雙擊這一個檔即可，不需開終端機、不需貼指令。

REPO="https://github.com/samwang38/store-report.git"
DEST="$HOME/Desktop/門市報表製作-app"
PORT=8783
URL="http://127.0.0.1:$PORT/"

cd "$(dirname "$0")"

# ── 設定自訂圖示（背景執行，不阻塞啟動）─────────────────────
_set_icon() {
  local dir icon self
  dir="$(cd "$(dirname "$0")" && pwd)"
  self="$dir/$(basename "$0")"
  icon="$dir/app_icon.png"
  [ -f "$icon" ] || return 0
  ( osascript <<APPLESCRIPT 2>/dev/null
use framework "AppKit"
set img to (current application's NSImage's alloc()'s initWithContentsOfFile:"$icon")
(current application's NSWorkspace's sharedWorkspace()'s setIcon:img forFile:"$self" options:0)
APPLESCRIPT
  ) &
}
_set_icon

echo "=== 門市報表製作 ==="
echo ""

# ── 情況 A：這個檔是單獨下載的（所在資料夾沒有 server.py / .git）──
#    → 自動 clone 到桌面，然後改用桌面那份重新啟動
if [ ! -f "server.py" ] || [ ! -d ".git" ]; then
  if ! command -v git &>/dev/null; then
    echo "正在準備 git（需要 Xcode Command Line Tools）…"
    xcode-select --install 2>/dev/null || true
    echo "[提示] 安裝完成後，請再次雙擊本檔。"
    read -p "按 Enter 關閉"
    exit 1
  fi
  if [ -d "$DEST/.git" ]; then
    echo "偵測到既有安裝，更新中…"
    cd "$DEST" && git pull --quiet 2>/dev/null || true
  else
    echo "首次設定，下載工具中…"
    git clone --quiet "$REPO" "$DEST" || { echo "[錯誤] 下載失敗，請確認網路。"; read -p "按 Enter 關閉"; exit 1; }
    cd "$DEST"
  fi
  # 解除桌面那份的 Gatekeeper 隔離，之後從桌面雙擊就不再跳警告
  xattr -dr com.apple.quarantine "$DEST" 2>/dev/null || true
  chmod +x "$DEST/啟動門市報表製作.command"
  echo ""
  echo "已安裝到桌面資料夾「門市報表製作-app」。"
  echo "改用桌面那份啟動…"
  echo ""
  exec "$DEST/啟動門市報表製作.command"
fi

# ── 情況 B：已在 repo 內 → 自動更新 ──────────────────────────
if command -v git &>/dev/null && [ -d ".git" ]; then
  echo "檢查更新中…"
  if git pull --quiet 2>/dev/null; then
    echo "已是最新版本。"
  else
    echo "（無法連線更新，使用現有版本）"
  fi
  echo ""
fi

# ── 檢查 Python ───────────────────────────────────────────────
if ! command -v python3 &>/dev/null; then
  echo "[錯誤] 找不到 python3，請先安裝 Python 3。"
  read -p "按 Enter 關閉"
  exit 1
fi

# ── 缺套件才安裝 ──────────────────────────────────────────────
if ! python3 -c "import openpyxl, pandas" 2>/dev/null; then
  echo "安裝必要套件中…"
  pip3 install openpyxl pandas --quiet
fi

echo "啟動伺服器（port $PORT）…"
echo "請連好公司 VPN。瀏覽器將自動開啟：$URL"
echo "關閉此視窗即可停止。"
echo "-------------------------------------------"

# 伺服器起來後自動開瀏覽器
( for i in $(seq 1 30); do
    if curl -s -o /dev/null "$URL"; then
      if open -a "Google Chrome" "$URL" 2>/dev/null; then :
      else open "$URL"; fi
      break
    fi
    sleep 0.5
  done ) &

python3 server.py
