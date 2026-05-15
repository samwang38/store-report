#!/bin/bash
# 門市報表製作 — 一鍵安裝腳本
set -e

REPO="https://github.com/samwang38/store-report.git"
DEST="$HOME/Desktop/門市報表製作-app"

echo "==================================="
echo "  門市報表製作 安裝程式"
echo "==================================="
echo ""

# ── 檢查 git ─────────────────────────────────────────────────
if ! command -v git &>/dev/null; then
  echo "正在安裝 git（需要 Xcode Command Line Tools）…"
  xcode-select --install 2>/dev/null || true
  echo ""
  echo "[提示] 安裝完成後，請重新執行此腳本。"
  exit 1
fi

# ── 檢查 python3 ──────────────────────────────────────────────
if ! command -v python3 &>/dev/null; then
  echo "[錯誤] 找不到 Python 3。"
  echo "請至 https://www.python.org/downloads/ 下載安裝後重試。"
  exit 1
fi

# ── Clone 或更新 ──────────────────────────────────────────────
if [ -d "$DEST/.git" ]; then
  echo "偵測到現有安裝，更新至最新版本…"
  cd "$DEST"
  git pull
else
  echo "下載工具中…"
  git clone "$REPO" "$DEST"
fi

# ── 安裝 Python 套件 ──────────────────────────────────────────
echo "安裝必要套件（openpyxl, pandas）…"
pip3 install openpyxl pandas --quiet

# ── 解除 Gatekeeper 隔離 + 確保可執行 ─────────────────────────
xattr -dr com.apple.quarantine "$DEST" 2>/dev/null || true
chmod +x "$DEST/啟動門市報表製作.command"

echo ""
echo "==================================="
echo "  ✅ 安裝完成！"
echo "==================================="
echo ""
echo "桌面已建立「門市報表製作-app」資料夾。"
echo "雙擊其中的「啟動門市報表製作.command」即可使用。"
echo ""
echo "往後每次啟動會自動同步最新版本，無需重新安裝。"
echo "（使用前請先連好公司 VPN）"
