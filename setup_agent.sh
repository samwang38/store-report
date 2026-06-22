#!/bin/bash
# EPB 同步代理一鍵安裝腳本
# 使用方式：bash setup_agent.sh
# 需要：公司 VPN/內網、EPB_INGEST_SECRET（向管理員取得）

set -e

REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="$REPO_DIR/venv"
PYTHON_BIN=""
PLIST_LABEL="com.studioa.epb-pickup-sync"
PLIST_SRC="$REPO_DIR/com.studioa.epb-pickup-sync.plist"
PLIST_DEST="$HOME/Library/LaunchAgents/$PLIST_LABEL.plist"
LOG_FILE="/tmp/epb-pickup-sync.log"

echo ""
echo "=== EPB 同步代理安裝 ==="
echo ""

# ── 1. 確認 Python 3.10+ ──────────────────────────────────────────────────────
find_python() {
  for candidate in \
    /opt/homebrew/bin/python3.14 \
    /opt/homebrew/bin/python3.13 \
    /opt/homebrew/bin/python3.12 \
    /opt/homebrew/bin/python3.11 \
    /opt/homebrew/bin/python3.10 \
    /Library/Frameworks/Python.framework/Versions/3.14/bin/python3 \
    /Library/Frameworks/Python.framework/Versions/3.13/bin/python3 \
    /Library/Frameworks/Python.framework/Versions/3.12/bin/python3; do
    if [ -x "$candidate" ]; then
      echo "$candidate"
      return
    fi
  done
}

PYTHON_BIN=$(find_python)

if [ -z "$PYTHON_BIN" ]; then
  echo "[1/5] Python 3.10+ 未找到，正在安裝 Homebrew + Python 3.14..."
  if ! command -v brew &>/dev/null; then
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    # 讓 brew 指令在此 shell 可用
    eval "$(/opt/homebrew/bin/brew shellenv)" 2>/dev/null || true
  fi
  brew install python@3.14
  PYTHON_BIN=$(find_python)
fi

echo "[1/5] Python：$PYTHON_BIN  ($(${PYTHON_BIN} --version 2>&1))"

# ── 2. 建立虛擬環境並安裝套件 ────────────────────────────────────────────────
echo "[2/5] 建立虛擬環境並安裝套件..."
if [ ! -f "$VENV_DIR/bin/python" ]; then
  "$PYTHON_BIN" -m venv "$VENV_DIR"
fi
"$VENV_DIR/bin/pip" install --quiet --upgrade pip
"$VENV_DIR/bin/pip" install --quiet openpyxl pandas
echo "      套件安裝完成"

# ── 3. 建立 local_config.json ─────────────────────────────────────────────────
CONFIG_FILE="$REPO_DIR/local_config.json"

if [ ! -f "$CONFIG_FILE" ]; then
  echo ""
  echo "[3/5] 尚未建立 local_config.json，需輸入以下設定："
  echo "      （ingestSecret 請向管理員索取，或從已安裝的機器複製）"
  echo ""
  read -p "      EPB_INGEST_SECRET：" INGEST_SECRET
  echo ""

  cat > "$CONFIG_FILE" << EOF
{
  "epb": {
    "workerUrl": "https://studioa-reservation.samwang775.workers.dev",
    "ingestSecret": "$INGEST_SECRET",
    "shops": ["004", "046", "068"],
    "windowDays": 90
  }
}
EOF
  echo "      已建立 local_config.json"
else
  echo "[3/5] local_config.json 已存在，略過"
fi

# ── 4. 測試連線 ───────────────────────────────────────────────────────────────
echo ""
echo "[4/5] 測試同步（需 VPN/內網）..."
cd "$REPO_DIR"
if "$VENV_DIR/bin/python" epb_pickup_sync.py --force; then
  echo "      同步成功"
else
  echo ""
  echo "  !! 同步失敗，請確認："
  echo "     - VPN 或內網是否已連線"
  echo "     - ingestSecret 是否正確"
  echo ""
  echo "  安裝中止。確認後重新執行此腳本。"
  exit 1
fi

# ── 5. 安裝 launchd ───────────────────────────────────────────────────────────
echo ""
echo "[5/5] 安裝 launchd 定時排程..."

# 停用舊版（若存在），忽略錯誤
launchctl bootout "gui/$(id -u)/$PLIST_LABEL" 2>/dev/null || true

# 寫入更新路徑的 plist
sed \
  -e "s|/Library/Frameworks/Python.framework/Versions/3.14/bin/python3|$VENV_DIR/bin/python|g" \
  -e "s|/opt/homebrew/bin/python3\..*|$VENV_DIR/bin/python|g" \
  -e "s|<string>/Users/.*/live-store-report-app</string>|<string>$REPO_DIR</string>|g" \
  -e "s|<string>/Users/.*/store-report</string>|<string>$REPO_DIR</string>|g" \
  "$PLIST_SRC" > "$PLIST_DEST"

launchctl bootstrap "gui/$(id -u)" "$PLIST_DEST"

# 確認
if launchctl list | grep -q "$PLIST_LABEL"; then
  echo "      排程已啟動（每 5 分鐘自動同步）"
else
  echo "  !! launchd 啟動失敗，請檢查 $LOG_FILE"
  exit 1
fi

echo ""
echo "=== 安裝完成 ==="
echo ""
echo "  查看 log：tail -f $LOG_FILE"
echo "  停用：    launchctl bootout gui/\$(id -u) $PLIST_DEST"
echo ""
