#!/bin/bash
# 在「常開的那台 Mac」上安裝週報小幫手 Mac 端工作程式（launchd 每 60 秒執行 report_worker.py）。
# 會依本機實際路徑產生 plist，所以換電腦、換使用者名稱都不用改檔案。
#
# 用法（在本資料夾內）：
#   ./install_report_worker.sh            檢查環境並安裝
#   ./install_report_worker.sh uninstall  停用並移除
# 注意：同一時間只讓一台 Mac 執行，避免驗證碼信重複寄出。

set -u
cd "$(dirname "$0")"
APP="$(pwd)"
LABEL="com.studioa.report-weekly-worker"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
LOG="/tmp/report-weekly-worker.log"
PY="$APP/venv/bin/python"

if [ "${1:-}" = "uninstall" ]; then
  launchctl unload "$PLIST" 2>/dev/null; rm -f "$PLIST"
  echo "已停用並移除 $LABEL"; exit 0
fi

ok=1
check() { if eval "$2" >/dev/null 2>&1; then echo "✓ $1"; else echo "✗ $1 —— $3"; ok=0; fi; }

echo "== 檢查環境 =="
check "Python 環境（venv）" "[ -x '$PY' ] && '$PY' -c 'import openpyxl, pandas'" "請先雙擊「啟動門市報表製作.command」完成首次安裝"
check "Playwright（來客數）" "'$PY' -c 'import playwright'" "請先雙擊「啟動門市報表製作.command」完成首次安裝"
check "Java（EPB 查詢）" "[ -x \"\${EPB_JAVA_HOME:-/Library/Java/JavaVirtualMachines/jdk1.8.0_251.jdk/Contents/Home}/bin/javac\" ]" "需要安裝 JDK 1.8（與原本那台相同版本）"
check "EPBrowser" "[ -d /Library/EPBrowser ]" "需要安裝公司 EPB 用戶端（/Library/EPBrowser）"
check "local_config.json 的 smtp 設定" "'$PY' -c \"import json;assert json.load(open('local_config.json'))['smtp']['host']\"" "請從原本那台複製 local_config.json，或執行 python3 smtp_test.py 並存檔"
check "local_config.json 的 reportWeekly 設定" "'$PY' -c \"import json;c=json.load(open('local_config.json'))['reportWeekly'];assert c['workerUrl'] and c['macSecret']\"" "請從原本那台複製 local_config.json"
check "公司內網／VPN（EPB 主機）" "'$PY' -c 'import report_worker as r;assert r.epb_reachable()'" "請確認 VPN 已連線"

if [ $ok -ne 1 ]; then echo; echo "有項目未通過，修正後再執行一次。"; exit 1; fi

mkdir -p "$HOME/Library/LaunchAgents"
cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key><array><string>$PY</string><string>report_worker.py</string></array>
  <key>WorkingDirectory</key><string>$APP</string>
  <key>StartInterval</key><integer>60</integer>
  <key>RunAtLoad</key><true/>
  <key>StandardOutPath</key><string>$LOG</string>
  <key>StandardErrorPath</key><string>$LOG</string>
</dict>
</plist>
EOF
launchctl unload "$PLIST" 2>/dev/null
launchctl load "$PLIST"
sleep 3
echo
if launchctl list | grep -q "$LABEL"; then
  echo "✓ 已安裝並啟動，每 60 秒執行一次。看紀錄：tail -f $LOG"
else
  echo "✗ 啟動失敗，請把 $LOG 的內容給 Claude 看"; exit 1
fi
