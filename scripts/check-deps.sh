#!/usr/bin/env bash
# 环境检查 + 确保 CDP Proxy 就绪

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
LOG_DIR="$HOME/.claude/logs"
mkdir -p "$LOG_DIR"

# Node.js
if command -v node &>/dev/null; then
  NODE_VER=$(node --version 2>/dev/null)
  NODE_MAJOR=$(echo "$NODE_VER" | sed 's/v//' | cut -d. -f1)
  if [ "$NODE_MAJOR" -ge 22 ] 2>/dev/null; then
    echo "node: ok ($NODE_VER)"
  else
    echo "node: warn ($NODE_VER, 建议升级到 22+)"
  fi
else
  echo "node: missing — 请安装 Node.js 22+"
  exit 1
fi

# CDP Proxy — 确保 proxy 运行且已连接 Chrome
HEALTH=$(curl -s --connect-timeout 2 "http://127.0.0.1:3456/health" 2>/dev/null)

if echo "$HEALTH" | grep -q '"connected":true'; then
  # Proxy 已运行且已连接
  echo "chrome: ok (via proxy)"
  echo "proxy: ready"
  exit 0
fi

# Proxy 未运行或未连接 → 杀掉旧实例，重新启动
if echo "$HEALTH" | grep -q '"ok"'; then
  echo "proxy: running but not connected, restarting..."
  # 找到旧 proxy 的 PID 并杀掉
  OLD_PID=$(netstat -ano 2>/dev/null | grep ':3456.*LISTENING' | awk '{print $5}' | head -1)
  if [ -n "$OLD_PID" ]; then
    taskkill //F //PID "$OLD_PID" > /dev/null 2>&1 || kill "$OLD_PID" 2>/dev/null
    sleep 2
  fi
else
  echo "proxy: starting..."
fi

node "$SCRIPT_DIR/cdp-proxy.mjs" > "$LOG_DIR/cdp-proxy.log" 2>&1 &

for ((i=1; i<=30; i++)); do
  sleep 1
  curl -s http://localhost:3456/health | grep -q '"connected":true' && echo "chrome: ok (via proxy)" && echo "proxy: ready" && exit 0
  [ $i -eq 5 ] && echo "[INFO] 等待 CDP Proxy 启动 Chrome 并连接..."
  [ $i -eq 15 ] && echo "[INFO] 仍在等待连接（Chrome 自动启动中）..."
done

echo "[ERROR] 连接超时（30秒）。请检查 Chrome 是否已安装。"
exit 1
