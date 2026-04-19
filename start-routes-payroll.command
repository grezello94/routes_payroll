#!/bin/zsh
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

PORT=5501
URL="http://127.0.0.1:${PORT}"
LOG_FILE="/tmp/routes-payroll-server.log"
ENV_FILE="$SCRIPT_DIR/.env.local"
STARTUP_WAIT_SECONDS="${APP_STARTUP_WAIT_SECONDS:-20}"

is_running_server_healthy() {
  local health
  health="$(curl -fsS "$URL/api/health" 2>/dev/null || true)"
  [ -n "$health" ] && [[ "$health" != *'"degraded":true'* ]]
}

stop_existing_server() {
  local pids
  pids="$(lsof -tiTCP:${PORT} -sTCP:LISTEN -n -P || true)"
  if [ -n "$pids" ]; then
    echo "$pids" | xargs kill
    sleep 1
  fi
}

start_server() {
  nohup node server.js >"$LOG_FILE" 2>&1 </dev/null &
  disown
}

wait_for_health() {
  local attempt=0
  while [ "$attempt" -lt "$STARTUP_WAIT_SECONDS" ]; do
    if is_running_server_healthy; then
      return 0
    fi
    attempt=$((attempt + 1))
    sleep 1
  done
  return 1
}

if lsof -iTCP:${PORT} -sTCP:LISTEN -n -P >/dev/null 2>&1; then
  if is_running_server_healthy; then
    open "$URL"
    exit 0
  fi
  stop_existing_server
fi

start_server

if wait_for_health; then
  open "$URL"
  exit 0
fi

open "$URL"
exit 0
