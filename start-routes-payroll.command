#!/bin/zsh
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

PORT=5501
URL="http://127.0.0.1:${PORT}"
LOG_FILE="/tmp/routes-payroll-server.log"
ENV_FILE="$SCRIPT_DIR/.env.local"

if [ -f "$ENV_FILE" ]; then
  set -a
  source "$ENV_FILE"
  set +a
fi

if lsof -iTCP:${PORT} -sTCP:LISTEN -n -P >/dev/null 2>&1; then
  open "$URL"
  exit 0
fi

nohup npm start >"$LOG_FILE" 2>&1 &

for _ in {1..20}; do
  if curl -s "$URL/api/health" >/dev/null 2>&1; then
    open "$URL"
    exit 0
  fi
  sleep 1
done

open "$URL"
exit 0
