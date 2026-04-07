#!/bin/zsh
set -e

PORT=5501
PIDS="$(lsof -tiTCP:${PORT} -sTCP:LISTEN -n -P || true)"

if [ -n "$PIDS" ]; then
  echo "$PIDS" | xargs kill
fi
