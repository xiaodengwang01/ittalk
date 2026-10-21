#!/bin/zsh
set -eu
TAILKI_REVIEW_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
TAILKI_REVIEW_PYTHON="$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
cd "$TAILKI_REVIEW_ROOT"
if ! "$TAILKI_REVIEW_PYTHON" -B scripts/review_center.py build --open; then
  echo "今日审核未能更新。原资料已保留，请回到聊天查看处理进度。"
  read -r "TAILKI_REVIEW_ACK?按回车关闭"
fi
