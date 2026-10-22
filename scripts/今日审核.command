#!/bin/zsh
set -u
TAILKI_REVIEW_ROOT="$(cd "$(dirname "$0")/.." && pwd)"

tailki_review_fail() {
  local result="$1"
  print -u2 -- "$2"
  if [[ -t 0 ]]; then
    read -r "TAILKI_REVIEW_ACK?按回车关闭" || true
  fi
  exit "$result"
}

cd "$TAILKI_REVIEW_ROOT" || tailki_review_fail 1 "无法打开项目目录，请确认今日审核.command 仍位于项目的 scripts 文件夹。"
[[ -f scripts/review_center.py ]] || tailki_review_fail 1 "今日审核程序缺失，请回到聊天修复项目文件。"

# Finder may provide a smaller PATH than a terminal. Prefer the bundled runtime,
# then accept a Python on PATH only when it has the document dependency.
TAILKI_REVIEW_PYTHON="${TAILKI_REVIEW_PYTHON:-$HOME/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3}"
if ! "$TAILKI_REVIEW_PYTHON" -B -c 'import sys; assert sys.version_info >= (3, 9); import docx' >/dev/null 2>&1; then
  TAILKI_REVIEW_PYTHON="$(command -v python3 || true)"
  if [[ -z "$TAILKI_REVIEW_PYTHON" ]] || ! "$TAILKI_REVIEW_PYTHON" -B -c 'import sys; assert sys.version_info >= (3, 9); import docx' >/dev/null 2>&1; then
    tailki_review_fail 1 "今日审核缺少可用的 Python 3.9+ 或 Word 文档依赖（python-docx）。请打开 Codex 后回到聊天修复运行环境。原资料已保留。"
  fi
fi

if "$TAILKI_REVIEW_PYTHON" -B scripts/review_center.py build --open; then
  exit 0
else
  tailki_review_fail "$?" "今日审核未能更新。请查看上方具体报错，或回到聊天处理。原资料已保留。"
fi
