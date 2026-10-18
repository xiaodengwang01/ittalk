#!/bin/zsh
set -u
project_dir="${0:A:h}"
cd "$project_dir" || exit 1
/usr/bin/env python3 "$project_dir/scripts/setup_git_pat.py"
result=$?
if [[ -t 0 ]]; then
  read -r "?按回车键关闭窗口。"
fi
exit "$result"
