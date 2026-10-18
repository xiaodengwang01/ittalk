#!/usr/bin/env python3
"""Read a PAT privately, store it in macOS Keychain, then use git_sync.py."""
import fcntl
import json
import os
from pathlib import Path
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
OWNER = "xiaodengwang01"
REPO_PATH = OWNER + "/ittalk.git"
REMOTE = "https://github.com/" + REPO_PATH
STATE = ROOT / ".git" / "tailki-pat-state.json"
ENV = {
    key: value for key, value in os.environ.items()
    if not key.startswith("GIT_TRACE") and key != "GIT_CURL_VERBOSE"
}
ENV.update(GIT_TERMINAL_PROMPT="0", GIT_ASKPASS="/usr/bin/false")
PROMPT = '''
activate
set answer to display dialog "请粘贴用于 xiaodengwang01/ittalk 的 GitHub Personal Access Token。\n\n输入内容会隐藏，并保存在本机 macOS 钥匙串。确认后将自动检查并上传 Tailki。" default answer "" with hidden answer buttons {"取消", "授权并上传"} default button "授权并上传" cancel button "取消" with title "Tailki · GitHub 授权"
return text returned of answer
'''


def run(command, *, data=None, timeout=30):
    return subprocess.run(
        command, cwd=ROOT, env=ENV, input=data, text=True,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout,
    )


def git(*args, data=None):
    result = run(["git", *args], data=data)
    if result.returncode:
        # Credential helper output is deliberately never included in errors.
        raise RuntimeError("Git 配置或钥匙串操作未完成，请检查本机钥匙串是否已解锁。")
    return result.stdout.strip()


def state(stage, message):
    STATE.write_text(json.dumps({"stage": stage, "message": message}, ensure_ascii=False), encoding="utf-8")
    print(message, flush=True)


def notify(message):
    # The message is fixed program text, never a credential or subprocess output.
    script = 'on run argv\nactivate\ndisplay dialog (item 1 of argv) with title "Tailki Git 上传" buttons {"好"} default button "好"\nend run'
    run(["/usr/bin/osascript", "-e", script, message], timeout=None)


def preflight():
    if sys.platform != "darwin":
        raise RuntimeError("这个授权入口需要 macOS 钥匙串。")
    if git("remote", "get-url", "origin") != REMOTE:
        raise RuntimeError("目标仓库与约定不符，已停止。")
    if git("remote", "get-url", "--push", "origin") != REMOTE:
        raise RuntimeError("上传地址与约定不符，已停止。")
    helper = Path(git("--exec-path")) / "git-credential-osxkeychain"
    if not os.access(helper, os.X_OK):
        raise RuntimeError("未找到 macOS Git 钥匙串助手，请先安装 Apple 命令行工具。")
    result = run([sys.executable, str(ROOT / "scripts" / "git_sync.py"), "--check-only"])
    if result.returncode:
        raise RuntimeError("上传范围或模拟数据检查未通过，请先运行 scripts/git_sync.py --check-only 查看原因。")
    return helper


def configure_keychain():
    # An empty entry resets inherited helpers; these settings affect this repo only.
    git("config", "--local", "--replace-all", "credential.helper", "")
    git("config", "--local", "--add", "credential.helper", "osxkeychain")
    git("config", "--local", "credential.https://github.com.useHttpPath", "true")
    git("config", "--local", "credential.https://github.com.username", OWNER)


def read_pat():
    result = run(["/usr/bin/osascript", "-e", PROMPT], timeout=None)
    if result.returncode:
        if "-128" in result.stderr:
            return None
        raise RuntimeError("无法打开安全输入框，请在本机双击“首次授权并上传.command”重试。")
    pat = result.stdout.strip()
    if len(pat) < 30 or not re.fullmatch(r"(?:github_pat_|ghp_)[A-Za-z0-9_]+", pat):
        raise RuntimeError("输入不是完整的 GitHub PAT。请复制生成的 Token 本身，再重新打开授权入口。")
    return pat


def authorize_and_upload():
    helper = preflight()
    state("waiting", "等待你在 macOS 安全输入框中粘贴 Token……")
    pat = read_pat()
    if pat is None:
        state("cancelled", "已取消授权，本地提交保留。")
        return 0
    # Pass the secret through stdin, never through argv, a URL, a file or a log.
    stored = run([str(helper), "store"], data=(
        "protocol=https\nhost=github.com\npath=" + REPO_PATH
        + "\nusername=" + OWNER + "\npassword=" + pat + "\n\n"
    ))
    del pat
    if stored.returncode:
        raise RuntimeError("Token 未能存入 macOS 钥匙串。请解锁本机钥匙串后重试。")
    configure_keychain()
    state("uploading", "授权已存入本机钥匙串，正在检查并上传……")
    result = run([
        sys.executable, str(ROOT / "scripts" / "git_sync.py"),
        "Add secure PAT authorization and upload entry point",
    ], timeout=120)
    # Only fixed status text is logged here; git_sync.py retains all local commits.
    if result.returncode:
        detail = (result.stderr + result.stdout).lower()
        if any(word in detail for word in ("non-fast-forward", "fetch first", "[rejected]")):
            reason = "远端存在本机没有的提交，已停止覆盖，请让 Codex 检查分支差异。"
        elif any(word in detail for word in ("authentication", "invalid username", "could not read username", "could not read password")):
            reason = "GitHub 未接受授权，请检查 Token 是否完整、有效且未过期。"
        elif any(word in detail for word in ("403", "permission", "write access", "repository not found")):
            reason = "仓库写权限未通过，请检查 Token 是否选中 ittalk，并授予 Contents 读写权限。"
        elif any(word in detail for word in ("timed out", "timeoutexpired", "failed to connect", "could not resolve", "ssl", "connection")):
            reason = "GitHub 网络连接未完成，请检查网络或本机代理后重试。"
        else:
            reason = "请运行“一键上传.command”查看详细错误。"
        message = "授权已保存，但上传未完成。本地提交已保留。\n\n" + reason
        state("failed", message)
        notify(message)
        return 1
    local = git("rev-parse", "HEAD")
    remote = git("ls-remote", "origin", "refs/heads/main").split()
    if not remote or remote[0] != local:
        raise RuntimeError("上传后远端提交核对未通过，本地提交保留，请重新检查。")
    state("success", "上传成功：origin/main 与本地提交一致（" + local[:7] + "）。")
    notify("Tailki 已成功上传到 GitHub！\n\nxiaodengwang01/ittalk · main\n本机和远端提交已核对一致。\n以后可直接双击“一键上传.command”。")
    return 0


def main():
    # Prevent two password prompts from racing to update the same credential.
    with open(ROOT / ".git" / "tailki-pat.lock", "w") as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            print("已有授权窗口打开，请在那个窗口中完成操作。", flush=True)
            return 1
        try:
            return authorize_and_upload()
        except (RuntimeError, OSError, subprocess.SubprocessError) as error:
            message = str(error) if isinstance(error, RuntimeError) else "本机授权或网络操作未完成，请重新运行授权入口。本地提交保留。"
            state("failed", message)
            notify(message)
            return 1


if __name__ == "__main__":
    sys.exit(main())
