#!/usr/bin/env python3
"""Publish one selected area of ittalk; keep local business data private."""
import argparse
from datetime import datetime
import fcntl
import fnmatch
import io
import json
import os
from pathlib import Path, PurePosixPath
import re
import subprocess
import sys
import tempfile
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parent.parent
REMOTE = "https://github.com/xiaodengwang01/ittalk.git"
UPSTREAM = "https://github.com/ZJU-REAL/Easel.git"
UPSTREAM_REMOTE = "easel-upstream"
BASELINE_REF = "refs/easel/upstream-baseline"
SCANNER = ROOT / "scripts" / "gitleaks"

# 2026-11-06: the user authorized publishing only the consolidated Tailki
# workflow. Exact paths keep archive documents, live workbooks and run data out.
HANDBOOK = "KOL搜集与邮件回复/Tailki_KOL全流程图文合集.docx"
SKILL_ROOT = "KOL搜集与邮件回复/参考资料/tailki-kol-workflow/"
TAILKI_SCRIPTS = frozenset("""
candidate_pool.py create_mock_workbook.mjs daily_cells.mjs daily_workbook.py
daily_workflow.py followup_workbook.mjs git_sync.py local_resources.py
mail_drafts.py mail_inbox.py mail_keychain.py mail_postprocess.py mail_records.py
mail_send.py mail_text.py review_center.py review_docx.py review_page.py
slim_workbook.mjs slim_workbook.py workbook_metadata.py workflow.py
workflow_rules.py write_workbook.mjs 今日审核.command
test_candidate_pool.py test_command_launchers.py test_compact_followup.py
test_daily_workflow.py test_git_sync.py test_local_resources.py
test_mail_drafts.py test_mail_inbox.py test_mail_postprocess.py test_mail_send.py
test_mail_text.py test_review_center.py test_slim_workbook.py
test_workflow_adapter.py test_workflow_rules.py
""".split())
# These binary versions were inspected: the consolidated handbook, three real
# reference images, and the SIM/.invalid-only test workbook. A changed binary
# needs a fresh content review before updating its approved Git blob below.
AUDITED_TAILKI_BLOBS = {
    HANDBOOK: "ffec8bf8a599314dd05a7e85417b1afd46d9dbe4",
    SKILL_ROOT + "assets/pet-avatar-style.jpeg": "25811c4180c6ef0d371e206cc5b669b326916325",
    SKILL_ROOT + "assets/tailki-collar.png": "20b268309ce118805753a5d850ccb9e5530552dd",
    SKILL_ROOT + "assets/wearing-interaction.png": "a5f294d0f670ca8939ab647bdafddf13201017c7",
    "scripts/examples/Tailki_模拟.xlsx": "7002408f306151a669605845318fcb9dd31fa378",
}
TAILKI_FILES = frozenset({".gitignore", "AGENTS.md", "README.md",
                        SKILL_ROOT + "SKILL.md", SKILL_ROOT + "agents/openai.yaml",
                        *AUDITED_TAILKI_BLOBS,
                        *("scripts/" + name for name in TAILKI_SCRIPTS)})


def in_scope(name, scope):
    if scope == "tailki":
        return name in TAILKI_FILES
    return name.startswith(("docs" if scope == "docs" else "Easel") + "/")


def scope_paths(scope):
    if scope != "tailki":
        return [("docs" if scope == "docs" else "Easel") + "/"]
    tracked = set(names(git("ls-files", "-z").stdout))
    return sorted(name for name in TAILKI_FILES
                  if name in tracked or (ROOT / name).exists() or (ROOT / name).is_symlink())

# Audited upstream source paths matching broad privacy patterns. These are source
# or example files, not actual local credentials; scanning still applies.
SOURCE_EXCEPTIONS = frozenset({
    "Easel/.env.example",
    "Easel/skills/builtin/ecom-details-image/.env.example",
    "Easel/skills/builtin/skill-wechat-publisher/scripts/wechat_token.py",
    "Easel/skills/builtin/skill-xhs-analyzer/src/lib/cdp-cookies.ts",
    "Easel/skills/builtin/skill-xhs-analyzer/src/lib/cookies.ts",
})
PRIVATE_PATTERNS = (
    ".ds_store", ".env", ".env.*", "*token*", "*credential*", "*cookie*",
    "*.pem", "*.key", "*.sqlite*", "*.db", "*.log", "*.xlsx", "*.xls",
    "*.xlsm", "*.csv", "*.tsv", "邮件模板*.txt",
)
PRIVATE_DIRS = frozenset({".git", ".venv", ".openclaw", ".scratch", "node_modules", "__pycache__"})
MEDIA_SUFFIXES = frozenset({".png", ".jpg", ".jpeg", ".webp", ".gif", ".svg",
                            ".mp4", ".mov", ".webm", ".mp3", ".wav", ".pdf",
                            ".ttf", ".otf", ".woff", ".woff2", ".zip"})
# The three matches in the exact upstream baseline were inspected: a signing
# constant, a synthetic test value, and a browser localStorage key. Accept only
# the precise source blob, rule and line; later edits require a fresh review.
AUDITED_FINDINGS = {
    ("skills/builtin/skill-xhs-analyzer/src/lib/signing.ts", "generic-api-key", 27): "06a8d04eef37e8acf233d8ff0b89d124670f4725",
    ("skills/shared/scripts/content_guard.py", "generic-api-key", 326): "192c66520dea6719c2f9211ab8b03baa64b7bc8b",
    ("web/frontend/src/lib/whoami.ts", "generic-api-key", 5): "664d5b29ce3d3e03b7de462fb054be6e7242676c",
}


def git(*args, check=True, timeout=30, data=None):
    result = subprocess.run(
        ["git", *args], cwd=ROOT, input=data, text=True,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout,
        env={**os.environ, "GIT_TERMINAL_PROMPT": "0"},
    )
    if check and result.returncode:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "Git 操作失败。")
    return result


def names(output):
    return [name for name in output.split("\0") if name]


def approved_path(name):
    path = PurePosixPath(name)
    return (not path.is_absolute() and ".." not in path.parts and
            (name in TAILKI_FILES or (len(path.parts) > 1 and path.parts[0] in {"docs", "Easel"})))


def private_path(name):
    if not approved_path(name):
        return True
    parts = PurePosixPath(name).parts
    if any(part.casefold() in PRIVATE_DIRS for part in parts):
        return True
    if any(part.casefold().endswith((".pyc", ".pyo")) for part in parts):
        return True
    if name in AUDITED_TAILKI_BLOBS:
        return False
    if parts[0] == "Easel":
        if len(parts) > 1 and parts[1] == "outputs" and parts != ("Easel", "outputs", ".gitkeep"):
            return True
        if len(parts) > 2 and parts[1] == "profiles" and parts[2] != "_template":
            return True
        if len(parts) > 2 and parts[1:3] == ("docker", "data"):
            return True
        if name in SOURCE_EXCEPTIONS:
            return False
    return any(fnmatch.fnmatchcase(part.casefold(), pattern) for part in parts for pattern in PRIVATE_PATTERNS)


def official_assets():
    """Only media already present in a verified upstream tree may be published."""
    assets = {}
    for ref in (BASELINE_REF, f"refs/remotes/{UPSTREAM_REMOTE}/main"):
        if git("rev-parse", "--verify", ref, check=False).returncode:
            continue
        for entry in names(git("ls-tree", "-r", "-z", ref).stdout):
            metadata, name = entry.split("\t", 1)
            mode, kind, blob = metadata.split()
            if kind == "blob" and mode in {"100644", "100755"}:
                assets.setdefault("Easel/" + name, set()).add(blob)
    return assets


def validate_entry(name, mode, blob, assets):
    if not approved_path(name) or mode not in {"100644", "100755"}:
        raise RuntimeError("待上传的提交包含允许范围之外的文件或链接，已停止：" + name)
    if private_path(name):
        raise RuntimeError("待上传的提交包含被隐私规则排除的文件：" + name)
    if name in AUDITED_TAILKI_BLOBS and blob != AUDITED_TAILKI_BLOBS[name]:
        raise RuntimeError("Tailki 合集、配图或模拟表版本尚未完成内容核对，已停止：" + name)
    source_only = (name.startswith("Easel/assets/") or name.startswith("Easel/profiles/_template/") or
                   name in {"Easel/.env.example", "Easel/skills/builtin/ecom-details-image/.env.example"} or
                   (name.startswith("Easel/") and PurePosixPath(name).suffix.casefold() in MEDIA_SUFFIXES))
    if source_only and blob not in assets.get(name, set()):
        raise RuntimeError("Easel 素材或示例配置不是已核实的官方文件，已停止上传：" + name)


def check_workspace(scope="docs"):
    if scope not in {"docs", "easel", "tailki"}:
        raise RuntimeError("未知上传范围。")
    for options in (("--all",), ("--all", "--push")):
        if git("remote", "get-url", *options, "origin").stdout.splitlines() != [REMOTE]:
            raise RuntimeError("origin 的读取或上传地址不是约定的 GitHub 仓库，已停止。")
    if git("branch", "--show-current").stdout.strip() != "main":
        raise RuntimeError("当前分支不是 main，已停止。")
    for state in ("MERGE_HEAD", "CHERRY_PICK_HEAD", "REVERT_HEAD", "rebase-merge", "rebase-apply"):
        state_path = Path(git("rev-parse", "--git-path", state).stdout.strip())
        if not state_path.is_absolute():
            state_path = ROOT / state_path
        if state_path.exists():
            raise RuntimeError("仓库仍有未完成的合并或历史操作，请先处理后再上传。")
    area = {"docs": "docs", "easel": "Easel", "tailki": "Tailki"}[scope]
    directory = ROOT if scope == "tailki" else ROOT / area
    if not directory.is_dir() or directory.is_symlink():
        raise RuntimeError(f"{area} 必须是本机的实际目录。")
    if scope == "easel" and (directory / ".git").exists():
        raise RuntimeError("Easel 仍是嵌套 Git 仓库；请先完成子树迁移。")
    staged = names(git("diff", "--cached", "--name-only", "-z").stdout)
    if staged:
        raise RuntimeError("存在已暂存的改动，已停止以保留原有暂存内容：" + "、".join(staged))
    candidates = (scope_paths(scope) if scope == "tailki" else
                  names(git("ls-files", "--cached", "--others", "--exclude-standard", "-z", "--", area + "/").stdout))
    existing = [name for name in candidates if (ROOT / name).exists() or (ROOT / name).is_symlink()]
    if existing:
        ignored = git("check-ignore", "--no-index", "-z", "--stdin", data="\0".join(existing) + "\0", check=False)
        if ignored.returncode not in (0, 1):
            raise RuntimeError(ignored.stderr.strip() or "无法检查隐私忽略规则。")
        if ignored.stdout:
            raise RuntimeError("已跟踪文件被隐私规则排除：" + "、".join(names(ignored.stdout)))
    for name in existing:
        path = ROOT / name
        if private_path(name) or path.is_symlink() or not path.is_file() or not path.resolve().is_relative_to(directory.resolve()):
            raise RuntimeError(f"{area} 中只能上传允许范围内的普通文件：" + name)


def check_tree(revision, *, subtree=False):
    assets = official_assets()
    entries = names(git("ls-tree", "-r", "-z", revision).stdout)
    for entry in entries:
        metadata, name = entry.split("\t", 1)
        mode, kind, blob = metadata.split()
        if kind != "blob":
            raise RuntimeError("待上传的提交包含不支持的对象：" + name)
        validate_entry(("Easel/" + name) if subtree else name, mode, blob, assets)


def subtree_source(revision):
    message = git("show", "-s", "--format=%B", revision).stdout
    dirs = re.findall(r"(?m)^git-subtree-dir: (.+)$", message)
    splits = re.findall(r"(?m)^git-subtree-split: ([0-9a-f]{40})$", message)
    if dirs != ["Easel"] or len(splits) != 1:
        return None
    parents = git("rev-list", "--parents", "-n", "1", revision).stdout.split()
    # Squash commits have at most one previous squash parent, never a merge.
    if len(parents) > 2:
        return None
    return splits[0]


def check_outgoing():
    base = "refs/remotes/origin/main"
    if git("merge-base", "--is-ancestor", base, "HEAD", check=False).returncode:
        raise RuntimeError("远端存在本机尚未包含的提交，已停止；不会强制覆盖或自动处理冲突。")
    revisions = git("rev-list", base + "..HEAD").stdout.splitlines()
    squash = []
    for revision in revisions:
        try:
            check_tree(revision)
        except RuntimeError:
            source = subtree_source(revision)
            if source is None:
                raise
            squash.append((revision, source))
    if squash:
        if git("remote", "get-url", UPSTREAM_REMOTE, check=False).stdout.strip() != UPSTREAM:
            raise RuntimeError("官方 Easel 上游地址不正确，无法核实子树历史。")
        git("fetch", "--no-tags", UPSTREAM_REMOTE,
            f"+refs/heads/main:refs/remotes/{UPSTREAM_REMOTE}/main", timeout=180)
        for revision, source in squash:
            if git("merge-base", "--is-ancestor", source, f"refs/remotes/{UPSTREAM_REMOTE}/main", check=False).returncode:
                raise RuntimeError("子树版本不属于已核实的官方主分支：" + source)
            if git("rev-parse", revision + "^{tree}").stdout.strip() != git("rev-parse", source + "^{tree}").stdout.strip():
                raise RuntimeError("子树历史与官方版本内容不一致：" + revision)
            check_tree(revision, subtree=True)
    return bool(revisions)


def check_outgoing_scope(scope):
    """Do not publish another area's existing unpushed commits by accident."""
    area = {"docs": "docs", "easel": "Easel", "tailki": "Tailki"}[scope]
    base = "refs/remotes/origin/main"
    for revision in git("rev-list", "--first-parent", base + "..HEAD").stdout.splitlines():
        parents = git("rev-list", "--parents", "-n", "1", revision).stdout.split()
        if len(parents) < 2:
            raise RuntimeError("待上传历史缺少父提交，已停止。")
        changed = names(git("diff", "--name-only", "-z", parents[1], revision, "--").stdout)
        other = [name for name in changed if not in_scope(name, scope)
                 and (approved_path(name) or scope == "tailki")]
        if other:
            raise RuntimeError(f"待上传的已有提交还包含 {area} 以外的改动，已停止：" + "、".join(other))


def scan_secrets(*, staged=False, base=None):
    if not SCANNER.is_file() or not os.access(SCANNER, os.X_OK):
        raise RuntimeError("缺少本机密钥检查工具 scripts/gitleaks，已停止上传。")
    with tempfile.TemporaryDirectory(prefix="tailki-git-scan-") as directory:
        report = Path(directory) / "report.json"
        args = [str(SCANNER), "git", "--no-banner", "--no-color", "--redact=100",
                "--ignore-gitleaks-allow", "--gitleaks-ignore-path", directory,
                "--report-format", "json", "--report-path", str(report)]
        args += ["--staged"] if staged else ["--log-opts", f"{base}..HEAD"]
        args.append(str(ROOT))
        environment = {key: value for key, value in os.environ.items()
                       if key not in {"GITLEAKS_CONFIG", "GITLEAKS_CONFIG_TOML"}}
        result = subprocess.run(args, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                text=True, timeout=180, env=environment)
        if result.returncode not in (0, 1):
            raise RuntimeError("密钥检查工具运行失败，已停止上传。")
        if not report.exists():
            if result.returncode:
                raise RuntimeError("密钥检查没有返回可核对的结果，已停止上传。")
            return
        try:
            findings = json.loads(report.read_text() or "[]")
        except (OSError, ValueError) as error:
            raise RuntimeError("无法读取密钥检查结果，已停止上传。") from error
        unapproved = []
        for finding in findings:
            name = finding.get("File", "")
            source_name = name.removeprefix("Easel/")
            key = (source_name, finding.get("RuleID"), finding.get("StartLine"))
            expected = AUDITED_FINDINGS.get(key)
            revision = finding.get("Commit", "")
            spec = (":" + name) if staged else (revision + ":" + name)
            actual = git("rev-parse", "--verify", spec, check=False).stdout.strip() if expected else ""
            if not expected or actual != expected:
                unapproved.append(f"{name}:{finding.get('StartLine', '?')} ({finding.get('RuleID', 'unknown')})")
        if unapproved:
            raise RuntimeError("密钥检查发现需人工核实的内容，未上传：" + "、".join(unapproved))
        if result.returncode and not findings:
            raise RuntimeError("密钥检查返回异常，已停止上传。")


def scan_document_xml(revision):
    """Scan document XML too; Git scanners otherwise see only ZIP binaries."""
    if not SCANNER.is_file() or not os.access(SCANNER, os.X_OK):
        raise RuntimeError("缺少本机密钥检查工具 scripts/gitleaks，已停止上传。")
    with tempfile.TemporaryDirectory(prefix="tailki-document-scan-") as directory:
        content = Path(directory) / "content"
        content.mkdir()
        for artifact, name in enumerate((HANDBOOK, "scripts/examples/Tailki_模拟.xlsx")):
            spec = revision + ":" + name
            if git("cat-file", "-e", spec, check=False).returncode:
                continue
            result = subprocess.run(["git", "cat-file", "blob", spec], cwd=ROOT,
                                    capture_output=True, timeout=30)
            if result.returncode:
                raise RuntimeError("无法读取待上传文档进行检查：" + name)
            with ZipFile(io.BytesIO(result.stdout)) as archive:
                for index, member in enumerate(archive.namelist()):
                    if member.endswith((".xml", ".rels")):
                        (content / f"{artifact}-{index}.xml").write_bytes(archive.read(member))
        environment = {key: value for key, value in os.environ.items()
                       if key not in {"GITLEAKS_CONFIG", "GITLEAKS_CONFIG_TOML"}}
        result = subprocess.run([
            str(SCANNER), "dir", "--no-banner", "--no-color", "--redact=100",
            "--ignore-gitleaks-allow", "--gitleaks-ignore-path", directory,
            str(content),
        ], cwd=ROOT, capture_output=True, text=True, timeout=180, env=environment)
        if result.returncode:
            raise RuntimeError("合集或模拟表的内部内容未通过密钥检查，未上传。")


def sync(message, scope="docs"):
    area = {"docs": "docs", "easel": "Easel", "tailki": "Tailki"}[scope]
    git("fetch", "--no-tags", "origin", "+refs/heads/main:refs/remotes/origin/main", timeout=120)
    outgoing = check_outgoing()
    if outgoing:
        check_outgoing_scope(scope)
        scan_secrets(base="refs/remotes/origin/main")
    # Retire legacy tracked files outside the approved roots, preserving local copies.
    local_only = ([] if scope == "tailki" else
                  [name for name in names(git("ls-files", "-z").stdout) if not approved_path(name)])
    for start in range(0, len(local_only), 100):
        git("rm", "--cached", "--", *local_only[start:start + 100])
    paths = scope_paths(scope)
    if paths and names(git("ls-files", "--cached", "--others", "--exclude-standard", "-z", "--", *paths).stdout):
        git("add", "-A", "--", *paths)
    try:
        staged = names(git("diff", "--cached", "--name-only", "-z").stdout)
        unexpected = [name for name in staged if name not in local_only and not in_scope(name, scope)]
        if unexpected:
            raise RuntimeError("另一范围出现已暂存改动，已停止并保留：" + "、".join(unexpected))
        approved_tree = git("write-tree").stdout.strip()
        check_tree(approved_tree)
        changed = git("diff", "--cached", "--quiet", check=False)
        if changed.returncode == 1:
            scan_secrets(staged=True)
            if scope == "tailki":
                scan_document_xml(approved_tree)
            if git("write-tree").stdout.strip() != approved_tree:
                raise RuntimeError("检查期间暂存内容发生变化，已停止上传。")
            git("commit", "--quiet", "-m", message)
            print(f"已创建 {area} 提交：" + git("rev-parse", "--short", "HEAD").stdout.strip(), flush=True)
            if git("rev-parse", "HEAD^{tree}").stdout.strip() != approved_tree:
                raise RuntimeError("提交内容与检查时不一致，已保留本地提交并停止上传。")
        elif changed.returncode:
            raise RuntimeError(changed.stderr.strip() or "无法检查暂存内容。")
    except (RuntimeError, subprocess.SubprocessError, OSError, ValueError):
        # Restore only paths this invocation staged. Preserve unrelated staging
        # added by another process while this sync was running.
        git("reset", "--quiet", "--", *paths, *local_only, check=False)
        raise
    check_tree("HEAD")
    if check_outgoing():
        check_outgoing_scope(scope)
        scan_secrets(base="refs/remotes/origin/main")
        if scope == "tailki":
            for revision in git("rev-list", "refs/remotes/origin/main..HEAD").stdout.splitlines():
                scan_document_xml(revision)
    print("正在上传到 GitHub……", flush=True)
    result = git("-c", "push.followTags=false", "push", "--porcelain", "--set-upstream", "origin", "HEAD:refs/heads/main", timeout=600)
    print(result.stdout.strip())
    remote = git("ls-remote", "origin", "refs/heads/main", timeout=60).stdout.split()
    if not remote or remote[0] != git("rev-parse", "HEAD").stdout.strip():
        raise RuntimeError("上传后的远端提交核对未通过，请重新检查。")
    print(f"上传成功：本次同步范围为 {area}；本机私人资料已保留。")


def main():
    parser = argparse.ArgumentParser(description="按范围提交上传；Tailki 只上传整理后的合集、技能与程序，历史文档和真实业务资料仅留本机。")
    parser.add_argument("message", nargs="?", default="Update " + datetime.now().strftime("%Y-%m-%d %H:%M"))
    parser.add_argument("--scope", choices=("docs", "easel", "tailki"), default="docs", help="Tailki 整理内容使用 --scope tailki；默认 docs 保留旧入口兼容性")
    parser.add_argument("--check-only", action="store_true", help="仅检查，不暂存、不提交、不上传")
    args = parser.parse_args()
    lock_path = Path(git("rev-parse", "--git-path", "tailki-sync.lock").stdout.strip())
    if not lock_path.is_absolute():
        lock_path = ROOT / lock_path
    with open(lock_path, "w") as lock:
        try:
            fcntl.flock(lock.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as error:
            raise RuntimeError("已有上传任务运行，请稍后重试。") from error
        check_workspace(args.scope)
        if args.check_only:
            print(f"检查通过：目标为约定仓库，范围为 {args.scope}；私人资料保留在本机。")
            return
        sync(args.message, args.scope)


if __name__ == "__main__":
    try:
        main()
    except (RuntimeError, subprocess.SubprocessError, OSError, ValueError) as error:
        print(f"未完成上传：{error}\n本地文件和已创建的提交均已保留。", file=sys.stderr)
        sys.exit(1)
