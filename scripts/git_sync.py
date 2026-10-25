#!/usr/bin/env python3
"""Commit and push this project's approved files without overwriting remote work."""
import argparse
from datetime import datetime
import fcntl
import os
from pathlib import Path
import subprocess
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parent.parent
REMOTE = "https://github.com/xiaodengwang01/ittalk.git"
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
MOCK_WORKBOOK = Path("scripts/examples/Tailki_模拟.xlsx")
PUBLIC_ROOT_FILES = {".gitignore", ".gitattributes", "README.md", "AGENTS.md", "一键上传.command", "首次授权并上传.command"}
# These reference images were explicitly approved for publication.
PUBLIC_COUNTING_IMAGES = {
    "docs/计数聚合题素材/第一批/01_双色胶囊.jpg",
    "docs/计数聚合题素材/第一批/02_彩色胶囊.png",
    "docs/计数聚合题素材/第一批/03_彩色纽扣.jpg",
    "docs/计数聚合题素材/第一批/04_彩色珠子.jpg",
    "docs/计数聚合题素材/第二批/01_彩色骰子.jpg",
    "docs/计数聚合题素材/第二批/02_多米诺骨牌.jpg",
    "docs/计数聚合题素材/第二批/03_彩色积木.jpg",
    "docs/计数聚合题素材/第二批/04_混合意面.webp",
    "docs/计数聚合题素材/第二批/05_彩色木夹.jpg",
    "docs/计数聚合题素材/第三批/01_国际象棋.jpg",
    "docs/计数聚合题素材/第三批/02_金属钥匙.jpg",
    "docs/计数聚合题素材/第三批/03_混合坚果.jpg",
    "docs/计数聚合题素材/第三批/04_贝壳与海星.jpg",
    "docs/计数聚合题素材/第三批/05_五金零件.jpg",
    "docs/计数聚合题素材/第四批/01_玩具汽车.jpg",
    "docs/计数聚合题素材/第四批/02_彩色图钉.jpg",
    "docs/计数聚合题素材/第四批/03_恐龙主题玩具.jpg",
    "docs/计数聚合题素材/第四批/04_秋叶叶片.jpg",
    "docs/计数聚合题素材/第四批/05_混合浆果.jpg",
    "docs/计数聚合题素材/第五批/01_彩色铅笔.jpg",
    "docs/计数聚合题素材/第五批/02_彩色缝纫线轴.jpg",
    "docs/计数聚合题素材/第五批/03_缤纷甜甜圈.jpg",
    "docs/计数聚合题素材/第五批/04_蘑菇与菌簇.jpg",
    "docs/计数聚合题素材/第五批/05_欧元教学硬币.jpg",
}


def git(*args, check=True, timeout=30):
    result = subprocess.run(
        ["git", *args], cwd=ROOT, text=True, stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, timeout=timeout,
        env={**os.environ, "GIT_TERMINAL_PROMPT": "0"},
    )
    if check and result.returncode:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "Git 操作失败。")
    return result


def check_mock_workbook():
    file = ROOT / MOCK_WORKBOOK
    # The sample is optional; deleting it must not block source-only uploads.
    if not file.exists():
        return
    with ZipFile(file) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            strings = ["".join(item.itertext()) for item in root]

        def value(cell):
            if cell.attrib.get("t") == "inlineStr":
                return "".join(cell.find("s:is", NS).itertext())
            node = cell.find("s:v", NS)
            if node is None or node.text is None:
                return ""
            return strings[int(node.text)] if cell.attrib.get("t") == "s" else node.text

        book = ET.fromstring(archive.read("xl/workbook.xml"))
        if [s.attrib["name"] for s in book.find("s:sheets", NS)] != ["达人名单", "发送记录", "邮件往来"]:
            raise RuntimeError("模拟表工作表结构已改变，请先核查。")
        for index in range(1, 4):
            sheet = ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml"))
            for row in sheet.findall("s:sheetData/s:row", NS):
                if int(row.attrib["r"]) == 1:
                    continue
                cells = {"".join(c for c in cell.attrib["r"] if c.isalpha()): value(cell) for cell in row}
                if not any(cells.values()):
                    continue
                if index != 1:
                    raise RuntimeError("模拟表已含邮件记录，停止上传，请先核查。")
                if not (
                    cells.get("A", "").startswith("SIM-")
                    and cells.get("B", "").startswith("模拟达人")
                    and cells.get("D", "").endswith(".invalid")
                    and cells.get("K") == "否"
                ):
                    raise RuntimeError("模拟表包含非模拟联系人或开启了发送许可，停止上传。")


def approved_path(name):
    path = Path(name)
    if name in PUBLIC_ROOT_FILES:
        return True
    if name in PUBLIC_COUNTING_IMAGES:
        return True
    if path == MOCK_WORKBOOK:
        return True
    return path.parts[0] in {"scripts", "docs"} and path.suffix in {".py", ".mjs", ".js", ".sh", ".command", ".md", ".txt"}


def main():
    parser = argparse.ArgumentParser(description="提交并上传 Tailki 项目。")
    parser.add_argument("message", nargs="?", default="Update Tailki workspace " + datetime.now().strftime("%Y-%m-%d %H:%M"))
    parser.add_argument("--check-only", action="store_true", help="仅检查，不暂存、不提交、不上传")
    args = parser.parse_args()
    if git("remote", "get-url", "origin").stdout.strip() != REMOTE:
        raise RuntimeError("origin 不是约定的 GitHub 仓库，已停止。")
    if git("branch", "--show-current").stdout.strip() != "main":
        raise RuntimeError("当前分支不是 main，已停止。")
    check_mock_workbook()
    tracked = git("ls-files", "--cached", "--others", "--exclude-standard", "-z").stdout.split("\0")
    rejected = sorted({name for name in tracked if name and (ROOT / name).exists() and not approved_path(name)})
    if rejected:
        raise RuntimeError("存在未批准上传的文件：" + "、".join(rejected))
    if args.check_only:
        print("检查通过：目标仓库、模拟数据和上传文件范围均正确。")
        return
    with open(ROOT / ".git" / "tailki-sync.lock", "w") as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as error:
            raise RuntimeError("已有上传任务运行，请稍后重试。") from error
        git("add", "-A", "--", ".")
        if git("diff", "--cached", "--quiet", check=False).returncode == 1:
            print(git("commit", "-m", args.message).stdout.strip(), flush=True)
        print("正在上传到 GitHub……", flush=True)
        result = git("push", "--porcelain", "--set-upstream", "origin", "main")
        print(result.stdout.strip())
        print("上传成功。")


if __name__ == "__main__":
    try:
        main()
    except (RuntimeError, subprocess.TimeoutExpired, OSError, ValueError) as error:
        print(f"未完成上传：{error}\n本地文件和已创建的提交均已保留。", file=sys.stderr)
        sys.exit(1)
