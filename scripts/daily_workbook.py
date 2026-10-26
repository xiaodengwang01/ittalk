"""Narrow cell edits for the current two-sheet workbook, without a full rewrite.

Artifact Tool authors typed cells/formulas. XML splicing retains all unrelated
parts, rows, native styles, hidden columns, shared formulas and annotations.
"""
from __future__ import annotations

import copy
import os
from pathlib import Path
import re
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from zipfile import ZipFile

import mail_drafts as mail
import workflow

NS = workflow.NS
S = NS["s"]
ET.register_namespace("", S)
MASTER = "达人开发主表"
SHEETS = [MASTER, "长期合作意愿"]
COMPACT_HEADERS = ["达人ID", "达人昵称", "当前建联状况", "沟通摘要", "Email", "WhatsApp联系方式", "可沟通时间",
                   "收件信息", "待发送回复梗概", "往来邮件梗概", "历史批次", "首联内容分类", "初次建联日期",
                   "账号链接", "平台", "宠物类型", "内容类型", "粉丝数", "近10条最高播放", "近10条中位播放",
                   "内容能力(1-5)", "宠物人群匹配度(1-5)", "播放能力(1-5)", "联系方式完整度(1-5)",
                   "综合评分", "是否低粉高播放", "是否带货类似产品", "备注"]


def compact_controls(path):
    """Native outline/filter controls missing from the supported authoring API.

    Called only on a freshly authored compact workbook. No source data or raw
    mail is copied here. Ordinary cell updates preserve these native controls.
    """
    from io import BytesIO
    buffer = BytesIO()
    with ZipFile(path) as source, ZipFile(buffer, "w") as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", info.filename):
                root = ET.fromstring(data)
                for c in root.findall("s:cols/s:col", NS):
                    if int(c.get("min")) >= 11:
                        c.set("outlineLevel", "1"); c.set("hidden", "1")
                cols = root.find("s:cols", NS)
                ET.SubElement(cols, "{"+S+"}col", {"min":"29","max":"29","width":"3","customWidth":"1","collapsed":"1"})
                prop = root.find("s:sheetPr", NS)
                if prop is None: prop=ET.Element("{"+S+"}sheetPr");root.insert(0,prop)
                ET.SubElement(prop,"{"+S+"}outlinePr",{"summaryRight":"1","showOutlineSymbols":"1"})
                fmt = root.find("s:sheetFormatPr", NS)
                if fmt is not None: fmt.set("outlineLevelCol","1")
                rows = root.find("s:sheetData",NS)
                end = max(int(r.get("r")) for r in rows)
                root.insert(list(root).index(rows)+1,ET.Element("{"+S+"}autoFilter",{"ref":f"A1:AB{end}"}))
                for view in root.findall("s:sheetViews/s:sheetView",NS):
                    for node in list(view):
                        if node.tag in ("{"+S+"}pane","{"+S+"}selection"): view.remove(node)
                    view.insert(0, ET.Element("{"+S+"}pane", {"xSplit":"2","ySplit":"1","topLeftCell":"C2","activePane":"bottomRight","state":"frozen"}))
                    ET.SubElement(view,"{"+S+"}selection",{"pane":"bottomRight","activeCell":"C2","sqref":"C2"})
                data=ET.tostring(root,encoding="utf-8",xml_declaration=True)
            target.writestr(info,data)
    Path(path).write_bytes(buffer.getvalue())


def read(path):
    with tempfile.TemporaryDirectory(prefix="tailki-read-") as work:
        book = workflow.read_book(path, work)
    if [s["name"] for s in book["sheets"]] != SHEETS:
        raise mail.DraftError("当前入口只支持达人开发主表、长期合作意愿两张表；已停止。")
    main = book["sheets"][0]
    creators = {}
    for sheet in book["sheets"]:
        sheet["creators"] = []
        for row in sheet["rows"]:
            if not re.fullmatch(r"(?:KOL|SIM)-\d+", str(row.get("达人ID", ""))): continue
            row["_sheet"], row["_headers"] = sheet["name"], sheet["headers"]
            if "当前建联状况" in sheet["headers"]:
                row["联系状态"] = row.get("当前建联状况")  # Read-only compatibility alias.
            identifier = row["达人ID"]
            if identifier in creators:
                if creators[identifier]["_sheet"] == sheet["name"] or "往来邮件梗概" in sheet["headers"]:
                    raise mail.DraftError("两表存在重复达人 ID，先合并记录。")
            else:
                creators[identifier] = row
            sheet["creators"].append(row)
    return {**main, "main_creators": main["creators"], "creators": list(creators.values()),
            "sheets": book["sheets"]}


def locate(path, creator_id):
    book = read(path)
    matches = [r for r in book["creators"] if r["达人ID"] == creator_id]
    if len(matches) != 1: raise mail.DraftError("未找到唯一达人 ID：" + creator_id)
    row = matches[0]
    return next(s for s in book["sheets"] if s["name"] == row["_sheet"]), row


def address(headers, header, number):
    if headers.count(header) != 1: raise mail.DraftError("表头缺失或重复：" + header)
    return workflow.col_name(headers.index(header) + 1) + str(number)


def score_formula(headers, number):
    names = ("内容能力(1-5)", "宠物人群匹配度(1-5)", "播放能力(1-5)", "联系方式完整度(1-5)")
    refs = ",".join(address(headers, h, number) for h in names)
    return f'=IF(COUNTA({refs})=0,"",AVERAGE({refs}))'


def update_fields(path, creator_id, fields, *, expected_hash):
    sheet, row = locate(path, creator_id)
    cells = {address(sheet["headers"], h, row["_row"]): {"value": v} for h, v in fields.items()}
    patch_cells(path, cells, expected_hash=expected_hash, backup_dir=None,
                style_row=row["_row"], sheet_name=sheet["name"])


def main_part(z, sheet_name=MASTER):
    if sheet_name not in SHEETS:
        raise mail.DraftError("需要明确指定现有工作表。")
    book = ET.fromstring(z.read("xl/workbook.xml"))
    rels = workflow.rels(z, "xl/workbook.xml")
    sheet = next(s for s in book.find("s:sheets", NS) if s.get("name") == sheet_name)
    return workflow.resolve_part("xl/workbook.xml", rels[sheet.get("{" + NS["r"] + "}id")].get("Target"))


def patch_cells(path, cells, *, expected_hash, backup_dir, style_row, new_rows=(), sheet_name=MASTER):
    """Call under the shared state lock. Reject stale sources and unknown layouts."""
    path = Path(path)
    original = path.read_bytes()
    if mail.digest(original) != expected_hash:
        raise mail.DraftError("表格已被修改，停止覆盖；请重新读取。")
    with ZipFile(path) as z:
        part = main_part(z, sheet_name)
        raw = z.read(part).decode("utf-8")
        root = ET.fromstring(raw)
        if root.find("s:tableParts", NS) is not None or root.find("s:sheetProtection", NS) is not None:
            raise mail.DraftError("表格新增了表对象或保护设置，需要核实后再扩展。")
        if re.search(r"<(?:\w+:)?sheetData\b", raw).group() != "<sheetData":
            raise mail.DraftError("工作表命名空间布局已变化，停止覆盖。")
        native_rows = {int(r.get("r")): r for r in root.findall("s:sheetData/s:row", NS)}
        sheet = next(s for s in read(path)["sheets"] if s["name"] == sheet_name)
        score_column = sheet["headers"].index("综合评分") + 1 if "综合评分" in sheet["headers"] else None
        template = native_rows[style_row]
        styles = {workflow.col_number(c.get("r")): c.get("s", "0") for c in template}
        for number in new_rows:
            existing = native_rows.get(number)
            if existing is not None:
                for cell in existing:
                    if cell.find("s:is", NS) is not None or cell.find("s:v", NS) is not None:
                        if "".join(cell.itertext()).strip():
                            raise mail.DraftError("拟新增行已有内容，保留人工记录。")
                    if cell.find("s:f", NS) is not None and workflow.col_number(cell.get("r")) != score_column:
                        raise mail.DraftError("拟新增行含非评分公式，保留人工记录。")
    with tempfile.TemporaryDirectory(prefix="tailki-cells-") as tmp:
        tmp = Path(tmp)
        spec, authored = tmp / "cells.json", tmp / "cells.xlsx"
        spec.write_text(__import__("json").dumps(cells, ensure_ascii=False))
        node = os.environ.get("TAILKI_NODE", str(Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node"))
        subprocess.run([node, str(Path(__file__).with_name("daily_cells.mjs")), str(spec), str(authored)],
                       check=True, capture_output=True, text=True, timeout=120)
        with ZipFile(authored) as z:
            strings = [copy.deepcopy(s) for s in ET.fromstring(z.read("xl/sharedStrings.xml"))] if "xl/sharedStrings.xml" in z.namelist() else []
            generated = {c.get("r"): c for c in ET.fromstring(z.read("xl/worksheets/sheet1.xml")).findall("s:sheetData/s:row/s:c", NS)}
        changed = {}
        for address in cells:
            cell = copy.deepcopy(generated[address])
            number = int(re.search(r"\d+", address).group())
            native = native_rows.get(number)
            old = next((c for c in native if c.get("r") == address), None) if native is not None else None
            cell.set("s", old.get("s", "0") if old is not None and number not in new_rows
                     else styles.get(workflow.col_number(address), "0"))
            if cell.get("t") == "s":
                value = cell.find("s:v", NS)
                inline = copy.deepcopy(strings[int(value.text)])
                inline.tag = "{" + S + "}is"
                cell.remove(value)
                cell.append(inline)
                cell.set("t", "inlineStr")
            if cell.get("t") == "e":
                raise mail.DraftError("生成单元格含计算错误，未修改原表。")
            changed.setdefault(number, []).append(cell)
        for number, edits in sorted(changed.items()):
            pattern = rf'<row\b[^>]*\br="{number}"[^>]*?(?:/>|>.*?</row>)'
            match = re.search(pattern, raw, flags=re.S)
            if number in new_rows:
                # Empty preformatted rows often have the default 15pt height.
                # Extend the last real creator's layout, without copying its data.
                shell = copy.deepcopy(template)
                shell.set("r", str(number))
                shell.attrib.pop("spans", None)
                for cell in shell:
                    cell.set("r", re.sub(r"\d+$", str(number), cell.get("r")))
                    cell.attrib.pop("t", None)
                    for child in list(cell):
                        cell.remove(child)
                row = ET.tostring(shell, encoding="unicode")
            elif match:
                row = match.group()
                if row.endswith("/>"):
                    row = row[:-2] + "></row>"
            else:
                attrs = {k: v for k, v in template.attrib.items() if k not in ("r", "spans")}
                attrs["r"] = str(number)
                row = ET.tostring(ET.Element("{" + S + "}row", attrs), encoding="unicode")
                row = row.replace(" />", "></row>")
            for cell in edits:
                address = cell.get("r")
                xml = ET.tostring(cell, encoding="unicode")
                # Stop at a self-closing tag before considering a paired tag.
                # A greedy prefix can consume its slash, then match through the
                # next cell's closing tag and silently remove adjacent data.
                cp = rf'<c\b[^>]*\br="{address}"[^>]*?(?:/>|>.*?</c>)'
                if re.search(cp, row, flags=re.S):
                    row = re.sub(cp, lambda _: xml, row, count=1, flags=re.S)
                else:
                    # Keep cells in ascending column order without reserializing old cells.
                    later = next((c for c in re.finditer(r'<c\b[^>]*\br="([A-Z]+)\d+"', row)
                                  if workflow.col_number(c.group(1)) > workflow.col_number(address)), None)
                    at = later.start() if later else row.rfind("</row>")
                    row = row[:at] + xml + row[at:]
            if match:
                raw = raw[:match.start()] + row + raw[match.end():]
            else:
                later = next((r for r in re.finditer(r'<row\b[^>]*\br="(\d+)"', raw) if int(r.group(1)) > number), None)
                at = later.start() if later else raw.index("</sheetData>")
                raw = raw[:at] + row + raw[at:]
        if new_rows:
            status_header = "当前建联状况" if "当前建联状况" in sheet["headers"] else "联系状态"
            status_col = workflow.col_name(sheet["headers"].index(status_header) + 1)
            # Only the state validation has a finite populated-row range in this layout.
            def extend(match):
                xml = match.group()
                if not re.search(r'\bsqref="' + status_col + r'\d', xml):
                    return xml
                if "待首联确认" not in xml:
                    xml = xml.replace("<formula1>&quot;", "<formula1>&quot;待首联确认,")
                    xml = xml.replace('<formula1>"', '<formula1>"待首联确认,')
                return re.sub(r'sqref="([^"]*)"', lambda m: 'sqref="' + m[1] + " " +
                              " ".join(f"{status_col}{n}" for n in new_rows) + '"', xml, count=1)
            raw = re.sub(r"<dataValidation\b.*?</dataValidation>", extend, raw, flags=re.S)
            raw = re.sub(r'(<autoFilter\b[^>]*ref="[A-Z]+\d+:[A-Z]+)\d+("[^>]*>)',
                         lambda m: m[1] + str(max(new_rows)) + m[2], raw)
            dimension = root.find("s:dimension", NS)
            if dimension is not None:
                tail = dimension.get("ref").split(":")[-1]
                end = max(int(re.search(r"\d+", tail).group()), max(new_rows))
                ref = f"A1:{workflow.col_name(workflow.col_number(tail))}{end}"
                raw = re.sub(r'<dimension\b[^>]*/>', f'<dimension ref="{ref}"/>', raw, count=1)
        ET.fromstring(raw)
        with ZipFile(path) as source, ZipFile(tmp / "result.xlsx", "w") as output:
            for info in source.infolist():
                output.writestr(info, raw.encode() if info.filename == part else source.read(info.filename))
        # Both sheets still parse; a concurrent edit must never be silently lost.
        read(tmp / "result.xlsx")
        if mail.digest(path.read_bytes()) != expected_hash:
            raise mail.DraftError("写入前发现人工修改，已保留原表。")
        if backup_dir is not None and "往来邮件梗概" not in sheet["headers"]:
            backup_dir = Path(backup_dir)
            backup_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
            backup = backup_dir / (expected_hash + ".xlsx")
            if not backup.exists():
                backup.write_bytes(original)
                os.chmod(backup, 0o600)
        fd, replacement = tempfile.mkstemp(prefix=".tailki-", suffix=".xlsx", dir=path.parent)
        try:
            with os.fdopen(fd, "wb") as out:
                out.write((tmp / "result.xlsx").read_bytes())
                out.flush()
                os.fsync(out.fileno())
            os.chmod(replacement, path.stat().st_mode & 0o777)
            if mail.digest(path.read_bytes()) != expected_hash:
                raise mail.DraftError("替换前发现人工修改，已保留原表。")
            os.replace(replacement, path)
        finally:
            if os.path.exists(replacement):
                os.unlink(replacement)
