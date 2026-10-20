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


def read(path):
    with tempfile.TemporaryDirectory(prefix="tailki-read-") as work:
        book = workflow.read_book(path, work)
    if [s["name"] for s in book["sheets"]] != SHEETS:
        raise mail.DraftError("当前入口只支持达人开发主表、长期合作意愿两张表；已停止。")
    sheet = book["sheets"][0]
    sheet["creators"] = [r for r in sheet["rows"] if re.fullmatch(r"(?:KOL|SIM)-\d+", str(r.get("达人ID", "")))]
    ids = [r["达人ID"] for r in sheet["creators"]]
    if len(ids) != len(set(ids)):
        raise mail.DraftError("表内达人 ID 重复，先核实。")
    return sheet


def main_part(z):
    book = ET.fromstring(z.read("xl/workbook.xml"))
    rels = workflow.rels(z, "xl/workbook.xml")
    sheet = next(s for s in book.find("s:sheets", NS) if s.get("name") == MASTER)
    return workflow.resolve_part("xl/workbook.xml", rels[sheet.get("{" + NS["r"] + "}id")].get("Target"))


def patch_cells(path, cells, *, expected_hash, backup_dir, style_row, new_rows=()):
    """Call under the shared state lock. Reject stale sources and unknown layouts."""
    path = Path(path)
    original = path.read_bytes()
    if mail.digest(original) != expected_hash:
        raise mail.DraftError("表格已被修改，停止覆盖；请重新读取。")
    with ZipFile(path) as z:
        part = main_part(z)
        raw = z.read(part).decode("utf-8")
        root = ET.fromstring(raw)
        if root.find("s:tableParts", NS) is not None or root.find("s:sheetProtection", NS) is not None:
            raise mail.DraftError("表格新增了表对象或保护设置，需要核实后再扩展。")
        if re.search(r"<(?:\w+:)?sheetData\b", raw).group() != "<sheetData":
            raise mail.DraftError("工作表命名空间布局已变化，停止覆盖。")
        native_rows = {int(r.get("r")): r for r in root.findall("s:sheetData/s:row", NS)}
        template = native_rows[style_row]
        styles = {workflow.col_number(c.get("r")): c.get("s", "0") for c in template}
        for number in new_rows:
            existing = native_rows.get(number)
            if existing is not None:
                for cell in existing:
                    if cell.find("s:is", NS) is not None or cell.find("s:v", NS) is not None:
                        if "".join(cell.itertext()).strip():
                            raise mail.DraftError("拟新增行已有内容，保留人工记录。")
                    if cell.find("s:f", NS) is not None and workflow.col_number(cell.get("r")) != 47:
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
            pattern = rf'<row\b[^>]*\br="{number}"[^>]*(?:/>|>.*?</row>)'
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
                cp = rf'<c\b[^>]*\br="{address}"[^>]*(?:/>|>.*?</c>)'
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
            # Only the state validation has a finite populated-row range in this layout.
            def extend(match):
                xml = match.group()
                if not re.search(r'\bsqref="C\d', xml):
                    return xml
                if "待首联确认" not in xml:
                    xml = xml.replace("<formula1>&quot;", "<formula1>&quot;待首联确认,")
                    xml = xml.replace('<formula1>"', '<formula1>"待首联确认,')
                return re.sub(r'sqref="([^"]*)"', lambda m: 'sqref="' + m[1] + " " +
                              " ".join(f"C{n}" for n in new_rows) + '"', xml, count=1)
            raw = re.sub(r"<dataValidation\b.*?</dataValidation>", extend, raw, flags=re.S)
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
