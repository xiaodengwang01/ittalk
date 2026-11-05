"""Private fields removed from the visible tracker, scoped to its local path.

Call writes under mail.State's lock. Workbook values always take precedence;
archived workbook values are not promoted to verified email facts.
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import tempfile

FIELDS = {"历史批次", "初次建联日期", "平台", "宠物类型", "近10条最高播放", "近10条中位播放",
          "内容能力(1-5)", "宠物人群匹配度(1-5)", "播放能力(1-5)", "联系方式完整度(1-5)",
          "综合评分", "是否低粉高播放", "是否带货类似产品", "备注",
          "WhatsApp联系方式", "可沟通时间", "收件信息", "待发送回复梗概"}


def key(book, creator_id):
    scope = hashlib.sha256(str(Path(book).resolve()).encode()).hexdigest()[:24]
    return f"workbook-fields:{scope}:{creator_id}"


def load(state, book, creator_id):
    if not state.db.execute("SELECT 1 FROM sqlite_master WHERE name='daily_state'").fetchone():
        return {}
    result = state.db.execute("SELECT data FROM daily_state WHERE key=?", (key(book, creator_id),)).fetchone()
    return json.loads(result[0]) if result else {}


def overlay(state, book, row):
    hidden = load(state, book, row.get("达人ID", "")).get("fields", {})
    pending = load(state, book, "__pending_write__")
    if pending:
        current = hashlib.sha256(Path(book).read_bytes()).hexdigest()
        if current == pending["after_hash"]:
            hidden = {**hidden, **pending["fields"].get(row.get("达人ID"), {})}
        elif current != pending["before_hash"]:
            raise ValueError("表格在本机字段写回中断后又被修改，需核对后继续。")
    return {**hidden, **row}


def update(state, book, creator_id, fields, *, archive_hash=None):
    unknown = set(fields) - FIELDS
    if unknown:
        raise ValueError("不是允许保存在本机的工作簿字段：" + ",".join(sorted(unknown)))
    if not fields:
        return
    value = load(state, book, creator_id)
    if archive_hash:
        snapshot = {"source_hash": archive_hash, "fields": fields,
                    "source": "用户授权删列前的工作簿记录；未重新核实邮件事实"}
        if value.get("archive") and value["archive"] != snapshot:
            raise ValueError("已有不同删列存档，停止覆盖。")
        value["archive"] = snapshot
    value.update(book=str(Path(book).resolve()), creator_id=creator_id)
    value["fields"] = {**value.get("fields", {}), **fields}
    with state.db:
        state.db.execute("CREATE TABLE IF NOT EXISTS daily_state (key TEXT PRIMARY KEY, data TEXT NOT NULL)")
        state.db.execute("INSERT OR REPLACE INTO daily_state VALUES (?,?)",
                         (key(book, creator_id), json.dumps(value, ensure_ascii=False)))


def may_store(headers, field):
    compact_followup = {"达人ID", "当前建联状况", "往来邮件梗概"}.issubset(headers)
    return (field in FIELDS and field not in headers
            and ("新增批次" in headers
                 or (compact_followup and field == "待发送回复梗概")))


def finish_write(state, book):
    pending = load(state, book, "__pending_write__")
    if not pending: return
    current = hashlib.sha256(Path(book).read_bytes()).hexdigest()
    if current == pending["before_hash"] and current != pending["after_hash"]: return
    if current != pending["after_hash"]:
        raise ValueError("中断后表格与预期版本不一致，保留人工修改。")
    for identifier, fields in pending["fields"].items():
        update(state, book, identifier, fields)
    with state.db:
        state.db.execute("DELETE FROM daily_state WHERE key=?", (key(book, "__pending_write__"),))


def stage_write(state, book, before_hash, after_hash, fields):
    """Bridge the atomic XLSX replacement and SQLite commit without replaying mail."""
    finish_write(state, book)
    if hashlib.sha256(Path(book).read_bytes()).hexdigest() != before_hash:
        raise ValueError("写入前表格变化，停止覆盖。")
    if not fields: return
    if any(set(values) - FIELDS for values in fields.values()):
        raise ValueError("包含不允许写入的本机字段。")
    payload = {"before_hash": before_hash, "after_hash": after_hash, "fields": fields}
    with state.db:
        state.db.execute("CREATE TABLE IF NOT EXISTS daily_state (key TEXT PRIMARY KEY, data TEXT NOT NULL)")
        state.db.execute("INSERT OR REPLACE INTO daily_state VALUES (?,?)",
                         (key(book, "__pending_write__"), json.dumps(payload, ensure_ascii=False)))


def publish_bytes(state, book, before_hash, data, fields):
    book = Path(book)
    stage_write(state, book, before_hash, hashlib.sha256(data).hexdigest(), fields)
    fd, temporary = tempfile.mkstemp(prefix=".tailki-fields-", suffix=".xlsx", dir=book.parent)
    try:
        with os.fdopen(fd, "wb") as out:
            out.write(data); out.flush(); os.fsync(out.fileno())
        os.chmod(temporary, book.stat().st_mode & 0o777)
        if hashlib.sha256(book.read_bytes()).hexdigest() != before_hash:
            raise ValueError("替换前表格发生变化，已保留原文件。")
        os.replace(temporary, book)
        finish_write(state, book)
    finally:
        if os.path.exists(temporary): os.unlink(temporary)
