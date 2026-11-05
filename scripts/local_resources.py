"""Resolve byte-identical archived documents without rewriting historical evidence."""
from __future__ import annotations

from datetime import datetime
import hashlib
import json
from pathlib import Path
import sqlite3
from zoneinfo import ZoneInfo

DEFAULT_STATE = Path.home() / "Library/Application Support/Tailki/mail-drafts.sqlite3"


def checksum(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def archive_key(path):
    return "document-archive:" + hashlib.sha256(str(Path(path).resolve()).encode()).hexdigest()


def mapping(path, state=None):
    db = state.db if state else None
    owned = False
    try:
        if db is None:
            if not DEFAULT_STATE.is_file(): return None
            db = sqlite3.connect(DEFAULT_STATE.as_uri() + "?mode=ro", uri=True)
            owned = True
        if not db.execute("SELECT name FROM sqlite_master WHERE name='daily_state'").fetchone(): return None
        row = db.execute("SELECT data FROM daily_state WHERE key=?", (archive_key(path),)).fetchone()
        return json.loads(row[0]) if row else None
    finally:
        if owned: db.close()


def resolve(path, state=None, expected_hash=None):
    original = Path(path).expanduser().resolve()
    if original.is_file():
        if expected_hash and checksum(original) != expected_hash:
            raise ValueError("原文档与历史校验值不符：" + str(original))
        return original
    record = mapping(original, state)
    if not record:
        raise FileNotFoundError("未找到文档或归档映射：" + str(original))
    target = Path(record["archive"])
    if checksum(target) != record["hash"] or (expected_hash and record["hash"] != expected_hash):
        raise ValueError("归档原件校验不符：" + str(original))
    return target


def archive(state, source, target):
    """Move one file unchanged; an existing destination is never overwritten."""
    source, target = Path(source).resolve(), Path(target).resolve()
    previous = mapping(source, state)
    if not source.exists() and previous:
        resolve(source, state, previous["hash"])
        return previous
    if source == target or target.exists():
        raise ValueError("归档目标已存在，未覆盖：" + str(target))
    record = {"original": str(source), "archive": str(target), "hash": checksum(source),
              "bytes": source.stat().st_size,
              "archived_at": datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(timespec="seconds")}
    if previous and any(previous[k] != record[k] for k in ("original", "archive", "hash")):
        raise ValueError("已有归档映射不一致。")
    target.parent.mkdir(parents=True, exist_ok=True)
    with state.db:
        state.db.execute("CREATE TABLE IF NOT EXISTS daily_state (key TEXT PRIMARY KEY, data TEXT NOT NULL)")
        state.db.execute("INSERT OR REPLACE INTO daily_state VALUES (?,?)",
                         (archive_key(source), json.dumps(record, ensure_ascii=False)))
    source.rename(target)
    if checksum(target) != record["hash"]: raise ValueError("归档读验失败。")
    return record
