"""One local message ledger; projections and delivery remain separate operations."""
from __future__ import annotations

from email.utils import getaddresses
from datetime import datetime
import json

import mail_drafts as mail


def initialize(state):
    state.db.execute("CREATE TABLE IF NOT EXISTS mail_records (id TEXT PRIMARY KEY, data TEXT NOT NULL)")
    state.db.execute("CREATE TABLE IF NOT EXISTS mail_actions (id TEXT PRIMARY KEY, data TEXT NOT NULL)")
    state.db.commit()


def put(state, record):
    initialize(state)
    for key in ("account", "message_id", "direction", "source"):
        if not record.get(key):
            raise mail.DraftError("单封记录缺少：" + key)
    if record["direction"] not in ("incoming", "outgoing"):
        raise mail.DraftError("邮件方向不明确。")
    record = dict(record, account=record["account"].lower())
    identifier = mail.digest([record["account"], record["direction"], record["message_id"]])
    row = state.db.execute("SELECT data FROM mail_records WHERE id=?", (identifier,)).fetchone()
    previous = json.loads(row[0]) if row else {}
    # A header-only observation can later be completed, never silently replaced.
    for key in ("creator_id", "raw_text", "mime_hash", "in_reply_to"):
        if previous.get(key) and record.get(key) and previous[key] != record[key]:
            raise mail.DraftError("同一邮件的内容或归属冲突，保留原记录：" + key)
    merged = {**previous, **{k: v for k, v in record.items() if v not in (None, "", [])},
              "id": identifier, "updated_at": mail.now()}
    merged["sources"] = sorted(set(previous.get("sources", [])) | {record["source"]})
    if merged.get("raw_text"):
        merged["body_status"] = "collected"
    if previous.get("verification") == "sent_folder":
        merged["verification"] = "sent_folder"
    with state.db:
        state.db.execute("INSERT OR REPLACE INTO mail_records VALUES (?,?)",
                         (identifier, json.dumps(merged, ensure_ascii=False)))
    return identifier


def incoming(state, record, creator_id=""):
    return put(state, {**record, "direction": "incoming", "creator_id": creator_id})


def action(state, key, **updates):
    initialize(state)
    row = state.db.execute("SELECT data FROM mail_actions WHERE id=?", (key,)).fetchone()
    value = json.loads(row[0]) if row else {}
    if updates:
        value.update(updates, updated_at=mail.now())
        with state.db:
            state.db.execute("INSERT OR REPLACE INTO mail_actions VALUES (?,?)",
                             (key, json.dumps(value, ensure_ascii=False)))
    return value


def creator_for(record, rows):
    """Shared mailboxes do not imply one creator; leave ambiguous mail unassigned."""
    addresses = {a.lower() for _, a in getaddresses([record.get("from", "")])}
    matches = [r["达人ID"] for r in rows if str(r.get("Email", "")).lower() in addresses]
    return matches[0] if len(matches) == 1 else ""


def context(state, creator_id):
    initialize(state)
    messages = [json.loads(r[0]) for r in state.db.execute("SELECT data FROM mail_records")]
    messages = [r for r in messages if r.get("creator_id") == creator_id]
    def order(record):
        try:
            moment = datetime.fromisoformat(record.get("received_at") or record.get("sent_at", ""))
            return (0, moment.timestamp(), record["id"]) if moment.tzinfo else (1, 0, record["id"])
        except ValueError:
            return (1, 0, record["id"])
    messages.sort(key=order)
    return {"creator_id": creator_id, "coverage": "仅统计本机已采集或已核验的单封记录，未补扫已读正文",
            "incoming": sum(r["direction"] == "incoming" for r in messages),
            "outgoing": sum(r["direction"] == "outgoing" for r in messages), "messages": messages}
