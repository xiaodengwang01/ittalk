#!/usr/bin/env python3
"""Collect unread inbox messages locally without marking them read or sending."""
from __future__ import annotations

import argparse
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
import imaplib
import json
from pathlib import Path
import re
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
if __package__:
    from . import daily_workflow as daily
else:
    import daily_workflow as daily

mail = daily.mail


def unread_flags(mailbox, uid):
    code, values = mailbox.client.uid("FETCH", uid, "(UID FLAGS)")
    rows = [x for x in (values or []) if isinstance(x, bytes)]
    if code != "OK" or len(rows) != 1:
        raise mail.DraftError("无法核实单封未读状态，停止读取正文。")
    matched = re.search(rb"\bUID\s+(\d+)\b", rows[0])
    flags = re.search(rb"\bFLAGS\s*\(([^)]*)\)", rows[0])
    if not matched or matched[1] != uid or not flags:
        raise mail.DraftError("邮件标识或未读状态不完整，停止读取正文。")
    return b"\\seen" not in flags[1].lower().split()


def collect_unread(state, mailbox):
    """UNSEEN + a fresh FLAGS check before each BODY.PEEK. No STORE/APPEND."""
    daily.initialize(state)
    folder = b"INBOX"
    validity = mailbox.select(folder)
    selected = mailbox.search("UNSEEN")
    records, skipped = [], []
    for uid in selected:
        if not unread_flags(mailbox, uid):
            skipped.append(uid.decode("ascii"))
            continue
        key = "inbox:" + mail.digest([mailbox.account.lower(), "INBOX", validity, uid.decode("ascii")])
        record = daily.get(state, key)
        if record is None:
            raw = mailbox.fetch(uid)
            message = BytesParser(policy=policy.default).parsebytes(raw)
            source = "INBOX:UIDVALIDITY=" + validity + ":UID=" + uid.decode("ascii")
            try:
                received = parsedate_to_datetime(str(message.get("Date", "")))
                received_at = received.isoformat() if received.tzinfo else ""
            except (TypeError, ValueError):
                received_at = ""
            record = {"account": mailbox.account, "folder": "INBOX", "uidvalidity": validity,
                      "uid": uid.decode("ascii"), "source": source,
                      "message_id": str(message.get("Message-ID", "")).strip() or source,
                      "from": str(message.get("From", "")), "to": str(message.get("To", "")),
                      "received_at": received_at, "subject": str(message.get("Subject", "")),
                      "raw_text": mail.snapshot(raw)["body"], "raw_hash": mail.digest(raw),
                      "collected_at": mail.now(), "unread_when_collected": True}
            daily.put(state, key, record)
        records.append(record)
    after = mailbox.search("UNSEEN")
    return {"account": mailbox.account, "checked_at": mail.now(), "scope": "INBOX UNSEEN",
            "unread_before": [x.decode("ascii") for x in selected],
            "unread_after": [x.decode("ascii") for x in after],
            "skipped_read_in_meantime": skipped, "messages": records}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--state", type=Path, default=mail.DEFAULT_STATE)
    args = parser.parse_args()
    state, mailbox = mail.State(args.state), None
    try:
        mailbox = mail.connect(state.config())
        result = collect_unread(state, mailbox)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    except (mail.DraftError, mail.CredentialError, OSError, imaplib.IMAP4.error) as error:
        # Do not log raw server responses or credentials.
        print(str(error) if isinstance(error, mail.DraftError) else "未读邮件读取未完成；已保存进度，未改变邮件状态。", file=sys.stderr)
        return 1
    finally:
        if mailbox:
            mailbox.close()
        state.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
