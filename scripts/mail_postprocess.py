"""Recoverable post-send work. This module never connects to SMTP."""
from __future__ import annotations

import base64
from datetime import datetime
from email.utils import parsedate_to_datetime
import imaplib
import json
import re
from pathlib import Path
from zoneinfo import ZoneInfo

import daily_workbook as workbook
import mail_records as records


def normalized_body(value):
    return re.sub(r"\s+", " ", value).strip()


def sent_synopsis(state, event):
    """Use a semantic summary only when it is bound to the actual sent body."""
    mail = records.mail
    if event.get("summary") and event.get("summary_body_hash") == mail.digest(event["raw_text"]):
        return event["summary"], False
    table = state.db.execute("SELECT name FROM sqlite_master WHERE name='review_packets'").fetchone()
    saved = state.db.execute("SELECT data FROM review_packets WHERE creator_id=?", (event["creator_id"],)).fetchone() if table else None
    if saved:
        packet = json.loads(saved[0]); draft = packet.get("draft") or {}
        body = "\n\n".join(b.get("text", "") for b in draft.get("blocks", []) if b.get("text"))
        if (packet.get("pending_summary") and not packet.get("unresolved_comments")
                and normalized_body(body) == normalized_body(event["raw_text"])
                and draft.get("subject") == event.get("subject")
                and draft.get("recipient", "").lower() == event.get("to", "").lower()):
            return packet["pending_summary"], True
    prefix = "首联已发送" if not event.get("in_reply_to") else "已回复"
    return prefix + "；主题：" + event.get("subject", "") + "；内容梗概待整理。", False


def history_line(existing, day, direction, summary):
    line = f"{day}｜{direction}：{summary.strip()}"
    lines = (existing or "").splitlines()
    # Message-ID idempotence is handled by mail_actions; identical semantic
    # messages may be separate mail, so do not deduplicate solely by wording.
    lines.append(line)
    return "\n".join(sorted(lines, key=lambda x: re.match(r"\d{4}-\d{2}-\d{2}", x)[0]
                          if re.match(r"\d{4}-\d{2}-\d{2}", x) else "9999"))


def verified_message(state, sender, review, entry, mailbox):
    mail = sender.mail
    approval = review.get("approval", {})
    if (sender.review_hash(review) != review.get("hash") or approval.get("review_hash") != review["hash"]
            or approval.get("count") != len(review["entries"]) or not approval.get("user_instruction")):
        raise mail.DraftError("收尾缺少该完整版本的有效发送确认。")
    result = sender.attempt(state, entry["draft_id"])
    if not result or result["status"] != "sent_verified" or result["review_id"] != review["id"]:
        raise mail.DraftError("尚未核实实际发送，不能写表或清理草稿。")
    if result.get("receipt", {}).get("mime_hash") != entry["sent_hash"]:
        raise mail.DraftError("回执内容与确认版不同。")
    found = sender.find_sent(mailbox, result)
    if len(found) != 1 or mail.digest(sender.content(found[0]["raw"], hide_bcc=True)) != entry["sent_hash"]:
        raise mail.DraftError("无法重新核实唯一已发送副本，保留表格和草稿。")
    actual = found[0]
    message = sender.parse(actual["raw"])
    stamp = parsedate_to_datetime(str(message.get("Date", "")))
    if (stamp.tzinfo is None or stamp < datetime.fromisoformat(approval["user_instruction_at"])
            or stamp > datetime.fromisoformat(mail.now()) + sender.timedelta(minutes=5)):
        raise mail.DraftError("已发送时间与本批确认不符。")
    recipient = sender.addresses(message, "To")
    if recipient != entry["details"]["envelope"]["to"]:
        raise mail.DraftError("已发送收件人与确认版不符。")
    source = actual["folder"] + ":UIDVALIDITY=" + actual["uidvalidity"] + ":UID=" + actual["uid"]
    event = {"account": review["account"], "direction": "outgoing", "creator_id": entry["creator_id"],
             "message_id": result["message_id"], "source": source, "sent_at": stamp.isoformat(),
             "to": recipient[0], "subject": str(message.get("Subject", "")),
             "raw_text": mail.snapshot(actual["raw"])["body"], "mime_hash": entry["sent_hash"],
             "body_status": "collected",
             "in_reply_to": str(message.get("In-Reply-To", "")), "references": str(message.get("References", "")),
             "review_id": review["id"], "draft_id": entry["draft_id"], "verification": "sent_folder"}
    records.put(state, event)
    import candidate_pool
    candidate_pool.mark_sent(state, entry["creator_id"], event)
    return event


def write_workbook(state, sender, review, entry, event, book):
    """Header-addressed, write-ahead values only. Never add columns or rewrite a sheet."""
    mail = sender.mail
    book = Path(book).resolve()
    key = "workbook:" + mail.digest([review["account"], event["message_id"], str(book)])
    progress = records.action(state, key)
    if progress.get("status") == "done":
        return progress
    original_hash = mail.digest(book.read_bytes())
    sheet = workbook.read(book)
    rows = [r for r in sheet["creators"] if r["达人ID"] == entry["creator_id"]]
    if len(rows) != 1 or str(rows[0].get("Email", "")).strip().lower() != event["to"]:
        raise mail.DraftError("主表达人或邮箱已变更，停止写回。")
    row = rows[0]
    sheet = next((s for s in sheet.get("sheets", []) if s["name"] == row.get("_sheet")), sheet)
    def empty(value): return value in (None, "")
    if not progress.get("fields"):
        changes = {}
        def change(header, value, *, date=False):
            if header not in sheet["headers"] or sheet["headers"].count(header) != 1:
                raise mail.DraftError("主表表头缺失或重复：" + header)
            changes[header] = {"before": row.get(header, ""), "after": value, "date": date}
        day = datetime.fromisoformat(event["sent_at"]).astimezone(ZoneInfo("Asia/Shanghai")).date().isoformat()
        simple = "往来邮件梗概" in sheet["headers"]
        state_header = "当前建联状况" if simple else "联系状态"
        if not event["in_reply_to"]:
            prior = row.get("初次建联日期", "")
            if not empty(prior) and str(prior)[:10] != day:
                raise mail.DraftError("主表首联日期与实际发送日期冲突，保留人工记录。")
            if empty(prior): change("初次建联日期", day, date=True)
            if row.get(state_header) == "待首联确认":
                change(state_header, "初次建联未回复")
        elif not simple and all(empty(row.get(h)) for h in ("我方第一次回复邮件日期", "我们的第一次回复邮件原始文本", "我们的第一次回复邮件摘要")):
            change("我方第一次回复邮件日期", day, date=True)
            change("我们的第一次回复邮件原始文本", event["raw_text"])
            # No invented semantic summary; retain later rounds in the message ledger.
        if simple:
            summary, matched = sent_synopsis(state, event)
            change("往来邮件梗概", history_line(row.get("往来邮件梗概"), day, "我方", summary))
            if matched and row.get("待发送回复梗概") == "未发送｜" + summary:
                change("待发送回复梗概", "")
                change("沟通摘要", day + " 我方已回复：" + summary)
                if row.get(state_header) != "已结束": change(state_header, "我方已回复")
        note = str(row.get("备注", "") or "")
        marker = "每日批次：" + entry["batch_id"] + "\n首联未发送；"
        if not event["in_reply_to"] and note.startswith(marker):
            note = "每日批次：" + entry["batch_id"] + "\n首联已发送（" + day + "）；" + note[len(marker):]
        receipt_note = "已核验发送：" + event["message_id"]
        if receipt_note not in note:
            note = note.rstrip() + ("\n" if note else "") + receipt_note + "；" + event["sent_at"]
        if note != str(row.get("备注", "") or ""):
            change("备注", note)
        progress = records.action(state, key, status="prepared", fields=changes,
                                  creator_id=entry["creator_id"], message_id=event["message_id"], book=str(book))
    cells = {}
    for header, spec in progress["fields"].items():
        current = row.get(header, "")
        if current == spec["after"] or (spec["date"] and str(current)[:10] == spec["after"]):
            continue
        if current != spec["before"] and not (empty(current) and empty(spec["before"])):
            raise mail.DraftError("待写回字段有人工修改，保留原值：" + header)
        address = workbook.workflow.col_name(sheet["headers"].index(header) + 1) + str(row["_row"])
        cells[address] = {"value": spec["after"], "date": spec["date"]}
    if cells:
        workbook.patch_cells(book, cells, expected_hash=original_hash,
                             backup_dir=state.path.parent / "backups", style_row=row["_row"],
                             sheet_name=row.get("_sheet", workbook.MASTER))
    # Verify after replacement before recording completion or touching any draft.
    saved = next(r for r in workbook.read(book)["creators"] if r["达人ID"] == entry["creator_id"])
    for header, spec in progress["fields"].items():
        value = saved.get(header, "")
        if not (value == spec["after"] or (empty(value) and empty(spec["after"])) or (spec["date"] and str(value)[:10] == spec["after"])):
            raise mail.DraftError("写回读验不一致，停止草稿清理。")
    return records.action(state, key, status="done", workbook_hash=mail.digest(book.read_bytes()))


def trash_folder(mailbox, mail):
    code, entries = mailbox.client.list()
    if code != "OK": raise mail.DraftError("无法识别已删除邮件文件夹。")
    folders = mail.parse_folders(entries)
    found = [f for f in folders if b"\\trash" in f["flags"]]
    if not found:
        found = [f for f in folders if f["name"].lower() in ("trash", "deleted", "deleted items", "已删除邮件", "已删除", "垃圾箱")]
    if len(found) != 1 or found[0]["raw"] in (mailbox.drafts["raw"], mailbox.sent["raw"], b"INBOX"):
        raise mail.DraftError("无法唯一识别回收站，保留草稿。")
    return found[0]


def cleanup_draft(state, sender, review, entry, mailbox):
    """Move only the approved UID to Trash; persist before any mutation, never recopy an unknown result."""
    mail = sender.mail
    key = "cleanup:" + mail.digest([review["account"], entry["draft_id"]])
    progress = records.action(state, key)
    if progress.get("status") in ("done", "absent"):
        return progress
    validity, headers = mailbox.header_index(mailbox.drafts["raw"], refresh=True)
    if validity != entry["uidvalidity"]:
        raise mail.DraftError("草稿箱 UIDVALIDITY 变化，保留邮件等待核对。")
    uid = entry["uid"].encode("ascii")
    present = uid in headers
    if not present and not progress:
        return records.action(state, key, status="absent", reason="原草稿已不在原位置，未重建或处理其他草稿。")
    if present:
        raw = mailbox.fetch(uid)
        if mail.digest(sender.content(raw)) != entry["content_hash"]:
            raise mail.DraftError("原草稿已修改，保留修改稿。")
    original = sender.parse(base64.b64decode(entry["raw_b64"]))
    target = trash_folder(mailbox, mail)

    def target_copies():
        tv, hs = mailbox.header_index(target["raw"], refresh=True)
        matches = []
        for tid, header in hs.items():
            if ((original.get("Message-ID") and str(header.get("Message-ID", "")) == str(original["Message-ID"]))
                    or str(header.get("X-Tailki-Draft-ID", "")) == entry["draft_id"]):
                if mail.digest(sender.content(mailbox.fetch(tid))) == entry["content_hash"]:
                    matches.append({"uid": tid.decode("ascii"), "uidvalidity": tv, "folder": target["name"]})
        return matches

    copies = target_copies()
    if progress:
        # After a lost response, inspect both sides and never repeat COPY/MOVE.
        if not present and len(copies) == 1:
            return records.action(state, key, status="done", trash=copies[0])
        if progress.get("phase") not in ("copied", "delete_reserved") or len(copies) != 1 or not present:
            raise mail.DraftError("草稿移动结果待核实；未自动再次移动。")
    else:
        if copies: raise mail.DraftError("回收站已有关联副本，先核实，未重复移动。")
        capabilities = {x.decode().upper() if isinstance(x, bytes) else x.upper()
                        for x in getattr(mailbox.client, "capabilities", ())}
        operation = "MOVE" if "MOVE" in capabilities else "COPY"
        records.action(state, key, status="pending", phase="move_reserved", operation=operation,
                       uid=entry["uid"], uidvalidity=validity, content_hash=entry["content_hash"])
        select_writable(mailbox, entry, mail)
        if mail.digest(sender.content(mailbox.fetch(uid))) != entry["content_hash"]:
            raise mail.DraftError("移动前草稿已修改，保留原稿。")
        code, _ = mailbox.client.uid(operation, uid, mailbox.quoted(target["raw"]))
        if code != "OK": raise mail.DraftError("草稿移动未获确认，暂停查证。")
        # Some Alibaba servers remove the source on COPY-to-Trash.
        validity, hs = mailbox.header_index(mailbox.drafts["raw"], refresh=True)
        present = uid in hs
        copies = target_copies()
        if validity != entry["uidvalidity"] or len(copies) != 1:
            raise mail.DraftError("草稿移动读验不完整，未继续删除。")
        if not present:
            return records.action(state, key, status="done", trash=copies[0])
        if operation == "MOVE": raise mail.DraftError("MOVE 后源草稿仍存在，暂停核对。")
        progress = records.action(state, key, phase="copied", trash=copies[0])
    capabilities = {x.decode().upper() if isinstance(x, bytes) else x.upper()
                    for x in getattr(mailbox.client, "capabilities", ())}
    if "UIDPLUS" not in capabilities:
        raise mail.DraftError("回收站副本已核实，但服务器不支持按 UID 清理，保留源草稿待人工处理。")
    select_writable(mailbox, entry, mail)
    if mail.digest(sender.content(mailbox.fetch(uid))) != entry["content_hash"]:
        raise mail.DraftError("删除前草稿已修改，保留原稿。")
    records.action(state, key, phase="delete_reserved")
    code, _ = mailbox.client.uid("STORE", uid, "+FLAGS.SILENT", r"(\Deleted)")
    if code != "OK": raise mail.DraftError("源草稿标记结果待核实。")
    code, _ = mailbox.client.uid("EXPUNGE", uid)
    if code != "OK": raise mail.DraftError("源草稿按 UID 清理结果待核实。")
    validity, hs = mailbox.header_index(mailbox.drafts["raw"], refresh=True)
    copies = target_copies()
    if validity != entry["uidvalidity"] or uid in hs or len(copies) != 1:
        raise mail.DraftError("草稿清理最终核验未通过。")
    return records.action(state, key, status="done", trash=copies[0])


def select_writable(mailbox, entry, mail):
    code, _ = mailbox.client.select(mailbox.quoted(mailbox.drafts["raw"]), readonly=False)
    _, values = mailbox.client.response("UIDVALIDITY")
    if code != "OK" or not values or values[0].decode("ascii") != entry["uidvalidity"]:
        raise mail.DraftError("无法安全打开原草稿位置，停止清理。")


def finalize(state, sender, identifier, mailbox, book):
    review = sender.get_review(state, identifier)
    if mailbox.account != review["account"]:
        raise sender.mail.DraftError("收尾邮箱账号与原批次不同。")
    outcomes = []
    for entry in review["entries"]:
        result = sender.attempt(state, entry["draft_id"])
        if not result or result["status"] != "sent_verified":
            outcomes.append({"creator_id": entry["creator_id"], "status": "awaiting_verified_send"})
            continue
        try:
            event = verified_message(state, sender, review, entry, mailbox)
            written = write_workbook(state, sender, review, entry, event, book)
            cleaned = cleanup_draft(state, sender, review, entry, mailbox)
            outcomes.append({"creator_id": entry["creator_id"], "status": "done",
                             "workbook": written["status"], "cleanup": cleaned["status"]})
        except (sender.mail.DraftError, records.mail.DraftError, OSError, ValueError, KeyError,
                imaplib.IMAP4.error) as error:
            reason = str(error) if isinstance(error, (sender.mail.DraftError, records.mail.DraftError)) else "收尾未完成，保留进度；先核对，不重新发送。"
            outcomes.append({"creator_id": entry["creator_id"], "status": "pending", "reason": reason})
    return {"review": identifier, "sent_by_this_command": False, "items": outcomes,
            "complete": bool(outcomes) and all(x["status"] == "done" for x in outcomes)}
