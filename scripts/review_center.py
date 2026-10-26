"""Private mentor review; explicit reviewed updates write summaries, never send mail."""
from __future__ import annotations

import argparse
from copy import deepcopy
from datetime import datetime
from email.utils import getaddresses
from io import BytesIO
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parent))
import daily_workflow as daily
import mail_send as sender
import mail_records as ledger

mail = daily.mail


def initialize(state):
    daily.initialize(state)
    sender.initialize(state)
    state.db.execute("CREATE TABLE IF NOT EXISTS review_packets (creator_id TEXT PRIMARY KEY, data TEXT NOT NULL)")
    state.db.commit()


def packets(state):
    return {r[0]: json.loads(r[1]) for r in state.db.execute("SELECT creator_id,data FROM review_packets")}


def source_rows(book):
    result = daily.workbook.read(book)
    main, longterm = result["sheets"]
    return {r["达人ID"]: r for r in result["creators"]}, {r["达人ID"]: r for r in longterm["creators"]}


def basis(state, creator_id):
    context = ledger.context(state, creator_id)
    incoming = [m for m in context["messages"] if m["direction"] == "incoming"]
    latest = incoming[-1] if incoming else None
    return {"message_id": latest["message_id"] if latest else "", "message_hash": mail.digest(
        [latest.get("raw_text", ""), latest.get("received_at", ""), latest.get("source", "")]) if latest else ""}


def row_basis(row, appendix=None):
    if "往来邮件梗概" in row:
        return mail.digest({k: row.get(k, "") for k in ("Email", "当前建联状况", "沟通摘要", "往来邮件梗概",
                           "待发送回复梗概", "WhatsApp联系方式", "可沟通时间", "收件信息")})
    return mail.digest({"main": {k: row.get(k, "") for k in ("Email", "联系状态", "沟通摘要", "建议回复策略", "预期发送邮件（待确认）", "预期回复摘要", "处理状态")},
                        "appendix": {k: (appendix or {}).get(k, "") for k in ("最新来信日期", "最新来信摘要", "下一步与待决事项", "往来轮次与定位")}})


def atomic_bytes(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    fd, temporary = tempfile.mkstemp(prefix=".tailki-review-", dir=path.parent)
    try:
        with os.fdopen(fd, "wb") as out:
            out.write(data); out.flush(); os.fsync(out.fileno())
        os.chmod(temporary, 0o600)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary): os.unlink(temporary)


def asset(state, data, suffix):
    if suffix.lower() not in (".png", ".jpg", ".jpeg", ".gif"):
        raise mail.DraftError("审核配图格式暂不支持，保留原文档。")
    path = state.path.parent / "review/assets" / (mail.digest(data) + suffix.lower())
    if not path.exists(): atomic_bytes(path, data)
    return {"path": str(path), "hash": mail.digest(data)}


def adopt_document(state, source, book=daily.BOOK, updates=None):
    """Import explicit local mentor sections and pictures, preserving manual edits."""
    from docx import Document
    from docx.oxml.ns import qn
    source = Path(source).resolve()
    book_hash = mail.digest(Path(book).read_bytes())
    rows, appendix = source_rows(book)
    content = source.read_bytes()
    document = Document(BytesIO(content))
    # Read the final visible text: insertions stay, deletions/move-from do not.
    # The source DOCX itself is never rewritten or stripped of review markup.
    for tag in ("w:del", "w:moveFrom"):
        for node in list(document.element.iter(qn(tag))): node.getparent().remove(node)
    comments = {}
    for rel in document.part.rels.values():
        if rel.reltype.endswith("/comments"):
            from lxml import etree
            root = etree.fromstring(rel.target_part.blob)
            comments = {n.get(qn("w:id")): "".join(n.itertext()) for n in root}
    for tag in ("w:ins", "w:moveTo"):
        for node in list(document.element.iter(qn(tag))):
            parent = node.getparent(); index = parent.index(node)
            for child in list(node): parent.insert(index, child); index += 1
            parent.remove(node)
    groups, group = [], None
    from docx.text.paragraph import Paragraph
    for node in document.element.body.iter(qn("w:p")):
        paragraph = Paragraph(node, document._body)
        if paragraph.style.name == "Heading 1":
            group = {"title": paragraph.text, "parts": []}
            groups.append(group)
        elif group is not None:
            images = []
            for blip in paragraph._p.iter(qn("a:blip")):
                rid = blip.get(qn("r:embed"))
                if not rid or rid not in document.part.related_parts:
                    raise mail.DraftError("图文报告含无法解析的图片，未覆盖审核资料。")
                part = document.part.related_parts[rid]
                images.append((part.blob, Path(str(part.partname)).suffix))
            anchored = {n.get(qn("w:id")) for n in paragraph._p.iter() if n.tag in
                        (qn("w:commentRangeStart"), qn("w:commentReference"))}
            group["parts"].append({"style": paragraph.style.name, "text": paragraph.text, "images": images,
                                   "comments": [{"id": i, "text": comments[i], "anchor": paragraph.text}
                                                for i in sorted(anchored) if i in comments]})
    anchored_ids = {comment["id"] for g in groups for part in g["parts"] for comment in part["comments"]}
    unlocated = [{"id":i,"text":value,"anchor":"文档批注，需结合原文定位"} for i,value in comments.items() if i not in anchored_ids]
    prepared, previous = {}, packets(state)
    for group in groups:
        text = "\n".join(p["text"] for p in group["parts"])
        outgoing_index = next((i for i,p in enumerate(group["parts"]) if p["style"] == "Heading 2" and p["text"] == "我方完整回复正文"), -1)
        recipient_index = next((i for i,p in enumerate(group["parts"]) if i > outgoing_index and p["text"].startswith("收件人：")), -1)
        explicit = group["parts"][recipient_index]["text"].split("：",1)[1].strip() if recipient_index>=0 else ""
        addresses = {a.lower() for _, a in getaddresses([explicit])} if explicit else set(re.findall(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text))
        if explicit and len(addresses) != 1:
            raise mail.DraftError("每封审核回复稿需要唯一主收件人。")
        stable_ids = [p["text"].split("：",1)[1].strip() for p in group["parts"] if p["text"].startswith("达人ID：")]
        matches = [r for r in rows.values() if (r["达人ID"] in stable_ids if stable_ids else str(r.get("Email", "")).lower() in addresses)]
        if stable_ids and (len(stable_ids)!=1 or len(matches)!=1 or (explicit and str(matches[0].get("Email","")).lower() not in addresses)):
            raise mail.DraftError("文稿达人 ID 与收件人不一致，请核实。")
        if not matches:
            # Introductory daily overview sections contain no creator packet.
            if explicit: raise mail.DraftError("报告收件人未能关联主表，原文件保留。")
            continue
        if len(matches) != 1:
            raise mail.DraftError("报告分节涉及多个达人或共享邮箱，需先核实归属。")
        row = matches[0]
        identifier = row["达人ID"]
        if identifier in prepared: raise mail.DraftError("报告同一达人出现多个一级分节，未自动合并。")
        packet = {"creator_id": identifier, "title": group["title"], "source": str(source),
                  "source_hash": mail.digest(content), "basis": basis(state, identifier),
                  "row_basis": row_basis(row, appendix.get(identifier)), "recorded_at": mail.now(), "details": [], "draft": None}
        packet["comments"] = list({c["id"]: c for p in group["parts"] for c in p["comments"]}.values()) + unlocated
        packet["unresolved_comments"] = bool(packet["comments"])
        section, sections = None, {}
        for part in group["parts"]:
            if part["style"] == "Heading 2" and part["text"] in ("对方来信原文", "来信要点", "我方回复要点及应对策略", "我方完整回复正文"):
                section = part["text"]; sections[section] = []
            elif section and section != "我方完整回复正文" and part["text"]:
                sections[section].append(part["text"])
        packet["sections"] = {k: "\n".join(v) for k,v in sections.items()}
        # Incoming originals can also contain Subject/主题; take the outgoing
        # subject only after its explicit recipient marker.
        subject_index = next((i for i,p in enumerate(group["parts"]) if i > recipient_index and p["text"].startswith("主题：")), -1)
        subject = group["parts"][subject_index]["text"].split("：",1)[1].strip() if subject_index >= 0 else ""
        in_body = False
        blocks = []
        for part_index, part in enumerate(group["parts"]):
            value = part["text"]
            if part_index == subject_index and explicit:
                in_body = True
                continue
            if in_body:
                if value.startswith("审核状态："): break
                if value: blocks.append({"text": value})
                if part["images"]:
                    blocks.append({"images": [asset(state, data, suffix) for data, suffix in part["images"]]})
            elif value and not value.startswith(("收件人：", "抄送：", "密送：", "审核状态：")):
                packet["details"].append({"text": value, "heading": part["style"].startswith("Heading")})
        if explicit:
            if not subject or not any(b.get("text") for b in blocks):
                raise mail.DraftError("报告回复稿缺主题或完整正文，未导入。")
            if any(token in subject for token in ("{{", "}}")) or any("{{" in b.get("text", "") or "}}" in b.get("text", "") for b in blocks):
                raise mail.DraftError("回复稿仍有占位符，未导入。")
            packet["draft"] = {"recipient": explicit, "subject": subject, "blocks": blocks,
                               "status": "awaiting_confirmation", "hash": mail.digest([explicit, subject, blocks])}
            for label, key in (("抄送：", "cc"), ("密送：", "bcc")):
                value = next((p["text"][len(label):].strip() for p in group["parts"] if p["text"].startswith(label)), "")
                if value:
                    for _, address in getaddresses([value]): mail.email_address(address)
                    packet["draft"][key] = value
            packet["draft"]["hash"] = mail.digest({k: v for k, v in packet["draft"].items() if k != "hash"})
        old = previous.get(identifier)
        if old and old.get("draft") == packet.get("draft") and old.get("pending_summary") and old.get("comments") == packet.get("comments"):
            packet["pending_summary"] = old["pending_summary"]
            packet["unresolved_comments"] = old.get("unresolved_comments", False)
        if old and old.get("basis") != packet["basis"]:
            # Formatting/manual document edits cannot silently attest that an
            # older reply draft covers a newly arrived message.
            packet["basis"] = old.get("basis", {})
        packet["history"] = (old or {}).get("history", [])
        if old and any(old.get(k) != packet.get(k) for k in ("details", "draft")):
            packet["history"] = packet["history"] + [{k: v for k, v in old.items() if k != "history"}]
        prepared[identifier] = packet
    if not prepared:
        raise mail.DraftError("报告没有可关联的达人分节，未导入或覆盖。")
    if updates is not None:
        apply_reviewed_updates(state, source, content, book, book_hash, prepared, rows, updates)
    # Validate all sections before replacing any packet. Removed drafts in an
    # explicitly refreshed source become inactive rather than reappearing.
    with state.db:
        for identifier, old in previous.items():
            if old.get("source") == str(source) and identifier not in prepared:
                old.update(inactive=True)
                state.db.execute("UPDATE review_packets SET data=? WHERE creator_id=?", (json.dumps(old, ensure_ascii=False), identifier))
        for identifier, packet in prepared.items():
            state.db.execute("INSERT OR REPLACE INTO review_packets VALUES (?,?)", (identifier, json.dumps(packet, ensure_ascii=False)))
    return {"imported": list(prepared), "source": str(source), "source_hash": mail.digest(content), "sent": False,
            "book_changed": updates is not None, "comments": {k: p["comments"] for k,p in prepared.items() if p["comments"]}}


def apply_reviewed_updates(state, source, content, book, book_hash, prepared, rows, updates):
    """The assistant supplies freshly reviewed summaries, bound to this DOCX.

    Arbitrary prose/comments cannot be safely converted to semantic summaries
    by a string parser. No old summary is silently reused after a body change.
    """
    updates = deepcopy(updates)
    if updates.get("source_hash") != mail.digest(content) or updates.get("book_hash") != book_hash:
        raise mail.DraftError("文稿或表格依据已变化，请重新核对修改稿与梗概。")
    allowed = {"当前建联状况", "沟通摘要", "待发送回复梗概", "WhatsApp联系方式", "可沟通时间", "收件信息"}
    specs = updates.get("creators", {})
    if not specs or not set(specs) <= set(prepared): raise mail.DraftError("更新内容未关联本次文稿的达人。")
    for identifier, spec in specs.items():
        packet = prepared[identifier]; fields = spec.get("fields", {})
        if not packet["draft"] or not fields or set(fields) - allowed or any(not isinstance(v,str) for v in fields.values()):
            raise mail.DraftError("只接受已核对回复稿及简表语义字段。")
        current_basis = basis(state, identifier)
        if packet["basis"] != current_basis:
            if spec.get("message_basis") != current_basis: raise mail.DraftError("文稿来信依据已过期。")
            packet["basis"] = current_basis
        if packet["comments"] and spec.get("comments_reviewed") is not True:
            raise mail.DraftError("文稿含批注，需先逐条处理，再更新梗概。")
        summary = fields.get("待发送回复梗概", "")
        if not summary.startswith("未发送｜") or not summary[4:].strip():
            raise mail.DraftError("需要本次修改后的待发梗概，并标明未发送。")
        for header in fields: daily.workbook.address(rows[identifier]["_headers"], header, rows[identifier]["_row"])
        packet.update(pending_summary=summary[4:], unresolved_comments=False)
        if spec.get("incoming_summary"):
            if not isinstance(spec["incoming_summary"],str) or spec.get("message_basis") != current_basis or not current_basis["message_id"]:
                raise mail.DraftError("来信梗概须对应已核对的单封最新来信。")
            marker="已整理来信：" + current_basis["message_id"]
            note=rows[identifier].get("备注") or ""
            if marker not in note:
                messages=ledger.context(state, identifier)["messages"]
                incoming=next(m for m in messages if m["direction"]=="incoming" and m["message_id"]==current_basis["message_id"])
                from mail_postprocess import history_line
                day=daily.timestamp(incoming["received_at"]).date().isoformat()
                fields["往来邮件梗概"]=history_line(rows[identifier].get("往来邮件梗概"),day,"对方",spec["incoming_summary"])
                fields["备注"]=note.rstrip() + ("\n" if note else "") + marker
    # Stage every tab change in a disposable file; only publish after all pass.
    # The temporary copy is deleted on exit and is never an original-text backup.
    with tempfile.TemporaryDirectory(prefix="tailki-reviewed-") as temp:
        staging = Path(temp)/"updated.xlsx"; staging.write_bytes(Path(book).read_bytes())
        for identifier, spec in specs.items():
            daily.workbook.update_fields(staging, identifier, spec["fields"], expected_hash=mail.digest(staging.read_bytes()))
            updated = {**rows[identifier], **spec["fields"]}
            prepared[identifier]["row_basis"] = row_basis(updated)
        if mail.digest(Path(book).read_bytes()) != book_hash or source.read_bytes() != content:
            raise mail.DraftError("写回前文稿或表格发生变化，保留人工修改。")
        atomic_bytes(book, staging.read_bytes())


def model(state, day, book=daily.BOOK):
    """Read-only snapshot. Sent/attempted mail is never offered as pending again."""
    datetime.strptime(day, "%Y-%m-%d")
    rows, appendix = source_rows(book)
    notes = packets(state)
    replies = {x["message"]["creator_id"]: x for x in sorted(daily.all_items(state, "reply:"),
               key=lambda x: daily.timestamp(x["message"]["received_at"]))}
    day_state = daily.get(state, "day:" + day, {})
    added = [rows[x] for x in day_state.get("creators", []) if x in rows]
    cards, warnings = [], []
    route_check = daily.get(state, "pool-routes:" + day, {})
    limitations = route_check.get("limitations", [])
    if isinstance(limitations, list):
        warnings.extend(str(item) for item in limitations if item)
    if limitations and day_state:
        warnings.append("今日新增 " + str(len(added)) + " 位，距每日最多 " + str(daily.DAILY_TARGET)
                        + " 位仍缺 " + str(max(0, daily.DAILY_TARGET-len(added))) + " 位；搜集尚未完成。")
    active_rows = {k for k,r in rows.items() if r.get("待发送回复梗概") or r.get("当前建联状况") in ("收到新回信", "报价沟通中", "合作条款确认中")}
    for identifier in sorted(set(replies) | set(notes) | set(appendix) | active_rows):
        if identifier not in rows: continue
        row = rows[identifier]
        if row.get("联系状态") == "已结束": continue
        packet = notes.get(identifier, {})
        if packet.get("inactive"): packet = {}
        summary_entry = replies.get(identifier, {})
        message = summary_entry.get("message", {})
        current_basis = basis(state, identifier)
        context = ledger.context(state, identifier)
        incoming = [m for m in context["messages"] if m["direction"] == "incoming"]
        latest = incoming[-1] if incoming else {}
        answered = any(m["direction"] == "outgoing" and (m.get("in_reply_to") == latest.get("message_id")
                       or (latest.get("message_id") and latest["message_id"] in m.get("references", "")))
                       for m in context["messages"])
        if answered: continue
        message_changed = bool(packet) and packet.get("basis") != current_basis
        stale = bool(packet) and (message_changed or packet.get("row_basis") != row_basis(row, appendix.get(identifier)))
        long = appendix.get(identifier, {})
        decisions = [long["下一步与待决事项"]] if long.get("下一步与待决事项") else message.get("decisions", [])
        if not isinstance(decisions, list): decisions = [str(decisions)]
        current_summary = row.get("沟通摘要") or message.get("summary", "")
        current_strategy = row.get("建议回复策略") or row.get("对方回复风格判断和应对策略") or message.get("reason", "")
        fresh = bool(message.get("message_id")) and message.get("message_id") == current_basis["message_id"]
        if message_changed or fresh:
            current_summary = message.get("summary", "") if fresh else "新来信或记录已变化，待补齐本轮摘要。"
            current_strategy = message.get("reason", "") if fresh else "待核实最新内容后判断。"
            decisions = message.get("decisions", []) if fresh else ["请先核实本轮来信及合作条件变化。"]
        if not isinstance(decisions, list): decisions = [str(decisions)]
        received_at = latest.get("received_at", message.get("received_at", ""))
        if received_at:
            received_at = daily.timestamp(received_at).isoformat()
        card = {"creator_id": identifier, "name": row.get("达人昵称", identifier), "email": row.get("Email", ""),
                "status": row.get("联系状态", ""), "summary": current_summary,
                "strategy": current_strategy,
                "decisions": decisions, "received_at": received_at,
                "incoming": context["incoming"], "outgoing": context["outgoing"], "coverage": context["coverage"],
                "raw_text": latest.get("raw_text", ""), "message_id": current_basis["message_id"],
                "source": latest.get("source", message.get("source", "")), "stale": stale,
                "title": packet.get("title", row.get("达人昵称", identifier)),
                "details": [] if stale else packet.get("details", []),
                "draft": None if stale else deepcopy(packet.get("draft")), "kind": "mentor" if identifier in appendix else "reply"}
        if stale:
            warnings.append(card["name"] + "：来信或表格已变化，历史汇报和回复稿已暂停展示，请按最新信息重新整理。")
        sections = {} if stale else packet.get("sections", {})
        card["incoming_points"] = sections.get("来信要点") or current_summary
        card["reply_points"] = sections.get("我方回复要点及应对策略") or current_strategy
        if not card["raw_text"]: card["raw_text"] = sections.get("对方来信原文", "")
        if sections: card["details"] = []
        if packet.get("unresolved_comments"):
            card["draft"] = None
            warnings.append(card["name"] + "：文稿含带教批注，请处理后更新回复和梗概。")
        if card["draft"]:
            for block in card["draft"]["blocks"]:
                for image in block.get("images", []):
                    path = Path(image["path"])
                    if not path.is_file() or mail.digest(path.read_bytes()) != image["hash"]:
                        card["draft"] = None
                        warnings.append(card["name"] + "：配图缺失或已修改，暂停展示完整回复稿。")
                        break
                if card["draft"] is None: break
        cards.append(card)
    cards.sort(key=lambda c: (c["kind"] != "mentor", c["creator_id"]))
    attempts = {r[0]: json.loads(r[1]) for r in state.db.execute("SELECT draft_id,data FROM send_attempts")}
    mails, represented = [], set()
    reviews = [json.loads(r[0]) for r in state.db.execute("SELECT data FROM send_reviews")]
    for draft_id, attempt in attempts.items():
        if attempt["status"] != "sent_verified":
            warnings.append("一封邮件的发送状态待处理（" + attempt["status"] + "），先核对实际记录，不自动重发。")
    for review in reviews:
        for item in review.get("postprocess", {}).get("items", []):
            if item.get("status") == "pending":
                warnings.append(item.get("creator_id", "") + "：发送收尾待处理。" + item.get("reason", ""))
        if review.get("status") == "halted":
            remaining = sum(e["draft_id"] not in attempts for e in review["entries"])
            if remaining: warnings.append("一个已暂停批次还有 " + str(remaining) + " 封未提交，需核对原确认范围后继续。")
    for review in sorted(reviews, key=lambda r: r.get("reviewed_at", ""), reverse=True):
        if review.get("status") not in ("awaiting_confirmation", "approved"): continue
        if any(rows.get(e["creator_id"], {}).get("联系状态") == "已结束" for e in review["entries"]): continue
        if any(e["draft_id"] in attempts for e in review["entries"]): continue
        if any(e["draft_id"] in represented for e in review["entries"]): continue
        valid = sender.review_hash(review) == review.get("hash")
        path = Path(review.get("preview", ""))
        valid = valid and path.is_file() and mail.digest(path.read_bytes()) == review.get("preview_hash")
        if not valid:
            warnings.append("一批历史完整预览已变化或缺失，需重新读取邮箱草稿后再确认。")
            continue
        for entry in review["entries"]:
            represented.add(entry["draft_id"])
            mails.append({"creator_id": entry["creator_id"], "name": entry["name"], "recipient": entry["details"]["envelope"]["to"][0],
                          "subject": entry["details"]["subject"], "body": "", "raw_b64": entry["raw_b64"],
                          "review_id": review["id"], "hash": review["hash"], "kind": "smtp",
                          "label": "已核对的完整稿；发送前仍须刷新邮箱版本", "read_at": review["reviewed_at"]})
    for item in daily.drafts_snapshot(state):
        if rows.get(item["creator_id"], {}).get("联系状态") == "已结束": continue
        if item["id"] in attempts or item.get("smtp_attempt") or item["id"] in represented: continue
        if item["status"] not in ("prepared", "saved", "edited"): continue
        # Edited/remote snapshots lack full MIME and may omit attachments. Show
        # them only as text preparation, never as an approved final-version source.
        current = item.get("remote") or {}
        mails.append({"creator_id": item["creator_id"], "name": item["name"],
                      "recipient": ", ".join(current.get("to", [])) or item["recipient"],
                      "subject": current.get("subject", item["subject"]), "body": current.get("body", item["body"]),
                      "cc": current.get("cc", []), "bcc": current.get("bcc", []),
                      "attachments_incomplete": current.get("attachments", 0),
                      "review_id": item["batch_id"], "hash": item["content_hash"], "kind": "prepared",
                      "label": "待核对邮箱最终稿" if current else "本机首联初稿，待确认", "read_at": item.get("updated_at", "")})
    generated_at = mail.now()
    import candidate_pool
    mentor_brief = daily.get(state, "mentor-brief:" + day, {})
    result = {"date": day, "generated_at": generated_at, "book": str(Path(book).resolve()),
              "book_hash": mail.digest(Path(book).read_bytes()), "added": added, "cards": cards, "mails": mails,
              "warnings": warnings, "schedule": daily.get(state, "activation", {}),
              "candidate_pool": candidate_pool.summary(state, day), "daily_limit": daily.DAILY_TARGET,
              "mentor_brief": mentor_brief,
              "verified_total": sum(a["status"] == "sent_verified" for a in attempts.values()),
              "sent_today": sum(a["status"] == "sent_verified" and a.get("receipt", {}).get("sent_at", "")[:10] == day for a in attempts.values())}
    result["hash"] = mail.digest({k: v for k, v in result.items() if k != "generated_at"})
    result["document_model_hash"] = mail.digest({k: result[k] for k in ("date", "added", "cards", "warnings", "mentor_brief")})
    return result


def build(state, day=None, book=daily.BOOK, output=None):
    initialize(state)
    day = day or daily.current_time().date().isoformat()
    output = Path(output or state.path.parent / "review/今日审核.html").resolve()
    if output.suffix != ".html" or (output.is_relative_to(daily.ROOT) and not output.is_relative_to(daily.BUSINESS)):
        raise mail.DraftError("审核页面须留在本机私有目录。")
    document = Path(book).resolve().parent / "参考资料" / (day + "_KOL跟进.docx")
    manifest_key = "review-center:" + day
    previous = daily.get(state, manifest_key, {})
    source_packets = packets(state)
    sources = {p["source"]: p.get("source_hash") for p in source_packets.values() if p.get("source")}
    # Human edits to a generated document are imported before rebuilding it.
    if document.exists() and previous and mail.digest(document.read_bytes()) != previous.get("document_hash"):
        adopt_document(state, document, book)
    elif document.exists() and not previous:
        raise mail.DraftError("同名带教文档已存在且未登记，先导入审核资料；未覆盖原文件。")
    for source, expected in sources.items():
        source = Path(source)
        if source == document: continue
        if source.exists() and mail.digest(source.read_bytes()) != expected:
            adopt_document(state, source, book)
    snapshot = model(state, day, book)
    preserve_document = document.exists() and any(p.get("unresolved_comments") and p.get("source") == str(document)
        for p in packets(state).values() if not p.get("inactive"))
    original_document_hash = mail.digest(document.read_bytes()) if document.exists() else None
    from review_docx import write_document
    from review_page import render_page
    old_document_model = previous.get("document_model_hash")
    if old_document_model is None and previous.get("model_hash") == mail.digest({k: v for k, v in snapshot.items()
            if k not in {"generated_at", "hash", "document_model_hash", "candidate_pool", "daily_limit"}}):
        old_document_model = snapshot["document_model_hash"]
    if not preserve_document and (not previous or old_document_model != snapshot["document_model_hash"] or not document.exists()):
        with tempfile.TemporaryDirectory(prefix="tailki-report-") as temp:
            temporary = Path(temp) / "report.docx"
            write_document(snapshot, temporary)
            if mail.digest(Path(book).read_bytes()) != snapshot["book_hash"]:
                raise mail.DraftError("生成期间表格发生变化，未发布旧审核内容。")
            current_document_hash = mail.digest(document.read_bytes()) if document.exists() else None
            if current_document_hash != original_document_hash:
                raise mail.DraftError("生成期间带教文档有人工修改，已保留原文件，请重新整理。")
            atomic_bytes(document, temporary.read_bytes())
    doc_hash = mail.digest(document.read_bytes())
    page = render_page(snapshot, document)
    atomic_bytes(output, page.encode("utf-8"))
    manifest = {"date": day, "page": str(output), "document": str(document), "document_hash": doc_hash,
                "model_hash": snapshot["hash"], "generated_at": snapshot["generated_at"],
                "document_model_hash": old_document_model if preserve_document else snapshot["document_model_hash"],
                "document_has_unresolved_comments": preserve_document,
                "counts": {"added": len(snapshot["added"]), "mentor": sum(c["kind"] == "mentor" for c in snapshot["cards"]),
                           "drafts": len(snapshot["mails"]) + sum(bool(c["draft"]) for c in snapshot["cards"])},
                "warnings": snapshot["warnings"], "sent": False, "book_changed": False}
    # A changed document always needs fresh visual QA; unchanged exports retain it.
    if previous.get("document_hash") == doc_hash and previous.get("qa"):
        manifest["qa"] = previous["qa"]
    daily.put(state, manifest_key, manifest)
    daily.put(state, "review-center:latest", manifest)
    return manifest


def refresh_after_operation(state, result, book=daily.BOOK):
    """Artifact failure cannot disguise the outcome of an already-completed operation."""
    try:
        return build(state, book=book)
    except Exception:
        return {"status": "needs_refresh", "reason": "操作记录已保留，审核页面生成未完成；重新打开今日审核恢复。", "sent": False}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--state", type=Path, default=mail.DEFAULT_STATE)
    parser.add_argument("--book", type=Path, default=daily.BOOK)
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("build"); p.add_argument("--day"); p.add_argument("--open", action="store_true")
    p = sub.add_parser("adopt-report"); p.add_argument("source", type=Path); p.add_argument("--updates", type=Path)
    args = parser.parse_args()
    state = mail.State(args.state)
    try:
        initialize(state)
        if args.command == "adopt-report": result = adopt_document(state, args.source, args.book,
            json.loads(args.updates.read_text()) if args.updates else None)
        else: result = build(state, args.day, args.book)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if getattr(args, "open", False): subprocess.run(["open", result["page"]], check=True)
    finally: state.close()


if __name__ == "__main__":
    try: main()
    except (mail.DraftError, OSError, ValueError, KeyError) as error:
        print("审核入口未更新：" + str(error), file=sys.stderr)
        sys.exit(1)
