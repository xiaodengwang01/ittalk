"""Build a private review page and mentor document. Never sends mail or edits Excel."""
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
    with tempfile.TemporaryDirectory(prefix="tailki-review-read-") as temp:
        result = daily.workbook.workflow.read_book(book, temp)
    if [s["name"] for s in result["sheets"]] != daily.workbook.SHEETS:
        raise mail.DraftError("审核入口只支持当前两张工作表，未改写文件。")
    main, longterm = result["sheets"]
    rows = [r for r in main["rows"] if re.fullmatch(r"(?:KOL|SIM)-\d+", str(r.get("达人ID", "")))]
    if len({r["达人ID"] for r in rows}) != len(rows):
        raise mail.DraftError("达人 ID 重复，先核实。")
    return {r["达人ID"]: r for r in rows}, {r["达人ID"]: r for r in longterm["rows"] if r.get("达人ID")}


def basis(state, creator_id):
    context = ledger.context(state, creator_id)
    incoming = [m for m in context["messages"] if m["direction"] == "incoming"]
    latest = incoming[-1] if incoming else None
    return {"message_id": latest["message_id"] if latest else "", "message_hash": mail.digest(
        [latest.get("raw_text", ""), latest.get("received_at", ""), latest.get("source", "")]) if latest else ""}


def row_basis(row, appendix=None):
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


def adopt_document(state, source, book=daily.BOOK):
    """Import explicit local mentor sections and pictures, preserving manual edits."""
    from docx import Document
    from docx.oxml.ns import qn
    source = Path(source).resolve()
    rows, appendix = source_rows(book)
    content = source.read_bytes()
    document = Document(BytesIO(content))
    groups, group = [], None
    for paragraph in document.paragraphs:
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
            group["parts"].append({"style": paragraph.style.name, "text": paragraph.text, "images": images})
    prepared, previous = {}, packets(state)
    for group in groups:
        text = "\n".join(p["text"] for p in group["parts"])
        explicit = next((p["text"].split("：", 1)[1].strip() for p in group["parts"] if p["text"].startswith("收件人：")), "")
        addresses = {a.lower() for _, a in getaddresses([explicit])} if explicit else set(re.findall(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text))
        if explicit and len(addresses) != 1:
            raise mail.DraftError("每封审核回复稿需要唯一主收件人。")
        matches = [r for r in rows.values() if str(r.get("Email", "")).lower() in addresses]
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
        subject = next((p["text"].split("：", 1)[1].strip() for p in group["parts"] if p["text"].startswith("主题：")), "")
        in_body = False
        blocks = []
        for part in group["parts"]:
            value = part["text"]
            if value.startswith("主题：") and explicit:
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
    # Validate all sections before replacing any packet. Removed drafts in an
    # explicitly refreshed source become inactive rather than reappearing.
    with state.db:
        for identifier, old in previous.items():
            if old.get("source") == str(source) and identifier not in prepared:
                old.update(inactive=True)
                state.db.execute("UPDATE review_packets SET data=? WHERE creator_id=?", (json.dumps(old, ensure_ascii=False), identifier))
        for identifier, packet in prepared.items():
            state.db.execute("INSERT OR REPLACE INTO review_packets VALUES (?,?)", (identifier, json.dumps(packet, ensure_ascii=False)))
    return {"imported": list(prepared), "source": str(source), "source_hash": mail.digest(content), "sent": False}


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
    for identifier in sorted(set(replies) | set(notes) | set(appendix)):
        if identifier not in rows: continue
        row = rows[identifier]
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
        if message_changed:
            fresh = message.get("message_id") == current_basis["message_id"]
            current_summary = message.get("summary", "") if fresh else "新来信或记录已变化，待补齐本轮摘要。"
            current_strategy = message.get("reason", "") if fresh else "待核实最新内容后判断。"
            decisions = message.get("decisions", []) if fresh else ["请先核实本轮来信及合作条件变化。"]
        if not isinstance(decisions, list): decisions = [str(decisions)]
        card = {"creator_id": identifier, "name": row.get("达人昵称", identifier), "email": row.get("Email", ""),
                "status": row.get("联系状态", ""), "summary": current_summary,
                "strategy": current_strategy,
                "decisions": decisions, "received_at": latest.get("received_at", message.get("received_at", "")),
                "incoming": context["incoming"], "outgoing": context["outgoing"], "coverage": context["coverage"],
                "raw_text": latest.get("raw_text", ""), "message_id": current_basis["message_id"],
                "source": latest.get("source", message.get("source", "")), "stale": stale,
                "title": packet.get("title", row.get("达人昵称", identifier)),
                "details": [] if stale else packet.get("details", []),
                "draft": None if stale else deepcopy(packet.get("draft")), "kind": "mentor" if identifier in appendix else "reply"}
        if stale:
            warnings.append(card["name"] + "：来信或表格已变化，历史汇报和回复稿已暂停展示，请按最新信息重新整理。")
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
    result = {"date": day, "generated_at": generated_at, "book": str(Path(book).resolve()),
              "book_hash": mail.digest(Path(book).read_bytes()), "added": added, "cards": cards, "mails": mails,
              "warnings": warnings, "schedule": daily.get(state, "activation", {}),
              "candidate_pool": candidate_pool.summary(state, day), "daily_limit": daily.DAILY_TARGET,
              "verified_total": sum(a["status"] == "sent_verified" for a in attempts.values()),
              "sent_today": sum(a["status"] == "sent_verified" and a.get("receipt", {}).get("sent_at", "")[:10] == day for a in attempts.values())}
    result["hash"] = mail.digest({k: v for k, v in result.items() if k != "generated_at"})
    result["document_model_hash"] = mail.digest({k: result[k] for k in ("date", "added", "cards", "warnings")})
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
    original_document_hash = mail.digest(document.read_bytes()) if document.exists() else None
    from review_docx import write_document
    from review_page import render_page
    old_document_model = previous.get("document_model_hash")
    if old_document_model is None and previous.get("model_hash") == mail.digest({k: v for k, v in snapshot.items()
            if k not in {"generated_at", "hash", "document_model_hash", "candidate_pool", "daily_limit"}}):
        old_document_model = snapshot["document_model_hash"]
    if not previous or old_document_model != snapshot["document_model_hash"] or not document.exists():
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
                "document_model_hash": snapshot["document_model_hash"],
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
    p = sub.add_parser("adopt-report"); p.add_argument("source", type=Path)
    args = parser.parse_args()
    state = mail.State(args.state)
    try:
        initialize(state)
        if args.command == "adopt-report": result = adopt_document(state, args.source, args.book)
        else: result = build(state, args.day, args.book)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if getattr(args, "open", False): subprocess.run(["open", result["page"]], check=True)
    finally: state.close()


if __name__ == "__main__":
    try: main()
    except (mail.DraftError, OSError, ValueError, KeyError) as error:
        print("审核入口未更新：" + str(error), file=sys.stderr)
        sys.exit(1)
