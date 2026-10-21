#!/usr/bin/env python3
"""Review final IMAP drafts, then send only a separately approved SMTP batch."""
from __future__ import annotations

import argparse
import base64
from collections import Counter
from datetime import datetime, timedelta
from email import policy
from email.parser import BytesParser
from email.utils import format_datetime, getaddresses, parsedate_to_datetime
from html import escape
import imaplib
import json
import os
from pathlib import Path
import smtplib
import ssl
import sys
import time
import uuid

sys.path.insert(0, str(Path(__file__).resolve().parent))
if __package__:
    from . import mail_drafts as mail
else:
    import mail_drafts as mail

SMTP_HOST = "smtp.qiye.aliyun.com"
SMTP_PORT = 465
ADDRESS_HEADERS = ("From", "To", "Cc", "Bcc", "Reply-To")
DROP_HEADERS = ("Date", "Message-ID", "Bcc", "Return-Path", "Received", "Authentication-Results",
                "DKIM-Signature", "DomainKey-Signature", "X-Tailki-Draft-ID", "X-Tailki-Batch-ID",
                "X-Tailki-Send-ID")


def initialize(state):
    state.db.executescript("""
        CREATE TABLE IF NOT EXISTS send_reviews (id TEXT PRIMARY KEY, data TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS send_attempts (
            draft_id TEXT PRIMARY KEY, review_id TEXT NOT NULL, data TEXT NOT NULL);
    """)


def get_review(state, identifier):
    row = state.db.execute("SELECT data FROM send_reviews WHERE id=?", (identifier,)).fetchone()
    if not row:
        raise mail.DraftError("未找到最终发送预览。")
    return json.loads(row[0])


def put_review(state, review):
    with state.db:
        state.db.execute("INSERT OR REPLACE INTO send_reviews VALUES (?,?)",
                         (review["id"], json.dumps(review, ensure_ascii=False)))


def attempt(state, draft_id):
    row = state.db.execute("SELECT data FROM send_attempts WHERE draft_id=?", (draft_id,)).fetchone()
    return json.loads(row[0]) if row else None


def put_attempt(state, entry, result):
    with state.db:
        state.db.execute("UPDATE send_attempts SET data=? WHERE draft_id=?",
                         (json.dumps(result, ensure_ascii=False), entry["draft_id"]))


def parse(raw):
    message = BytesParser(policy=policy.default).parsebytes(raw)
    if any(part.defects for part in message.walk()):
        raise mail.DraftError("邮件格式不完整，停止处理，保留原草稿。")
    for header in (*ADDRESS_HEADERS, "Subject", "Sender", "Message-ID", "In-Reply-To", "References"):
        values = message.get_all(header, [])
        if len(values) > 1 or any(getattr(value, "defects", ()) for value in values):
            raise mail.DraftError("邮件头重复或格式不明，停止处理。")
    if any(key.lower().startswith("resent-") for key in message.keys()):
        raise mail.DraftError("首联草稿含重新投递邮件头，需单独核实。")
    if any(part.get_content_type() in ("multipart/signed", "multipart/encrypted") for part in message.walk()):
        raise mail.DraftError("签名或加密邮件需单独处理，未改写或发送。")
    return message


def addresses(message, header):
    result = []
    for _, value in getaddresses(message.get_all(header, [])):
        value = mail.email_address(value).lower()
        if not value.isascii():
            raise mail.DraftError("当前发送入口只支持 ASCII 邮箱地址，显示名称可使用中文。")
        result.append(value)
    return result


def mime_content(part):
    """Fingerprint decoded MIME contents, ignoring harmless transfer encodings."""
    value = {"type": part.get_content_type(), "disposition": part.get_content_disposition(),
             "filename": part.get_filename(), "cid": str(part.get("Content-ID", ""))}
    if part.is_multipart():
        value["parts"] = [mime_content(child) for child in part.iter_parts()]
    elif part.get_content_maintype() == "text":
        value["text"] = mail.normalized_body(part.get_content())
    else:
        value["payload_hash"] = mail.digest(part.get_payload(decode=True) or b"")
    return value


def content(raw, *, hide_bcc=False):
    message = parse(raw)
    result = {key.lower(): str(message.get(key, "")) for key in
              (*ADDRESS_HEADERS, "Subject", "Sender", "In-Reply-To", "References")}
    if hide_bcc:
        result["bcc"] = ""
    result["mime"] = mime_content(message)
    return result


def inspect(raw, account):
    message = parse(raw)
    if addresses(message, "From") != [account.lower()]:
        raise mail.DraftError("草稿发件人与当前认证邮箱不同，需人工核实。")
    if message.get("Sender") and addresses(message, "Sender") != [account.lower()]:
        raise mail.DraftError("草稿 Sender 与当前邮箱不同，需人工核实。")
    envelope = {key.lower(): addresses(message, key) for key in ("To", "Cc", "Bcc")}
    if len(envelope["to"]) != 1:
        raise mail.DraftError("本批每位达人须有一个明确的主收件人。")
    if not str(message.get("Subject", "")).strip() or not message.get_body():
        raise mail.DraftError("草稿缺少主题或正文，未纳入发送。")
    for part in message.walk():
        if not part.is_multipart() and part.get_content_type() in ("text/plain", "text/html"):
            text = part.get_content()
            if "{{" in text or "}}" in text:
                raise mail.DraftError("草稿仍有未填占位符，未纳入发送。")
    if "{{" in str(message["Subject"]) or "}}" in str(message["Subject"]):
        raise mail.DraftError("主题仍有未填占位符。")
    return {"headers": {key: str(message.get(key, "")) for key in ADDRESS_HEADERS},
            "subject": str(message["Subject"]), "envelope": envelope,
            "hash": mail.digest(content(raw)), "sent_hash": mail.digest(content(raw, hide_bcc=True))}


def read_final(state, item, mailbox):
    found = mailbox.find(item, include_raw=True)
    if len(found) != 1:
        raise mail.DraftError(item["creator_id"] + "：草稿缺失或存在多个关联版本，未重新创建。")
    remote = found[0]
    details = inspect(remote["raw"], mailbox.account)
    state.update(item, **{key: remote[key] for key in ("uid", "uidvalidity", "folder", "remote")},
                 status="saved" if mail.matches_original(item, remote["remote"]) else "edited")
    return remote, details


def history_check(mailbox, details, *, draft_uid=None):
    for recipient in dict.fromkeys(sum(details["envelope"].values(), [])):
        # Copies to the sender are intentional and do not represent KOL history.
        if recipient == mailbox.account.lower():
            continue
        if mailbox.history(recipient):
            raise mail.DraftError("收件人已有收件或已发送历史，停止重复首联，需先核实。")
        others = mailbox.existing_drafts(recipient)
        if any(uid != draft_uid for uid in others):
            raise mail.DraftError("收件人另有草稿，停止处理，保留原稿。")


def review_hash(review):
    return mail.digest({"account": review["account"], "batches": review["batches"],
                        "entries": review["entries"]})


def prepare_review(state, mailbox, batches, output, *, expected_count):
    initialize(state)
    if not batches or len(set(batches)) != len(batches):
        raise mail.DraftError("需要不重复的原草稿批次。")
    items, excluded = [], []
    for batch in batches:
        state.bind(batch, mailbox.account)
        for item in state.items(batch):
            if attempt(state, item["id"]) or item.get("smtp_attempt"):
                raise mail.DraftError("本批存在发送尝试，须先核对回执，不能重新批量发送。")
            if item["status"] in ("saved", "edited"):
                items.append(item)
            else:
                excluded.append({key: item.get(key, "") for key in ("creator_id", "name", "status", "reason")})
    if len(items) != expected_count or not items:
        raise mail.DraftError("可核对的远端草稿数与本批预期不一致：" + str(len(items)))
    entries, seen = [], set()
    for item in items:
        remote, details = read_final(state, item, mailbox)
        history_check(mailbox, details, draft_uid=remote["uid"])
        recipient = details["envelope"]["to"][0]
        if recipient in seen:
            raise mail.DraftError("最终草稿中出现重复主收件人，需先核实归属。")
        seen.add(recipient)
        entries.append({"draft_id": item["id"], "batch_id": item["batch_id"],
                        "creator_id": item["creator_id"], "name": item["name"], "category": item["category"],
                        "uid": remote["uid"], "uidvalidity": remote["uidvalidity"], "folder": remote["folder"],
                        "raw_b64": base64.b64encode(remote["raw"]).decode("ascii"), "details": details,
                        "content_hash": details["hash"], "sent_hash": details["sent_hash"], "edited": item["status"] == "edited"})
    review = {"account": mailbox.account, "batches": batches, "entries": entries, "excluded": excluded,
              "reviewed_at": mail.now(), "status": "awaiting_confirmation"}
    review["hash"] = review_hash(review)
    review["id"] = "smtp-" + review["hash"][:16]
    prior = state.db.execute("SELECT data FROM send_reviews WHERE id=?", (review["id"],)).fetchone()
    if prior and json.loads(prior[0])["status"] not in ("awaiting_confirmation", "approved", "stale"):
        raise mail.DraftError("此版本已有发送尝试，不重新创建发送批次。")
    # A new read is a new review; it never silently carries forward approval.
    render_review(review, output)
    review["preview"] = str(Path(output).resolve())
    review["preview_hash"] = mail.digest(Path(output).read_bytes())
    put_review(state, review)
    return summary(review)


def summary(review):
    return {"review": review["id"], "hash": review["hash"], "status": review["status"],
            "count": len(review["entries"]), "edited": sum(entry["edited"] for entry in review["entries"]),
            "categories": dict(Counter(entry["category"] for entry in review["entries"])),
            "preview": review.get("preview"), "excluded": review["excluded"]}


def render_review(review, output):
    cards, nav = [], []
    for index, entry in enumerate(review["entries"], 1):
        message = parse(base64.b64decode(entry["raw_b64"]))
        label = entry["name"] or entry["creator_id"]
        nav.append(f'<a href="#mail-{index}">{index}. {escape(label)}</a>')
        headers = "".join('<p><b>' + title + '：</b>' + escape(str(message.get(key, ""))) + '</p>'
                          for key, title in (("From", "发件人"), ("To", "收件人"), ("Cc", "抄送"),
                                             ("Bcc", "密送"), ("Reply-To", "回复地址")) if message.get(key))
        body = message.get_body(preferencelist=("plain",))
        plain = '<pre>' + escape(body.get_content()) + '</pre>' if body else ""
        html = message.get_body(preferencelist=("html",))
        rendered = ""
        if html:
            html_text = html.get_content()
            for part in message.walk():
                cid = str(part.get("Content-ID", "")).strip("<>")
                if cid and part.get_content_maintype() == "image":
                    uri = 'data:' + part.get_content_type() + ';base64,' + base64.b64encode(part.get_payload(decode=True) or b"").decode("ascii")
                    html_text = html_text.replace('cid:' + cid, uri)
            csp = "default-src 'none'; img-src data:; style-src 'unsafe-inline'; form-action 'none'; base-uri 'none'"
            source = '<meta http-equiv="Content-Security-Policy" content="' + escape(csp, quote=True) + '">' + html_text
            rendered = '<p>邮件排版（外部图片不加载）：</p><iframe sandbox="" referrerpolicy="no-referrer" srcdoc="' + escape(source, quote=True) + '"></iframe>'
        attachments = []
        for part in message.iter_attachments():
            payload = part.get_payload(decode=True)
            if payload is None:
                payload = part.as_bytes()
            name = part.get_filename() or "未命名附件"
            uri = "data:application/octet-stream;base64," + base64.b64encode(payload).decode("ascii")
            attachments.append('<li><a download="' + escape(name, quote=True) + '" href="' + uri + '">' + escape(name)
                               + '</a> · ' + escape(part.get_content_type()) + ' · ' + str(len(payload)) + ' 字节</li>')
        attachment_html = '<p>附件（将按当前文件发送）：</p><ul>' + ''.join(attachments) + '</ul>' if attachments else ""
        changed = '<p class="changed">已采用邮箱中的人工修改版本</p>' if entry["edited"] else ""
        cards.append(f'<article id="mail-{index}"><small>' + escape(entry["creator_id"] + ' · ' + entry["category"])
                     + '</small><h2>' + str(index) + '. ' + escape(label) + '</h2>' + changed + headers
                     + '<p><b>主题：</b>' + escape(str(message["Subject"])) + '</p>' + plain + rendered + attachment_html + '</article>')
    excluded = "".join('<li>' + escape(x["name"] + '（' + x["creator_id"] + '）：' + x["reason"]) + '</li>' for x in review["excluded"])
    page = '''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tailki SMTP 最终发送预览</title><style>
*{box-sizing:border-box}body{margin:0;background:#f5f6f8;color:#20242b;font:16px/1.7 -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif}
main{max-width:1080px;margin:auto;padding:36px 24px}h1{font-size:28px;line-height:1.3}h2{font-size:22px}nav{display:flex;flex-wrap:wrap;gap:8px 18px;margin:24px 0}
nav a,a{color:#31557d}article{background:white;border:1px solid #dfe3e8;border-radius:10px;padding:26px 30px;margin:22px 0;scroll-margin-top:18px}
pre{white-space:pre-wrap;overflow-wrap:anywhere;font:inherit;line-height:1.8;margin-top:24px}p{overflow-wrap:anywhere}.changed{background:#fff6e9;padding:10px}small{color:#606975}
iframe{width:100%;height:650px;border:1px solid #ddd;background:white}aside{background:#fff6e9;padding:12px 24px;border-radius:8px}
@media(max-width:640px){main{padding:20px 12px}article{padding:20px 18px}}@media print{body{background:white}nav{display:none}article{break-inside:avoid}iframe{height:900px}}
</style></head><body><main>'''
    page += '<h1>Tailki SMTP 最终发送预览 · ' + str(len(review["entries"])) + ' 封</h1>'
    page += '<p><b>等待你明确确认，尚未发送。</b>请核对以下每封邮件的全部地址、主题、正文及附件。内容变化后须重新确认。</p>'
    page += '<p>批次：<b>' + escape(review["id"]) + '</b><br>从邮箱读取时间：' + escape(review["reviewed_at"]) + '</p>'
    if excluded:
        page += '<aside><b>以下联系人不在本次发送范围：</b><ul>' + excluded + '</ul></aside>'
    page += '<nav>' + ''.join(nav) + '</nav>' + ''.join(cards) + '</main></body></html>'
    output = Path(output).expanduser().resolve()
    if output.suffix.lower() != ".html" or output.is_relative_to(mail.ROOT):
        raise mail.DraftError("完整预览须保存为仓库外的本机 HTML 文件，不进入 Git。")
    output.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    temporary = output.with_name('.' + output.name + '.tmp')
    descriptor = os.open(temporary, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
        handle.write(page)
    os.replace(temporary, output)
    os.chmod(output, 0o600)


def approve(state, identifier, instruction):
    """The caller must supply the human's actual instruction after showing review."""
    review = get_review(state, identifier)
    if review["status"] != "awaiting_confirmation":
        raise mail.DraftError("该预览当前不能批准；已有发送尝试或需要重新预览。")
    text = str(instruction.get("user_instruction", "")).strip()
    try:
        stamp = datetime.fromisoformat(instruction["user_instruction_at"])
        reviewed = datetime.fromisoformat(review["reviewed_at"])
        if stamp.tzinfo is None or stamp < reviewed or stamp > datetime.fromisoformat(mail.now()) + timedelta(minutes=5):
            raise ValueError()
    except (KeyError, TypeError, ValueError):
        raise mail.DraftError("确认时间须晚于本次完整预览，且具有时区。") from None
    if not text or instruction.get("review_hash") != review["hash"] or instruction.get("count") != len(review["entries"]):
        raise mail.DraftError("确认必须明确对应本批完整预览及全部邮件。")
    if mail.digest(Path(review["preview"]).read_bytes()) != review["preview_hash"]:
        raise mail.DraftError("预览文件已变化，请重新读取并展示最终版本。")
    review.update(status="approved", approval={**instruction, "recorded_at": mail.now()})
    put_review(state, review)
    return summary(review)


class DeliveryError(mail.DraftError):
    def __init__(self, status, reason, code=None):
        super().__init__(reason)
        self.status, self.code = status, code


class SMTPTransport:
    def __init__(self, account, secret, *, factory=None):
        self.account, self.client = account, None
        try:
            self.client = (factory or smtplib.SMTP_SSL)(SMTP_HOST, SMTP_PORT, timeout=30,
                                                       context=ssl.create_default_context())
            self.client.ehlo_or_helo_if_needed()
            self.client.login(account, secret)
        except (OSError, smtplib.SMTPException):
            self.close()
            raise mail.DraftError("SMTP 加密连接或认证失败，未投递任何邮件。请核对 SMTP 权限和本机专用密码。") from None

    def check(self):
        try:
            code, _ = self.client.noop()
        except (OSError, smtplib.SMTPException):
            raise mail.DraftError("SMTP 连通性核验失败；未投递邮件。") from None
        if code != 250:
            raise mail.DraftError("SMTP 连通性核验未通过；未投递邮件。")
        return {"smtp_authenticated": True, "host": SMTP_HOST, "port": SMTP_PORT, "sent": False}

    def close(self):
        if self.client:
            try:
                self.client.quit()
            except (OSError, smtplib.SMTPException):
                self.client.close()

    def submit(self, raw, recipients, stage):
        data_started = False
        try:
            stage("envelope")
            code, _ = self.client.mail(self.account)
            if code != 250:
                raise DeliveryError("failed", "SMTP 拒绝发件人，未提交正文，不自动重试。", code)
            for recipient in recipients:
                code, _ = self.client.rcpt(recipient)
                if code not in (250, 251):
                    self.client.rset()
                    raise DeliveryError("failed", "SMTP 拒绝收件人，整封未提交正文，不自动重试。", code)
            stage("data")  # Durable write before DATA; crashes never cause a retry.
            data_started = True
            code, _ = self.client.data(raw)
            if code != 250:
                raise DeliveryError("failed", "SMTP 明确拒绝邮件正文，不自动重试。", code)
            return {"smtp_code": code, "accepted_at": mail.now()}
        except DeliveryError:
            raise
        except smtplib.SMTPDataError as error:
            raise DeliveryError("failed", "SMTP 明确拒绝 DATA，不自动重试。", error.smtp_code) from None
        except (OSError, smtplib.SMTPException):
            status = "uncertain" if data_started else "failed"
            raise DeliveryError(status, "SMTP 连接中断，已停止整批；先核实已发送记录，不自动重试。") from None


def smtp_connect(config):
    if not config.get("account"):
        raise mail.DraftError("请先配置本机邮箱。")
    secret = mail.Keychain().get(config["account"])
    try:
        return SMTPTransport(config["account"], secret)
    finally:
        del secret


def current_item(state, entry):
    return next(item for item in state.items(entry["batch_id"]) if item["id"] == entry["draft_id"])


def check_current(state, review, entry, mailbox):
    if mailbox.account != review["account"]:
        raise mail.DraftError("发送邮箱与确认版不同。")
    item = current_item(state, entry)
    if item.get("smtp_attempt") or attempt(state, entry["draft_id"]):
        raise mail.DraftError("该邮件已有发送尝试，不能再次投递。")
    remote, details = read_final(state, item, mailbox)
    if (details["hash"] != entry["content_hash"] or remote["uid"] != entry["uid"]
            or remote["uidvalidity"] != entry["uidvalidity"]):
        raise mail.DraftError(entry["creator_id"] + "：草稿内容或位置已变化，须重新完整预览并确认。")
    history_check(mailbox, details, draft_uid=entry["uid"])


def outgoing(entry, review, send_id, sent_at):
    message = parse(base64.b64decode(entry["raw_b64"]))
    for key in DROP_HEADERS:
        if key in message:
            del message[key]
    message["Date"] = format_datetime(datetime.fromisoformat(sent_at))
    message["Message-ID"] = '<tailki.send.' + send_id + '@' + review["account"].split('@', 1)[1] + '>'
    message["X-Tailki-Send-ID"] = send_id
    raw = message.as_bytes(policy=policy.SMTP.clone(cte_type="7bit"))
    if mail.digest(content(raw, hide_bcc=True)) != entry["sent_hash"]:
        raise mail.DraftError("生成投递邮件时内容发生变化，停止发送。")
    return raw, str(message["Message-ID"])


def find_sent(mailbox, result):
    validity, headers = mailbox.header_index(mailbox.sent["raw"])
    uids = [uid for uid, message in headers.items()
            if str(message.get("Message-ID", "")).strip() == result["message_id"]]
    records = []
    for uid in uids:
        raw = mailbox.fetch(uid)
        records.append({"uid": uid.decode('ascii'), "uidvalidity": validity,
                        "folder": mailbox.sent["name"], "raw": raw})
    return records


def reconcile_one(state, review, entry, mailbox):
    result = attempt(state, entry["draft_id"])
    if not result or result["status"] == "sent_verified":
        return result
    found = find_sent(mailbox, result)
    if not found:
        return result
    if len(found) != 1 or mail.digest(content(found[0]["raw"], hide_bcc=True)) != entry["sent_hash"]:
        result.update(status="needs_review", reason="已发送中出现多个版本或内容不一致，未重发。")
    else:
        record = found[0]
        message = parse(record["raw"])
        try:
            sent_at = parsedate_to_datetime(str(message["Date"]))
            if sent_at.tzinfo is None or sent_at < datetime.fromisoformat(result["started_at"]) - timedelta(minutes=5):
                raise ValueError()
        except (TypeError, ValueError, KeyError):
            raise mail.DraftError("已发送邮件缺少可靠发送时间，保留待核实状态。") from None
        details = inspect(record["raw"], mailbox.account)
        body = mail.snapshot(record["raw"])
        receipt = {"result": "success", "verification": "sent_folder", "message_id": result["message_id"],
                   "source": record["folder"] + ':UIDVALIDITY=' + record["uidvalidity"] + ':UID=' + record["uid"],
                   "sent_at": sent_at.isoformat(), "recipient": details["envelope"]["to"][0],
                   "content_hash": mail.digest([details["envelope"]["to"][0], body["subject"], body["body"]]),
                   "mime_hash": entry["sent_hash"]}
        result.update(status="sent_verified", reason="", receipt=receipt, verified_at=mail.now())
        item = current_item(state, entry)
        state.update(item, status="sent_verified", sent_receipt=receipt)
    put_attempt(state, entry, result)
    return result


def deliver(state, identifier, mailbox, smtp_factory=None, *, pause=time.sleep, progress=None):
    review = get_review(state, identifier)
    if review["status"] != "approved" or not review.get("approval"):
        raise mail.DraftError("尚未取得该最终版本的明确发送确认，未发送。")
    if review_hash(review) != review["hash"] or review["approval"]["review_hash"] != review["hash"]:
        raise mail.DraftError("确认指纹不一致，未发送。")
    if mail.digest(Path(review["preview"]).read_bytes()) != review["preview_hash"]:
        raise mail.DraftError("已确认预览文件变化，须重新预览确认。")
    try:
        for entry in review["entries"]:
            check_current(state, review, entry, mailbox)
    except (mail.DraftError, OSError, imaplib.IMAP4.error):
        review["status"] = "stale"
        put_review(state, review)
        raise mail.DraftError("发送前复核未通过；草稿或历史有变化，请重新核对完整预览。") from None
    transport = (smtp_factory or smtp_connect)(state.config())
    try:
        if transport.account != review["account"]:
            raise mail.DraftError("SMTP 账号与确认批次不同。")
        transport.check()
        # Consume approval before the first submission. Interrupted batches need
        # reconciliation and a fresh explicit decision for any untouched rest.
        review["status"] = "dispatching"
        put_review(state, review)
        for entry in review["entries"]:
            check_current(state, review, entry, mailbox)
            send_id, stamp = uuid.uuid4().hex, mail.now()
            raw, message_id = outgoing(entry, review, send_id, stamp)
            result = {"status": "dispatching", "phase": "reserved", "started_at": stamp,
                      "send_id": send_id, "message_id": message_id, "review_id": identifier,
                      "payload_hash": mail.digest(raw), "reason": ""}
            with state.db:
                state.db.execute("INSERT INTO send_attempts VALUES (?,?,?)",
                                 (entry["draft_id"], identifier, json.dumps(result, ensure_ascii=False)))
            item = current_item(state, entry)
            state.update(item, smtp_attempt=send_id)

            def stage(phase):
                result.update(phase=phase, updated_at=mail.now())
                put_attempt(state, entry, result)

            recipients = list(dict.fromkeys(sum(entry["details"]["envelope"].values(), [])))
            try:
                accepted = transport.submit(raw, recipients, stage)
                result.update(status="smtp_accepted", **accepted)
            except DeliveryError as error:
                result.update(status=error.status, reason=str(error), smtp_code=error.code)
            put_attempt(state, entry, result)
            for retry in range(3):
                result = reconcile_one(state, review, entry, mailbox)
                if result["status"] == "sent_verified" or result["status"] in ("failed", "needs_review"):
                    break
                if retry < 2:
                    pause(2)
            if progress:
                progress({"creator_id": entry["creator_id"], "status": result["status"]})
            if result["status"] != "sent_verified":
                review["status"] = "halted"
                put_review(state, review)
                return delivery_report(state, identifier)
            pause(3)
        review["status"] = "sent_verified"
        put_review(state, review)
        return delivery_report(state, identifier)
    except BaseException:
        # Even process interruption leaves approval consumed and attempts durable.
        review["status"] = "halted"
        put_review(state, review)
        raise
    finally:
        transport.close()


def delivery_report(state, identifier):
    review = get_review(state, identifier)
    entries = [{"creator_id": entry["creator_id"], **(attempt(state, entry["draft_id"]) or {"status": "not_attempted"})}
               for entry in review["entries"]]
    return {"review": identifier, "status": review["status"],
            "states": dict(Counter(entry["status"] for entry in entries)), "items": entries}


def reconcile_delivery(state, identifier, mailbox):
    review = get_review(state, identifier)
    if mailbox.account != review["account"]:
        raise mail.DraftError("核验账号与发送账号不同。")
    for entry in review["entries"]:
        reconcile_one(state, review, entry, mailbox)
    if all(attempt(state, entry["draft_id"]) and attempt(state, entry["draft_id"])["status"] == "sent_verified"
           for entry in review["entries"]):
        review["status"] = "sent_verified"
        put_review(state, review)
    return delivery_report(state, identifier)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--state", type=Path, default=mail.DEFAULT_STATE)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("check-smtp")
    p = sub.add_parser("review")
    p.add_argument("--batches", required=True)
    p.add_argument("--expect-count", type=int, required=True)
    p.add_argument("--output", type=Path, required=True)
    p = sub.add_parser("approve")
    p.add_argument("review")
    p.add_argument("instruction", type=Path)
    for command in ("send", "reconcile", "status"):
        sub.add_parser(command).add_argument("review")
    args = parser.parse_args()
    state, mailbox, smtp = mail.State(args.state), None, None
    try:
        initialize(state)
        if args.command == "check-smtp":
            smtp = smtp_connect(state.config())
            result = smtp.check()
        elif args.command == "approve":
            result = approve(state, args.review, json.loads(args.instruction.read_text(encoding="utf-8")))
        elif args.command == "status":
            result = delivery_report(state, args.review)
        else:
            mailbox = mail.connect(state.config())
            if args.command == "review":
                result = prepare_review(state, mailbox, args.batches.split(','), args.output, expected_count=args.expect_count)
            elif args.command == "send":
                result = deliver(state, args.review, mailbox, progress=lambda value: print(json.dumps(value, ensure_ascii=False), flush=True))
            else:
                result = reconcile_delivery(state, args.review, mailbox)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    finally:
        if smtp:
            smtp.close()
        if mailbox:
            mailbox.close()
        state.close()


if __name__ == "__main__":
    try:
        main()
    except (mail.DraftError, mail.CredentialError) as error:
        print("操作未完成：" + str(error), file=sys.stderr)
        sys.exit(1)
    except (ValueError, KeyError, OSError, imaplib.IMAP4.error):
        print("操作未完成：请核对本机状态和邮箱；未自动重试。连接错误详情不输出，以免暴露邮箱或凭据。", file=sys.stderr)
        sys.exit(1)
