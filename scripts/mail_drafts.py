#!/usr/bin/env python3
"""Prepare local first-contact text and IMAP drafts. No SMTP or sending code."""
from __future__ import annotations

import argparse
import base64
from collections import Counter
from datetime import datetime
from email import policy
from email.headerregistry import Address
from email.message import EmailMessage
from email.parser import BytesParser
from email.utils import format_datetime, getaddresses
import fcntl
import hashlib
from html import unescape
from html.parser import HTMLParser
import imaplib
import json
import os
from pathlib import Path
import re
import sqlite3
import ssl
import sys
import tempfile
from urllib.parse import unquote, urlparse
import uuid
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from zoneinfo import ZoneInfo

# The existing read-only adapter uses sibling imports, both for script and -m use.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import workflow
from mail_keychain import CredentialError, Keychain, prompt


ROOT = Path(__file__).resolve().parent.parent
DEFAULT_STATE = Path.home() / "Library/Application Support/Tailki/mail-drafts.sqlite3"
CATEGORIES = ("宠物剧情类", "宠物用品类", "非宠物垂类")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
LOCAL_STATES = {"prepared", "cleared"}
UNCERTAIN_STATES = {"appending", "uncertain"}


class DraftError(RuntimeError):
    pass


def now():
    return datetime.now(ZoneInfo("Asia/Shanghai")).isoformat(timespec="seconds")


def digest(value):
    if not isinstance(value, bytes):
        value = json.dumps(value, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(value).hexdigest()


def email_address(value):
    value = str(value or "").strip()
    if not workflow.rules.is_valid_email(value) or any(c in value for c in "\r\n\x00"):
        raise DraftError("需要有效的单一邮箱地址。")
    try:
        Address(addr_spec=value)
    except (ValueError, IndexError) as error:
        raise DraftError("需要有效的单一邮箱地址。") from error
    return value


def paragraph_text(paragraph):
    out = []
    for node in paragraph.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "t":
            out.append(node.text or "")
        elif tag in ("br", "cr"):
            out.append("\n")
        elif tag == "tab":
            out.append("\t")
    return "".join(out)


def load_templates(path):
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    sections, current = {}, None
    for paragraph in root.findall("w:body/w:p", NS):
        text = paragraph_text(paragraph).strip()
        heading = re.fullmatch(r"\d+\s+初次建联\s+(.*)", text)
        if heading:
            category = heading.group(1).strip()
            if category not in CATEGORIES or category in sections:
                raise DraftError("模板分类未知或重复，请核实文档。")
            current = sections[category] = {"subject": "", "paragraphs": []}
        elif current is not None and text:
            if re.match(r"主题\s*[:：]", text):
                if current["subject"]:
                    raise DraftError("同一分类存在多个主题，请核实文档。")
                current["subject"] = re.split(r"[:：]", text, maxsplit=1)[1].strip()
            elif re.match(r"(?:策略|选择策略|收件人)\s*[:：]", text):
                continue
            elif current["subject"]:
                current["paragraphs"].append(text)
    if set(sections) != set(CATEGORIES):
        raise DraftError("文档必须包含三类首联模板。")
    for section in sections.values():
        section["body"] = "\n\n".join(section.pop("paragraphs"))
        if not section["subject"] or not section["body"]:
            raise DraftError("模板缺少主题或完整正文。")
    return sections


def fill_text(row, templates, planned_date):
    if row.get("联系状态") != "初次建联未回复":
        raise DraftError("联系状态不在本批首联范围。")
    if any(row.get(k) for k in ("对方第一次回信日期", "对方第一次回复邮件原始文本",
                                "我方第一次回复邮件日期", "我们的第一次回复邮件原始文本")):
        raise DraftError("已有回信或我方回复记录，先核实历史。")
    previous_date = str(row.get("初次建联日期") or "").strip()
    if previous_date and (not planned_date or previous_date[:10] != planned_date):
        raise DraftError("已有初次建联日期，尚未确认为计划日期。")
    category = str(row.get("首联内容分类") or "").strip()
    if category not in templates:
        raise DraftError("首联分类缺失或与模板不匹配。")
    recipient = email_address(row.get("Email"))
    url = urlparse(str(row.get("账号链接") or "").strip())
    if url.scheme not in ("https", "http") or url.hostname not in ("tiktok.com", "www.tiktok.com"):
        raise DraftError("缺少有效的 TikTok 账号链接。")
    match = re.fullmatch(r"/@([A-Za-z0-9._]+)/?", unquote(url.path))
    if not match:
        raise DraftError("无法从链接中明确提取 TikTok 账号。")
    replacements = {"邮箱": recipient, "邮箱前缀": recipient.split("@", 1)[0], "TikTok账号": match.group(1)}

    def fill(value):
        for key, replacement in replacements.items():
            value = value.replace("{{" + key + "}}", replacement)
        if "{{" in value or "}}" in value:
            raise DraftError("模板仍有未填写字段。")
        return value

    subject, body = fill(templates[category]["subject"]), fill(templates[category]["body"])
    if any(c in subject for c in "\r\n\x00") or "\x00" in body:
        raise DraftError("主题或正文含不支持的控制字符。")
    return {"creator_id": row["达人ID"], "name": str(row.get("达人昵称") or ""),
            "category": category, "recipient": recipient, "subject": subject, "body": body,
            "content_hash": digest([recipient, subject, body]), "status": "prepared", "reason": ""}


class State:
    def __init__(self, path=DEFAULT_STATE):
        self.path = Path(path).expanduser().resolve()
        if self.path.suffix.lower() not in {".sqlite3", ".sqlite", ".db"}:
            raise DraftError("本机状态库须使用 .sqlite3、.sqlite 或 .db 后缀，避免被当作公开文件上传。")
        self.path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
        # macOS flock and SQLite record locks interact on the same inode.
        # A private sidecar coordinates commands without locking the database itself.
        lock_path = self.path.with_suffix(self.path.suffix + ".lock")
        self.lock = lock_path.open("a+b")
        os.chmod(lock_path, 0o600)
        try:
            fcntl.flock(self.lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as error:
            self.lock.close()
            raise DraftError("已有邮件草稿任务运行，请完成后重试。") from error
        self.db = None
        old_umask = os.umask(0o077)
        try:
            self.db = sqlite3.connect(self.path)
            os.chmod(self.path, 0o600)
            self.db.execute("PRAGMA foreign_keys=ON")
            self.db.executescript("""
                CREATE TABLE IF NOT EXISTS config (key TEXT PRIMARY KEY, value TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS batches (id TEXT PRIMARY KEY, data TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS drafts (
                    id TEXT PRIMARY KEY, batch_id TEXT NOT NULL REFERENCES batches(id),
                    scope TEXT UNIQUE NOT NULL, data TEXT NOT NULL);
            """)
        except Exception:
            if self.db:
                self.db.close()
            self.lock.close()
            raise
        finally:
            os.umask(old_umask)

    def close(self):
        self.db.close()
        self.lock.close()

    def config(self):
        return dict(self.db.execute("SELECT key,value FROM config"))

    def set_config(self, values):
        with self.db:
            self.db.executemany("INSERT OR REPLACE INTO config VALUES (?,?)", values.items())

    def batch(self, batch_id):
        row = self.db.execute("SELECT data FROM batches WHERE id=?", (batch_id,)).fetchone()
        if not row:
            raise DraftError("未找到指定批次，请先准备文本。")
        return json.loads(row[0])

    def items(self, batch_id):
        self.batch(batch_id)
        return [json.loads(r[0]) for r in self.db.execute(
            "SELECT data FROM drafts WHERE batch_id=? ORDER BY rowid", (batch_id,))]

    def update(self, item, **changes):
        item.update(changes, updated_at=now())
        with self.db:
            self.db.execute("UPDATE drafts SET data=? WHERE id=?", (json.dumps(item, ensure_ascii=False), item["id"]))

    def bind(self, batch_id, account):
        batch = self.batch(batch_id)
        if batch.get("account") and batch["account"] != account:
            raise DraftError("本批已绑定另一个邮箱；请切回原账户，不迁移或重建草稿。")
        batch["account"] = account
        with self.db:
            self.db.execute("UPDATE batches SET data=? WHERE id=?", (json.dumps(batch, ensure_ascii=False), batch_id))


def prepare(state, book, template, ids, *, batch_id, planned_date="", expected_count=None):
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,80}", batch_id):
        raise DraftError("批次编号只能包含字母、数字、下划线或短横线。")
    if not ids or len(set(ids)) != len(ids):
        raise DraftError("需要明确、不重复的本批达人编号。")
    if expected_count is not None and len(ids) != expected_count:
        raise DraftError("选择人数与本批预期人数不一致。")
    if planned_date:
        try:
            datetime.strptime(planned_date, "%Y-%m-%d")
        except ValueError as error:
            raise DraftError("计划日期须为 YYYY-MM-DD。") from error
    book, template = Path(book).resolve(), Path(template).resolve()
    sources = {"book": str(book), "template": str(template),
               "book_hash": digest(book.read_bytes()), "template_hash": digest(template.read_bytes()),
               "ids": ids, "planned_date": planned_date}
    previous = state.db.execute("SELECT data FROM batches WHERE id=?", (batch_id,)).fetchone()
    if previous:
        if json.loads(previous[0])["sources"] != sources:
            raise DraftError("源文件或选择名单已变化；保留已有草稿，请先核实本批内容。")
        return report(state, batch_id)
    templates = load_templates(template)
    with tempfile.TemporaryDirectory(prefix="tailki-read-") as workdir:
        workbook = workflow.read_book(book, workdir, extract_images=False)
    sheet = next((s for s in workbook["sheets"] if s["name"] == "达人开发主表"), None)
    if not sheet:
        raise DraftError("没有找到达人开发主表。")
    rows = {}
    for row in sheet["rows"]:
        identifier = row.get("达人ID")
        if identifier in ids:
            if identifier in rows:
                raise DraftError("所选达人编号在表内重复，已停止。")
            rows[identifier] = row
    prepared, blocked = [], []
    for identifier in ids:
        try:
            if identifier not in rows:
                raise DraftError("表内没有此达人编号。")
            item = fill_text(rows[identifier], templates, planned_date)
            item.update(id=uuid.uuid4().hex, batch_id=batch_id, created_at=now())
            prepared.append(item)
        except DraftError as error:
            blocked.append({"creator_id": identifier, "reason": str(error)})
    addresses = Counter(x["recipient"].lower() for x in prepared)
    for item in prepared[:]:
        if addresses[item["recipient"].lower()] > 1:
            prepared.remove(item)
            blocked.append({"creator_id": item["creator_id"], "reason": "本批存在共享邮箱，先核实达人归属。"})
    if digest(book.read_bytes()) != sources["book_hash"] or digest(template.read_bytes()) != sources["template_hash"]:
        raise DraftError("读取期间源文件已变化，未保存本批。")
    batch = {"id": batch_id, "created_at": now(), "sources": sources, "blocked": blocked}
    try:
        with state.db:
            state.db.execute("INSERT INTO batches VALUES (?,?)", (batch_id, json.dumps(batch, ensure_ascii=False)))
            for item in prepared:
                scope = digest([str(book), item["creator_id"], "first-contact"])
                state.db.execute("INSERT INTO drafts VALUES (?,?,?,?)", (
                    item["id"], batch_id, scope, json.dumps(item, ensure_ascii=False)))
    except sqlite3.IntegrityError as error:
        raise DraftError("所选达人已有首联草稿批次，保留原稿，不重复创建。") from error
    return report(state, batch_id)


def source_check(state, batch_id):
    sources = state.batch(batch_id)["sources"]
    for key in ("book", "template"):
        if digest(Path(sources[key]).read_bytes()) != sources[key + "_hash"]:
            raise DraftError("源表或模板在准备后发生变化；请先核实，未继续上传草稿。")


def decode_mailbox(raw):
    text = raw.decode("utf-8")
    def decode(match):
        value = match.group(1)
        if not value:
            return "&"
        encoded = value.replace(",", "/")
        return base64.b64decode(encoded + "=" * (-len(encoded) % 4)).decode("utf-16-be")
    return re.sub(r"&([A-Za-z0-9+,]*)-", decode, text)


def parse_folders(entries):
    folders = []
    for entry in entries:
        if not entry:
            continue
        header = entry[0] if isinstance(entry, tuple) else entry
        match = re.fullmatch(rb'\(([^)]*)\)\s+(?:NIL|"(?:\\.|[^"\\])*")\s+(.+)', header)
        if not match:
            raise DraftError("无法可靠识别邮箱文件夹，未写入草稿。")
        name = entry[1] if isinstance(entry, tuple) else match.group(2)
        if not isinstance(entry, tuple) and name.startswith(b'"') and name.endswith(b'"'):
            name = re.sub(rb"\\(.)", rb"\1", name[1:-1])
        folders.append({"raw": name, "name": decode_mailbox(name),
                        "flags": {flag.lower() for flag in match.group(1).split()}})
    return folders


def choose_folder(folders, role, configured=""):
    if configured:
        selected = [f for f in folders if f["name"] == configured]
    else:
        flag = b"\\drafts" if role == "drafts" else b"\\sent"
        selected = [f for f in folders if flag in f["flags"]]
        if not selected:
            names = {"drafts", "draft", "草稿箱", "草稿"} if role == "drafts" else {"sent", "sent items", "sent messages", "已发送", "已发送邮件"}
            selected = [f for f in folders if f["name"].casefold() in names]
    selected = [f for f in selected if b"\\noselect" not in f["flags"]]
    if len(selected) != 1:
        label = "草稿箱" if role == "drafts" else "已发送"
        raise DraftError("无法唯一确定" + label + "，请用 setup 的文件夹选项指定现有文件夹；没有创建新文件夹。")
    return selected[0]


class _PlainHTML(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        self.parts.append(data)

    def handle_starttag(self, tag, attrs):
        if tag in ("br", "div", "p", "li"):
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in ("div", "p", "li"):
            self.parts.append("\n")


def normalized_body(body):
    return body.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")


def snapshot(raw):
    message = BytesParser(policy=policy.default).parsebytes(raw)
    body = message.get_body(preferencelist=("plain", "html"))
    text = body.get_content() if body is not None else ""
    if body is not None and body.get_content_type() == "text/html":
        parser = _PlainHTML()
        parser.feed(text)
        text = unescape("".join(parser.parts))
    result = {"subject": str(message.get("Subject", "")), "body": normalized_body(text),
              "to": [address.lower() for _, address in getaddresses(message.get_all("To", []))],
              "cc": [address.lower() for _, address in getaddresses(message.get_all("Cc", []))],
              "bcc": [address.lower() for _, address in getaddresses(message.get_all("Bcc", []))],
              "attachments": sum(1 for _ in message.iter_attachments()),
              "message_id": str(message.get("Message-ID", "")),
              "draft_id": str(message.get("X-Tailki-Draft-ID", ""))}
    result["hash"] = digest({k: result[k] for k in ("subject", "body", "to", "cc", "bcc", "attachments")})
    return result


def matches_original(item, remote):
    return (remote["to"] == [item["recipient"].lower()] and not remote["cc"] and not remote["bcc"]
            and not remote["attachments"] and remote["subject"] == item["subject"]
            and remote["body"] == normalized_body(item["body"]))


def make_message(item, account):
    message = EmailMessage(policy=policy.SMTP)
    message["From"] = Address("Sara", addr_spec=account)
    message["To"] = item["recipient"]
    message["Subject"] = item["subject"]
    message["Date"] = format_datetime(datetime.now(ZoneInfo("Asia/Shanghai")))
    message["Message-ID"] = item["message_id"]
    message["X-Tailki-Draft-ID"] = item["id"]
    message["X-Tailki-Batch-ID"] = item["batch_id"]
    message.set_content(item["body"], charset="utf-8")
    return message.as_bytes()


class Mailbox:
    """IMAP-only transport. Reads use EXAMINE and BODY.PEEK, writes only APPEND."""
    def __init__(self, account, secret, config=None, factory=None):
        config = config or {}
        self.account = account
        self.client = None
        try:
            self.client = (factory or imaplib.IMAP4_SSL)(
                "imap.qiye.aliyun.com", 993, ssl_context=ssl.create_default_context(), timeout=30)
            self.client.login(account, secret)
            code, entries = self.client.list()
            if code != "OK":
                raise DraftError("无法列出邮箱文件夹，未写入草稿。")
            folders = parse_folders(entries)
            self.drafts = choose_folder(folders, "drafts", config.get("drafts_folder", ""))
            self.sent = choose_folder(folders, "sent", config.get("sent_folder", ""))
        except Exception:
            self.close()
            raise

    def close(self):
        if self.client:
            try:
                self.client.logout()
            except (OSError, imaplib.IMAP4.error):
                pass

    @staticmethod
    def quoted(value):
        raw = value.encode("utf-8") if isinstance(value, str) else value
        return b'"' + raw.replace(b"\\", b"\\\\").replace(b'"', b'\\"') + b'"'

    def select(self, folder):
        code, _ = self.client.select(self.quoted(folder), readonly=True)
        if code != "OK":
            raise DraftError("无法只读打开邮箱文件夹，未继续操作。")
        _, values = self.client.response("UIDVALIDITY")
        if not values or not values[0]:
            raise DraftError("邮箱未返回 UIDVALIDITY，无法安全关联草稿。")
        return values[0].decode("ascii")

    def search(self, *criteria):
        code, values = self.client.uid("SEARCH", None, *criteria)
        if code != "OK":
            raise DraftError("邮箱检索未完成，未继续上传。")
        return list(dict.fromkeys(b" ".join(x for x in values if x).split()))

    def fetch(self, uid, *, headers=False):
        part = "(BODY.PEEK[HEADER])" if headers else "(BODY.PEEK[])"
        code, values = self.client.uid("FETCH", uid, part)
        chunks = [x[1] for x in (values or []) if isinstance(x, tuple) and isinstance(x[1], bytes)]
        if code != "OK" or len(chunks) != 1:
            raise DraftError("无法完整读回邮件，未继续上传。")
        return chunks[0]

    def history(self, recipient):
        found = []
        for name, folder in (("收件箱", b"INBOX"), ("已发送", self.sent["raw"])):
            self.select(folder)
            uids = set()
            for field in ("FROM", "TO", "CC"):
                uids.update(self.search(field, self.quoted(recipient)))
            for uid in sorted(uids):
                raw = self.fetch(uid, headers=True)
                message = BytesParser(policy=policy.default).parsebytes(raw)
                addresses = {address.lower() for _, address in getaddresses(
                    message.get_all("From", []) + message.get_all("To", []) + message.get_all("Cc", []))}
                if recipient.lower() in addresses:
                    found.append({"folder": name, "uid": uid.decode("ascii")})
        return found

    def find(self, item, *, sent=False):
        folder = self.sent if sent else self.drafts
        validity = self.select(folder["raw"])
        uids = set()
        for header, value in (("X-Tailki-Draft-ID", item["id"]), ("Message-ID", item.get("message_id", ""))):
            if not value:
                continue
            for uid in self.search("HEADER", header, self.quoted(value)):
                raw = self.fetch(uid, headers=True)
                message = BytesParser(policy=policy.default).parsebytes(raw)
                if str(message.get(header, "")).strip() == value:
                    uids.add(uid)
        # A web edit may retain its UID while changing Message-ID or custom headers.
        if not sent and not uids and item.get("uidvalidity") == validity and item.get("uid"):
            candidates = self.search("UID", item["uid"])
            uids.update(candidates)
        results = []
        for uid in sorted(uids):
            results.append({"uid": uid.decode("ascii"), "uidvalidity": validity,
                            "folder": folder["name"], "remote": snapshot(self.fetch(uid))})
        return results

    def existing_drafts(self, recipient):
        self.select(self.drafts["raw"])
        found = []
        for uid in self.search("TO", self.quoted(recipient)):
            message = BytesParser(policy=policy.default).parsebytes(self.fetch(uid, headers=True))
            addresses = {address.lower() for _, address in getaddresses(message.get_all("To", []))}
            if recipient.lower() in addresses:
                found.append(uid.decode("ascii"))
        return found

    def append(self, raw):
        code, _ = self.client.append(self.quoted(self.drafts["raw"]), r"(\Draft)", None, raw)
        if code != "OK":
            raise DraftError("草稿上传未获确认，已暂停；先查证，不自动重传。")


def reconcile(state, item, mailbox):
    found = mailbox.find(item)
    if len(found) > 1:
        state.update(item, status="needs_review", reason="发现多个关联草稿，未改动任何邮件。")
    elif found:
        remote = found[0]
        same = matches_original(item, remote["remote"])
        state.update(item, **remote, status="saved" if same else "edited",
                     reason="" if same else "邮箱中的修改已保留；未用初稿覆盖。")
    elif mailbox.find(item, sent=True):
        state.update(item, status="sent_manually", reason="在已发送中找到本批邮件，未重新创建。")
    elif item["status"] in UNCERTAIN_STATES:
        state.update(item, status="uncertain", reason="上传结果尚未核实，未自动重传。")
    else:
        state.update(item, status="removed", reason="原草稿已不在原位置，未自动重建。")
    return item


def audit(state, batch_id, mailbox):
    state.bind(batch_id, mailbox.account)
    for item in state.items(batch_id):
        if item["status"] not in LOCAL_STATES:
            continue
        hits = mailbox.history(item["recipient"])
        if hits:
            state.update(item, status="skipped_history", history=hits,
                         reason="收件箱或已发送存在相关历史，先核实，不重复首联。")
        elif mailbox.existing_drafts(item["recipient"]):
            state.update(item, status="skipped_existing_draft", reason="此收件人已有草稿，保留原稿，先核实。")
        else:
            state.update(item, status="cleared", reason="", audited_at=now())
    return report(state, batch_id)


def save(state, batch_id, mailbox):
    source_check(state, batch_id)
    state.bind(batch_id, mailbox.account)
    # Reconcile interrupted APPENDs before considering any new upload.
    for item in state.items(batch_id):
        if item["status"] in UNCERTAIN_STATES | {"saved", "edited"}:
            reconcile(state, item, mailbox)
        if item["status"] in UNCERTAIN_STATES | {"needs_review"}:
            return report(state, batch_id, halted=True)
    audit(state, batch_id, mailbox)
    source_check(state, batch_id)
    for item in state.items(batch_id):
        if item["status"] != "cleared":
            continue
        # Check again immediately before writing, including drafts created in the meantime.
        if mailbox.history(item["recipient"]) or mailbox.existing_drafts(item["recipient"]):
            state.update(item, status="skipped_history", reason="上传前发现相关邮件或草稿，先核实。")
            continue
        source_check(state, batch_id)
        message_id = "<tailki." + item["id"] + "@" + mailbox.account.split("@", 1)[1] + ">"
        state.update(item, status="appending", message_id=message_id)
        try:
            mailbox.append(make_message(item, mailbox.account))
            reconcile(state, item, mailbox)
            if item["status"] != "saved":
                return report(state, batch_id, halted=True)
        except (OSError, imaplib.IMAP4.error, DraftError):
            state.update(item, status="uncertain", reason="上传或读回未获完整确认，已暂停，未自动重传。")
            return report(state, batch_id, halted=True)
    return report(state, batch_id)


def verify(state, batch_id, mailbox):
    state.bind(batch_id, mailbox.account)
    for item in state.items(batch_id):
        if item["status"] not in LOCAL_STATES | {"skipped_history", "skipped_existing_draft"}:
            reconcile(state, item, mailbox)
    return report(state, batch_id)


def report(state, batch_id, *, halted=False, show=False):
    batch, items = state.batch(batch_id), state.items(batch_id)
    fields = ("creator_id", "category", "status", "reason")
    if show:
        fields += ("recipient", "subject", "body", "remote")
    return {"batch": batch_id, "prepared": len(items), "categories": dict(Counter(x["category"] for x in items)),
            "states": dict(Counter(x["status"] for x in items)), "blocked": batch.get("blocked", []),
            "halted": halted, "items": [{k: item[k] for k in fields if k in item} for item in items]}


def connect(config):
    if not config.get("account"):
        raise DraftError("尚未配置邮箱；本地文本已保留，请先运行 setup --gui。")
    secret = Keychain().get(config["account"])
    try:
        return Mailbox(config["account"], secret, config)
    finally:
        del secret


def setup(state, args):
    existing = state.config()
    if args.account:
        account = email_address(args.account)
    else:
        label = "请输入用于存放这批草稿的完整企业邮箱地址（包含 @ 和企业域名）。\n此处只填邮箱地址，下个窗口才输入专用密码。"
        for attempt in range(3):
            try:
                account = email_address(prompt(label, default=existing.get("account", ""), gui=args.gui))
                break
            except DraftError:
                if attempt == 2:
                    raise
                label = "刚才的内容不是完整邮箱地址。请填写形如 sara@企业域名 的地址，不填姓名或密码。\n只保存草稿，不发送邮件。"
    secret = prompt(
        "请粘贴阿里邮箱的三方客户端安全密码。\n设置 → 账户与安全 → 账户安全 → 三方客户端安全密码。\n"
        "内容会隐藏，只存入本机钥匙串，不会发送到聊天或 Git。", secret=True, gui=args.gui)
    if not secret or any(c in secret for c in "\r\n\x00"):
        raise DraftError("未输入有效专用密码；本地草稿保留。")
    previous_folders = existing if existing.get("account") == account else {}
    config = {"account": account,
              "drafts_folder": args.drafts_folder or previous_folders.get("drafts_folder", ""),
              "sent_folder": args.sent_folder or previous_folders.get("sent_folder", "")}
    mailbox = None
    try:
        mailbox = Mailbox(account, secret, config)
        Keychain().set(account, secret)
        state.set_config({"account": account, "drafts_folder": mailbox.drafts["name"], "sent_folder": mailbox.sent["name"]})
    finally:
        del secret
        if mailbox:
            mailbox.close()
    return {"configured": True, "message": "IMAP 连接及文件夹识别成功，专用密码已存入本机钥匙串；未发送邮件。"}


def main(argv=None):
    parser = argparse.ArgumentParser(description="本地首联文本及 IMAP 草稿，不包含 SMTP 或邮件发送。")
    parser.add_argument("--state", type=Path, default=DEFAULT_STATE)
    sub = parser.add_subparsers(dest="command", required=True)
    prepare_parser = sub.add_parser("prepare", help="读取指定表格和模板，填好本地邮件文本")
    prepare_parser.add_argument("--book", type=Path, required=True)
    prepare_parser.add_argument("--template", type=Path, required=True)
    prepare_parser.add_argument("--ids", required=True, help="逗号分隔的明确达人编号")
    prepare_parser.add_argument("--batch", required=True)
    prepare_parser.add_argument("--planned-date", default="")
    prepare_parser.add_argument("--expect-count", type=int)
    prepare_parser.add_argument("--show", action="store_true", help="在本机显示完整收件人、主题及正文")
    setup_parser = sub.add_parser("setup", help="在本机输入凭据并验证 IMAP")
    setup_parser.add_argument("--gui", action="store_true")
    setup_parser.add_argument("--account")
    setup_parser.add_argument("--drafts-folder")
    setup_parser.add_argument("--sent-folder")
    for command in ("audit", "save", "verify"):
        item_parser = sub.add_parser(command)
        item_parser.add_argument("batch")
        if command == "verify":
            item_parser.add_argument("--show", action="store_true")
            item_parser.add_argument("--local", action="store_true", help="只查看本地记录，不连接邮箱")
    args = parser.parse_args(argv)
    state, mailbox = None, None
    try:
        state = State(args.state)
        if args.command == "prepare":
            result = prepare(state, args.book, args.template, args.ids.split(","), batch_id=args.batch,
                             planned_date=args.planned_date, expected_count=args.expect_count)
            if args.show:
                result = report(state, args.batch, show=True)
        elif args.command == "setup":
            result = setup(state, args)
        elif args.command == "verify" and args.local:
            result = report(state, args.batch, show=args.show)
        else:
            mailbox = connect(state.config())
            result = {"audit": audit, "save": save, "verify": verify}[args.command](state, args.batch, mailbox)
            if args.command == "verify" and args.show:
                result = report(state, args.batch, show=True)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 1 if result.get("halted") else 0
    except (DraftError, CredentialError) as error:
        print(str(error), file=sys.stderr)
        return 1
    except (OSError, imaplib.IMAP4.error, sqlite3.Error, ValueError, ET.ParseError) as error:
        # Server, prompt and Keychain output are deliberately never included.
        print("本机文件或 IMAP 操作未完成（" + type(error).__name__ + "）。本地草稿保留；未自动重传或发送。", file=sys.stderr)
        return 1
    finally:
        if mailbox:
            mailbox.close()
        if state:
            state.close()


if __name__ == "__main__":
    sys.exit(main())
