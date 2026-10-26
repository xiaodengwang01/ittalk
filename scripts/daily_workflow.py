#!/usr/bin/env python3
"""Daily preparation, local records and read-only history audit. Never sends mail."""
from __future__ import annotations

import argparse
from datetime import datetime, timedelta
import json
from pathlib import Path
import re
import statistics
import sys
from urllib.parse import urlparse
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parent))
import mail_drafts as mail
import daily_workbook as workbook
import mail_records as records

TZ = ZoneInfo("Asia/Shanghai")
ROOT = Path(__file__).resolve().parent.parent
BUSINESS = ROOT / "KOL搜集与邮件回复"
BOOK = BUSINESS / "达人跟进表.xlsx"
TEMPLATE = BUSINESS / "Tailki_KOL初次建联邮件_三个版本.docx"
STRATEGY = BUSINESS / "Tailki_KOL回复策略-第二版.docx"
REFERENCE = BUSINESS / "Tailki_KOL达人回复参考.docx"
DAILY_TARGET = 5
RUN_HOUR = 10
RESEARCH_COLUMNS = {"达人ID", "达人昵称", "Email", "首联内容分类", "账号链接", "宠物类型", "内容类型",
                    "粉丝数", "近10条最高播放", "近10条中位播放", "内容能力(1-5)", "宠物人群匹配度(1-5)",
                    "播放能力(1-5)", "联系方式完整度(1-5)", "是否低粉高播放", "是否带货类似产品"}


def timestamp(value):
    result = datetime.fromisoformat(value)
    if result.tzinfo is None:
        raise mail.DraftError("时间必须含时区。")
    return result.astimezone(TZ)


def current_time():
    return datetime.now(TZ)


def initialize(state):
    state.db.execute("CREATE TABLE IF NOT EXISTS daily_state (key TEXT PRIMARY KEY, data TEXT NOT NULL)")
    state.db.commit()
    records.initialize(state)


def get(state, key, default=None):
    row = state.db.execute("SELECT data FROM daily_state WHERE key=?", (key,)).fetchone()
    return json.loads(row[0]) if row else default


def put(state, key, value):
    with state.db:
        state.db.execute("INSERT OR REPLACE INTO daily_state VALUES (?,?)",
                         (key, json.dumps(value, ensure_ascii=False)))


def all_items(state, prefix):
    return [json.loads(r[0]) for r in state.db.execute(
        "SELECT data FROM daily_state WHERE key LIKE ? ORDER BY key", (prefix + "%",))]


def source_fingerprints(book=BOOK, template=TEMPLATE, strategy=STRATEGY, reference=REFERENCE):
    return {key: mail.digest(Path(path).read_bytes()) for key, path in
            (("book", book), ("template", template), ("strategy", strategy), ("reference", reference))}


def drafts_snapshot(state):
    return [json.loads(r[0]) for r in state.db.execute("SELECT data FROM drafts ORDER BY id")]


def preflight(state, book=BOOK, template=TEMPLATE, strategy=STRATEGY, reference=REFERENCE):
    """Read-only readiness inventory, not an assertion of actual delivery."""
    sheet = workbook.read(book)
    templates = mail.load_templates(template)
    snapshot = drafts_snapshot(state)
    return {
        "mode": get(state, "activation", {}).get("status", "paused"),
        "sources": source_fingerprints(book, template, strategy, reference),
        "creator_ids": sorted(r["达人ID"] for r in sheet["creators"]),
        "draft_ids": sorted(r["id"] for r in snapshot),
        "draft_snapshot": mail.digest(snapshot),
        "categories": list(templates),
        "requires_live_reconciliation": True,
        "next_step": "每日只读核对联系历史，准备名单、完整预览和未读摘要；发送须另获该版本确认。",
    }


def require_instruction(record):
    if not isinstance(record.get("user_instruction"), str) or not record["user_instruction"].strip() or not record.get("user_instruction_at"):
        raise mail.DraftError("缺少当前用户明确指令的原文和时间；不能代替用户确认。")
    moment = timestamp(record["user_instruction_at"])
    if moment > current_time() + timedelta(minutes=5):
        raise mail.DraftError("用户指令时间不能在未来。")


def activate(state, evidence, **paths):
    require_instruction(evidence)
    before = preflight(state, **paths)
    if evidence.get("pipeline_ready") is not True or evidence.get("backlog_cleared") is not True:
        raise mail.DraftError("发送链路和积压邮件尚未全部完成，保持暂停。")
    if (not evidence.get("pipeline") or not evidence.get("account") or evidence.get("unresolved") != []
            or not isinstance(evidence.get("sent_contacts"), list)):
        raise mail.DraftError("发送链路、邮箱或历史核对仍不完整，保持暂停。")
    for field in ("sources", "creator_ids", "draft_ids", "draft_snapshot"):
        if evidence.get(field) != before[field]:
            raise mail.DraftError("启动核对已过期或不完整，请重新核对：" + field)
    now = current_time()
    checked = timestamp(evidence["checked_at"])
    if not timedelta(0) <= now - checked <= timedelta(hours=24):
        raise mail.DraftError("启动前核对必须在最近 24 小时内完成。")
    receipts = evidence.get("backlog_receipts", {})
    for item in drafts_snapshot(state):
        receipt = receipts.get(item["id"], {})
        verify_receipt(item, receipt)
    first = now.replace(hour=RUN_HOUR, minute=0, second=0, microsecond=0)
    if first <= now:
        first += timedelta(days=1)
    value = {"status": "active", "activated_at": now.isoformat(), "first_run": first.isoformat(),
             "evidence": evidence}
    if get(state, "activation", {}).get("status") == "active":
        raise mail.DraftError("已经启用，不能重置起始日期。")
    put(state, "activation", value)
    return value


def configure_schedule(state, evidence, **paths):
    """Authorize preparation at ten, independently of backlog or send approval."""
    require_instruction(evidence)
    if evidence.get("scope") != "prepare_only" or evidence.get("hour") != RUN_HOUR or evidence.get("timezone") != "Asia/Shanghai":
        raise mail.DraftError("定时授权仅允许北京时间每天 10:00 的准备任务。")
    account = state.config().get("account", "")
    if not account or evidence.get("account", "").lower() != account.lower():
        raise mail.DraftError("定时授权邮箱与本机配置不一致。")
    snapshot = preflight(state, **paths)
    if evidence.get("sources") != snapshot["sources"]:
        raise mail.DraftError("定时授权的资料快照已变化，请重新读取。")
    now = current_time()
    first = now.replace(hour=RUN_HOUR, minute=0, second=0, microsecond=0)
    if first <= now: first += timedelta(days=1)
    previous = get(state, "activation", {})
    if previous.get("evidence") == evidence and previous.get("scope") == "prepare_only":
        return previous
    value = {"status": "pending_scheduler", "scope": "prepare_only", "hour": RUN_HOUR,
             "timezone": "Asia/Shanghai", "account": account, "first_run": first.isoformat(),
             "activated_at": now.isoformat(), "evidence": evidence,
             "previous": {k: v for k, v in previous.items() if k != "previous"}}
    put(state, "activation", value)
    return value


def confirm_schedule(state, automation_id):
    value = get(state, "activation", {})
    if value.get("scope") != "prepare_only" or value.get("status") not in ("pending_scheduler", "active") or not automation_id:
        raise mail.DraftError("先配置定时准备授权，再核实调度器启用结果。")
    value.update(status="active", automation_id=automation_id)
    put(state, "activation", value)
    return value


def audit_history(state, mailbox):
    """Minimal headers only, including seen mail. No body backscan or flag changes."""
    if mailbox.account.lower() != state.config().get("account", "").lower():
        raise mail.DraftError("历史核对邮箱不一致。")
    folders, contacts = {}, set()
    from email.utils import getaddresses
    for name, folder in (("INBOX", b"INBOX"), ("Sent", mailbox.sent["raw"]), ("Drafts", mailbox.drafts["raw"])):
        validity, headers = mailbox.header_index(folder, refresh=True)
        folders[name] = {"uidvalidity": validity, "uids": [u.decode("ascii") for u in headers]}
        if name == "Sent":
            for message in headers.values():
                contacts.update(a.lower() for _, a in getaddresses(message.get_all("To", []) + message.get_all("Cc", []))
                                if a and a.lower() != mailbox.account.lower())
    value = {"account": mailbox.account, "checked_at": current_time().isoformat(), "folders": folders,
             "sent_contacts": [{"email": a} for a in sorted(contacts)]}
    put(state, "history:" + current_time().date().isoformat(), value)
    return value


def sync_records(state, book=BOOK):
    """Reuse already-collected local bodies. Does not access the mailbox or change Excel."""
    rows = workbook.read(book)["creators"]
    collected = []
    accounts = {}
    from email.utils import parsedate_to_datetime
    for audit in all_items(state, "targeted-thread-audit:"):
        account = audit.get("account")
        for folder, messages in audit.get("folders", {}).items():
            for item in messages:
                headers = item.get("headers", {})
                message_id = headers.get("Message-ID", "")
                if not account or not message_id: continue
                accounts.setdefault(message_id, set()).add(account)
                addresses = headers.get("From", "") if folder == "INBOX" else headers.get("To", "")
                creator_id = records.creator_for({"from": addresses}, rows)
                source = folder + ":UIDVALIDITY=" + item["uidvalidity"] + ":UID=" + item["uid"]
                try:
                    moment = parsedate_to_datetime(headers.get("Date", ""))
                    stamp = moment.isoformat() if moment.tzinfo else ""
                except (TypeError, ValueError): stamp = ""
                collected.append(records.put(state, {"account": account, "message_id": message_id,
                    "direction": "incoming" if folder == "INBOX" else "outgoing", "creator_id": creator_id,
                    "source": source, "from": headers.get("From", ""), "to": headers.get("To", ""),
                    "subject": headers.get("Subject", ""), "in_reply_to": headers.get("In-Reply-To", ""),
                    "references": headers.get("References", ""), "received_at" if folder == "INBOX" else "sent_at": stamp,
                    "body_status": "not_collected", "verification": "observed_headers"}))
    for item in all_items(state, "inbox:"):
        collected.append(records.incoming(state, item, records.creator_for(item, rows)))
    for item in all_items(state, "reply:"):
        message = item["message"]
        collected.append(records.incoming(state, message, message["creator_id"]))
    # Named historical reads were explicitly authorized at collection time.
    for group in all_items(state, "targeted-latest-messages:"):
        for item in group.values():
            if isinstance(item, dict) and item.get("raw_text") and item.get("headers", {}).get("Message-ID"):
                headers = item["headers"]
                observed_accounts = accounts.get(headers["Message-ID"], set())
                if len(observed_accounts) != 1: continue
                message = {**item, "account": next(iter(observed_accounts)), "message_id": headers["Message-ID"],
                           "from": headers.get("From", ""), "in_reply_to": headers.get("In-Reply-To", ""),
                           "references": headers.get("References", ""), "body_status": "collected"}
                collected.append(records.incoming(state, message, records.creator_for(message, rows)))
    return {"records": len(set(collected)), "workbook_changed": False, "mailbox_read": False}


def new_day(day):
    return {"date": day, "batch": "tailki-" + day.replace("-", ""), "creators": [],
            "reserved": [], "prepared": False, "rejected": [], "routes": []}


def start_today(state, day, evidence, **paths):
    """Authorize one day's preparation without activating the recurring schedule."""
    require_instruction(evidence)
    now = current_time()
    if day != now.date().isoformat() or timestamp(evidence["user_instruction_at"]).date() != now.date():
        raise mail.DraftError("单次运行需要北京时间当天的明确指令，不补跑往日。")
    if evidence.get("scope") != "prepare_only":
        raise mail.DraftError("单次入口仅准备名单和预览，不授予发送许可。")
    if not evidence.get("account") or evidence["account"].lower() != state.config().get("account", "").lower():
        raise mail.DraftError("历史核对账号与当前配置不一致。")
    checked = timestamp(evidence["checked_at"])
    if checked.date() != now.date() or not timedelta(0) <= now - checked <= timedelta(hours=24):
        raise mail.DraftError("单次运行需要当天实际邮箱历史核对。")
    if (evidence.get("history_folders") != ["INBOX", "Sent", "Drafts"]
            or evidence.get("unresolved") != [] or not isinstance(evidence.get("sent_contacts"), list)):
        raise mail.DraftError("收件、已发送和草稿历史核对不完整。")
    key = "day:" + day
    previous = get(state, key)
    if previous is not None:
        if previous.get("manual_authorization") != evidence:
            raise mail.DraftError("当天已有批次，不能替换单次授权或重置进度。")
        return previous
    before = preflight(state, **paths)
    for field in ("sources", "creator_ids", "draft_ids", "draft_snapshot"):
        if evidence.get(field) != before[field]:
            raise mail.DraftError("单次运行快照已变化，请重新核对：" + field)
    value = {**new_day(day), "manual_authorization": evidence}
    put(state, key, value)
    return value


def day_state(state, day):
    now = current_time()
    if day != now.date().isoformat():
        raise mail.DraftError("只处理北京时间当天；不补跑等待期间或往日任务。")
    key = "day:" + day
    value = get(state, key)
    manual = (value or {}).get("manual_authorization")
    if manual:
        require_instruction(manual)
        if (manual.get("scope") != "prepare_only" or timestamp(manual["user_instruction_at"]).date() != now.date()
                or manual.get("account", "").lower() != state.config().get("account", "").lower()):
            raise mail.DraftError("当天单次授权已失效或邮箱账号已变化。")
        return value
    activation = get(state, "activation", {})
    if activation.get("status") != "active" or now < timestamp(activation["first_run"]) or now.hour < activation.get("hour", RUN_HOUR):
        raise mail.DraftError("计划尚未到明确启用后的运行时间。")
    if activation.get("scope") == "prepare_only":
        if activation.get("account", "").lower() != state.config().get("account", "").lower():
            raise mail.DraftError("定时授权邮箱已变化，停止准备。")
        audit = get(state, "history:" + day, {})
        if (audit.get("account", "").lower() != activation["account"].lower()
                or not audit.get("checked_at") or not timedelta(0) <= now - timestamp(audit["checked_at"]) <= timedelta(hours=4)):
            raise mail.DraftError("先运行 audit-history 完成当天邮箱联系历史核对。")
    if value is None:
        value = new_day(day)
        put(state, key, value)
    return value


def start_supplement(state, day, evidence, **paths):
    """A new instruction may fill today's gap after the prior round was sent."""
    require_instruction(evidence)
    value = day_state(state, day)
    if timestamp(evidence["user_instruction_at"]).date().isoformat() != day:
        raise mail.DraftError("补充搜集需要当天的最新明确指令。")
    if (evidence.get("scope") != "prepare_only" or not evidence.get("account")
            or evidence["account"].lower() != state.config().get("account", "").lower()
            or evidence.get("history_folders") != ["INBOX", "Sent", "Drafts"]
            or evidence.get("unresolved") != [] or not isinstance(evidence.get("sent_contacts"), list)):
        raise mail.DraftError("补充批次的准备范围、账号或历史核对不完整。")
    now, checked = current_time(), timestamp(evidence["checked_at"])
    if checked.date() != now.date() or not timedelta(0) <= now - checked <= timedelta(hours=24):
        raise mail.DraftError("补充批次需要当天实际邮箱历史核对。")
    if value.get("supplement_authorization") == evidence:
        return value
    if not value["prepared"] or value["reserved"] or len(value["creators"]) >= DAILY_TARGET:
        raise mail.DraftError("当前批次尚未完成准备，或当天已达人数上限，不能新建补充批次。")
    items = state.items(value["batch"])
    prior_ids = value["creators"][value.get("round_start", 0):]
    if not items or {item["creator_id"] for item in items} != set(prior_ids):
        raise mail.DraftError("当前批次与当天名单不一致，先核对历史。")
    for item in items:
        if item.get("smtp_attempt"):
            # SMTP final versions may differ from the initial template. Validate
            # their own approved MIME receipt rather than the legacy draft hash.
            import mail_send
            result = mail_send.attempt(state, item["id"])
            if not result or result["status"] != "sent_verified" or result.get("receipt") != item.get("sent_receipt"):
                raise mail.DraftError("上一批 SMTP 发送仍未核实，不能建立补充批次。")
            review = mail_send.get_review(state, result["review_id"])
            entries = [e for e in review["entries"] if e["draft_id"] == item["id"]]
            if (len(entries) != 1 or mail_send.review_hash(review) != review["hash"]
                    or review.get("approval", {}).get("review_hash") != review["hash"]
                    or result["receipt"].get("mime_hash") != entries[0]["sent_hash"]):
                raise mail.DraftError("上一批 SMTP 回执与最终确认版不一致。")
            continue
        receipt = get(state, "receipt:" + item["id"])
        if item["status"] != "sent_verified" or not isinstance(receipt, dict) or receipt != item.get("sent_receipt"):
            raise mail.DraftError("上一批次仍有未核实发送，不能重置名单或确认。")
        verify_receipt(item, receipt)
    before = preflight(state, **paths)
    for field in ("sources", "creator_ids", "draft_ids", "draft_snapshot"):
        if evidence.get(field) != before[field]:
            raise mail.DraftError("补充批次核对快照已变化：" + field)
    rounds = value.setdefault("completed_rounds", [])
    next_batch = "tailki-" + day.replace("-", "") + "-part" + str(len(rounds) + 2)
    if state.db.execute("SELECT 1 FROM batches WHERE id=?", (next_batch,)).fetchone():
        raise mail.DraftError("补充批次编号已有历史，停止覆盖。")
    rounds.append({"batch": value["batch"], "creators": prior_ids, "completed_at": now.isoformat()})
    value.update(batch=next_batch, round_start=len(value["creators"]), prepared=False,
                 rejected=[], supplement_authorization=evidence)
    put(state, "day:" + day, value)
    return value


def profile(value):
    value = workbook.workflow.rules.canonical_account_url("TikTok", value)
    if not re.fullmatch(r"https://tiktok\.com/@[a-z0-9._]+", value):
        raise mail.DraftError("需要可核实的 TikTok 账号主页。")
    return value


def public_url(value):
    parts = urlparse(str(value or ""))
    return parts.scheme in ("http", "https") and bool(parts.hostname) and not parts.username


def validate_candidate(candidate, day):
    result = dict(candidate)
    result["账号链接"] = profile(candidate.get("账号链接"))
    result["Email"] = mail.email_address(candidate.get("Email")).lower()
    for key in ("达人昵称", "内容类型", "英语市场依据", "内容判断", "宠物匹配判断", "播放判断", "活跃判断"):
        if not str(candidate.get(key, "")).strip():
            raise mail.DraftError("缺少：" + key)
    if candidate.get("合格") is not True:
        raise mail.DraftError("尚未完成内容、宠物匹配、播放和活跃度评估。")
    if candidate.get("发现路径") not in ("FastMoss", "TikTok"):
        raise mail.DraftError("发现路径不明确。")
    if candidate.get("首联内容分类") not in mail.CATEGORIES:
        raise mail.DraftError("首联内容分类与三版本模板不匹配。")
    if candidate.get("宠物类型") not in ("Dog", "Cat", "Dog+Cat", "Cat+Dog", "Other"):
        raise mail.DraftError("缺少宠物类型。")
    for key in ("邮箱公开来源", "数据来源"):
        if not public_url(candidate.get(key)):
            raise mail.DraftError("需要公开来源链接：" + key)
    observed = timestamp(candidate["核实时间"])
    if observed.date().isoformat() != day or observed > current_time() + timedelta(minutes=5):
        raise mail.DraftError("需要当天实际核实时间。")
    if candidate.get("已联系") is not False or not candidate.get("历史核对依据"):
        raise mail.DraftError("需核对当前邮箱发送/收信/草稿历史，并确认没有既有联系。")
    history_time = timestamp(candidate["历史核对时间"])
    if history_time.date().isoformat() != day or history_time > current_time() + timedelta(minutes=5):
        raise mail.DraftError("需要当天的实际邮箱历史核对记录。")
    posts = candidate.get("近期视频", [])
    if not 1 <= len(posts) <= 10 or len({p["url"] for p in posts}) != len(posts):
        raise mail.DraftError("需要不重复的近期视频样本，最多 10 条。")
    for post in posts:
        if (not public_url(post.get("url")) or not isinstance(post.get("views"), int)
                or isinstance(post["views"], bool) or post["views"] < 0):
            raise mail.DraftError("视频来源或播放数据缺失。")
        if post.get("date"):
            posted = datetime.fromisoformat(post["date"]).date()
            if posted > observed.date():
                raise mail.DraftError("视频发布日期不能晚于核实日期。")
        else:
            # Preserve the site's visible precision; never invent an exact day
            # for a label such as "1 week ago".
            if not re.fullmatch(r"\d+\s*(分钟|小时|天|周)前", str(post.get("date_text", ""))):
                raise mail.DraftError("需要真实发布日期或网页显示的相对日期。")
            captured = timestamp(post["observed_at"])
            if captured.date() != observed.date() or captured > current_time() + timedelta(minutes=5):
                raise mail.DraftError("相对日期需要当天实际读取时间。")
    if len(posts) < 10 and not candidate.get("样本不足说明"):
        raise mail.DraftError("不足 10 条近期视频时需要说明样本限制。")
    for key in ("粉丝数", "内容能力(1-5)", "宠物人群匹配度(1-5)", "播放能力(1-5)", "联系方式完整度(1-5)"):
        value = candidate.get(key)
        if not isinstance(value, int) or isinstance(value, bool) or value < 0 or (key != "粉丝数" and not 1 <= value <= 5):
            raise mail.DraftError("数据缺失或评分不在 1–5 范围：" + key)
    result["近10条最高播放"] = max(p["views"] for p in posts)
    result["近10条中位播放"] = statistics.median(p["views"] for p in posts)
    for key in ("是否低粉高播放", "是否带货类似产品"):
        if candidate.get(key) not in ("是", "否", ""):
            raise mail.DraftError("加分项只能填是、否或留空。")
        if candidate.get(key) == "是" and not candidate.get(key + "依据"):
            raise mail.DraftError("加分项需要依据：" + key)
    return result


def recover_reserved(state, value, book):
    """Write-ahead rows make an interruption after file replacement resumable."""
    if not value["reserved"]:
        return
    book_hash = mail.digest(Path(book).read_bytes())
    sheet = workbook.read(book)
    existing = {r["达人ID"]: r for r in sheet["creators"]}
    missing = []
    for candidate in value["reserved"]:
        old = existing.get(candidate["达人ID"])
        if old:
            if profile(old.get("账号链接")) != candidate["账号链接"] or str(old.get("Email")).lower() != candidate["Email"]:
                raise mail.DraftError("上次保留的达人编号与人工记录冲突，停止恢复。")
        else:
            missing.append(candidate)
    if missing:
        last = max(r["_row"] for r in sheet["rows"])
        style_row = max((r["_row"] for r in sheet["main_creators"]), default=2)
        cells, new_rows = {}, []
        for offset, candidate in enumerate(missing, 1):
            number = last + offset
            new_rows.append(number)
            record = {**{k: v for k, v in candidate.items() if k in RESEARCH_COLUMNS},
                      "联系状态": "待首联确认", "初次建联日期": "", "平台": "TikTok",
                      "备注": "每日批次：" + value["batch"] + "\n首联未发送；" +
                      json.dumps({k: v for k, v in candidate.items() if k not in sheet["headers"]},
                                 ensure_ascii=False)}
            record.update(当前建联状况="待首联确认", WhatsApp联系方式="未提供", 可沟通时间="未提供",
                          收件信息="姓名：未提供\n电话：未提供\n地址及邮编：未提供", 历史批次=value["date"])
            for i, header in enumerate(sheet["headers"], 1):
                if header == "综合评分":
                    cells[f"{workbook.workflow.col_name(i)}{number}"] = {
                        "formula": workbook.score_formula(sheet["headers"], number)}
                elif header and record.get(header) not in ("", None):
                    cells[f"{workbook.workflow.col_name(i)}{number}"] = {"value": record[header]}
        workbook.patch_cells(book, cells, expected_hash=book_hash,
                             backup_dir=state.path.parent / "backups", style_row=style_row, new_rows=new_rows)
    value["creators"].extend(x["达人ID"] for x in value["reserved"] if x["达人ID"] not in value["creators"])
    value["reserved"] = []
    put(state, "day:" + value["date"], value)


def add_candidates(state, day, data, book=BOOK):
    value = day_state(state, day)
    recover_reserved(state, value, book)
    if set(data.get("routes_checked", [])) != {"FastMoss", "TikTok"}:
        raise mail.DraftError("需要同时记录 FastMoss 和 TikTok 两路搜索；不可用时先报告原因。")
    sheet = workbook.read(book)
    accounts = {profile(r["账号链接"]) for r in sheet["creators"] if r.get("账号链接")}
    emails = {str(r.get("Email", "")).strip().lower() for r in sheet["creators"]}
    emails.update(x["recipient"].lower() for x in drafts_snapshot(state))
    # Activation audit includes contacts already sent to but absent from this workbook.
    history = value.get("supplement_authorization", value.get("manual_authorization",
                        get(state, "activation", {}).get("evidence", {}))).get("sent_contacts", [])
    history = history + get(state, "history:" + day, {}).get("sent_contacts", [])
    emails.update(x["email"].lower() for x in history if x.get("email"))
    accounts.update(profile(x["profile"]) for x in history if x.get("profile"))
    next_id = max(int(re.search(r"\d+$", r["达人ID"]).group()) for r in sheet["creators"]) + 1
    candidates, rejected = [], []
    for raw in data.get("creators", []):
        try:
            candidate = validate_candidate(raw, day)
            if candidate["账号链接"] in accounts or candidate["Email"] in emails:
                raise mail.DraftError("已有账号、已有联系记录或共享邮箱，跳过并保留人工核对。")
            accounts.add(candidate["账号链接"])
            emails.add(candidate["Email"])
            candidates.append(candidate)
        except (mail.DraftError, ValueError, KeyError) as error:
            rejected.append({"account": raw.get("账号链接", ""), "reason": str(error)})
    # Content, audience, play and activity are assessed before numeric tie-breaks.
    candidates.sort(key=lambda c: (c["内容能力(1-5)"], c["宠物人群匹配度(1-5)"],
                                   c["播放能力(1-5)"], c.get("是否低粉高播放") == "是",
                                   c.get("是否带货类似产品") == "是"), reverse=True)
    remaining = max(0, DAILY_TARGET - len(value["creators"]))
    if value["prepared"] and candidates and remaining:
        raise mail.DraftError("当天批次已生成预览，不能静默更改名单；保留已准备的批次。")
    for candidate in candidates[:remaining]:
        candidate["达人ID"] = f"KOL-{next_id:06d}"
        next_id += 1
        value["reserved"].append(candidate)
    value["rejected"] = rejected
    value["routes"] = data["routes_checked"]
    put(state, "day:" + day, value)
    recover_reserved(state, value, book)
    return {"date": day, "added": value["creators"], "shortfall": max(0, DAILY_TARGET-len(value["creators"])),
            "rejected": rejected}


def prepare_day(state, day, book=BOOK, template=TEMPLATE):
    value = day_state(state, day)
    recover_reserved(state, value, book)
    creators = value["creators"][value.get("round_start", 0):]
    if not creators:
        return {"batch": None, "shortfall": max(0, DAILY_TARGET-len(value["creators"]))}
    exists = state.db.execute("SELECT 1 FROM batches WHERE id=?", (value["batch"],)).fetchone()
    if exists:
        result = mail.report(state, value["batch"])
    else:
        if len(creators) > DAILY_TARGET:
            raise mail.DraftError("既有未冻结名单超过当前每日 5 位上限，保留记录，请先核实；未新增草稿。")
        result = mail.prepare(state, book, template, creators, batch_id=value["batch"],
                              expected_count=len(creators))
    value["prepared"] = True
    put(state, "day:" + day, value)
    return {**result, "preview": mail.render_preview(state, value["batch"]),
            "content_hashes": {i["id"]: i["content_hash"] for i in state.items(value["batch"])},
            "shortfall": max(0, DAILY_TARGET-len(value["creators"]))}


def summarize_replies(state, messages):
    """No workbook path and no drafting/sending dependency: summaries stay pending."""
    required = ("account", "message_id", "creator_id", "received_at", "source",
                "raw_text", "summary", "changes", "strategy", "reason", "response_direction", "decisions")
    for message in messages:
        if any(k not in message for k in required) or any(not message[k] for k in
                ("account", "message_id", "creator_id", "source", "raw_text", "summary", "reason", "response_direction")):
            raise mail.DraftError("回信需单封定位、原文关联、摘要、策略理由和待决事项。")
        timestamp(message["received_at"])
        if message["strategy"] not in ("A", "B", "C", "D", "待判断"):
            raise mail.DraftError("建议策略须为 A/B/C/D 或待判断。")
        if any(k in message for k in ("reply_body", "reply_html", "send", "archived")):
            raise mail.DraftError("每日回信入口只接收摘要，不接收回复草稿、发送或归档指令。")
    changed = []
    for message in messages:
        key = "reply:" + mail.digest([message["account"].lower(), message["message_id"]])
        previous = get(state, key)
        if previous and previous["message"] == message:
            continue
        value = {"message": message, "status": "summarized_pending_instruction", "updated_at": mail.now(),
                 "history": (previous or {}).get("history", [])}
        if previous:
            value["history"].append({k: previous[k] for k in ("message", "status", "updated_at")})
        put(state, key, value)
        records.incoming(state, message, message["creator_id"])
        changed.append(message["message_id"])
    return {"new_or_changed": changed, "pending": pending_replies(state)}


def pending_replies(state):
    return [{k: v for k, v in item["message"].items() if k != "raw_text"} | {"status": item["status"]}
            for item in all_items(state, "reply:") if item["status"] == "summarized_pending_instruction"]


def approve_batch(state, batch, instruction):
    require_instruction(instruction)
    items = state.items(batch)
    hashes = {x["id"]: x["content_hash"] for x in items}
    if not items or state.batch(batch).get("blocked") or instruction.get("content_hashes") != hashes:
        raise mail.DraftError("确认必须对应完整预览的全部邮件及当前内容。")
    if any(x["status"] != "prepared" for x in items):
        raise mail.DraftError("此批已有远端草稿或发送状态，需核对最终版本后单独处理。")
    mail.refresh_source(state, batch, state.batch(batch)["sources"]["book"])
    put(state, "approval:" + batch, {**instruction, "approved_at": mail.now()})
    return {"batch": batch, "approved": True, "sent": False}


def verify_receipt(item, receipt):
    if receipt.get("result") != "success" or receipt.get("verification") != "sent_folder":
        raise mail.DraftError("必须实际核实已发送记录；失败或结果不明不能记为已发送。")
    for key in ("message_id", "source", "sent_at"):
        if not receipt.get(key):
            raise mail.DraftError("实际发送记录缺少：" + key)
    if timestamp(receipt["sent_at"]) > current_time() + timedelta(minutes=5):
        raise mail.DraftError("实际发送时间不能在未来。")
    if receipt.get("recipient", "").lower() != item["recipient"].lower() or receipt.get("content_hash") != item["content_hash"]:
        raise mail.DraftError("已发送的收件人或内容与确认版不一致，先核实。")


def claim_send(state, batch):
    """Reserve before the external pipeline runs; interrupted attempts never retry."""
    approval = get(state, "approval:" + batch)
    items = state.items(batch)
    if not approval or not items:
        raise mail.DraftError("缺少本批完整预览的明确确认。")
    if any(i["status"] != "prepared" for i in items):
        raise mail.DraftError("已有发送尝试、远端草稿或结果待核实；不得自动再次提交。")
    if approval["content_hashes"] != {i["id"]: i["content_hash"] for i in items}:
        raise mail.DraftError("批准的内容已变化，停止发送。")
    mail.refresh_source(state, batch, state.batch(batch)["sources"]["book"])
    with state.db:
        for item in items:
            item.update(status="dispatching", dispatch_started_at=mail.now())
            state.db.execute("UPDATE drafts SET data=? WHERE id=?", (json.dumps(item, ensure_ascii=False), item["id"]))
    return {"batch": batch, "handoff_to_user_pipeline": True, "items": items}


def record_sent(state, batch, receipts, book=BOOK):
    approval = get(state, "approval:" + batch)
    if not approval:
        raise mail.DraftError("本批未取得对应完整预览的发送确认。")
    items = {i["id"]: i for i in state.items(batch)}
    book_hash = mail.digest(Path(book).read_bytes())
    sheet = workbook.read(book)
    rows = {r["达人ID"]: r for r in sheet["creators"]}
    cells_by_sheet, accepted = {}, []
    for identifier, receipt in receipts.items():
        item = items[identifier]
        if approval["content_hashes"].get(identifier) != item["content_hash"]:
            raise mail.DraftError("确认后内容已变化，需要重新预览确认。")
        verify_receipt(item, receipt)
        if timestamp(receipt["sent_at"]) < timestamp(approval["user_instruction_at"]):
            raise mail.DraftError("发送记录早于本批授权，需要单独核对历史。")
        previous = get(state, "receipt:" + identifier)
        if previous and previous != receipt:
            raise mail.DraftError("已有不同发送记录，不能覆盖或重复发送。")
        row = rows[item["creator_id"]]
        target = next(s for s in sheet["sheets"] if s["name"] == row["_sheet"])
        cells = cells_by_sheet.setdefault(row["_sheet"], {})
        def set_field(header, value, **options):
            cells[workbook.address(target["headers"], header, row["_row"])] = {"value": value, **options}
        if str(row.get("Email", "")).lower() != item["recipient"].lower():
            raise mail.DraftError("主表邮箱已变更，停止写回。")
        day = timestamp(receipt["sent_at"]).date().isoformat()
        if row.get("初次建联日期") and str(row["初次建联日期"])[:10] != day:
            raise mail.DraftError("主表已有不同日期，保留人工记录。")
        if not row.get("初次建联日期"):
            set_field("初次建联日期", day, date=True)
        if row.get("联系状态") == "待首联确认":
            set_field("当前建联状况" if "当前建联状况" in target["headers"] else "联系状态", "初次建联未回复")
        # Replace only our own preparation marker, preserving all source notes
        # and any human edits around it.
        pending_note = "每日批次：" + batch + "\n首联未发送；"
        note = row.get("备注", "")
        if isinstance(note, str) and note.startswith(pending_note):
            set_field("备注", "每日批次：" + batch + "\n首联已发送（" + day + "）；" + note[len(pending_note):])
        receipt_marker = "已核验发送：" + receipt["message_id"]
        if "往来邮件梗概" in target["headers"] and not previous and receipt_marker not in str(note):
            from mail_postprocess import history_line
            set_field("往来邮件梗概", history_line(row.get("往来邮件梗概"), day, "我方",
                      "首联已发送；主题：" + item["subject"] + "。"))
            note_address = workbook.address(target["headers"], "备注", row["_row"])
            new_note = cells.get(note_address, {}).get("value", note or "")
            set_field("备注", new_note.rstrip() + "\n" + receipt_marker + "；" + receipt["sent_at"])
        accepted.append((item, receipt))
    for name, cells in cells_by_sheet.items():
        if not cells: continue
        row_number = int(re.search(r"\d+$", next(iter(cells)))[0])
        workbook.patch_cells(book, cells, expected_hash=book_hash, sheet_name=name,
                             backup_dir=state.path.parent / "backups", style_row=row_number)
        book_hash = mail.digest(Path(book).read_bytes())
    for item, receipt in accepted:
        put(state, "receipt:" + item["id"], receipt)
        mail.State.update(state, item, status="sent_verified", sent_receipt=receipt)
        import candidate_pool
        candidate_pool.mark_sent(state, item["creator_id"], receipt)
    return {"recorded": [i["creator_id"] for i, _ in accepted], "sent_by_this_command": False}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--state", type=Path, default=mail.DEFAULT_STATE)
    parser.add_argument("--book", type=Path, default=BOOK)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("status")
    sub.add_parser("preflight")
    for cmd in ("activate", "summarize", "configure-schedule"):
        sub.add_parser(cmd).add_argument("input", type=Path)
    sub.add_parser("confirm-schedule").add_argument("automation_id")
    sub.add_parser("audit-history")
    sub.add_parser("sync-records")
    sub.add_parser("review")
    sub.add_parser("pool-status")
    sub.add_parser("pool-ingest").add_argument("input", type=Path)
    sub.add_parser("select-pool").add_argument("day")
    sub.add_parser("context").add_argument("creator_id")
    for cmd in ("begin", "prepare"):
        sub.add_parser(cmd).add_argument("day")
    sub.add_parser("claim-send").add_argument("batch")
    p = sub.add_parser("add")
    p.add_argument("day")
    p.add_argument("input", type=Path)
    p = sub.add_parser("start-today")
    p.add_argument("day")
    p.add_argument("input", type=Path)
    p = sub.add_parser("start-supplement")
    p.add_argument("day")
    p.add_argument("input", type=Path)
    for cmd in ("approve", "record-sent"):
        p = sub.add_parser(cmd)
        p.add_argument("batch")
        p.add_argument("input", type=Path)
    args = parser.parse_args()
    state = mail.State(args.state)
    try:
        initialize(state)
        data = json.loads(args.input.read_text()) if hasattr(args, "input") else None
        if args.command == "status":
            import candidate_pool
            result = {"activation": get(state, "activation", {"status": "paused"}),
                      "days": all_items(state, "day:"), "pending_replies": pending_replies(state),
                      "candidate_pool": candidate_pool.summary(state)}
        elif args.command == "pool-status":
            import candidate_pool
            result = candidate_pool.summary(state)
        elif args.command == "pool-ingest":
            import candidate_pool
            result = candidate_pool.ingest(state, data, args.book)
        elif args.command == "preflight":
            result = preflight(state, book=args.book)
        elif args.command == "activate":
            result = activate(state, data, book=args.book)
        elif args.command == "configure-schedule":
            result = configure_schedule(state, data, book=args.book)
        elif args.command == "confirm-schedule":
            result = confirm_schedule(state, args.automation_id)
        elif args.command == "audit-history":
            mailbox = mail.connect(state.config())
            try:
                result = audit_history(state, mailbox)
            finally:
                mailbox.close()
        elif args.command == "sync-records":
            result = sync_records(state, book=args.book)
        elif args.command == "context":
            result = records.context(state, args.creator_id)
        elif args.command == "review":
            import review_center
            result = review_center.build(state, book=args.book)
        elif args.command == "begin":
            result = day_state(state, args.day)
        elif args.command == "start-today":
            result = start_today(state, args.day, data, book=args.book)
        elif args.command == "start-supplement":
            result = start_supplement(state, args.day, data, book=args.book)
        elif args.command in ("add", "select-pool"):
            import candidate_pool
            if args.command == "add": candidate_pool.ingest(state, data, args.book)
            mailbox = mail.connect(state.config())
            try:
                audit_history(state, mailbox)
                result = candidate_pool.select_day(state, args.day, mailbox, args.book)
            finally:
                mailbox.close()
        elif args.command == "prepare":
            result = prepare_day(state, args.day, book=args.book)
        elif args.command == "summarize":
            result = summarize_replies(state, data)
        elif args.command == "approve":
            result = approve_batch(state, args.batch, data)
        elif args.command == "claim-send":
            result = claim_send(state, args.batch)
        elif args.command == "record-sent":
            result = record_sent(state, args.batch, data, book=args.book)
        if args.command in ("prepare", "summarize", "record-sent", "pool-ingest", "select-pool", "add"):
            import review_center
            result["review_center"] = review_center.refresh_after_operation(state, result, book=args.book)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    finally:
        state.close()


if __name__ == "__main__":
    try:
        main()
    except (mail.DraftError, ValueError, KeyError, OSError) as error:
        print("未执行：" + str(error), file=sys.stderr)
        sys.exit(1)
