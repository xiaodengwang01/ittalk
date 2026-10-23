"""Persistent local research pool and daily selection. Never sends mail."""
from __future__ import annotations

import argparse
from collections import Counter
from copy import deepcopy
from email.utils import getaddresses
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
import daily_workflow as daily

mail = daily.mail
PROTECTED = {"reserved", "selected", "sent_verified"}
LABELS = {"available": "资料齐全，待入选复核", "needs_info": "待补齐资料", "needs_recheck": "资料过期，待复核",
          "held_shared_email": "共享邮箱，待核实归属", "excluded_contacted": "已有记录，排除首联",
          "excluded_quality": "上次筛查未达标，暂缓",
          "reserved": "入选处理中", "selected": "已入选", "sent_verified": "已核验发送", "needs_review": "待核实"}


def initialize(state):
    daily.initialize(state)
    state.db.execute("CREATE TABLE IF NOT EXISTS candidate_pool (id TEXT PRIMARY KEY, profile TEXT UNIQUE NOT NULL, email TEXT NOT NULL, data TEXT NOT NULL)")
    state.db.commit()


def entries(state):
    if not state.db.execute("SELECT 1 FROM sqlite_master WHERE name='candidate_pool'").fetchone(): return []
    return [json.loads(r[0]) for r in state.db.execute("SELECT data FROM candidate_pool ORDER BY id")]


def store(state, item):
    item["updated_at"] = mail.now()
    with state.db:
        state.db.execute("INSERT OR REPLACE INTO candidate_pool VALUES (?,?,?,?)", (item["id"], item["profile"],
                         item.get("email", ""), json.dumps(item, ensure_ascii=False)))


def known_contacts(state, book):
    sheet = daily.workbook.read(book)
    profiles, emails = set(), set()
    for row in sheet["creators"]:
        if row.get("账号链接"):
            try: profiles.add(daily.profile(row["账号链接"]))
            except mail.DraftError: pass
        if row.get("Email"): emails.add(str(row["Email"]).strip().lower())
    emails.update(x["recipient"].lower() for x in daily.drafts_snapshot(state))
    for history in daily.all_items(state, "history:"):
        emails.update(x["email"].lower() for x in history.get("sent_contacts", []) if x.get("email"))
    for day in daily.all_items(state, "day:"):
        for key in ("manual_authorization", "supplement_authorization"):
            emails.update(x["email"].lower() for x in day.get(key, {}).get("sent_contacts", []) if x.get("email"))
    for item in daily.all_items(state, "inbox:"):
        emails.update(a.lower() for _, a in getaddresses([item.get("from", "")]) if a)
    if state.db.execute("SELECT 1 FROM sqlite_master WHERE name='mail_records'").fetchone():
        for record in state.db.execute("SELECT data FROM mail_records"):
            record = json.loads(record[0])
            emails.update(a.lower() for _, a in getaddresses([record.get("from", ""), record.get("to", "")]) if a)
    if state.db.execute("SELECT 1 FROM sqlite_master WHERE name='send_reviews'").fetchone():
        attempted = {r[0] for r in state.db.execute("SELECT draft_id FROM send_attempts")}
        for raw in state.db.execute("SELECT data FROM send_reviews"):
            for entry in json.loads(raw[0])["entries"]:
                if entry["draft_id"] in attempted:
                    emails.update(sum(entry["details"]["envelope"].values(), []))
    return profiles, emails


def validate_research(raw):
    data = deepcopy(raw)
    data.pop("达人ID", None)
    day = daily.timestamp(data["核实时间"]).date().isoformat()
    return daily.validate_candidate(data, day)


def ingest(state, data, book=daily.BOOK):
    initialize(state)
    if not isinstance(data, dict) or not isinstance(data.get("creators", []), list) or any(not isinstance(x, dict) for x in data.get("creators", [])):
        raise mail.DraftError("候选输入须为 creators 对象数组。")
    routes = data.get("routes_checked", [])
    if not isinstance(routes, list) or not set(routes) <= {"FastMoss", "TikTok"}:
        raise mail.DraftError("候选来源只能记录实际检查过的 FastMoss/TikTok。")
    checked = data.get("checked_at")
    if checked and daily.timestamp(checked) > daily.current_time() + daily.timedelta(minutes=5):
        raise mail.DraftError("搜索核对时间不能在未来。")
    profiles, contacted = known_contacts(state, book)
    previous = {x["profile"]: x for x in entries(state)}
    accepted, rejected = [], []
    for raw in data.get("creators", []):
        try: profile = daily.profile(raw.get("账号链接"))
        except (mail.DraftError, TypeError, ValueError) as error:
            rejected.append({"profile": raw.get("账号链接", ""), "reason": str(error)})
            continue
        old = previous.get(profile)
        item = deepcopy(old) if old else {"id": "POOL-" + mail.digest(profile)[:16], "profile": profile,
                                         "first_seen": mail.now(), "history": [], "aliases": []}
        try: email = mail.email_address(raw.get("Email")).lower()
        except mail.DraftError: email = ""
        # Once allocated, new observations do not mutate the frozen recipient or
        # research snapshot. They remain available for a later human review.
        if old and old["status"] in PROTECTED:
            if mail.digest(raw) != mail.digest(old["data"]):
                observations = item.setdefault("observations", [])
                if not observations or observations[-1]["data"] != raw:
                    observations.append({"at": mail.now(), "data": raw})
                    store(state, item)
            accepted.append(item["id"])
            continue
        if old and old.get("data") != raw:
            item["history"].append({"data": old["data"], "status": old["status"], "at": old.get("updated_at")})
        item.update(data=deepcopy(raw), email=email, status="needs_info", reason="资料尚未完整核实。")
        item["data"]["账号链接"] = profile
        item["data"].pop("达人ID", None)
        item["aliases"] = sorted(set(item["aliases"]) | ({email} if email else set()))
        try:
            verified = validate_research(raw)
            item["data"] = verified
            today = daily.current_time().date().isoformat()
            current = daily.timestamp(verified["核实时间"]).date().isoformat() == today
            item.update(status="available" if current else "needs_recheck", reason="入选前仍需实时核对联系历史。" if current else "需在入选当天重新核实邮箱、近期视频及评估依据。")
        except (mail.DraftError, KeyError, ValueError, TypeError) as error:
            item.update(reason=str(error))
            if raw.get("筛查备注"): item["reason"] = "上次筛查：" + str(raw["筛查备注"]) + "；待补齐：" + item["reason"]
        if raw.get("合格") is False and raw.get("不合格原因"):
            item.update(status="excluded_quality", reason=str(raw["不合格原因"]))
        if item.get("ever_contacted") or raw.get("已联系") is True or profile in profiles or set(item["aliases"]) & contacted:
            item["ever_contacted"] = True
            item.update(status="excluded_contacted", reason="主表、草稿或本机联系记录中已存在该账号或邮箱，不进入新首联。")
        store(state, item)
        previous[profile] = item
        accepted.append(item["id"])
    # Keep different accounts separate; a shared address never merges them or
    # lets whichever row arrived first win the outreach slot.
    pool = entries(state)
    for item in pool:
        shared = any(other["profile"] != item["profile"] and set(item["aliases"]) & set(other["aliases"]) for other in pool)
        if shared and item["status"] not in PROTECTED | {"excluded_contacted", "excluded_quality"}:
            item.update(status="held_shared_email", reason="不同账号共用当前或历史邮箱，分别保留，先核实归属。")
            store(state, item)
    # Old research imports cannot pretend that today's two search routes ran.
    if checked:
        stamp, now = daily.timestamp(checked), daily.current_time()
        if stamp.date() == now.date():
            key = "pool-routes:" + stamp.date().isoformat()
            before = daily.get(state, key, {})
            daily.put(state, key, {"routes": sorted(set(before.get("routes", [])) | set(routes)),
                                  "checked_at": checked, "limitations": data.get("limitations", [])})
    return {"ingested": list(dict.fromkeys(accepted)), "rejected": rejected, "pool": summary(state), "book_changed": False, "sent": False}


def rank(item):
    c = item["data"]
    def score(key):
        value = c.get(key, 0)
        return value if isinstance(value, (int, float)) and not isinstance(value, bool) and 1 <= value <= 5 else 0
    return (-score("内容能力(1-5)"), -score("宠物人群匹配度(1-5)"), -score("播放能力(1-5)"),
            c.get("是否低粉高播放") != "是", c.get("是否带货类似产品") != "是", item["first_seen"], item["profile"])


def link_selected(state, value, book):
    rows = daily.workbook.read(book)["creators"]
    for item in entries(state):
        if item["status"] != "reserved" or item.get("selected_day") != value["date"]: continue
        found = [r for r in rows if r["达人ID"] in value["creators"] and r.get("账号链接")
                 and daily.profile(r["账号链接"]) == item["profile"] and str(r.get("Email", "")).lower() == item["email"]]
        if len(found) == 1:
            item.update(status="selected", creator_id=found[0]["达人ID"], batch=value["batch"], reason="已入选并写入原主表，邮件待确认。")
            store(state, item)


def select_day(state, day, mailbox, book=daily.BOOK, template=daily.TEMPLATE):
    initialize(state)
    if mailbox.account.lower() != state.config().get("account", "").lower():
        raise mail.DraftError("候选复核邮箱与本机配置不一致。")
    value = daily.day_state(state, day)
    daily.recover_reserved(state, value, book)
    link_selected(state, value, book)
    if value["prepared"]:
        return {"date": day, "frozen": True, "selection": value["creators"], "limit": daily.DAILY_TARGET,
                "shortfall": max(0, daily.DAILY_TARGET-len(value["creators"])), "sent": False,
                "preparation": daily.prepare_day(state, day, book, template), "pool": summary(state, day)}
    routes = daily.get(state, "pool-routes:" + day, {}).get("routes", [])
    if set(routes) != {"FastMoss", "TikTok"}:
        raise mail.DraftError("尚无当天两路搜索核对记录；保留候选，请如实报告缺失来源。")
    remaining = max(0, daily.DAILY_TARGET-len(value["creators"]))
    profiles, contacted = known_contacts(state, book)
    pool = entries(state)
    pending = [x for x in pool if x["status"] == "reserved" and x.get("selected_day") == day]
    candidates = pending if pending else sorted([x for x in pool if x["status"] in {"available", "needs_recheck", "needs_info"}], key=rank)
    chosen = []
    for item in candidates:
        if len(chosen) >= remaining: break
        planned = item.get("planned_day") or item["data"].get("计划批次日期", "")
        if planned and planned > day:
            continue
        if item["profile"] in profiles or set(item["aliases"]) & contacted:
            item.update(status="excluded_contacted", ever_contacted=True, reason="入选前发现已有账号或本机联系记录。")
            store(state, item); continue
        if any(other["profile"] != item["profile"] and set(other["aliases"]) & set(item["aliases"]) for other in pool):
            item.update(status="held_shared_email", reason="共享邮箱归属未核实，保留两个账号。")
            store(state, item); continue
        try:
            verified = daily.validate_candidate(item["data"], day)
        except (mail.DraftError, KeyError, ValueError, TypeError) as error:
            item.update(status="needs_recheck", reason=str(error)); store(state, item); continue
        conflicts = []
        for email in item["aliases"]:
            conflicts.extend(mailbox.history(email))
            conflicts.extend({"folder": "Drafts", "uid": uid} for uid in mailbox.existing_drafts(email))
        if conflicts:
            item.update(status="excluded_contacted", ever_contacted=True, reason="实际邮箱中已有收件、已发送或草稿记录。", history_conflicts=conflicts)
            store(state, item); continue
        checked = daily.current_time().isoformat()
        verified.update(已联系=False, 历史核对时间=checked, 历史核对依据="本批入选前只读核对当前邮箱收件、已发送、草稿；所有已知邮箱均无联系记录。")
        item["data"] = verified
        chosen.append(item)
    if chosen:
        # Persist the exact reservation before any workbook write. Interrupted
        # file replacement is recovered via daily_workflow's existing journal.
        with state.db:
            for item in chosen:
                item.update(status="reserved", selected_day=day, selection_basis=mail.digest(item["data"]),
                            reason="已按内容、受众匹配、近期播放及加分项排序并完成当天联系历史复核。")
                state.db.execute("UPDATE candidate_pool SET data=? WHERE id=?", (json.dumps(item, ensure_ascii=False), item["id"]))
        daily.add_candidates(state, day, {"routes_checked": routes, "creators": [x["data"] for x in chosen]}, book)
        value = daily.day_state(state, day)
        link_selected(state, value, book)
    preparation = daily.prepare_day(state, day, book, template)
    return {"date": day, "limit": daily.DAILY_TARGET, "selection": value["creators"],
            "shortfall": max(0, daily.DAILY_TARGET-len(value["creators"])), "preparation": preparation,
            "pool": summary(state, day), "sent": False}


def summary(state, day=None):
    day = day or daily.current_time().date().isoformat()
    result = []
    for item in entries(state):
        status = item["status"]
        if status == "available":
            try:
                if daily.timestamp(item["data"]["核实时间"]).date().isoformat() != day: status = "needs_recheck"
            except (KeyError, ValueError): status = "needs_info"
        result.append({"id": item["id"], "profile": item["profile"], "name": item["data"].get("达人昵称") or "@" + item["profile"].rsplit("@", 1)[-1],
                       "email": item["email"], "status": status, "label": LABELS[status], "reason": item.get("reason", ""),
                       "checked_at": item["data"].get("核实时间") or item["data"].get("历史筛查日期", ""), "creator_id": item.get("creator_id", ""),
                       "planned_day": item.get("planned_day") or item["data"].get("计划批次日期", ""),
                       "selected_day": item.get("selected_day", "")})
    return {"limit": daily.DAILY_TARGET, "total": len(result), "counts": dict(Counter(x["status"] for x in result)),
            "available_today": sum(x["status"] == "available" and (not x["planned_day"] or x["planned_day"] <= day) for x in result), "items": result}


def mark_sent(state, creator_id, event):
    if event.get("verification") != "sent_folder" or not all(event.get(k) for k in ("message_id", "sent_at")):
        raise mail.DraftError("候选池只能记录已核实的实际发送事实。")
    recipient = (event.get("to") or event.get("recipient", "")).lower()
    for item in entries(state):
        if item.get("creator_id") == creator_id and item["email"] == recipient:
            item.update(status="sent_verified", ever_contacted=True, sent_at=event["sent_at"],
                        message_id=event["message_id"], reason="已核验实际发送；后续不再进入首联候选。")
            store(state, item)


def research(state, limit=10):
    if not 1 <= limit <= 20: raise mail.DraftError("每次复核候选数量须在 1–20 之间。")
    pool = sorted([x for x in entries(state) if x["status"] in {"available", "needs_recheck", "needs_info"}], key=rank)
    return {"candidates": pool[:limit], "remaining": max(0, len(pool)-limit),
            "next_step": "按公开来源复核近期普通视频、英语市场、邮箱和分类，再用当天核实时间入池；不能只修改日期沿用旧证据。"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--state", type=Path, default=mail.DEFAULT_STATE)
    parser.add_argument("--book", type=Path, default=daily.BOOK)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("status")
    sub.add_parser("ingest").add_argument("input", type=Path)
    sub.add_parser("select").add_argument("day")
    sub.add_parser("research").add_argument("--limit", type=int, default=10)
    args = parser.parse_args()
    state, mailbox = mail.State(args.state), None
    try:
        initialize(state)
        if args.command == "status": result = summary(state)
        elif args.command == "research": result = research(state, args.limit)
        elif args.command == "ingest": result = ingest(state, json.loads(args.input.read_text()), args.book)
        else:
            mailbox = mail.connect(state.config())
            daily.audit_history(state, mailbox)
            result = select_day(state, args.day, mailbox, args.book)
        if args.command not in ("status", "research"):
            import review_center
            result["review_center"] = review_center.refresh_after_operation(state, result, args.book)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    finally:
        if mailbox: mailbox.close()
        state.close()


if __name__ == "__main__":
    try: main()
    except (mail.DraftError, mail.CredentialError, OSError, ValueError, KeyError, mail.imaplib.IMAP4.error) as error:
        print("候选池操作未完成：" + (str(error) if isinstance(error, mail.DraftError) else "已保留本机进度，请核对资料或连接后继续。"), file=sys.stderr)
        sys.exit(1)
