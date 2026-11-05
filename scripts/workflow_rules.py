"""Pure local rules for the creator workbook. No mailbox access or email sending.

Public functions mutate the supplied Chinese-field dictionaries/lists in place.
Dates accept datetime/date or ISO strings; missing dates never imply "today".
The workbook adapter is responsible for persisting returned conflict notes.
"""

from __future__ import annotations

import hashlib
import json
import re
from datetime import date, datetime, time, timedelta, timezone
from decimal import Decimal, InvalidOperation
from urllib.parse import unquote, urlsplit, urlunsplit


NEXT_ACTIONS = ("补资料", "核实分类", "首联草稿", "回复草稿", "跟进草稿", "等待对方", "暂停联系")
EMAIL_FIELDS = (
    "邮件记录ID", "达人ID", "方向", "邮件时间", "发件人", "收件人", "主题", "正文或草稿",
    "原邮件定位", "邮箱消息ID", "会话ID", "处理状态", "模板及版本", "草稿版本", "草稿指纹",
    "发送确认批次", "确认时间", "确认指纹", "实际发送结果", "记录来源",
)
FOLLOWUP_TEMPLATE = (
    "Hi {name},\n\nI wanted to follow up on my earlier email about a potential "
    "collaboration. Would you be interested in learning more?\n\nBest regards"
)
_SHANGHAI = timezone(timedelta(hours=8))
_PLATFORMS = {
    "youtube": "YouTube", "yt": "YouTube", "油管": "YouTube",
    "instagram": "Instagram", "ig": "Instagram", "ins": "Instagram",
    "tiktok": "TikTok", "tik tok": "TikTok", "抖音国际版": "TikTok",
    "twitter": "X", "x": "X", "facebook": "Facebook", "fb": "Facebook",
}
_EMAIL_RE = re.compile(
    r"^[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+)*@"
    r"(?:[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+"
    r"[A-Za-z](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?$"
)
_BLOCKED_CONTACT_STATES = {"暂停联系", "已拒绝", "不合作", "拒绝", "勿联系", "已完成"}


def _is_metric_field(field):
    return field in {"粉丝数", "近10条中位播放", "近10条最高播放"} or "粉丝" in field or "播放量" in field


def _contact_blocked(creator):
    return creator.get("下一步动作") == "暂停联系" or creator.get("联系状态") in _BLOCKED_CONTACT_STATES


def _text(value):
    return "" if value is None else str(value).strip()


def normalize_platform(value):
    text = _text(value)
    return _PLATFORMS.get(text.lower(), text)


def canonical_account_url(platform, value):
    """Remove tracking noise and normalize known social profile URL aliases."""
    raw = _text(value)
    if not raw:
        return ""
    if raw.startswith("@"):
        bases = {"YouTube": "www.youtube.com/", "Instagram": "www.instagram.com/",
                 "TikTok": "www.tiktok.com/", "X": "x.com/"}
        base = bases.get(normalize_platform(platform))
        if not base:
            return ""
        raw = base + (raw if normalize_platform(platform) in ("YouTube", "TikTok") else raw[1:])
    if "://" not in raw:
        raw = "https://" + raw
    try:
        parts = urlsplit(raw)
        host = (parts.hostname or "").lower()
        if parts.scheme.lower() not in ("http", "https") or not host or "." not in host:
            return ""
        if parts.username or parts.password or parts.port not in (None, 80, 443):
            return ""
    except ValueError:
        return ""
    host = host.removeprefix("www.").removeprefix("m.")
    host = {"mobile.twitter.com": "x.com", "twitter.com": "x.com"}.get(host, host)
    path = re.sub(r"/+", "/", unquote(parts.path)).rstrip("/")
    if host in ("instagram.com", "tiktok.com", "x.com", "facebook.com"):
        path = path.lower()
    if host == "youtube.com":
        if path.startswith("/@"):
            path = path.lower()
        # Profile tabs are the same account; channel IDs remain case sensitive.
        path = re.sub(r"/(videos|shorts|streams|featured|about|community|releases)$", "", path)
    # Keep non-tracking query identifiers for profile.php-style account URLs.
    query = ""
    if host == "facebook.com" and path.lower() == "/profile.php":
        from urllib.parse import parse_qs, urlencode
        identifiers = parse_qs(parts.query).get("id")
        query = urlencode({"id": identifiers[0]}) if identifiers else ""
    return urlunsplit(("https", host, path, query, ""))


def is_valid_email(value):
    """A single plain email address; URLs, display names and lists are rejected."""
    text = _text(value)
    return bool(len(text) <= 254 and text.count("@") == 1
                and len(text.split("@", 1)[0]) <= 64 and _EMAIL_RE.fullmatch(text))


def parse_metric(value):
    """Return a numeric count, or None when a count is missing/ambiguous."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        raw = str(value)
    else:
        raw = _text(value).lower().replace(",", "").replace("，", "").replace(" ", "")
    match = re.fullmatch(r"(\d+(?:\.\d+)?)(k|m|b|万|亿)?", raw)
    if not match:
        return None
    try:
        count = Decimal(match.group(1)) * {None: 1, "k": 1000, "m": 1000000,
                                           "b": 1000000000, "万": 10000, "亿": 100000000}[match.group(2)]
    except InvalidOperation:
        return None
    if not count.is_finite():
        return None
    return int(count) if count == count.to_integral_value() else float(count)


def _datetime(value):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        result = value
    elif isinstance(value, date):
        result = datetime.combine(value, time())
    elif isinstance(value, str):
        try:
            result = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
        except ValueError:
            return None
    else:
        return None
    return result.astimezone(_SHANGHAI).replace(tzinfo=None) if result.tzinfo else result


def _now(value=None):
    result = _datetime(value) if value is not None else datetime.now(_SHANGHAI).replace(tzinfo=None)
    if result is None:
        raise ValueError("需要有效的日期或时间")
    return result


def _timestamp_text(value):
    """Preserve date-only source precision instead of inventing midnight."""
    parsed = _datetime(value)
    if parsed is None:
        return ""
    if (isinstance(value, date) and not isinstance(value, datetime)) or (
            isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip())):
        return parsed.date().isoformat()
    return parsed.isoformat(sep=" ", timespec="seconds")


def add_workdays(value, days=7):
    """Exclude the starting day; workdays are Monday through Friday."""
    parsed = _datetime(value)
    if parsed is None:
        return None
    if not isinstance(days, int) or days < 0:
        raise ValueError("工作日数量必须是非负整数")
    remaining = days
    while remaining:
        parsed += timedelta(days=1)
        if parsed.weekday() < 5:
            remaining -= 1
    return parsed.date()


def _next_id(records, field, prefix, width):
    used = {_text(record.get(field)) for record in records}
    numbers = [int(value[len(prefix):]) for value in used
               if value.startswith(prefix) and value[len(prefix):].isdigit()]
    number = max(numbers, default=0) + 1
    while f"{prefix}{number:0{width}d}" in used:
        number += 1
    return f"{prefix}{number:0{width}d}"


def ensure_creator_ids(creators):
    """Reuse existing 编号, and allocate immutable IDs only to missing IDs."""
    used = set()
    for creator in creators:
        identifier = _text(creator.get("达人ID")) or _text(creator.get("编号"))
        if identifier and identifier in used:
            raise ValueError(f"达人ID重复：{identifier}")
        if identifier:
            creator["达人ID"] = identifier
            used.add(identifier)
    for creator in creators:
        if not _text(creator.get("达人ID")):
            creator["达人ID"] = _next_id(creators, "达人ID", "KOL-", 6)
        if "编号" in creator:
            creator["编号"] = creator["达人ID"]
    return creators


def _creator_key(creator):
    platform = normalize_platform(creator.get("平台"))
    url = canonical_account_url(platform, creator.get("账号链接"))
    return (platform, url) if platform and url else None


def merge_duplicate_creators(creators, messages):
    """Merge only identical platform/profile accounts; retain conflicts explicitly."""
    ensure_creator_ids(creators)
    by_key, kept, id_map, conflicts = {}, [], {}, []
    for creator in creators:
        key = _creator_key(creator)
        if key is None or key not in by_key:
            kept.append(creator)
            if key:
                by_key[key] = creator
            continue
        target = by_key[key]
        id_map[creator["达人ID"]] = target["达人ID"]
        for field, value in creator.items():
            if field in ("达人ID", "编号") or value in (None, ""):
                continue
            if target.get(field) in (None, ""):
                target[field] = value
            elif target[field] != value:
                conflicts.append({"达人ID": target["达人ID"], "原达人ID": creator["达人ID"],
                                  "字段": field, "保留值": target[field], "待核实值": value})
    creators[:] = kept
    for message in messages:
        if message.get("达人ID") in id_map:
            message["达人ID"] = id_map[message["达人ID"]]
    return {"merged": len(id_map), "id_map": id_map, "conflicts": conflicts}


def merge_creator_candidates(creators, candidates, event_time=None):
    """Merge scrape candidates without replacing existing human-owned information.

    Only an explicit valid scrape event_time plus a source URL writes provenance.
    Existing contacts, classification and arbitrary user fields are never replaced.
    Counts can refresh on a dated scrape; conflicts are returned for review.
    """
    # Validate the whole batch before allocating IDs or mutating any records.
    normalized = []
    timestamp = _datetime(event_time)
    if event_time is not None and timestamp is None:
        raise ValueError("抓取事件时间无效")
    for source in candidates:
        candidate = dict(source)
        candidate.pop("达人ID", None)
        candidate.pop("编号", None)
        supplied_source = candidate.pop("来源链接", None)
        candidate.pop("最近抓取时间", None)
        candidate["平台"] = normalize_platform(candidate.get("平台"))
        candidate["账号链接"] = canonical_account_url(candidate["平台"], candidate.get("账号链接"))
        if not _creator_key(candidate):
            raise ValueError("抓取候选必须提供平台和有效账号链接；本批次未导入")
        if candidate.get("Email") and not is_valid_email(candidate["Email"]):
            value = _text(candidate["Email"])
            if canonical_account_url("", value) and "@" not in value:
                candidate.setdefault("联系页链接", value)
            candidate["Email"] = ""
        for field in tuple(candidate):
            if _is_metric_field(field):
                parsed = parse_metric(candidate[field])
                if parsed is not None:
                    candidate[field] = parsed
        normalized.append((candidate, supplied_source))
    ensure_creator_ids(creators)
    created, updated, conflicts = [], [], []
    index = {_creator_key(item): item for item in creators if _creator_key(item)}
    for candidate, supplied_source in normalized:
        key = _creator_key(candidate)
        target = index.get(key) if key else None
        if target is None:
            target = {"达人ID": _next_id(creators, "达人ID", "KOL-", 6), **candidate}
            target.setdefault("分类核实状态", "初判" if target.get("首联内容分类") not in (None, "", "待核实") else "待核实")
            target["联系状态"] = "待联系"
            target["下一步动作"] = "补资料" if not is_valid_email(target.get("Email")) else "核实分类"
            creators.append(target)
            created.append(target["达人ID"])
            if key:
                index[key] = target
        else:
            for field, value in candidate.items():
                if value in (None, ""):
                    continue
                if target.get(field) in (None, ""):
                    target[field] = value
                elif target[field] != value:
                    if timestamp and _is_metric_field(field) and isinstance(value, (int, float)):
                        target[field] = value
                    else:
                        conflicts.append({"达人ID": target["达人ID"], "字段": field,
                                          "保留值": target[field], "待核实值": value})
            updated.append(target["达人ID"])
        if timestamp and canonical_account_url("", supplied_source):
            target["来源链接"] = _text(supplied_source)
            target["最近抓取时间"] = _timestamp_text(event_time)
    return {"created": created, "updated": updated, "conflicts": conflicts}


def _hash(fields):
    return hashlib.sha256(json.dumps(fields, ensure_ascii=False, sort_keys=True,
                                     default=str, separators=(",", ":")).encode("utf-8")).hexdigest()


def _mail_fingerprint(message):
    fields = ["方向", "邮件时间", "发件人", "收件人", "主题", "正文或草稿"]
    # A partial subject/date is not a complete mail identity.
    if any(message.get(field) in (None, "") for field in fields):
        return None
    payload = {field: _text(message.get(field)) for field in fields}
    parsed = _datetime(message.get("邮件时间"))
    if parsed is None:
        return None
    payload["邮件时间"] = _timestamp_text(message.get("邮件时间"))
    payload["发件人"] = payload["发件人"].lower()
    payload["收件人"] = payload["收件人"].lower()
    return _hash(payload)


def _same_message(left, right):
    a, b = _text(left.get("邮箱消息ID")), _text(right.get("邮箱消息ID"))
    if a and b:
        # Distinct provider IDs must never collapse even if subjects/bodies match.
        return a == b and _text(left.get("邮箱账户")).lower() == _text(right.get("邮箱账户")).lower()
    a, b = _text(left.get("原邮件定位")), _text(right.get("原邮件定位"))
    if a and b:
        return a == b
    fingerprint = _mail_fingerprint(left)
    return bool(fingerprint and fingerprint == _mail_fingerprint(right))


def _resolve_creator(creators, messages, incoming):
    known = {item["达人ID"]: item for item in creators}
    candidates = []
    identifier = _text(incoming.get("达人ID"))
    if identifier:
        if identifier not in known:
            return None
        candidates.append({identifier})
    key = _creator_key(incoming)
    if key:
        matches = {item["达人ID"] for item in creators if _creator_key(item) == key}
        if not matches:
            return None
        candidates.append(matches)
    thread = _text(incoming.get("会话ID"))
    if thread:
        matches = {item.get("达人ID") for item in messages
                   if _text(item.get("会话ID")) == thread and item.get("达人ID") in known
                   and _text(item.get("邮箱账户")).lower() == _text(incoming.get("邮箱账户")).lower()}
        if matches:
            candidates.append(matches)
    if candidates:
        intersection = set.intersection(*candidates)
        return next(iter(intersection)) if len(intersection) == 1 else None
    address = _text(incoming.get("发件人" if incoming.get("方向") == "收信" else "收件人")).lower()
    matches = {item["达人ID"] for item in creators
               if is_valid_email(address) and _text(item.get("Email")).lower() == address}
    return next(iter(matches)) if len(matches) == 1 else None


def _clear_confirmation(record):
    for field in ("发送确认批次", "确认时间", "确认指纹"):
        record[field] = ""


def _is_followup(record):
    return record.get("草稿用途") == "跟进" or record.get("邮件类型") == "跟进" or _text(record.get("模板及版本")).startswith("跟进")


def _is_initial(record):
    return record.get("草稿用途") == "首联" or record.get("邮件类型") == "首联" or _text(record.get("模板及版本")).startswith("首联")


def _set_latest(creator, field, timestamp):
    parsed, old = _datetime(timestamp), _datetime(creator.get(field))
    if parsed and (not old or parsed > old):
        creator[field] = _timestamp_text(timestamp)


def import_message(creators, messages, incoming):
    """Import once, resolve ownership conservatively, and invalidate reminders."""
    ensure_creator_ids(creators)
    record = {field: "" for field in EMAIL_FIELDS}
    record.update(dict(incoming))
    record["邮箱消息ID"] = _text(record.get("邮箱消息ID")) or _text(incoming.get("provider_message_id"))
    record["会话ID"] = _text(record.get("会话ID")) or _text(incoming.get("thread_id"))
    if record.get("方向") not in ("收信", "发信"):
        raise ValueError("方向必须是收信或发信")
    existing = next((item for item in messages if _same_message(item, record)), None)
    if existing:
        return {"imported": False, "record": existing, "ambiguous": existing.get("处理状态") == "待人工归属"}
    record["邮件记录ID"] = _next_id(messages, "邮件记录ID", "MAIL-", 6)
    identifier = _resolve_creator(creators, messages, record)
    record["达人ID"] = identifier or ""
    record["记录来源"] = record.get("记录来源") or "邮件导入"
    if not identifier:
        record["处理状态"] = "待人工归属"
    elif record["方向"] == "收信":
        record["处理状态"] = "待回复"
        creator = next(item for item in creators if item["达人ID"] == identifier)
        _set_latest(creator, "最近收信时间", record.get("邮件时间"))
        creator["联系状态"] = "已回复"
        if creator.get("下一步动作") != "暂停联系":
            creator["下一步动作"] = "回复草稿"
        creator["计划跟进日期"] = ""
        for draft in messages:
            if draft.get("达人ID") == identifier and _is_followup(draft) and draft.get("实际发送结果") not in ("成功", "结果待核实"):
                draft["处理状态"] = "已作废"
                _clear_confirmation(draft)
    elif record.get("实际发送结果") == "成功":
        creator = next(item for item in creators if item["达人ID"] == identifier)
        _set_latest(creator, "最近发信时间", record.get("邮件时间"))
        record["处理状态"] = "已发送"
    messages.append(record)
    return {"imported": True, "record": record, "ambiguous": not bool(identifier)}


def draft_fingerprint(record):
    return _hash({field: _text(record.get(field)) for field in
                  ("邮件记录ID", "达人ID", "收件人", "主题", "正文或草稿", "草稿版本")})


def create_draft(creators, messages, creator_id, subject, body, template_version,
                 purpose="首联", in_reply_to=""):
    """Create a local unsent draft; its creation time is not a mail send time."""
    ensure_creator_ids(creators)
    creator = next((item for item in creators if item["达人ID"] == creator_id), None)
    if creator is None:
        raise ValueError("未找到达人ID")
    if purpose not in ("首联", "回复", "跟进"):
        raise ValueError("草稿用途必须为首联、回复或跟进")
    if not is_valid_email(creator.get("Email")):
        raise ValueError("需要有效的单一邮箱")
    if _contact_blocked(creator):
        raise ValueError("暂停、拒绝、勿联系或已完成的达人不能生成草稿")
    if not _text(subject) or not _text(body):
        raise ValueError("需要完整主题和正文")
    related = [item for item in messages if item.get("达人ID") == creator_id]
    if any(item.get("实际发送结果") == "结果待核实" for item in related):
        raise ValueError("存在发送结果待核实的邮件，不能重新生成发信草稿")
    if purpose == "首联":
        if creator.get("联系状态") not in ("待联系", "未联系") or creator.get("最近发信时间"):
            raise ValueError("只有未联系的达人可以生成首联草稿")
        if creator.get("分类核实状态") == "待核实" or creator.get("首联内容分类") in (None, "", "待核实"):
            raise ValueError("首联分类尚待核实，先补充分类再生成首联草稿")
        if any(_is_initial(item) or (item.get("方向") == "发信" and (
                item.get("实际发送结果") == "成功" or item.get("处理状态") == "已发送" or item.get("邮件时间")))
               for item in related):
            raise ValueError("已有首联或发信记录，不能重复生成首联草稿")
    if purpose == "跟进" and any(_is_followup(item) for item in related):
        raise ValueError("已经存在跟进记录，不重复生成")
    original = None
    if in_reply_to:
        original = next((item for item in related if item.get("邮件记录ID") == in_reply_to), None)
        if original is None:
            raise ValueError("回复记录不存在或不属于当前达人")
    if purpose == "回复":
        existing = [item for item in related if item.get("方向") == "发信" and item.get("草稿用途") == "回复"
                    and _text(item.get("回复邮件记录ID")) == _text(in_reply_to)
                    and item.get("处理状态") in ("待确认草稿", "已确认待发送", "发送失败")
                    and item.get("实际发送结果") not in ("成功", "结果待核实")]
        if existing:
            draft = existing[-1]
            for obsolete in existing[:-1]:
                obsolete["处理状态"] = "已作废"
                _clear_confirmation(obsolete)
            update_draft(draft, **{"收件人": _text(creator.get("Email")), "主题": _text(subject),
                                  "正文或草稿": _text(body), "模板及版本": template_version})
            creator["下一步动作"] = "回复草稿"
            return draft
    if purpose == "跟进":
        generated = generate_followup_drafts([creator], messages, template=body)
        if not generated:
            raise ValueError("尚未到跟进日期、已收到回复或缺少已确认首联记录")
        draft = generated[0]
        draft.update({"主题": _text(subject), "正文或草稿": _text(body), "模板及版本": template_version})
        if original:
            draft["会话ID"] = original.get("会话ID", "")
            draft["回复邮件记录ID"] = in_reply_to
        draft["草稿指纹"] = draft_fingerprint(draft)
        return draft
    draft = {field: "" for field in EMAIL_FIELDS}
    draft.update({"邮件记录ID": _next_id(messages, "邮件记录ID", "MAIL-", 6),
                  "达人ID": creator_id, "方向": "发信", "收件人": _text(creator.get("Email")),
                  "主题": _text(subject), "正文或草稿": _text(body), "模板及版本": template_version,
                  "处理状态": "待确认草稿", "草稿版本": 1, "草稿用途": purpose,
                  "记录来源": "本地草稿生成", "草稿生成时间": _now().isoformat(sep=" ", timespec="seconds")})
    if original:
        draft["会话ID"] = original.get("会话ID", "")
        draft["回复邮件记录ID"] = in_reply_to
    draft["草稿指纹"] = draft_fingerprint(draft)
    messages.append(draft)
    creator["下一步动作"] = {"首联": "首联草稿", "回复": "回复草稿", "跟进": "跟进草稿"}[purpose]
    return draft


def confirmation_valid(record):
    return bool(record.get("发送确认批次") and _datetime(record.get("确认时间"))
                and record.get("处理状态") == "已确认待发送"
                and is_valid_email(record.get("收件人"))
                and record.get("实际发送结果") not in ("成功", "结果待核实")
                and record.get("草稿指纹") == record.get("确认指纹") == draft_fingerprint(record))


def update_draft(record, **changes):
    if record.get("实际发送结果") in ("成功", "结果待核实"):
        raise ValueError("已发送或结果待核实的记录不可修改为新草稿")
    allowed = {"收件人", "主题", "正文或草稿", "模板及版本", "达人ID"}
    if not set(changes).issubset(allowed):
        raise ValueError("只能修改草稿的收件人、主题、正文、模板或达人归属")
    changed = any(record.get(field) != value for field, value in changes.items())
    if changed:
        record.update(changes)
        record["草稿版本"] = int(record.get("草稿版本") or 0) + 1
        record["草稿指纹"] = draft_fingerprint(record)
        record["处理状态"] = "待确认草稿"
        record["实际发送结果"] = ""
        _clear_confirmation(record)
    return record


def confirm_batch(messages, record_ids, batch_id, confirmed_at=None):
    """Persist a user-confirmed batch; the caller must already hold authorization."""
    identifiers = list(record_ids)
    if not identifiers or len(set(identifiers)) != len(identifiers) or not _text(batch_id):
        raise ValueError("需要非空批次和不重复的草稿ID")
    selected = [item for identifier in identifiers for item in messages if item.get("邮件记录ID") == identifier]
    if len(selected) != len(identifiers):
        raise ValueError("未找到唯一的草稿记录")
    for record in selected:
        if record.get("方向") != "发信" or record.get("处理状态") not in ("待确认草稿", "已确认待发送", "发送失败"):
            raise ValueError("批次包含不可确认的记录")
        if record.get("实际发送结果") in ("成功", "结果待核实") or not is_valid_email(record.get("收件人")):
            raise ValueError("收件人无效或发送状态需要先核实")
        if not _text(record.get("主题")) or not _text(record.get("正文或草稿")):
            raise ValueError("确认前需要完整主题和正文")
    stamp = _now(confirmed_at).isoformat(sep=" ", timespec="seconds")
    for record in selected:
        old_fingerprint = record.get("草稿指纹")
        if old_fingerprint and old_fingerprint != draft_fingerprint(record):
            record["草稿版本"] = int(record.get("草稿版本") or 0) + 1
        else:
            record["草稿版本"] = int(record.get("草稿版本") or 1)
        record["草稿指纹"] = draft_fingerprint(record)
        record["发送确认批次"] = _text(batch_id)
        record["确认时间"] = stamp
        record["确认指纹"] = record["草稿指纹"]
        record["处理状态"] = "已确认待发送"
    return selected


def generate_followup_drafts(creators, messages, today=None, template=None):
    """Make at most one reminder for a confirmed successful initial contact."""
    ensure_creator_ids(creators)
    day = _now(today).date()
    added = []
    for creator in creators:
        if _contact_blocked(creator) or creator.get("联系状态") == "已回复":
            continue
        if not is_valid_email(creator.get("Email")):
            continue
        related = [item for item in messages if item.get("达人ID") == creator["达人ID"]]
        if any(item.get("实际发送结果") == "结果待核实" for item in related):
            continue
        sent = [item for item in related if item.get("方向") == "发信"
                and item.get("实际发送结果") == "成功"
                and item.get("草稿用途", item.get("邮件类型")) == "首联"
                and item.get("发送确认批次") and _datetime(item.get("确认时间"))
                and item.get("确认指纹") == draft_fingerprint(item)
                and _datetime(item.get("邮件时间"))]
        if not sent:
            continue
        initial = min(sent, key=lambda item: _datetime(item["邮件时间"]))
        initial_time = _datetime(initial["邮件时间"])
        received = [item for item in related if item.get("方向") == "收信"]
        last_received = _datetime(creator.get("最近收信时间"))
        if (last_received and last_received >= initial_time) or any(
                _datetime(item.get("邮件时间")) is None or _datetime(item["邮件时间"]) >= initial_time for item in received):
            continue
        due = add_workdays(initial_time, 7)
        creator["计划跟进日期"] = due.isoformat()
        # An existing second successful outbound (including legacy history) is
        # evidence that the single reminder opportunity was already consumed.
        second_sent = any(item is not initial and item.get("方向") == "发信"
                          and (item.get("实际发送结果") == "成功" or item.get("处理状态") == "已发送")
                          for item in related)
        if day < due or second_sent or any(_is_followup(item) for item in related):
            continue
        address = _text(creator.get("Email"))
        original_recipient = _text(initial.get("收件人"))
        if original_recipient and original_recipient.lower() != address.lower():
            # A changed contact needs a human decision before using an old conversation.
            creator["下一步动作"] = "补资料"
            continue
        name = _text(creator.get("昵称") or creator.get("达人名称") or creator.get("账号名称")) or "there"
        body = (template or FOLLOWUP_TEMPLATE).replace("{name}", name)
        original_subject = _text(initial.get("主题"))
        draft = {field: "" for field in EMAIL_FIELDS}
        draft.update({"邮件记录ID": _next_id(messages, "邮件记录ID", "MAIL-", 6),
                      "达人ID": creator["达人ID"], "方向": "发信", "收件人": address,
                      "发件人": initial.get("发件人", ""),
                      "主题": original_subject if original_subject.lower().startswith("re:") else "Re: " + original_subject,
                      "正文或草稿": body, "会话ID": initial.get("会话ID", ""),
                      "处理状态": "待确认草稿", "模板及版本": "跟进/v1", "草稿版本": 1,
                      "记录来源": "本地规则生成", "邮件类型": "跟进",
                      "草稿用途": "跟进",
                      "首联邮件记录ID": initial["邮件记录ID"], "草稿生成时间": _now(today).isoformat(sep=" ", timespec="seconds")})
        draft["草稿指纹"] = draft_fingerprint(draft)
        messages.append(draft)
        added.append(draft)
        creator["下一步动作"] = "跟进草稿"
    return added


def record_send_result(creators, messages, record_id, result, event_time=None):
    """Record an external observed outcome; never performs a send or retry."""
    if result not in ("成功", "失败", "结果待核实"):
        raise ValueError("实际发送结果必须为成功、失败或结果待核实")
    matches = [item for item in messages if item.get("邮件记录ID") == record_id]
    if len(matches) != 1:
        raise ValueError("未找到唯一的邮件记录")
    record = matches[0]
    previous = record.get("实际发送结果")
    if previous == "成功":
        if result != "成功":
            raise ValueError("已成功发送的记录不能改写为其他结果")
        return record
    # Resolving an unknown outcome still refers to the originally confirmed
    # content. An Excel edit must never become evidence for another sent body.
    if previous == "结果待核实":
        if not (record.get("发送确认批次") and _datetime(record.get("确认时间"))
                and record.get("草稿指纹") == record.get("确认指纹") == draft_fingerprint(record)):
            raise ValueError("待核实记录与原确认版本不一致，保留待核实状态")
    elif not confirmation_valid(record):
        raise ValueError("草稿缺少对应版本的有效批次确认")
    if previous == "结果待核实" and result == "结果待核实":
        return record
    stamp = _timestamp_text(event_time) if event_time is not None else _timestamp_text(_now())
    if not stamp:
        raise ValueError("需要有效的发送结果时间")
    record["实际发送结果"] = result
    record["处理状态"] = {"成功": "已发送", "失败": "发送失败", "结果待核实": "发送结果待核实"}[result]
    if result == "成功":
        record["邮件时间"] = stamp
        creator = next((item for item in creators if item.get("达人ID") == record.get("达人ID")), None)
        if creator:
            _set_latest(creator, "最近发信时间", stamp)
            creator["联系状态"] = "已联系"
            creator["下一步动作"] = "等待对方"
            creator["计划跟进日期"] = "" if _is_followup(record) else add_workdays(stamp, 7).isoformat()
    elif result == "失败":
        _clear_confirmation(record)
    return record
