"""Local, text-only mail archive and bounded per-creator reads.

No mailbox or sending API is used here. Public functions accept Chinese-field
dicts and mutate only the supplied lists/dicts. Drafts belong to the conversation,
never to this archive. A hidden spreadsheet column is not an access boundary:
``overview`` and ``context`` explicitly select what a reader receives.
"""
from __future__ import annotations

from datetime import date, datetime, time, timezone
import hashlib
import json
import re
from uuid import uuid4
from urllib.parse import parse_qs, urlsplit

try:
    from . import workflow_rules as rules
    from . import main_schema
except ImportError:
    import workflow_rules as rules
    import main_schema


ARCHIVE_FIELDS = (
    "邮件记录ID", "达人ID", "达人昵称", "方向", "邮件时间", "主题", "邮件摘要", "处理状态",
    "正文完整性", "邮件正文", "发件人", "收件人", "原邮件定位", "邮件原始时间", "核实说明",
    "记录来源", "邮箱消息ID", "会话ID", "摘要依据指纹", "摘要状态", "邮件用途", "实际发送结果",
    "发送确认批次", "确认时间", "确认指纹", "回复邮件记录ID", "邮箱账户",
)
_MAIL_SUMMARY_FIELDS = tuple(f for f in ARCHIVE_FIELDS
                             if f not in {"邮件摘要", "摘要依据指纹", "摘要状态"})
_CREATOR_SUMMARY_FIELDS = (
    "达人ID", "达人昵称", "联系状态", "计划跟进日期", "最近联系", "最近收信时间", "最近发信时间",
    "对方报价(原币)", "报价币种", "报价口径",
)
_STALE_SUMMARY_PREFIX = "【摘要待更新】"
_INCOMPLETE = {"部分", "待核实", "缺失"}
_DRAFT_STATES = {"待确认草稿", "已确认待发送", "草稿", "已作废", "发送失败"}
_INITIAL_PURPOSES = {"首联", "初次建联", "首次联系", "初次联系"}
_REPLY_PURPOSES = {"回复", "跟进", "后续回复"}
_HEADER = re.compile(r"^\s*(?:>\s*)?(From|Sent|Date|To|Subject|Cc|发件人|发送时间|日期|收件人|主题|抄送)\s*[:：]", re.I)
_HEADER_NAME = {"from": "from", "发件人": "from", "sent": "date", "date": "date",
                "发送时间": "date", "日期": "date", "to": "to", "收件人": "to",
                "subject": "subject", "主题": "subject", "cc": "cc", "抄送": "cc"}
_EXPLICIT_QUOTE = re.compile(
    r"^\s*(?:On\s+.+\s+wrote\s*[:：]|(?:在|于).+写道\s*[:：]|"
    r"[-_=]{2,}\s*(?:Original Message|Forwarded message|原始邮件|原邮件|转发邮件)\s*[-_=]{2,})\s*$", re.I)


def _text(value):
    return "" if value is None else str(value).strip()


def _raw(value):
    return "" if value is None else str(value)


def _hash(value):
    encoded = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"),
                         default=lambda v: v.isoformat() if isinstance(v, (datetime, date)) else str(v))
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def body_hash(body):
    """Hash exact text, including whitespace, rather than a stripped surrogate."""
    return hashlib.sha256(_raw(body).encode("utf-8")).hexdigest()


def _quote_boundary(lines):
    for index, line in enumerate(lines):
        # An author's normal > quotation is not itself an old-mail boundary.
        if line.lstrip().startswith(">"):
            continue
        if _EXPLICIT_QUOTE.fullmatch(line.rstrip("\r\n")):
            return index, index + 1
        header = _HEADER.match(line)
        if header and _HEADER_NAME[header.group(1).lower()] == "from":
            kinds = set()
            end = index + 1
            for position, candidate in enumerate(lines[index:index + 10], index):
                match = _HEADER.match(candidate)
                if match:
                    kinds.add(_HEADER_NAME[match.group(1).lower()])
                    end = position + 1
                elif not candidate.strip():
                    continue
                elif candidate.startswith((" ", "\t")) and kinds:
                    end = position + 1
                else:
                    break
            if "from" in kinds and "date" in kinds and kinds & {"to", "subject"}:
                return index, end
    return None


def extract_new_body(body):
    """Extract only structurally clear new text; report ambiguous unquoted tails.

    Inline and bottom replies after > quote runs are retained. If a mail chain
    has no > quote markers, its unquoted tail may contain either old text or a
    bottom reply. Preserve that evidence with a verification note, rather than
    silently deleting it or claiming it is confirmed new text.
    """
    body = _raw(body)
    lines = body.splitlines(keepends=True)
    boundary = _quote_boundary(lines)
    if boundary is None:
        return {"邮件正文": body, "需核实": False, "核实说明": "", "可确认新增正文": bool(body.strip())}
    first, after_header = boundary
    prefix = "".join(lines[:first]).rstrip("\r\n")
    tail = lines[after_header:]
    quoted = [line for line in tail if line.lstrip().startswith(">")]
    if quoted:
        first_quote = next(i for i, line in enumerate(tail) if line.lstrip().startswith(">"))
        ambiguous_prefix = any(line.strip() for line in tail[:first_quote])
        kept = [line for line in tail if not line.lstrip().startswith(">")]
        ambiguous_prefix = ambiguous_prefix or _quote_boundary(kept) is not None
        rest = "".join(kept).strip("\r\n")
        text = prefix + ("\n\n" if prefix and rest else "") + rest
        # Unmarked text before the first actual > quote may be part of old mail;
        # no content is guessed or dropped, and the record remains reviewable.
        note = "引用头后存在未标记引用的文字，已保留；请核对其中旧邮件与本次新增正文。" if ambiguous_prefix else ""
        return {"邮件正文": text, "需核实": ambiguous_prefix, "核实说明": note,
                "可确认新增正文": bool(prefix.strip()) or (bool(rest.strip()) and not ambiguous_prefix)}
    if not any(line.strip() for line in tail):
        return {"邮件正文": prefix, "需核实": False, "核实说明": "", "可确认新增正文": bool(prefix.strip())}
    return {"邮件正文": body, "需核实": True,
            "核实说明": "引用头之后的正文未用引用标记区分，已保留原文；请核对是否包含本次回复或旧邮件链。",
            "可确认新增正文": bool(prefix.strip())}


def strip_quoted_history(body):
    """Compatibility wrapper. canonical_message also persists ambiguity notes."""
    return extract_new_body(body)["邮件正文"]


def is_header_only(body):
    """Recognize even incomplete quoted mail headers without claiming a reply."""
    saw_header = False
    for raw in _raw(body).splitlines():
        if not raw.strip():
            continue
        line = re.sub(r"^(?:\s*>\s*)+", "", raw)
        if _HEADER.match(line) or _EXPLICIT_QUOTE.fullmatch(line):
            saw_header = True
        elif saw_header and line.startswith((" ", "\t")):
            continue  # Folded address/date/subject header continuation.
        else:
            return False
    return saw_header


def _purpose(record):
    purpose = _text(record.get("邮件用途")) or _text(record.get("草稿用途")) or _text(record.get("邮件类型"))
    if not purpose:
        template = _text(record.get("模板及版本"))
        purpose = next((p for p in ("首联", "回复", "跟进") if template.startswith(p)), "")
    if record.get("方向") == "收信":
        return "收信"
    return "回复" if purpose in _REPLY_PURPOSES else purpose


def canonical_message(record):
    """Drop image/draft/unknown fields and use 邮件正文 as the sole body field."""
    result = {field: record.get(field, "") for field in ARCHIVE_FIELDS}
    body = record.get("邮件正文") if "邮件正文" in record else record.get("正文或草稿")
    extracted = extract_new_body(body)
    result["邮件正文"] = extracted["邮件正文"]
    if extracted["需核实"]:
        old_note = _raw(result.get("核实说明"))
        if extracted["核实说明"] not in old_note:
            result["核实说明"] = (old_note + "\n" + extracted["核实说明"]).strip("\n")
    result["邮箱消息ID"] = _text(result["邮箱消息ID"]) or _text(record.get("provider_message_id"))
    result["会话ID"] = _text(result["会话ID"]) or _text(record.get("thread_id"))
    result["邮件用途"] = _purpose(record)
    completeness = _text(result["正文完整性"])
    if completeness not in {"完整", "部分", "待核实", "缺失"}:
        if not result["邮件正文"].strip():
            completeness = "缺失"
        elif "截图" in _text(result["记录来源"]):
            completeness = "待核实"
        else:
            completeness = "完整"
    if not result["邮件正文"].strip():
        completeness = "缺失"
    elif extracted["需核实"] and completeness == "完整":
        completeness = "待核实"
    result["正文完整性"] = completeness
    result["摘要状态"] = "有效" if result.get("摘要状态") == "有效" else "待更新"
    return result


def archive_reason(record):
    """An empty reason means an eligible received mail or successful reply."""
    direction = record.get("方向")
    if direction not in {"收信", "发信"}:
        return "收发方向待核实"
    purpose = _purpose(record)
    if direction == "发信" and purpose in _INITIAL_PURPOSES:
        return "初次建联只保留主表元信息"
    state = _text(record.get("处理状态"))
    if state in _DRAFT_STATES or ("草稿" in state and record.get("实际发送结果") != "成功"):
        return "未实际发出的草稿不入档"
    if direction == "收信":
        return ""
    if record.get("实际发送结果") != "成功":
        return "发信尚无成功发送结果"
    if purpose != "回复" and not _text(record.get("回复邮件记录ID")):
        return "发信用途待核实"
    body = record.get("邮件正文") if "邮件正文" in record else record.get("正文或草稿")
    if is_header_only(body):
        return "仅有邮件引用头，没有本封实际回复正文"
    extracted = extract_new_body(body)
    if not extracted["邮件正文"].strip():
        return "没有本封实际回复正文"
    if not extracted["可确认新增正文"]:
        return "引用链未区分，无法确认本封实际回复正文"
    return ""


def is_archive_message(record):
    return not archive_reason(record)


def initial_contact_metadata(record):
    """A small allowlist for the main sheet, with no subject or body leakage."""
    if record.get("方向") != "发信" or _purpose(record) not in _INITIAL_PURPOSES:
        return {}
    if record.get("实际发送结果") != "成功":
        return {}
    return {"达人ID": _text(record.get("达人ID")), "最近发信时间": record.get("邮件时间", ""),
            "最近联系": record.get("邮件时间", ""), "首联发送时间": record.get("邮件时间", ""),
            "首联发送结果": "成功"}


def message_fingerprint(record):
    canonical = canonical_message(record)
    # A direct edit to stored text, including a newly pasted quote chain or
    # whitespace, must invalidate the summary. Do not normalize the hash input.
    canonical["邮件正文"] = _raw(record.get("邮件正文") if "邮件正文" in record else record.get("正文或草稿"))
    return _hash({field: _stable_value(field, canonical.get(field)) for field in _MAIL_SUMMARY_FIELDS})


def set_mail_summary(record, summary):
    if not is_archive_message(record):
        raise ValueError("仅能为真实入档邮件设置摘要")
    record["邮件摘要"] = _raw(summary)
    record["摘要依据指纹"] = message_fingerprint(record)
    record["摘要状态"] = "有效" if _text(summary) else "待更新"
    return record


def creator_fingerprint(creator, messages):
    identifier = _text(creator.get("达人ID"))
    mail = sorted(((_text(m.get("邮件记录ID")), message_fingerprint(m)) for m in messages
                   if _text(m.get("达人ID")) == identifier and is_archive_message(m)))
    profile = {f: _stable_value(f, creator.get(f)) for f in _CREATOR_SUMMARY_FIELDS}
    # Current unsent text belongs to the main-sheet workspace, not mail history.
    # Bind summaries that mention it while retaining existing hashes for rows
    # with no draft. An edit or removal must mark that summary stale.
    draft_fields = ("草稿状态", "草稿收件人", "草稿主题", "当前回复草稿", "草稿来源")
    if any(_text(creator.get(f)) for f in draft_fields):
        profile["当前草稿"] = {f: _stable_value(f, creator.get(f)) for f in draft_fields}
    return _hash({"达人": profile, "邮件": mail})


def set_creator_summary(creator, messages, summary):
    creator["沟通摘要"] = _summary_without_stale_marker(summary)
    creator["沟通摘要依据指纹"] = creator_fingerprint(creator, messages)
    creator.pop("沟通摘要状态", None)
    return creator


def _summary_without_stale_marker(summary):
    summary = _raw(summary)
    while summary.startswith(_STALE_SUMMARY_PREFIX):
        summary = summary[len(_STALE_SUMMARY_PREFIX):].lstrip()
    return summary


def refresh_summary_states(creators, messages):
    """Detect direct spreadsheet edits too; never silently bless stale summaries."""
    for record in messages:
        if not is_archive_message(record):
            continue
        if (not _text(record.get("邮件摘要")) or not record.get("摘要依据指纹")
                or record["摘要依据指纹"] != message_fingerprint(record)):
            record["摘要状态"] = "待更新"
        elif record.get("摘要状态") != "待更新":
            record["摘要状态"] = "有效"
    for creator in creators:
        prior_state = creator.pop("沟通摘要状态", None)
        summary = _raw(creator.get("沟通摘要"))
        if (prior_state == "待更新" or not _text(summary) or not creator.get("沟通摘要依据指纹")
                or creator["沟通摘要依据指纹"] != creator_fingerprint(creator, messages)):
            creator["沟通摘要"] = _STALE_SUMMARY_PREFIX + _summary_without_stale_marker(summary)
    return {"达人": creators, "邮件": messages}


def _single_locator(value):
    raw = _text(value)
    if not raw or re.search(r"\s", raw):
        return ""
    try:
        parts = urlsplit(raw)
    except ValueError:
        return ""
    if not parts.scheme:
        return raw if re.fullmatch(r"(?:message|mail|消息ID|邮件ID):[^:]+", raw, re.I) else ""
    query = {k.lower(): v for k, v in parse_qs(parts.query).items()}
    if (re.search(r"/(?:messages?|readmail)/[^/]+/?$", parts.path, re.I)
            or any(query.get(k) for k in ("message_id", "messageid", "mailid", "mid"))):
        return raw
    return ""


def _datetime(value):
    if isinstance(value, datetime):
        parsed = value
    elif isinstance(value, date):
        parsed = datetime.combine(value, time.min)
    elif isinstance(value, str):
        try:
            parsed = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
        except ValueError:
            return None
    else:
        return None
    # Naive source timestamps are business wall time (Asia/Shanghai).
    if parsed.tzinfo is None:
        from datetime import timedelta
        parsed = parsed.replace(tzinfo=timezone(timedelta(hours=8)))
    return parsed.astimezone(timezone.utc)


def _stable_value(field, value):
    """Survive XLSX empty cells and ISO-space/T date serialization changes."""
    if value is None:
        return ""
    if field in {"邮件时间", "确认时间", "计划跟进日期", "最近联系", "最近收信时间", "最近发信时间", "首联发送时间"}:
        if ((isinstance(value, date) and not isinstance(value, datetime))
                or (isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", value.strip()))):
            return str(value).strip()
        parsed = _datetime(value)
        if parsed is not None:
            return parsed.isoformat()
    return value


def _dedup_fingerprint(record):
    fields = ("方向", "邮件时间", "发件人", "收件人", "主题", "邮件正文")
    if any(not _text(record.get(f)) for f in fields) or _datetime(record.get("邮件时间")) is None:
        return ""
    return _hash({"方向": record["方向"], "邮件时间": _datetime(record["邮件时间"]).isoformat(),
                  "发件人": _text(record["发件人"]).lower(), "收件人": _text(record["收件人"]).lower(),
                  "主题": _raw(record["主题"]), "邮件正文": _raw(record["邮件正文"]),
                  "邮箱账户": _text(record.get("邮箱账户")).lower()})


def _same_message(left, right):
    left, right = canonical_message(left), canonical_message(right)
    a, b = _text(left.get("邮箱消息ID")), _text(right.get("邮箱消息ID"))
    if a and b:
        return a == b and _text(left.get("邮箱账户")).lower() == _text(right.get("邮箱账户")).lower()
    # A replay of a locally identified screenshot may have no provider ID, date
    # or locator. Its permanent record ID is usable only with compatible mail
    # identity, and never overrides distinct provider-message identities.
    local_a, local_b = _text(left.get("邮件记录ID")), _text(right.get("邮件记录ID"))
    if not a and not b and local_a and local_a == local_b:
        for field in ("方向", "达人ID", "发件人", "收件人"):
            first, second = _text(left.get(field)), _text(right.get(field))
            if field in {"发件人", "收件人"}:
                first, second = first.lower(), second.lower()
            if first and second and first != second:
                return False
        return True
    a, b = _single_locator(left.get("原邮件定位")), _single_locator(right.get("原邮件定位"))
    if a and b:
        return a == b
    fingerprint = _dedup_fingerprint(left)
    return bool(fingerprint and fingerprint == _dedup_fingerprint(right))


def _note(record, note):
    old = _raw(record.get("核实说明"))
    if note not in old:
        record["核实说明"] = (old + "\n" + note).strip("\n")
    record["摘要状态"] = "待更新"


def _next_id(messages):
    used = {_text(m.get("邮件记录ID")) for m in messages}
    # Legacy numeric IDs remain untouched; removed draft IDs are never reused.
    identifier = "MAIL-" + uuid4().hex
    while identifier in used:
        identifier = "MAIL-" + uuid4().hex
    return identifier


def _resolve_creator(creators, messages, incoming):
    # Reuse conservative platform/thread/shared-agency matching, not the legacy
    # importer whose duplicate early return prevents explicit text completion.
    return rules._resolve_creator(creators, messages, incoming)


def _apply_contact_state(creators, record):
    creator = next((c for c in creators if c.get("达人ID") == record.get("达人ID")), None)
    if creator is None:
        record["处理状态"] = "待人工归属"
        return
    record["达人昵称"] = creator.get("达人昵称") or creator.get("昵称") or record.get("达人昵称", "")
    latest = "最近收信时间" if record["方向"] == "收信" else "最近发信时间"
    current, old = _datetime(record.get("邮件时间")), _datetime(creator.get(latest))
    if current and (old is None or current > old):
        creator[latest] = record["邮件时间"]
    if record["方向"] == "收信":
        creator["计划跟进日期"] = ""
        if not _text(record.get("处理状态")):
            record["处理状态"] = "待回复"
    else:
        record["处理状态"] = "已发送"


def import_message(creators, messages, incoming):
    """Import eligible mail only. Completion is an explicit separate operation."""
    raw_body = incoming.get("邮件正文") if "邮件正文" in incoming else incoming.get("正文或草稿")
    if extract_new_body(raw_body)["需核实"]:
        return {"已导入": False, "记录": None, "需人工归属": False,
                "原因": "需核实引用链；请明确分离本次新增正文后再导入", "补全待处理": False}
    reason = archive_reason(incoming)
    if reason:
        return {"已导入": False, "记录": None, "需人工归属": False, "原因": reason, "补全待处理": False,
                "首联元信息": initial_contact_metadata(incoming)}
    record = canonical_message(incoming)
    existing = next((m for m in messages if is_archive_message(m) and _same_message(m, record)), None)
    if existing is not None:
        previous_body = canonical_message(existing)["邮件正文"]
        changed_body = previous_body != record["邮件正文"]
        if changed_body and canonical_message(existing)["正文完整性"] == "完整":
            _note(existing, "同封邮件出现不同正文；已保留原文，需核实。候选正文指纹：" + body_hash(record["邮件正文"]))
        refresh_summary_states(creators, messages)
        return {"已导入": False, "记录": existing, "需人工归属": not bool(existing.get("达人ID")),
                "原因": "重复邮件；正文变更需显式补全" if changed_body else "重复邮件",
                "补全待处理": changed_body}
    rules.ensure_creator_ids(creators)
    # Preserve optional account/profile keys only during ownership resolution.
    resolution = {**dict(incoming), **record}
    record["达人ID"] = _resolve_creator(creators, messages, resolution) or ""
    supplied_id = _text(record.get("邮件记录ID"))
    if supplied_id and any(_text(m.get("邮件记录ID")) == supplied_id for m in messages):
        raise ValueError("邮件记录ID已存在且对应不同邮件")
    record["邮件记录ID"] = supplied_id or _next_id(messages)
    record["记录来源"] = record.get("记录来源") or "原邮件文本导入"
    record["邮件用途"] = "收信" if record["方向"] == "收信" else "回复"
    _apply_contact_state(creators, record)
    messages.append(record)
    main_schema.refresh_progress(creators, messages)
    refresh_summary_states(creators, messages)
    return {"已导入": True, "记录": record, "需人工归属": not bool(record["达人ID"]), "原因": "", "补全待处理": False}


def fill_message(creators, messages, record_id, incoming):
    """Explicitly complete one existing partial transcription, preserving its ID.

    An already complete differing body is retained and reported for review. No
    candidate body is placed in notes or accidentally exposed via an overview.
    """
    found = [m for m in messages if _text(m.get("邮件记录ID")) == _text(record_id)]
    if len(found) != 1:
        raise ValueError("需要唯一的邮件记录ID")
    existing = found[0]
    if not is_archive_message(existing):
        raise ValueError("只能补全真实入档邮件")
    candidate_fields = dict(incoming)
    if "邮件正文" not in candidate_fields and "正文或草稿" in candidate_fields:
        candidate_fields["邮件正文"] = candidate_fields["正文或草稿"]
    candidate_body = candidate_fields.get("邮件正文", existing.get("邮件正文") if "邮件正文" in existing else existing.get("正文或草稿"))
    if extract_new_body(candidate_body)["需核实"]:
        raise ValueError("需核实引用链；请明确分离本次新增正文后再补全")
    for field in ("邮件记录ID", "达人ID", "方向", "邮箱消息ID", "邮箱账户"):
        supplied = _text(candidate_fields.get(field))
        if supplied and _text(existing.get(field)) and supplied != _text(existing.get(field)):
            raise ValueError("补全不能改变原记录身份或归属：" + field)
    merged = canonical_message({**existing, **candidate_fields})
    if archive_reason(merged):
        raise ValueError("补全后不满足纯文本归档范围")
    old = canonical_message(existing)
    if old["正文完整性"] == "完整" and old["邮件正文"] != merged["邮件正文"]:
        _note(existing, "完整原文与补全文本冲突；已保留原文，需核实。候选正文指纹：" + body_hash(merged["邮件正文"]))
        refresh_summary_states(creators, messages)
        return {"已补全": False, "存在冲突": True, "记录": existing}
    if old["正文完整性"] == "完整" and merged["正文完整性"] in _INCOMPLETE:
        raise ValueError("不能用部分转录降低完整原文的完整性")
    # Fields absent in the incoming dict are preserved, not defaulted to empty.
    for field in ARCHIVE_FIELDS:
        if field in candidate_fields and field not in {"邮件记录ID", "达人ID", "邮件摘要", "摘要依据指纹", "摘要状态"}:
            existing[field] = merged[field]
    existing["邮件正文"] = merged["邮件正文"]
    existing["正文完整性"] = merged["正文完整性"]
    existing.pop("正文或草稿", None)
    _apply_contact_state(creators, existing)
    main_schema.refresh_progress(creators, messages)
    refresh_summary_states(creators, messages)
    return {"已补全": True, "存在冲突": False, "记录": existing}


def _mail_index(record):
    canonical = canonical_message(record)
    # No body, addresses, sources, raw time, notes or draft fields in overview.
    fields = ("邮件记录ID", "达人ID", "达人昵称", "方向", "邮件时间", "主题", "邮件摘要", "处理状态",
              "正文完整性", "摘要状态")
    return {field: canonical.get(field, "") for field in fields}


def overview(creators, messages):
    refresh_summary_states(creators, messages)
    fields = ("达人ID", "达人昵称", "沟通摘要", "联系状态", "计划跟进日期", "草稿状态", "最近联系",
              "最近收信时间", "最近发信时间")
    active = [m for m in messages if is_archive_message(m)]
    return {"达人进展": [{field: c.get(field, "") for field in fields} for c in creators],
            "邮件索引": [_mail_index(m) for m in active],
            "待人工归属数量": sum(not _text(m.get("达人ID")) for m in active)}


def context(creators, messages, creator_id, budget=24000, cursor=0):
    """Return complete mail units, sorted known times then unknown source order.

    The character budget applies to the JSON serialization of full mail units;
    a single oversize message is still returned whole and explicitly labelled.
    Cursors carry a snapshot fingerprint, so a changed archive cannot silently
    skip newly inserted mail. Numeric cursors remain accepted for CLI convenience.
    """
    if not isinstance(budget, int) or isinstance(budget, bool) or budget <= 0:
        raise ValueError("字符预算必须为正整数")
    creator = next((c for c in creators if _text(c.get("达人ID")) == _text(creator_id)), None)
    if creator is None or not _text(creator_id):
        raise ValueError("未找到达人ID")
    refresh_summary_states(creators, messages)
    own = [canonical_message(m) for m in messages
           if _text(m.get("达人ID")) == _text(creator_id) and is_archive_message(m)]
    known = sorted((m for m in own if _datetime(m.get("邮件时间")) is not None),
                   key=lambda m: _datetime(m["邮件时间"]))
    unknown = [m for m in own if _datetime(m.get("邮件时间")) is None]
    ordered = known + unknown
    snapshot = _hash([(m["邮件记录ID"], message_fingerprint(m)) for m in ordered])
    if isinstance(cursor, dict):
        if cursor.get("达人ID") != _text(creator_id) or cursor.get("快照") != snapshot:
            raise ValueError("邮件记录已变化或游标不属于该达人；请从第一页重新读取")
        start = cursor.get("位置")
    else:
        start = cursor
    if not isinstance(start, int) or isinstance(start, bool) or start < 0 or start > len(ordered):
        raise ValueError("分页游标超出范围")
    page, used, oversize = [], 0, []
    for record in ordered[start:]:
        length = len(json.dumps(record, ensure_ascii=False, default=str))
        if page and used + length > budget:
            break
        if length > budget:
            oversize.append(record["邮件记录ID"])
        page.append(record)
        used += length
    end = start + len(page)
    known_ids = {id(m) for m in known}
    return {"达人ID": _text(creator_id), "达人昵称": creator.get("达人昵称") or creator.get("昵称", ""),
            "沟通摘要": creator.get("沟通摘要", ""),
            "全部邮件数量": len(ordered), "已知时间邮件数量": len(known), "时间不明邮件数量": len(unknown),
            "本页范围": {"起始": start + 1 if page else 0, "结束": end}, "本页字符数": used,
            "字符预算": budget, "单封超过预算": oversize,
            "时间明确邮件": [m for m in page if id(m) in known_ids],
            "时间不明邮件": [m for m in page if id(m) not in known_ids],
            "下一页游标": {"达人ID": _text(creator_id), "位置": end, "快照": snapshot} if end < len(ordered) else None,
            "已读完": end >= len(ordered)}
